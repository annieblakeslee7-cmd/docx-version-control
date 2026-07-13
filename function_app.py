import base64
import hashlib
import io
import json
import logging
import zipfile
from typing import Any

import azure.functions as func
from lxml import etree

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

NS = {
    "w": WORD_NS
}


def get_attribute(element: etree._Element, name: str) -> str | None:
    """Read a WordprocessingML attribute."""
    return element.get(f"{{{WORD_NS}}}{name}")


def collect_text(element: etree._Element, change_type: str) -> str:
    """
    Collect visible text inside a revision.

    Insertions normally use w:t.
    Deletions normally use w:delText.
    """
    if change_type in {"Deletion", "MoveFrom"}:
        nodes = element.xpath(".//w:delText | .//w:t", namespaces=NS)
    else:
        nodes = element.xpath(".//w:t | .//w:delText", namespaces=NS)

    return "".join(node.text or "" for node in nodes)


def collect_paragraph_text(paragraph: etree._Element) -> str:
    """Collect inserted, deleted, and ordinary text from a paragraph."""
    nodes = paragraph.xpath(".//w:t | .//w:delText", namespaces=NS)
    return "".join(node.text or "" for node in nodes)


def nearest_paragraph(element: etree._Element) -> etree._Element | None:
    """Find the paragraph containing a tracked revision."""
    current = element

    while current is not None:
        if current.tag == f"{{{WORD_NS}}}p":
            return current

        current = current.getparent()

    return None


def create_fingerprint(
    change_type: str,
    author: str,
    revision_date: str,
    changed_text: str,
    paragraph_text: str,
    document_part: str,
) -> str:
    """
    Build a stable comparison key.

    Word revision IDs alone are not always reliable across saved versions.
    """
    raw_value = "|".join(
        [
            change_type.strip().lower(),
            author.strip().lower(),
            revision_date.strip().lower(),
            changed_text.strip().lower(),
            paragraph_text.strip().lower(),
            document_part.strip().lower(),
        ]
    )

    return hashlib.sha256(raw_value.encode("utf-8")).hexdigest()


def parse_xml_part(
    archive: zipfile.ZipFile,
    part_name: str,
    sequence_start: int,
) -> tuple[list[dict[str, Any]], int]:
    """Extract tracked changes from one XML part of the DOCX package."""
    if part_name not in archive.namelist():
        return [], sequence_start

    xml_bytes = archive.read(part_name)
    root = etree.fromstring(xml_bytes)

    revision_types = [
        ("Insertion", ".//w:ins"),
        ("Deletion", ".//w:del"),
        ("MoveFrom", ".//w:moveFrom"),
        ("MoveTo", ".//w:moveTo"),
    ]

    changes: list[dict[str, Any]] = []
    sequence = sequence_start

    for change_type, xpath in revision_types:
        for revision in root.xpath(xpath, namespaces=NS):
            changed_text = collect_text(revision, change_type)

            # Skip revisions that contain no readable text.
            if not changed_text.strip():
                continue

            paragraph = nearest_paragraph(revision)
            paragraph_text = (
                collect_paragraph_text(paragraph)
                if paragraph is not None
                else changed_text
            )

            revision_id = get_attribute(revision, "id") or ""
            author = get_attribute(revision, "author") or "Unknown"
            revision_date = get_attribute(revision, "date") or ""

            sequence += 1

            fingerprint = create_fingerprint(
                change_type=change_type,
                author=author,
                revision_date=revision_date,
                changed_text=changed_text,
                paragraph_text=paragraph_text,
                document_part=part_name,
            )

            changes.append(
                {
                    "sequence": sequence,
                    "revisionId": revision_id,
                    "changeType": change_type,
                    "author": author,
                    "revisionDate": revision_date,
                    "changedText": changed_text,
                    "paragraphText": paragraph_text,
                    "documentPart": part_name,
                    "fingerprint": fingerprint,
                }
            )

    return changes, sequence

def extract_current_paragraphs(
    archive: zipfile.ZipFile
) -> list[str]:
    """
    Extract paragraph text representing the document's current visible content.

    Inserted text is included.
    Deleted text is excluded.
    """
    part_name = "word/document.xml"

    if part_name not in archive.namelist():
        return []

    root = etree.fromstring(archive.read(part_name))
    paragraphs: list[str] = []

    for paragraph in root.xpath(".//w:p", namespaces=NS):
        pieces: list[str] = []

        for element in paragraph.iter():
            if element.tag == f"{{{WORD_NS}}}del":
                # Do not include deleted content in the visible snapshot.
                continue

            if element.tag == f"{{{WORD_NS}}}t":
                # Exclude text when it is inside a deletion.
                inside_deletion = any(
                    ancestor.tag == f"{{{WORD_NS}}}del"
                    for ancestor in element.iterancestors()
                )

                if not inside_deletion and element.text:
                    pieces.append(element.text)

        paragraph_text = "".join(pieces).strip()

        if paragraph_text:
            paragraphs.append(paragraph_text)

    return paragraphs


def parse_docx(file_bytes: bytes) -> list[dict[str, Any]]:
    """Open the DOCX ZIP package and scan relevant Word XML parts."""
    changes: list[dict[str, Any]] = []
    sequence = 0

    with zipfile.ZipFile(io.BytesIO(file_bytes), "r") as archive:
        parts_to_scan = ["word/document.xml"]

        # Include headers, footers, footnotes, and endnotes when present.
        for name in archive.namelist():
            if (
                name.startswith("word/header")
                or name.startswith("word/footer")
                or name == "word/footnotes.xml"
                or name == "word/endnotes.xml"
            ):
                if name.endswith(".xml"):
                    parts_to_scan.append(name)

        for part_name in sorted(set(parts_to_scan)):
            part_changes, sequence = parse_xml_part(
                archive=archive,
                part_name=part_name,
                sequence_start=sequence,
            )
            changes.extend(part_changes)
    print(f"Found {len(changes)} tracked changes.")
    return changes


@app.route(route="extract-docx-revisions", methods=["POST"])
def extract_docx_revisions(req: func.HttpRequest) -> func.HttpResponse:
    try:
        request_body = req.get_json()

        file_name = request_body.get("fileName", "")
        version_label = request_body.get("versionLabel", "")
        version_id = request_body.get("versionId", "")
        encoded_content = request_body.get("fileContent", "")

        if not encoded_content:
            return func.HttpResponse(
                json.dumps({"error": "fileContent is required"}),
                status_code=400,
                mimetype="application/json",
            )

        file_bytes = base64.b64decode(encoded_content, validate=True)
        print("Received:", file_name)
        print("Version:", version_label)
        print("Bytes:", len(file_bytes))

        if not zipfile.is_zipfile(io.BytesIO(file_bytes)):
            return func.HttpResponse(
                json.dumps(
                    {
                        "error": "The supplied content is not a valid DOCX package.",
                        "fileName": file_name,
                        "versionLabel": version_label,
                    }
                ),
                status_code=400,
                mimetype="application/json",
            )

        changes = parse_docx(file_bytes)

        response = {
            "fileName": file_name,
            "versionLabel": version_label,
            "versionId": version_id,
            "changeCount": len(changes),
            "changes": changes,
        }

        return func.HttpResponse(
            json.dumps(response, ensure_ascii=False),
            status_code=200,
            mimetype="application/json",
        )

    except ValueError as exc:
        logging.exception("Invalid Base64 or JSON input.")

        return func.HttpResponse(
            json.dumps({"error": f"Invalid request: {str(exc)}"}),
            status_code=400,
            mimetype="application/json",
        )

    except zipfile.BadZipFile:
        return func.HttpResponse(
            json.dumps({"error": "The supplied file is not a readable DOCX file."}),
            status_code=400,
            mimetype="application/json",
        )

    except Exception as exc:
        logging.exception("Unexpected document-processing error.")

        return func.HttpResponse(
            json.dumps({"error": f"Processing failed: {str(exc)}"}),
            status_code=500,
            mimetype="application/json",
        )