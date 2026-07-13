import base64
import json
import urllib.error
import urllib.request
from pathlib import Path


FUNCTION_URL = "http://localhost:7071/api/extract-docx-revisions"
DOCX_PATH = Path("test.docx")


def main() -> None:
    if not DOCX_PATH.exists():
        raise FileNotFoundError(f"Could not find {DOCX_PATH.resolve()}")

    file_content = base64.b64encode(DOCX_PATH.read_bytes()).decode("utf-8")

    payload = {
        "fileName": DOCX_PATH.name,
        "versionLabel": "TEST",
        "versionId": "TEST-1",
        "fileContent": file_content,
    }

    request = urllib.request.Request(
        FUNCTION_URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )

    try:
        with urllib.request.urlopen(request) as response:
            response_text = response.read().decode("utf-8")
            parsed_response = json.loads(response_text)
            print(json.dumps(parsed_response, indent=2))

    except urllib.error.HTTPError as error:
        print(f"HTTP {error.code}")
        print(error.read().decode("utf-8"))


if __name__ == "__main__":
    main()