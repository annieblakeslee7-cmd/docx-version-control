docx-version-control

As of now, this project is a proof of concept tool to track changes between shared document versions. 

TO RUN:
activate venv in vscode
run func start --python
in powershell, run 
python function_app.py

you will see a json file returned with specific changes from version to version.
example output:
{
  "fileName": "test.docx",
  "versionLabel": "TEST",
  "versionId": "TEST-1",
  "changeCount": 2,
  "changes": [
    {
      "sequence": 1,
      "revisionId": "2",
      "changeType": "Insertion",
      "author": "Annie Blakeslee",
      "revisionDate": "2026-07-13T10:04:00Z",
      "changedText": "Next update done.",
      "paragraphText": "Next update done.",
      "documentPart": "word/document.xml",
      "fingerprint": "e9d85e148913ea3a097036dd417feabef838e91623e216d3a7c6ec056a021998"
    },
    {
      "sequence": 2,
      "revisionId": "3",
      "changeType": "Insertion",
      "author": "Annie Blakeslee",
      "revisionDate": "2026-07-13T10:05:00Z",
      "changedText": "And again.",
      "paragraphText": "And again.",
      "documentPart": "word/document.xml",
      "fingerprint": "469cdd6688e59d83217ce46d869955442cb0a60bcf4831d7f7550c25bbe095d0"
    }
  ]
}
