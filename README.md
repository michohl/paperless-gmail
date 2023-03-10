# Gmail Emails as PDF

This script is meant to take emails from Gmail and convert them to PDF so they can be ingested by paperless-ngx


## Requirements

```shell
$ pip install -r requirements.txt
# apt install wkhtmltopdf
```

Make sure the labels set for fetching and marking messages as consumed are
already existing in Gmail otherwise the script will fail to copy the message to the consumed folder