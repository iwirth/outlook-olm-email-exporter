# outlook-olm-email-exporter

*extract and export e-mail addresses from outlook .olm files using python*

## Changes in this fork
- Switched from "Local" to "Accounts", because I'm using multiple Microsoft 365 accounts
- Added account selection
- program now extracts the olm file itself and deletes the artifacts at exit
- outputs the addresses in the already existing form (sender name, e-mail address) and now also outputs just the addresses in a seperate file

## How to use
1. Export all your e-mails using the `Export` menu in Outlook (at the time of writing, it did not work in the "New Outlook" for me)
2. Download or clone the script
3. Run `python extract.py` and obtain the emails.txt file
4. Input the path of the olm file
