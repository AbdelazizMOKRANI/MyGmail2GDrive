Gmail2GDrive
Gmail2GDrive is a Google Apps Script which automatically stores and sorts Gmail attachments into Google Drive folders, and can also save the thread as a PDF file.

It does so by defining a list of rules which consist of Gmail search filters and Google Drive destination folders. This way the attachments of periodic emails can be automatically organized in folders without the need to install and run anything on the client.

Features:
Automatically sorts your attachments in the background
Filter for relevant emails
Specify the destination folder
Rename attachments (using date format strings and email subject as filenames)
Save the thread as a PDF File
Setup
Open Google Apps Script.
Create an empty project.
Give the project a name (e.g. MyGmail2GDrive)
Replace the content of the created file Code.gs with the provided Code.gs and save the changes.
Create a new script file with the name 'Config' and replace its content with the provided Config.gs and save the changes.
Adjust the configuration to your needs. It is recommended to restrict the timeframe using 'newerThan' to prevent running into API quotas by Google.
Test the script by manually executing the function performGmail2GDrive.
Create a time based trigger which periodically executes 'Gmail2GDrive' (e.g. once per day) to automatically organize your Gmail attachments within Google Drive.


Global Configuration:
globalFilter: Global filter expression (see https://support.google.com/mail/answer/7190?hl=en for avialable search operators)
Example: "globalFilter": "has:attachment -in:trash -in:drafts -in:spam"
processedLabel: The GMail label to mark processed threads (will be created, if not existing)
Example: "processedLabel": "to-gdrive/processed"
sleepTime: Sleep time in milliseconds between processed messages
Example: "sleepTime": 100
maxRuntime: Maximum script runtime in seconds (Google Scripts will be killed after 5 minutes)
Example: "maxRuntime": 280
newerThan: Only process message newer than (leave empty for no restriction; use d, m and y for day, month and year)
Example: "newerThan": "1m"
timezone: Timezone for date/time operations
Example: "timezone": "GMT"
rules: List of rules to be processed
Example: "rules": [ {..rule1..}, {..rule2..}, ... ]



Rule Configuration:
A rule supports the following parameters documentation:

filter (String, mandatory): a typical gmail search expression (see http://support.google.com/mail/bin/answer.py?hl=en&answer=7190)
folder (String, mandatory): a path to an existing Google Drive folder (will be created, if not existing)
archive (boolean, optional): Should the gmail thread be archived after processing? (default: false)
filenameFrom (String, optional): The attachment filename that should be renamed when stored in Google Drive
filenameFromRegexp (String, optional): A regular expression to specify only relevant attachments
filenameTo (String, optional): The pattern for the new filename of the attachment. If 'filenameFrom' is not given then this will be the new filename for all attachments.
You can use '%s' to insert the email subject and date format patterns like 'yyyy' for year, 'MM' for month and 'dd' for day as pattern in the filename.
See https://developers.google.com/apps-script/reference/utilities/utilities#formatDate(Date,String,String) for more information on the possible date format strings.
saveThreadPDF (boolean, optional): Should the thread be saved as a PDF? (default: false)
