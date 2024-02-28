# OMRAH
Outlook Meeting Request Autohandler, a bare-bones Outlook add-in that lets you configure simple rules to automatically handle meeting requests.

OMRAH works by filtering meeting requests based on a few parameters (subject, recipient, etc.), responding in the desired way (e.g., accept, decline, etc.), and adding a pre-configured category to the resulting appointment.

As of now, OMRAH only support seeting one "rule" (i.e., way to handle meetings requests in terms of filters and action).

## Installation

1. Download the latest installer.zip from the [releases page](https://github.com/CaileanMParker/OMRAH/releases)
2. Extract the contents of the installer.zip
3. Make sure Outlook is no currently running on the target machine
3. Run setup.exe
4. Follow the installation wizard
5. Launch Outlook

## Configuration

OMRAH can be configured by editing the "OMRAH.dll.config" app settings file in the installation directory ("<ProgramFiles64Folder>\CaileanMParker\OMRAH\" by default). The following settings are available:

| Setting | Type | Description | Default |
| --- | --- | --- | --- |
| SubjectFilters | System.Collections.Specialized.StringCollection | A list of substrings to match against the subject of incoming meeting requests. | "ooo", "out of office", "oof", "out of facility" |
| RecipientFilters | System.Collections.Specialized.StringCollection | A list of email addresses to match against the recipients of incoming meeting requests. Partial matches are possible and an empty value will match any recipients. | (none) |
| UnlessDirect | bool | If true, only meeting requests that are not directly addressed to the user will be handled. | False |
| CategoryName | string | The name of the category to add to appointments that match the filters. | "Other OOF" |
| CategoryColor | Microsoft.Office.Interop.Outlook.OlCategoryColor | The color of the category to add to appointments that match the filters. | olCategoryColorGray |
| SendResponse | bool | If true, a response will be sent to the meeting organizer. | False |
| Response | Microsoft.Office.Interop.Outlook.OlMeetingResponse | The type of response to set for meeting requests which mach the filters. | olMeetingTentative |