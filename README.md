# Universal Email Cleaner

A powerful, user-friendly GUI tool designed to help administrators and users clean up emails and meetings from Microsoft Exchange Server and Exchange Online.

## Features

*   **Dual Protocol Support**:
    *   **EWS (Exchange Web Services)**: Supports Exchange Server 2010, 2013, 2016, 2019, and Exchange Online.
    *   **Microsoft Graph API**: The modern standard for Exchange Online (Microsoft 365).
*   **Flexible Graph API Configuration**:
    *   **Multi-Environment**: Supports both **Global (International)** and **China (21Vianet)** cloud environments.
    *   **Authentication Modes**:
        *   **Auto**: Automatically creates an Azure AD App and self-signed certificate for seamless setup.
        *   **Manual**: Connect using an existing App ID, Tenant ID, and Client Secret.
*   **Comprehensive Cleanup Options**:
    *   **Target**: Clean Emails or Meetings.
    *   **Criteria**: Filter by Subject, Sender, Body keywords, Date Range, and Message ID.
    *   **Meeting Scope**: Handle Single instances, Series, or only Cancelled meetings.
*   **Safety First**:
    *   **Report Only Mode**: Generate a CSV report of items to be deleted without actually deleting them.
    *   **Logging**: Detailed logs for auditing and troubleshooting.

## Requirements

*   Windows OS (Recommended)
*   Python 3.8+
*   Dependencies: `requests`, `msal`, `azure-identity`, `exchangelib`, `tk`

## Installation

1.  Clone the repository or download the source code.
2.  Install the required dependencies:
    ```bash
    pip install -r requirements.txt
    ```

## Usage

1.  Run the application:
    ```bash
    python universal_email_cleaner.py
    ```
    *Or run the compiled `.exe` if available.*

2.  **Connection Configuration**:
    *   Select **EWS** or **Graph API**.
    *   **For Graph API**:
        *   Choose Environment: **Global** or **China**.
        *   Choose Mode: **Auto** (requires Global Admin rights to set up) or **Manual** (Client Secret).
    *   **For EWS**:
        *   Enter Server URL (or use Autodiscover), User UPN, and Password.
        *   Select Access Type: **Impersonation** (recommended for admins) or **Delegate**.

3.  **Task Configuration**:
    *   Select the CSV file containing the list of user mailboxes (`UserPrincipalName` column required).
    *   Define cleanup criteria (Subject, Sender, Date, etc.).
    *   Choose **Report Only** to test first.

4.  **Run**: Click "Start Cleanup" to begin processing.

## Logging & Diagnostics

Logs and reports are written to your Documents folder:

- Log/report base directory: `%USERPROFILE%\Documents\UniversalEmailCleaner`
- Reports: `%USERPROFILE%\Documents\UniversalEmailCleaner\Reports`

The app supports three log levels (Tools → Log Level, also synced with the main UI):

- **Normal**: Standard operational logs.
- **Advanced**: Adds detailed Graph/EWS request diagnostics.
- **Expert**: Adds more verbose diagnostics; shows a warning prompt before enabling.

In Advanced/Expert mode, extra debug logs are written to separate daily files:

- `app_advanced_YYYY-MM-DD.log`
- `app_expert_YYYY-MM-DD.log`

EWS Expert mode also writes GetItem responses to:

- `ews_getitem_responses_expert_YYYY-MM-DD.log`

### Graph Authorization token logging (Expert only)

By default, Graph `Authorization` headers are masked in logs (`Bearer ***`).

If you must capture the token for troubleshooting, you can enable the dangerous option:

- Tools → Log Level → **“Graph Expert 保存 Authorization Token (危险)”**

This option is **OFF by default**, only effective in **Expert**, and will prompt for confirmation.

## License

MIT License
# UniversalEmailCleaner
