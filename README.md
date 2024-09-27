# WebUntisTimeTableToIcs

![Update ICS Calendar](https://github.com/Chaos02/WebUntisTimeTableToIcs/actions/workflows/GeneratePage.yml/badge.svg)


## Overview

This repository contains a GitHub Actions workflow and a PowerShell script to generate an ICS calendar file from a timetable. The workflow is scheduled to run at specific intervals and can also be triggered manually. The PowerShell script fetches data from a specified URL and generates the ICS file.

### Workflow: `GeneratePage.yml`

The GitHub Actions workflow `GeneratePage.yml` is designed to update the ICS calendar file at regular intervals and upon manual request. The workflow performs the following steps:

1. **Checkout Repository**: Checks out the repository to the GitHub Actions runner.
2. **Download Previous ICS**: Downloads the previous ICS artifact if available.
3. **Run PowerShell Script**: Executes the PowerShell script to generate a new ICS file.
4. **Compare ICS Files**: Compares the newly generated ICS file with the previous one.

#### Schedule

The workflow is scheduled to run at the following times:
- Every 6 hours
- Daily at 6:00 UTC
- Daily at 7:00 UTC

Additionally, the workflow can be triggered manually using the `workflow_dispatch` event.

### PowerShell Script: `timeTableToIcs.ps1`

The PowerShell script `timeTableToIcs.ps1` generates an ICS file from a timetable by making HTTP requests to a specified URL. The script uses headers and cookies that can be exported from the Chrome Dev Console.

#### Parameters

- `baseUrl`: The base URL for the HTTP requests.
- `elementType`: The type of element (default is `1`).
- `elementId`: The ID of the element.
- `date`: The date for the timetable (default is the current date).
- `OutputFilePath`: The output file path for the ICS file (default is `calendar.ics`).
- `cookie`: The cookie value for authentication.
- `tenantId`: The tenant ID for authentication.

#### Usage

To use the script, run the following command in PowerShell:

```powershell
./timeTableToIcs.ps1 -OutputFilePath "calendar.ics" -baseUrl "<BASE_URL>" -elementType "<ELEMENT_TYPE>" -elementId "<ELEMENT_ID>" -cookie "<COOKIE>" -tenantId "<TENANT_ID>"