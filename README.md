![Update ICS Calendar](https://github.com/Chaos02/WebUntisTimeTableToIcs/actions/workflows/GeneratePage.yml/badge.svg)

## Usage:
1. Fork this repository 
2. Enable GitHub Pages with workflow as source.
3. Store the environment variables `BASE_URL`, `ELEMENT_ID`, `COOKIE` and `TENNANT_ID` in the repos **secrets**
4. Adjust cronjob in ./.github/workflows/GeneratePage.yml to your needs.


## Overview

This repository contains a GitHub Actions workflow and a PowerShell script to generate an ICS calendar file from a WebUntis timetable. The workflow is scheduled to run at specific intervals and can also be triggered manually. The PowerShell script fetches data from a specified URL and generates the ICS file for subscription.

### PowerShell Script: `timeTableToIcs.ps1`

The PowerShell script `timeTableToIcs.ps1` generates an ICS file from a timetable by making HTTP requests to your WebUntis URL. The script uses headers and cookies that can be exported from the Chrome Dev Console.

#### Parameters

- `baseUrl`: The base URL for the HTTP requests.
- `elementType`: The return type of the request (default is `1`, I don't know what other values might return from the API).
- `elementId`: The ID of the classes timetable. (Get from URL or Chrome Dev tools)
- `date`: Accepts DateTime object(s) to determine and fetch the relevant week(s) content.
- `OutputFilePath`: The output file path for the ICS file (default is `calendar.ics`).
  (if left out the ICS content will be written to pipe and returned from the program)
- `cookie`: The cookie value for authentication.
- `tenantId`: The tenant ID for authentication.

### Workflow: `GeneratePage.yml`

The GitHub Actions workflow `GeneratePage.yml` is designed to update and deploy the ICS calendar file. The workflow performs the following steps:

1. **Download Previous ICS**: Downloads the previous ICS artifact if available.
2. **Run PowerShell Script**: Executes the PowerShell script to generate a new ICS file.
3. **Compare ICS Files**: Compares the newly generated ICS file with the previous one.
4. **Upload ICS Artifact**: Uploads the newly generated ICS file as an artifact.
5. **Setup Pages**: Configures GitHub Pages for deployment.
6. **Upload Artifact**: Uploads the artifacts as a Pages artifact.
7. **Deploy to GitHub Pages**: Deploys the uploaded artifact to GitHub Pages.

#### Schedule

The workflow is scheduled to run at the following times:
- Every 6 hours
- Daily at 6:00 UTC
- Daily at 7:00 UTC

The workflow can also be triggered manually.

### Secrets File: `secrets.ps1`

The `secrets.ps1` file contains the necessary parameters for running the PowerShell script **locally**. This file should be kept secure and not be committed to version control.

#### Sample `secrets.ps1`

```powershell
# secrets.ps1
[string]$baseUrl = "XXX.webuntis.com"
[string]$elementId = 2280 # Course ID
[string]$cookie = "_d3ZzcyBtYW5uaGVpbQ==" # Get by opening official page with dev tools and copy http request as powershell
[string]$tenantId = "5028200" # Get by opening official page with dev tools and copy http request as powershell
```
