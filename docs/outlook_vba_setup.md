# HVDC Outlook VBA Automation Setup

## Overview

This package delivers hardened Outlook VBA automation for the HVDC email capture pipeline. It exports every incoming message as `.msg`, preserves attachments, and emits JSON sidecars for downstream ingestion. The automation runs in real time with hourly catch-up protection and structured logging.

## Prerequisites

- Outlook LTSC 2021 on Windows 11 Enterprise
- Macro security configured for signed code or trusted location
- Access to the Windows account that will run Outlook continuously
- Write permissions to `%LOCALAPPDATA%\HVDC`

## Import Instructions

1. Launch Outlook with macros enabled.
2. Press `ALT+F11` to open the VBA editor.
3. In the Project pane, remove previous HVDC modules if present.
4. Choose **File → Import File…** and import each module/class from `outlook_vba`:
   - `Modules/modHVDC_Config.bas`
   - `Modules/modHVDC_FileIO.bas`
   - `Modules/modHVDC_Logging.bas`
   - `Modules/modHVDC_Export.bas`
   - `Classes/clsHVDC_InboxWatcher.cls`
   - `ThisOutlookSession.cls` (overwrite the existing `ThisOutlookSession` code)
5. Save the VBA project (`ALT+F S`).

## Macro Security

- Prefer signed VBA projects. If signatures are not available, add `%LOCALAPPDATA%\Microsoft\Outlook` to the trusted locations list and store the VBA project there.
- Confirm that **Trust access to the VBA project object model** is enabled.

## First-Run Checklist

1. Execute `InitializeHVDC` from the VBA Immediate window (`Ctrl+G`) to provision folders, logging, the processed cache, and hidden tasks.
2. Restart Outlook to ensure `Application_Startup` wires all watchers.
3. Send a test message to the monitored Inbox and verify that:
   - `.msg` and `.json` files appear in `%LOCALAPPDATA%\HVDC\exports\YYYY\MM\DD`.
   - Attachments are written under `...\attachments\<EntryID>`.
   - Logs are written to `%LOCALAPPDATA%\HVDC\logs\hvdc_pipeline.log`.
4. Inspect `%LOCALAPPDATA%\HVDC\state\catchup_checkpoint.json` for the initialized timestamp.

## Configuration Notes

- To watch additional folders under Inbox, edit `ADDITIONAL_FOLDERS` in `modHVDC_Config`. Separate multiple entries with `|` (example: `"Inbox\\Operations|Inbox\\Escalations"`).
- Adjust batching or throttling via `DEFAULT_BATCH_SIZE` and `DEFAULT_QUEUE_INTERVAL_SECONDS` constants in `modHVDC_Config`.
- Logs rotate automatically at ~10 MB across five files.

## Operational Guidance

- Hourly resilience: A hidden task named **HVDC Hourly Tick** drives the catch-up enumerator. Do not delete this task.
- Real-time queue: A hidden task named **HVDC Queue Pulse** schedules batched processing without blocking the UI.
- To reset the processed cache, delete `%LOCALAPPDATA%\HVDC\state\processed_index.json` and run `InitializeHVDC` again.

## Support

- Logs are JSON lines and include metrics per batch.
- Failures are non-modal; investigate the latest log file for diagnostic context.
- For disaster recovery, re-run `InitializeHVDC` and relaunch Outlook.


## Troubleshooting

- **"Too many line continuations" when importing modules**: Ensure the files retain Windows line endings. Clone or unzip the repository on Windows, or convert the `.bas/.cls` files to `CRLF` before import (the repository now enforces this automatically via `.gitattributes`).
