# RSR

RSR records a Windows workflow, generates a step-by-step PDF with screenshots, and creates a plan to replay the workflow later. You can add an intent note when an action needs more explanation or should adapt on future runs.

## What is included

- `DesktopSteps/`: C# source for the Windows Forms application.
- `bin/RSR/`: the latest framework-dependent Windows build.
- `.env.example`: a safe configuration template. The included `.env` also contains placeholders only and is ignored by Git.

This package contains no recordings, screenshots, personal Foundry endpoint, or backup builds.

## Requirements

- Windows with an unlocked interactive desktop session.
- .NET 10 Windows Desktop Runtime to run the included executable. The .NET 10 SDK is needed only to build from source.
- Azure CLI and an account permitted to invoke your configured Microsoft Foundry agent.
- A Microsoft Foundry project endpoint and agent name for generating execution plans.

## Configure and run

1. Replace the placeholder values in the repository-root `.env`, or copy `.env.example` to `.env` and edit it. Use the **project endpoint**, not a model endpoint. The resource, project, and agent names need not match. Sample content:

# Local configuration template. Replace these placeholders before running RSR.
# Never commit a .env file containing your real endpoint or other private values.

PROJECT_ENDPOINT=https://YOUR-RESOURCE.services.ai.azure.com/api/projects/YOUR-PROJECT
AGENT_NAME=YOUR-AGENT-NAME

2. Sign in with `az login` using an account with access to that Foundry project and agent.
3. Run `bin/RSR/DesktopSteps.exe`.

To build from source instead, run:

```powershell
dotnet run --project DesktopSteps/DesktopSteps.csproj
```

The application searches its own folder and parent folders for `.env`. New recordings are saved in `Recordings` beside the executable being run. For the included build, that is `bin/RSR/Recordings`. This folder is created when needed and is ignored by Git.

## Use

1. Start recording and perform the task on your desktop. Optionally press **Ctrl+Alt+Shift+I** during recording to add an intent note to the latest action.
2. Stop recording. RSR saves the captured events, execution plan, summary, screenshots, and `manual_steps.pdf` in a timestamped recording folder.
3. Review the plan and PDF, then choose **Replay selected** to repeat the workflow. Leave the desktop available while replay runs.

RSR captures actions and UI control information through Windows accessibility APIs. It saves screenshots for the PDF and human review; it does not run OCR or send screenshot pixels to Foundry. The planning request uses recorded action metadata and optional intent notes. Replay uses Windows UI Automation and stops or asks for help when it cannot safely identify an action. Elevated applications may require running RSR as administrator to record their controls.

The project has been tried with parts of Excel, DebugDiag 2 Collection, and Microsoft Edge. Results depend on each application's accessibility support and current UI state; review a workflow before relying on unattended replay.
