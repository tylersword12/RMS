# Repair Management Suite

Windows desktop app for tracking machine repairs, service orders and customer
sign-off, built for field and shop technicians.

**Current version: 1.3.3**

## What it does

- Repair log with active and past jobs, split by **Field** vs **Shop**
- Service orders with parts lookup, labor, totals and on-screen customer signature
- Up to 10 named labor rates (Settings > Labor Rates in a service order),
  shared by every computer and picked from a dropdown
- **Customer copy PDF** of completed work, with the signature rendered on it
- Opens an **Outlook draft** with that PDF attached, addressed to the customer
- Brother label printing and barcode / webcam scanning
- Excel and CSV exports
- **Offline-safe**: work done with no connection is merged back in, never
  overwritten, when the share returns
- **Self-updating**: checks GitHub for a new release on launch, installs on close

## Installing

Download `RepairManagementSuite.exe` from the
[latest release](../../releases/latest) and put it anywhere the user can write
to — Desktop or a user folder. **Do not put it in Program Files**: the app
replaces its own EXE when updating and cannot do that without write access.

No Python needed. First launch asks for a company name and a shared data folder.

## Shipping an update to every machine

1. Make your changes.
2. Bump `APP_VERSION` near the top of `repair_management_suite.py`.
3. Commit, then tag and push:

```bash
git add -A
git commit -m "Describe the change"
git push

git tag v1.4.0          # must match APP_VERSION exactly
git push origin v1.4.0
```

GitHub Actions builds the Windows EXE and publishes it as a release. Every
installed copy finds it on next launch, downloads it in the background, and
offers to install it when the tech closes the program.

The build **fails on purpose** if the tag and `APP_VERSION` disagree — otherwise
the new build would ship but no machine would recognise it as newer.

To test a build without releasing it, run the workflow by hand from the
**Actions** tab; the EXE is attached to the run as an artifact.

## How updating works

- On launch a background thread asks the GitHub releases API for the latest
  version. A slow or blocked network never delays startup.
- A newer release is downloaded to a local staging folder and size-checked. A
  truncated download is discarded rather than installed.
- On exit the tech picks: install and restart, install and stay closed, or skip.
- A running EXE cannot overwrite itself, so a helper script waits for the
  process to release the file, backs it up to `.bak`, and copies the new one in.
- If the new build does not start within 80 seconds, the backup is restored
  automatically and the tech is told on next launch.

Machines with no internet simply never see an update; everything else keeps
working.

## Building locally

Run `build_repair_suite.bat` on Windows with Python installed. It creates a
virtual environment, installs dependencies and produces
`dist\RepairManagementSuite.exe`.

## Data and privacy

No credentials or customer data are stored in this repository. Company name,
file locations and all job data are entered at runtime and live in the shared
folder you configure, plus a local cache under
`%LOCALAPPDATA%\RepairManagementSuite`.
