# Releasing a new version

Builds are produced by GitHub Actions on `windows-latest` and published as
release assets. Installed copies check for a newer release on launch and offer
to install it on close, so tagging a release is the only step needed to roll an
update out to every machine.

## Steps

1. Make the change.
2. Bump `APP_VERSION` near the top of `repair_management_suite.py`.
3. Commit and push:

```bash
git add -A
git commit -m "Describe the change"
git push
```

4. Tag and push the tag:

```bash
git tag v1.3.1          # must match APP_VERSION exactly
git push origin v1.3.1
```

The workflow builds the EXE and attaches it to a new release. The build fails
deliberately if the tag and `APP_VERSION` disagree, because a mismatch would
publish a build that no installed copy recognises as newer.

## Testing a build without releasing it

Run the workflow manually from the **Actions** tab. It builds and attaches the
EXE to the run as an artifact without creating a release, so nothing reaches
the field until a tag is pushed.

## Pulling a bad release

Mark the release as a pre-release or delete it — the update check ignores both.
Technicians who have already downloaded it can choose **Skip This Time** on the
exit prompt.

## Installation notes

The EXE must live somewhere the signed-in user can write to, such as the
Desktop or a user folder. It replaces its own executable when updating, which
cannot be done from `Program Files` without elevation. The app detects an
unwritable location and reports it rather than failing silently.

## Build dependencies

Pinned in `requirements.txt`. PyInstaller flags live in
`.github/workflows/build-windows.yml`. Dependency problems with OpenCV, pyzbar
and brother_ql surface on the runner rather than in the source, so check the
Actions log if a build fails.

## Local builds

`build_repair_suite.bat` produces the same EXE on a Windows machine with Python
installed. It creates its own virtual environment and writes to `dist\`.
