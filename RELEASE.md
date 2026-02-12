# Release Checklist

Use this checklist when publishing a GitHub Release for the Windows executable.

## Pre-Release

- [ ] Ensure the working tree is clean and all changes are committed
- [ ] Confirm version number (semantic versioning: `vX.Y.Z`)
- [ ] Update any user-facing docs or notes if needed

## Create Release Tag

- [ ] Create tag locally: `git tag -a vX.Y.Z`
- [ ] Push tag: `git push origin vX.Y.Z`

## GitHub Actions Build

- [ ] Verify the GitHub Actions workflow started for the tag
- [ ] Confirm build succeeded on Windows runner
- [ ] Confirm release assets were uploaded:
  - [ ] `PyPDF_Toolbox-vX.Y.Z-win64.zip`
  - [ ] `PyPDF_Toolbox-vX.Y.Z-win64.zip.sha256.txt`

## Post-Release

- [ ] Download ZIP and test launch on Windows
- [ ] Verify SHA256 checksum matches the downloaded ZIP
- [ ] Add release notes or highlights if needed
