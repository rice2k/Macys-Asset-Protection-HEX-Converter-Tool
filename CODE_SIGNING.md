# Code Signing Plan

Windows code signing is not enabled yet because it requires a trusted signing certificate.

When a certificate is available, the recommended release flow is:

1. Build the Windows EXE.
2. Sign the EXE with the certificate using Microsoft SignTool.
3. Generate a fresh SHA-256 checksum after signing.
4. Attach the signed EXE and checksum to the GitHub Release.
5. Keep the unsigned build only as a developer artifact, not the main user download.

This is tracked as a future trust improvement so Windows is less likely to warn users when launching the utility.
