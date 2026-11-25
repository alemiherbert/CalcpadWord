# Install

Run `CalcpadWord-Setup-<version>.exe`, then start Word.

- One file holds everything: the add-in, Calcpad's engine, a private .NET runtime and the fonts (math fonts, Libre Franklin).
- Per user, no administrator rights.
- Needs 64-bit Windows and 64-bit Word.
- **Update:** run the new Setup, then restart Word. An open Word keeps the old version until then.
- **Remove:** Settings → Apps → Calcpad for Word → Uninstall. The fonts stay.
- **Scripted:** `/quiet`, `/uninstall`.

## Blocked by Windows?

Setup isn't code-signed.
- SmartScreen: *More info → Run anyway*.
- "Windows cannot access the specified device…": right-click Setup → Properties → **Unblock**, or allow it in Windows Security.

[Home](../README.md)
