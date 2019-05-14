# RelocateOutlookOST

A logon script to force Outlook to relocate/rebuild the OST after a change of ForceOSTPath

After changing the registry key ForceOSTPath (either using a GPO or by direct  registry editing) the OST file is not relocated by Outlook. This script will force Outlook to rebuild the OST in the new location. Preferably run this script at logon.

## Limitations, pending features and assumptions

- Tested and developed on Windows 2016/10 with Outlook 2016/365
- Only the default Outlook profile is relocated
- Relocation is done by rebuilding the OST and not by moving the file
- No other files are relocated
- The script assumes Outlook is installed and does not test for that