# outlook-hrdrepair
Script to correct Outlook saved credentials in non HRD friendly format

Fixes user_hint being stored in netbios\ format which breaks HRD during Modern Authenication Go-Live for Office 365 Hybrid.

Outlook stores the user_hint value in hex in the registry. While this value appears to be originally derived from your manually entered creds, this particular value is independent.

When the user_hint value is in domain\user format - Home Realm Discovery cannot run properly (expects FQDN via UPN or Email). This setting is stored in this key: HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles<Profile Name><WEF Provider ID> with value 001f3d16 of type REG_BINARY.

This script recursively searches the Profiles Key for this value (which appears to be unique to this logon type) and if the data of that value contains a , converts the logon format to UPN and re-injects.

This does NOT appear to break existing saved-password logon workflow, removing saved credentials does not appear to be required.
