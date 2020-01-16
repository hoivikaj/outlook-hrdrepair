# outlook-hrdrepair
Script to correct Outlook saved credentials in non HRD friendly format

It is common for users to enter their credentials in Outlook Anywhere workflows in domain\username format when using legacy authentication. If you attempt to enable modern authentication while users have this value saved in the registry, Home-Realm Discovery will fail because the username string does not have the domain in the proper format for auto acceleration.

This script fixes the user_hint stored in netbios\ format and converts to UPN. This issue is really only seen during Legact to Modern Authenication Go-Live for Office 365 Hybrid.

## How it Works

Outlook stores the user_hint value in hex in the registry. While this value appears to be originally derived from your manually entered creds, this particular value is independent.

When the user_hint value is in domain\user format - Home Realm Discovery cannot run properly (expects FQDN via UPN or Email). This setting is stored in this key: HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles<Profile Name><WEF Provider ID> with value 001f3d16 of type REG_BINARY.

This script recursively searches the Profiles Key for this value (which appears to be unique to this logon type) and if the data of that value contains a , converts the logon format to UPN and re-injects.

This does NOT appear to break existing saved-password logon workflow, removing saved credentials does not appear to be required.
