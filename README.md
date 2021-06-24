# Intune Email Signature Manager for Outlook

This Intune Win32 app queries Azure Active Directory as the signed in user and creates Signatures in Outlook for the user.

## Requirements
- Requires Windows 10 Azure AD Joined devices managed with Microsoft Intune.
- Application must be deployed to run in **User** context.

* The app leverages the -AccountId parameter of the Connect-AzureAD cmdlet for Single Sign-On. Please note that this has only been tested on Azure AD Joined devices.

## How does it work?
1. Replace or update the files in the **Source\Signatures** folders with one or more Signature template(s) you would like to use. For example you can create a Signature in Outlook and obtain the files from %APPDATA%\Microsoft\Signatures.

2. Modify the Signature files to include placeholder values. Supported placeholder values for the templates are listed below. *Note: It is important that the actual values are available on the Azure AD user, either managed from Active Directory or directly in Office 365 / Azure AD.*

3. Package the source folder with the Microsoft Win32 Content Prep Tool, for example:
`IntuneWinAppUtil.exe -c '.\Source' -s '.\Source\install.ps1' -o '.\Package'`

4. Deploy the .intunewin app with Microsoft Intune to your users!

## Supported placeholder values

- %DisplayName%
- %GivenName%
- %Surname%
- %Mail%
- %Mobile%
- %TelephoneNumber%
- %JobTitle%
- %Department%
- %City%
- %Country%
- %StreetAddress%
- %PostalCode%
- %Country%
- %State%
- %PhysicalDeliveryOfficeName%

## Deploying the Win32 app

### Install command
`PowerShell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "install.ps1"`

### Uninstall command
`PowerShell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "uninstall.ps1"`

### Install behavior
User

## Detection rules
- Manually configure detection rules

Example:
 - Rule type: File
 - Path: %APPDATA%\Microsoft\Signatures
 - File or folder: Default signature.htm
 - Detection method: File or folder exists

You can change the signature's display name in Outlook by changing the file names in the **Source\Signatures** folder. Make sure to translate the changes into the detection rules!
