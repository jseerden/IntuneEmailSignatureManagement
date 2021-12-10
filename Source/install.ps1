# Win32 app runs PowerShell in 32-bit by default. AzureAD module requires PowerShell in 64-bit, so we are going to trigger a rerun in 64-bit.
if ($env:PROCESSOR_ARCHITEW6432 -eq "AMD64") {
    try {
        & "$env:WINDIR\SysNative\WindowsPowershell\v1.0\PowerShell.exe" -File $PSCommandPath
    }
    catch {
        throw "Failed to start $PSCommandPath"
    }
    exit
}

Start-Transcript -Path "$($env:TEMP)\IntuneSignatureManagerForOutlook-log.txt" -Force

# Install NuGet Package Provider
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force

# Install AzureAD module to retrieve the user information
Install-Module -Name AzureAD -Scope CurrentUser -Force

# Leverage Single Sign-on to sign into the AzureAD PowerShell module
$userPrincipalName = whoami -upn
Connect-AzureAD -AccountId $userPrincipalName

# Get the user information to update the signature
$userObject = Get-AzureADUser -ObjectId $userPrincipalName

# Create signatures folder if not exists
if (-not (Test-Path "$($env:APPDATA)\Microsoft\Signatures")) {
    $null = New-Item -Path "$($env:APPDATA)\Microsoft\Signatures" -ItemType Directory
}

# Get all signature files
$signatureFiles = Get-ChildItem -Path "$PSScriptRoot\Signatures"

foreach ($signatureFile in $signatureFiles) {
    if ($signatureFile.Name -like "*.htm" -or $signatureFile.Name -like "*.rtf" -or $signatureFile.Name -like "*.txt") {
        # Get file content with placeholder values
        $signatureFileContent = Get-Content -Path $signatureFile.FullName

        # Replace placeholder values
        $signatureFileContent = $signatureFileContent -replace "%DisplayName%", $userObject.DisplayName
        $signatureFileContent = $signatureFileContent -replace "%GivenName%", $userObject.GivenName
        $signatureFileContent = $signatureFileContent -replace "%Surname%", $userObject.Surname
        $signatureFileContent = $signatureFileContent -replace "%Mail%", $userObject.Mail
        $signatureFileContent = $signatureFileContent -replace "%Mobile%", $userObject.Mobile
        $signatureFileContent = $signatureFileContent -replace "%TelephoneNumber%", $userObject.TelephoneNumber
        $signatureFileContent = $signatureFileContent -replace "%JobTitle%", $userObject.JobTitle
        $signatureFileContent = $signatureFileContent -replace "%Department%", $userObject.Department
        $signatureFileContent = $signatureFileContent -replace "%City%", $userObject.City
        $signatureFileContent = $signatureFileContent -replace "%Country%", $userObject.Country
        $signatureFileContent = $signatureFileContent -replace "%StreetAddress%", $userObject.StreetAddress
        $signatureFileContent = $signatureFileContent -replace "%PostalCode%", $userObject.PostalCode
        $signatureFileContent = $signatureFileContent -replace "%Country%", $userObject.Country
        $signatureFileContent = $signatureFileContent -replace "%State%", $userObject.State
        $signatureFileContent = $signatureFileContent -replace "%PhysicalDeliveryOfficeName%", $userObject.PhysicalDeliveryOfficeName

        # Set file content with actual values in $env:APPDATA\Microsoft\Signatures
        Set-Content -Path "$($env:APPDATA)\Microsoft\Signatures\$($signatureFile.Name)" -Value $signatureFileContent -Force
    } elseif ($signatureFile.getType().Name -eq 'DirectoryInfo') {
        Copy-Item -Path $signatureFile.FullName -Destination "$($env:APPDATA)\Microsoft\Signatures\$($signatureFile.Name)" -Recurse -Force
    }
}

Stop-Transcript
