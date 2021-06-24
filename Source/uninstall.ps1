# Get all signature files in the Win32 app package
$signatureFiles = Get-ChildItem -Path "$PSScriptRoot\Signatures"

# Figure out the signature name and delete the corresponding items from the $env:APPDATA\Microsoft\Signatures folder
foreach ($signatureFile in $signatureFiles) {
    if ($signatureFile.getType().Name -eq 'DirectoryInfo') {
        $signatureName = $signatureFile.Name.Split("_")[0]
        Remove-Item -Path "$($env:APPDATA)\Microsoft\Signatures\$($signatureName)_files" -Recurse -Force
        Remove-Item -Path "$($env:APPDATA)\Microsoft\Signatures\$($signatureName).htm" -Force
        Remove-Item -Path "$($env:APPDATA)\Microsoft\Signatures\$($signatureName).rtf" -Force
        Remove-Item -Path "$($env:APPDATA)\Microsoft\Signatures\$($signatureName).txt" -Force
    }
}
