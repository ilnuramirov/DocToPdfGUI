if (-not (Get-Module -ListAvailable ps2exe)) {
    Install-Module ps2exe -Scope CurrentUser -Force
}

Invoke-PS2EXE `
    -InputFile ".\DocToPdfGUI.ps1" `
    -OutputFile ".\DocToPdfGUI.exe" `
    -NoConsole `
    -Title "DOC to PDF Converter" `
    -Product "DOC to PDF Converter"