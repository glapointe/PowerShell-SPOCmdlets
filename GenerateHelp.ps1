param(
  [string] $TargetDir = $(throw "TargetDir is required!"),
  [string] $ConfigurationName = $(throw "ConfigurationName is required!")
)
$path = Split-Path -parent $MyInvocation.MyCommand.Definition  
$helpAsm = "$($TargetDir)\Lapointe.PowerShell.MamlGenerator.dll"
$cmdletAsm = "$($TargetDir)\Lapointe.SharePointOnline.PowerShell.dll"
Write-Host "Help generation work path: $path"
Write-Host "Help generation maml assembly path: $helpAsm"
Write-Host "Help generation cmdlet assembly path: $cmdletAsm"

#Start-Process "C:\Program Files (x86)\Microsoft SDKs\Windows\v7.0A\bin\gacutil.exe" -ArgumentList "/uf","Lapointe.PowerShell.MamlGenerator"
#Start-Process "C:\Program Files (x86)\Microsoft SDKs\Windows\v7.0A\bin\gacutil.exe" -ArgumentList "/uf","Lapointe.SharePointOnline.PowerShell"

Write-Host "Loading help assembly..."
[System.Reflection.Assembly]::LoadFrom($helpAsm)
Write-Host "Loading cmdlet assembly..."
$asm = [System.Reflection.Assembly]::LoadFrom($cmdletAsm)
$asm
Write-Host "Generating help..."
[Lapointe.PowerShell.MamlGenerator.CmdletHelpGenerator]::GenerateHelp($asm, "$path", $true)

if ($ConfigurationName -eq "Release") {
	& "C:\Program Files (x86)\WiX Toolset v3.8\bin\candle.exe" -v -out "$TargetDir\install.wixobj" "$TargetDir\install.wxs"
	& "C:\Program Files (x86)\WiX Toolset v3.8\bin\light.exe" -ext WixUIExtension -out "$TargetDir\Lapointe.SharePointOnline.PowerShell.msi" "$TargetDir\install.wixobj"
}