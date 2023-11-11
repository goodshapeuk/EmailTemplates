<#
.SYNOPSIS
    PowerShell script to remediate an pre/non existing Email Signatures; from Set-OutlookSignatures script.

.EXAMPLE
    .\Remediate-EmailSignatures.ps1

.DESCRIPTION
    This PowerShell script is deployed as a remediation script using Microsoft Intune remediations.

.LINK
    https://github.com/alltimeuk/EmailTemplates/blob/main/Intune/Remediate-EmailSignatures.ps1

.LINK
    https://learn.microsoft.com/en-us/mem/intune/fundamentals/remediations

.NOTES
    Version:        1.0.2
    Creation Date:  2023-11-07
    Last Updated:   2023-11-08
    Author:         Simon Jackson / sjackson0109
    Organization:   Alltime Technologies Ltd
    Contact:        support@alltimetech.co.uk

#>
$temp = $(Get-Location).path
Start-Transcript $temp\Set-OutlookSignatures.log -Append

# Variables for Download and Extract
$githubProductOrg = "Set-OutlookSignatures"
$githubProductRepo = "Set-OutlookSignatures"
$githubTemplateOrg = "goodshapeuk"
$githubTemplateRepo = "EmailTemplates"

# Product Variables (standard)
$graphOnly = "true"
$SetOofMsg = "false"
$CreateRtfSignatures = "true"
$CreateTxtSignatures = "true"
$SignaturesForAutomappedAndAdditionalMailboxes = "true"
$DeleteUserCreatedSignatures = "false"  #REQ TRUE FOR GO-LIVE

# Product Variables (premium, req benefactor circle)
$DocxHighResImageConversion = "false"
$SetCurrentUserOutlookWebSignature = "true"
$MirrorLocalSignaturesToCloud = "true" #not used
$DeleteScriptCreatedSignaturesWithoutTemplate = "false" #not used


# Init
# Obtain the latest release off each github project  -- note: latest is always array item 0
$productUrl = "https://api.github.com/repos/$githubProductOrg/$githubProductRepo/tags"
$templateUrl = "https://api.github.com/repos/$githubTemplateOrg/$githubTemplateRepo/tags"
$productMeta = (Invoke-WebRequest $productUrl | ConvertFrom-Json)[0]
$templateMeta = (Invoke-WebRequest $templateUrl | ConvertFrom-Json)[0]

# Specify the file-system of the downloaded targets
$productZip = "$temp\Set-OutlookSignatures.zip"
$templateZip = "$temp\EmailTemplates.zip"
$productPath = "$temp\$githubProductOrg-$githubProductRepo-$($($productMeta.commit.sha).substring(0,7))"
$templatePath = "$temp\$githubTemplateOrg-$githubTemplateRepo-$($($templateMeta.commit.sha).substring(0,7))" 

Add-Type -AssemblyName System.IO.Compression.FileSystem


# Check if the latest version is already downloaded, clean up the file-system and download+extract, or just extract again
If (Test-Path $productPath){
    Write-Host "Cleaning up local path $productPath"
    Remove-Item $productPath -recurse -Force
} else {
    Write-Host "Downloading $productUrl to $productZip"
    Invoke-WebRequest "$($productMeta.zipball_url)" -Out $productZip
}

If (Test-Path $templatePath){
    Write-Host "Cleaning up local path $templatePath"
    Remove-Item $templatePath -recurse -Force
} else {
    Write-Host "Downloading $templateUrl to $templateZip"
    Invoke-WebRequest "$($templateMeta.zipball_url)" -Out $templateZip
}
Write-host "==============="
Get-ChildItem $temp
Write-host "==============="


# A fresh Extraction of the zipball files to the temp directory, filename encoding needs converting to ascii, not utf8.
# Note: some errors with file-name length when testing with my user docs area. C:\WINDOWS\IMECache\HealthScripts\(GUID)\ is just as long, so skip errors. Only signature samples anyway, don't need them.
Write-Host "Extracting $productZip to $temp"
[System.IO.Compression.ZipFile]::ExtractToDirectory("$productZip", "$temp", [System.Text.Encoding]::ascii) | out-null
Write-Host "Extracting $templateZip to $temp"
[System.IO.Compression.ZipFile]::ExtractToDirectory("$templateZip", "$temp", [System.Text.Encoding]::ascii) | out-null

Write-host "==============="
Get-ChildItem $temp
Write-host "==============="

# Gather some path data
$productTargetPath = "$temp\$githubProductOrg-$githubProductRepo-$($($productMeta.commit.sha).substring(0,7))\src_Set-OutlookSignatures"
$templateTargetPath = "$temp\$githubTemplateOrg-$githubTemplateRepo-$($($templateMeta.commit.sha).substring(0,7))"
$executionPath = "$githubProductOrg-$githubProductRepo-$($($productMeta.commit.sha).substring(0,7))\src_Set-OutlookSignatures"
# Clean up the downloaded content
#Remove-Item -Path $productZip -Force
#Remove-Item -Path $templateZip -Force


#Run product, with transcript logging, and args passed from variables above
Set-Location $temp\$executionPath
.\Set-OutlookSignatures.ps1 -graphonly $graphOnly -SignatureTemplatePath $templateTargetPath\Signatures -SignatureIniPath $templateTargetPath\Signatures\_Signatures.ini -SetCurrentUserOOFMessage $SetOofMsg -CreateRtfSignatures $CreateRtfSignatures -CreateTxtSignatures $CreateTxtSignatures -SignaturesForAutomappedAndAdditionalMailboxes $SignaturesForAutomappedAndAdditionalMailboxes -DisableRoamingSignatures $DisableRoamingSignatures -SetCurrentUserOutlookWebSignature $SetCurrentUserOutlookWebSignature -DeleteUserCreatedSignatures $DeleteUserCreatedSignatures -DeleteScriptCreatedSignaturesWithoutTemplate $DeleteScriptCreatedSignaturesWithoutTemplate
Set-Location $temp
Stop-Transcript 
exit 0