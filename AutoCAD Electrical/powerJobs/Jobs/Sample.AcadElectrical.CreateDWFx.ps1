#=============================================================================#
# PowerShell script sample for coolOrange powerJobs                           #
# Creates a DWFx file and add it to Autodesk Vault as Design Vizualization    #
#                                                                             #
# Copyright (c) coolOrange s.r.l. - All rights reserved.                      #
#                                                                             #
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER   #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.  #
#=============================================================================#

#- Debugging -----------------------------------------------------------------#
#Import-Module powerJobs
#Open-VaultConnection
#$file = Get-VaultFile -File "$/Designs/Test/ACADE/28871-E01revB/28871-E01.wdp"
#$file = Get-VaultFile -File "$/Designs/Test/ACADE/32155-E01/Test/32155-E01.wdp"
#$file = Get-VaultFile -File "$/Designs/Test/ACADE/931979-E12/931979-E12.wdp"
#Add-VaultJob -Name "Sample.CreateDWF.ElectricalProjects" -Description "Create DWF for AutoCAD Electrical project" -Parameters @{EntityClassId = "FILE"; EntityId = $file.Id}
#Add-VaultJob -Name "AM.CreateAcadeDWF" -Description "Create DWF for AutoCAD Electrical project" -Parameters @{EntityClassId = "FILE"; EntityId = $file.Id}
#-----------------------------------------------------------------------------#
#region Settings
# To include the Revision of the main file in the DWFX name set $true, otherwise $false
$dwfxFileNameWithRevision = $false

# The character used to separate file name and Revision label in the DWFX name such as hyphen (-) or underscore (_)
$dwfxFileNameRevisionSeparator = "_"

# To include the file extension of the main file in the DWFX name set $true, otherwise $false
$dwfFileNameWithExtension = $true

# To add the DWFX to Vault set $true, to keep it out set $false
$addDWFXToVault = $true

# To attach the DWFX to the main file set $true, otherwise $false
$attachDWFXToVaultFile = $true

# Specify a Vault folder in which the DWFX
x should be stored (e.g. $/Designs/DWFX), or leave the setting empty to store the DWFX next to the main file
$dwfxVaultFolder = ""

# Specify a network share into which the DWFX should be copied (e.g. \\SERVERNAME\Share\Public\DWF
xs\)
$dwfxNetworkFolder = ""

# To enable faster opening of released Inventor drawings without downloading and opening their model files set $true, otherwise $false
$openReleasedDrawingsFast = $true
#endregion

$dwfxFileName = [System.IO.Path]::GetFileNameWithoutExtension($file._Name)
if ($dwfxFileNameWithRevision) {
    $dwfxFileName += $dwfxFileNameRevisionSeparator + $file._Revision
}
if ($dwfxFileNameWithExtension) {
    $dwfxFileName += "." + $file._Extension
}
$dwfxFileName += ".dwfx"

if ([string]::IsNullOrWhiteSpace($dwfxVaultFolder)) {
    $dwfxVaultFolder = $file._EntityPath
}

Write-Host "Starting job 'Create DWFx as visualization attachment' for AutoCAD Electrical project '$($file._Name)' ..."

if( @("wdp") -notcontains $file._Extension ) {
    Write-Host "Files with extension: '$($file._Extension)' are not supported"
    return
}
if (-not $addDWFXToVault -and -not $dwfxNetworkFolder) {
    throw("No output for the DWFx is defined in ps1 file!")
}
if ($dwfxNetworkFolder -and -not (Test-Path $dwfxNetworkFolder)) {
    throw("The network folder '$dwfxNetworkFolder' does not exist! Correct dwfxNetworkFolder in ps1 file!")
}

$file = (Save-VaultFile -File $file._FullPath -DownloadDirectory $workingDirectory)[0]
$openResult = Open-Document -LocalFile $file.LocalPath

if($openResult) {
    $localDWFXfileLocation = "$workingDirectory\$dwfxFileName"    
    $configFile = "$($env:POWERJOBS_MODULESDIR)Export\PDF_AcadElectrical.ini"    
    $exportResult = Export-Document -Format 'DWFx' -To $localDWFfileLocation -Options $configFile

    if ($exportResult) {
        if ($addDWFXToVault) {
            $dwfxVaultFolder = $dwfxVaultFolder.TrimEnd('/')
            Write-Host "Add DWFx '$dwfxFileName' to Vault: $dwfxVaultFolder"
            $DWFXfile = Add-VaultFile -From $localDWFXfileLocation -To "$dwfxVaultFolder/$dwfxFileName" -FileClassification DesignVisualization
            if ($attachDWFXToVaultFile) {
                $file = Update-VaultFile -File $file._FullPath -AddAttachments @($DWFXfile._FullPath)
            }
        }
        if ($dwfxNetworkFolder) {
            Write-Host "Copy DWFx '$dwfxFileName' to network folder: $dwfxNetworkFolder"
            Copy-Item -Path $localDWFXfileLocation -Destination $dwfxNetworkFolder -ErrorAction Continue -ErrorVariable "ErrorCopyDWFXToNetworkFolder"
        }
    }
    $closeResult = Close-Document      
    if($exportResult) {       
        $DWFfile = Add-VaultFile -From $localDWFXfileLocation -To $vaultDWFfileLocation -FileClassification DesignVisualization -Hidden $hideDWF
        $file = Update-VaultFile -File $file._FullPath -AddAttachments @($DWFfile._FullPath)
    }
    $closeResult = Close-Document
}

if(-not $openResult) {
    throw("Failed to open document $($file.LocalPath)! Reason: $($openResult.Error.Message)")
}
if(-not $exportResult) {
    throw("Failed to export document $($file.LocalPath) to $localDWFXfileLocation! Reason: $($exportResult.Error.Message)")
}
if(-not $closeResult) {
    throw("Failed to close document $($file.LocalPath)! Reason: $($closeResult.Error.Message))")
}
if ($ErrorCopyPDFToNetworkFolder) {
    throw("Failed to copy DWFx file to network folder '$dwfxNetworkFolder'! Reason: $($ErrorCopyDWFXToNetworkFolder)")
}

Write-Host "Completed job 'Create DWFx as attachment'"