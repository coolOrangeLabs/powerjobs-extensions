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

$hideDWF = $false
$workingDirectory = "C:\temp\$($file._Name)"
$localDWFfileLocation = "$workingDirectory\$($file._Name).dwfx"
$vaultDWFfileLocation = $file._EntityPath +"/"+ (Split-Path -Leaf $localDWFfileLocation)

Write-Host "Starting job 'Create DWF as attachment' for AutoCAD Electrical project '$($file._Name)' ..."

if( @("wdp") -notcontains $file._Extension ) {
    Write-Host "Files with extension: '$($file._Extension)' are not supported"
    return
}

$downloadedFiles = Save-VaultFile -File $file._FullPath
$file = $downloadedFiles | select -First 1
$openResult = Open-Document -LocalFile $file.LocalPath

if($openResult) {
    $configFile = "$($env:POWERJOBS_MODULESDIR)Export\PDF_AcadElectrical.ini"
    if(-not (Test-Path -Path $workingDirectory)) { New-Item -Path $workingDirectory -ItemType "directory"}          
    $exportResult = Export-Document -Format 'DWFx' -To $localDWFfileLocation -Options $configFile
    if($exportResult) {       
        $DWFfile = Add-VaultFile -From $localDWFfileLocation -To $vaultDWFfileLocation -FileClassification DesignVisualization -Hidden $hideDWF
        $file = Update-VaultFile -File $file._FullPath -AddAttachments @($DWFfile._FullPath)
    }
    $closeResult = Close-Document
}

Clean-Up -folder $workingDirectory

if(-not $openResult) {
    throw("Failed to open document $($file.LocalPath)! Reason: $($openResult.Error.Message)")
}
if(-not $exportResult) {
    throw("Failed to export document $($file.LocalPath) to $localDWFfileLocation! Reason: $($exportResult.Error.Message)")
}
if(-not $closeResult) {
    throw("Failed to close document $($file.LocalPath)! Reason: $($closeResult.Error.Message))")
}
Write-Host "Completed job 'Create DWF as attachment'"