#=============================================================================#
# PowerShell script sample for coolOrange powerJobs                           #
# Creates a PDF file and add it to Autodesk Vault as Design Vizualization     #
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
#$file = Get-VaultFile -File "$/Designs/Test/ACADE/32155-E01/32155-E01.wdp"
#$file = Get-VaultFile -File "$/Designs/Test/ACADE/931979-E12/931979-E12.wdp"
#Add-VaultJob -Name "Sample.CreatePDF.ElectricalProjects" -Description "Create PDF for AutoCAD Electrical project" -Parameters @{EntityClassId = "FILE"; EntityId = $file.Id}
#-----------------------------------------------------------------------------#

$hidePDF = $false
$workingDirectory = "C:\temp\$($file._Name)"
$localPDFfileLocation = "$workingDirectory\$($file._Name).pdf"
$vaultPDFfileLocation = $file._EntityPath +"/"+ (Split-Path -Leaf $localPDFfileLocation)

Write-Host "Starting job 'Create PDF as attachment' for AutoCAD Electrical project '$($file._Name)' ..."

if( @("wdp") -notcontains $file._Extension ) {
    Write-Host "Files with extension: '$($file._Extension)' are not supported"
    return
}

$downloadedFiles = Save-VaultFile -File $file._FullPath
$file = $downloadedFiles | select -First 1
$openResult = Open-Document -LocalFile $file.LocalPath

if($openResult) {
    $configFile = "$($env:POWERJOBS_MODULESDIR)Export\PDF_AcadElectrical.ini"          
    $exportResult = Export-Document -Format 'PDF' -To $localPDFfileLocation -Options $configFile
    if($exportResult) {       
        $PDFfile = Add-VaultFile -From $localPDFfileLocation -To $vaultPDFfileLocation -FileClassification DesignVisualization -Hidden $hidePDF
        $file = Update-VaultFile -File $file._FullPath -AddAttachments @($PDFfile._FullPath)
    }
    $closeResult = Close-Document
}

Clean-Up -folder $workingDirectory

if(-not $openResult) {
    throw("Failed to open document $($file.LocalPath)! Reason: $($openResult.Error.Message)")
}
if(-not $exportResult) {
    throw("Failed to export document $($file.LocalPath) to $localPDFfileLocation! Reason: $($exportResult.Error.Message)")
}
if(-not $closeResult) {
    throw("Failed to close document $($file.LocalPath)! Reason: $($closeResult.Error.Message))")
}
Write-Host "Completed job 'Create PDF as attachment'"