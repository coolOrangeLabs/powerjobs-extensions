#==============================================================================#
# (c) 2011 coolOrange s.r.l.                                                   #
#                                                                              #
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER    #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES  #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.   #
#==============================================================================#

#region Settings
# To include the Revision of the main file in the PDF name set $true, otherwise $false
$pdfFileNameWithRevision = $false

# The character used to separate file name and Revision label in the PDF name such as hyphen (-) or underscore (_)
$pdfFileNameRevisionSeparator = "_"

# To include the file extension of the main file in the PDF name set $true, otherwise $false
$pdfFileNameWithExtension = $true

# To add the PDF to Vault set $true, to keep it out set $false
$addPDFToVault = $true

# To attach the PDF to the main file set $true, otherwise $false
$attachPDFToVaultFile = $true

# Specify a Vault folder in which the PDF should be stored (e.g. $/Designs/PDF), or leave the setting empty to store the PDF next to the main file
$pdfVaultFolder = ""

# Specify a network share into which the PDF should be copied (e.g. \\SERVERNAME\Share\Public\PDFs\)
$pdfNetworkFolder = ""

# To enable faster opening of released Inventor drawings without downloading and opening their model files set $true, otherwise $false
$openReleasedDrawingsFast = $true
#endregion

$pdfFileName = [System.IO.Path]::GetFileNameWithoutExtension($file._Name)
if ($pdfFileNameWithRevision) {
    $pdfFileName += $pdfFileNameRevisionSeparator + $file._Revision
}
if ($pdfFileNameWithExtension) {
    $pdfFileName += "." + $file._Extension
}
$pdfFileName += ".pdf"

if ([string]::IsNullOrWhiteSpace($pdfVaultFolder)) {
    $pdfVaultFolder = $file._EntityPath
}

Write-Host "Starting job 'Create PDF as visualization attachment' for file '$($file._Name)' ..."

if ( @("idw", "dwg","slddrw") -notcontains $file._Extension ) {
    Write-Host "Files with extension: '$($file._Extension)' are not supported"
    return
}
if (-not $addPDFToVault -and -not $pdfNetworkFolder) {
    throw("No output for the PDF is defined in ps1 file!")
}
if ($pdfNetworkFolder -and -not (Test-Path $pdfNetworkFolder)) {
    throw("The network folder '$pdfNetworkFolder' does not exist! Correct pdfNetworkFolder in ps1 file!")
}

$fastOpen = $openReleasedDrawingsFast -and $file._ReleasedRevision
$file = (Save-VaultFile -File $file._FullPath -DownloadDirectory $workingDirectory -ExcludeChildren:$fastOpen -ExcludeLibraryContents:$fastOpen)[0]
$openResult = Open-Document -LocalFile $file.LocalPath -Options @{ FastOpen = $fastOpen }

if ($openResult) {
    $localPDFfileLocation = "$workingDirectory\$pdfFileName"
    if ($openResult.Application.Name -like 'Inventor*') {
        $configFile = "$($env:POWERJOBS_MODULESDIR)Export\PDF_2D.ini"
    }
    else {
         $configFile = "$($env:POWERJOBS_MODULESDIR)Export\PDF.dwg"
    }

    $exportResult = Export-Document -Format 'PDF' -To $localPDFfileLocation -Options $configFile

    if ($exportResult) {
        if ($addPDFToVault) {
            $pdfVaultFolder = $pdfVaultFolder.TrimEnd('/')
            Write-Host "Add PDF '$pdfFileName' to Vault: $pdfVaultFolder"
            $PDFfile = Add-VaultFile -From $localPDFfileLocation -To "$pdfVaultFolder/$pdfFileName" -FileClassification DesignVisualization 
            if ($attachPDFToVaultFile) {
                $file = Update-VaultFile -File $file._FullPath -AddAttachments @($PDFfile._FullPath)
            }
        }
        if ($pdfNetworkFolder) {
            Write-Host "Copy PDF '$pdfFileName' to network folder: $pdfNetworkFolder"
            Copy-Item -Path $localPDFfileLocation -Destination $pdfNetworkFolder -ErrorAction Continue -ErrorVariable "ErrorCopyPDFToNetworkFolder"
        }
    }
    $closeResult = Close-Document
}

if (-not $openResult) {
    throw("Failed to open document $($file.LocalPath)! Reason: $($openResult.Error.Message)")
}
if (-not $exportResult) {
    throw("Failed to export document $($file.LocalPath) to $localPDFfileLocation! Reason: $($exportResult.Error.Message)")
}
if (-not $closeResult) {
    throw("Failed to close document $($file.LocalPath)! Reason: $($closeResult.Error.Message))")
}
if ($ErrorCopyPDFToNetworkFolder) {
    throw("Failed to copy PDF file to network folder '$pdfNetworkFolder'! Reason: $($ErrorCopyPDFToNetworkFolder)")
}

Write-Host "Completed job 'Create PDF as visualization attachment'"