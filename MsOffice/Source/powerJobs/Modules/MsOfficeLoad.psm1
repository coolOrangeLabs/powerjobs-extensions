Add-Type -Path "$($Env:POWERJOBS_MODULESDIR)\MsOfficeSupport\MsOfficeApplication.dll"
Register-Application ([MsOfficeApplication.Word.MsWordApplication])
Register-Application ([MsOfficeApplication.Excel.MsExcelApplication])
Register-Application ([MsOfficeApplication.PowerPoint.MsPowerPointApplication])
