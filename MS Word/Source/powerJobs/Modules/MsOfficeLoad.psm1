Add-Type -Path "$($Env:POWERJOBS_MODULESDIR)\MsOfficeSupport\MsOfficeApplication.dll"
Register-Application ([MsOfficeApplication.MsWordApplication])
Register-Application ([MsOfficeApplication.MsExcelApplication])
Register-Application ([MsOfficeApplication.MsPowerPointApplication])
