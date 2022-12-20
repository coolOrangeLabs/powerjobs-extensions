Add-Type -Path "$($Env:POWERJOBS_MODULESDIR)\Solidworks\Solidworks.dll"
Register-Application ([Solidworks.Application])