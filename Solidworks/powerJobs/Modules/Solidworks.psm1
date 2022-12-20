Add-Type -Path "$($Env:POWERJOBS_MODULESDIR)\Solidworks\coolOrange.Solidworks.dll"
Register-Application ([coolOrange.Solidworks.Application])