#==============================================================================#
# PowerShell Module - Application helper for coolOrange powerJobs applications #
# Copyright (c) coolOrange s.r.l. - All rights reserved.                       #
#                                                                              #
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER    #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES  #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.   #
#=============================================================================#


Add-Type -Path "$($Env:POWERJOBS_MODULESDIR)\coolOrange.AcadElectrical\coolOrange.AutoCADElectrical.dll"
Register-Application ([coolOrange.AutoCADElectrical.Application])