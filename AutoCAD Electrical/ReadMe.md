[![Windows](https://img.shields.io/badge/Platform-Windows-lightgray.svg)](https://www.microsoft.com/en-us/windows/)
[![PowerShell](https://img.shields.io/badge/PowerShell-5-blue.svg)](https://microsoft.com/PowerShell/)
[![Vault](https://img.shields.io/badge/Autodesk%20Vault-2021-yellow.svg)](https://www.autodesk.com/products/vault/)
[![powerJobs](https://img.shields.io/badge/coolOrange%20powerJobs-21-orange.svg)](https://www.coolorange.com/en-eu/enhance.html#powerJobs)
## Disclaimer
THE SAMPLE CODE ON THIS REPOSITORY IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.
THE USAGE OF THIS SAMPLE IS AT YOUR OWN RISK AND **THERE IS NO SUPPORT** RELATED TO IT.
# AutoCAD-Electrical-PowerJobs-Extension
Custom powerJobs application to export DWFx/PDFs from AutoCAD Electrical application.
### Getting Started
This is a custom powerJobs application which supports exporting DWFx/PDFs from AutoCAD Electrical application. 

### Prerequisite
Following applications are required for running this job:
- AutoCAD Electrical
- [powerJobs v22.0.20 and above](https://www.coolorange.com/powerjobs)
### Installation
Use the installer found in the release page and follow the instructions. After the installation is complete, you will find 
- new jobs "Sample.CreateDWFx.ElectricalProjects.ps1" and "Sample.CreatePDF.ElectricalProjects.ps1" added to the Jobs folder. 
- new module 'powerJobsAcadElectrical.psm1' added to the modules folder 
- new custom powerJobs application 'coolOrange.AutoCADElectrical.dll' and its other dependent assemblies in the subfolder 'coolOrange.AcadElectrical' under the modules folder.

Now you are good to go. Either directly use the new job and publish your DWFx/PDFs or you can change the provided sample job to your likings before using the job.

### Upgrade the application to support new Major version powerJobs
The custom application provided here will work with powerJob's <b><i>updates</i></b> versions which are equal and greater than 22.0.20. For newer version of powerJobs you will need to build and compile this solution. Follow the steps below to upgrade the custom application to support newer powerJobs version. 
Furhter information on how to create custom application for powerJobs can be found [on the powerJobs wiki](https://doc.coolorange.com/projects/coolorange-powerjobsprocessordocs/en/stable/jobprocessor/applications.html#custom-applications).

#### Compile the solution
- install or upgrade powerJobs Processor on your development machine. 
- Open the Visual Studio solution which can be found under the [AutoCAD Electrical/Source/coolOrange.AcadElectrical](AutoCAD%20Electrical/Source/coolOrange.AcadElectrical).
- In Visual Studio right-click on References and click “Add References”.
- Search for the assembly powerJobs.Common” in Assemblies-tab and add it to your project.
[see here](https://doc.coolorange.com/projects/coolorange-powerjobsprocessordocs/en/stable/_images/vs_add_reference.png)
- build the solution

If the build has been successful and you don't have any errors, you will find the new assembly in powerJobs Module folder - [coolOrange.AcadElectrical](AutoCAD%20Electrical/Source/powerJobs/Modules/coolOrange.AcadElectrical/)
#### Build the installer
- find and install the wix toolset found on the [Wix](https://wixtoolset.org/). This is necessary because the the installer solution is made using the wix toolset.
- Open the Visual Studio solution 'Installer.sln' found within the [Installer](AutoCAD%20Electrical/Installer).
- installer uses the heat to harvest the files and folders from the [AutoCAD Electrical/Source/powerJobs](AutoCAD%20Electrical/Source/powerJobs)
- replace the UpgradeCode with another GUID found in [ProductVariables.wxi](/Installer/powerJobs.AcadElectricalPlugin.Setup/Includes/)
- build the solution

If the build is successful, then you will find 'powerJobs.AcadElectricalPlugin.Setup_1.0.0.0_x64.msi' in the output folder.

## At your own risk
The usage of these samples is at your own risk. There is no free support related to the samples. However, if you have questions to powerJobs, then visit http://www.coolorange.com/wiki or start a conversation in our support forum at http://support.coolorange.com/support/discussions

## Author
coolOrange s.r.l.  
Channel Readiness Team

![coolOrange](https://user-images.githubusercontent.com/36075173/46519882-4b518880-c87a-11e8-8dab-dffe826a9630.png)
