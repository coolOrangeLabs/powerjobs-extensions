[![Windows](https://img.shields.io/badge/Platform-Windows-lightgray.svg)](https://www.microsoft.com/en-us/windows/)
[![PowerShell](https://img.shields.io/badge/PowerShell-5-blue.svg)](https://microsoft.com/PowerShell/)
[![Vault](https://img.shields.io/badge/Autodesk%20Vault-2021-yellow.svg)](https://www.autodesk.com/products/vault/)
[![powerJobs](https://img.shields.io/badge/coolOrange%20powerJobs-21-orange.svg)](https://www.coolorange.com/en-eu/enhance.html#powerJobs)
## Disclaimer
THE SAMPLE CODE ON THIS REPOSITORY IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.
THE USAGE OF THIS SAMPLE IS AT YOUR OWN RISK AND **THERE IS NO SUPPORT** RELATED TO IT.
# AutoCAD-Electrical-PowerJobs-Extension
Custom powerJobs extension to export DWFx/PDFs from AutoCAD Electrical application.
### Getting Started
This is a custom powerJobs application which supports exporting DWFx/PDFs from AutoCAD Electrical application. 

#### Prerequisite
Following applications are required for running this job:
- AutoCAD Electrical
- powerJobs 2022 stream (v22.0.20 and above)
#### Installation
Use the installer found in the release page and follow the instructions. After the installation is complete, you will find 
- New jobs "Sample.CreateDWFx.ElectricalProjects.ps1" and "Sample.CreatePDF.ElectricalProjects.ps1" added to the Jobs folder. 
- New module 'powerJobsAcadElectrical.psm1' added to the modules folder 
- New custom powerJobs application 'coolOrange.AutoCADElectrical.dll' and its other dependent assemblies in the subfolder 'coolOrange.AcadElectrical' under the modules folder.

#### Setting and running the job
Rename the installed sample jobs so that your job can be identified as custom job and will not be overwritten by future updates. You can modify this copied job as you please.

Depending on the needs, the job can be triggered to run during lifecycle state change or in some other ways.
Further information on how to setup and run the job on lifecycle state change can be found on [powerJob's Getting Started documentation](https://doc.coolorange.com/projects/coolorange-powerjobsprocessordocs/en/stable/getting_started.html#how-to-embed-the-job-in-a-status-change)

### Upgrade the solution
The custom application provided here will work with powerJobs **2022 streams** (equal or greater than 22.0.20). For newer powerJobs' streams you will need to build and compile this solution. Follow the steps below to upgrade the powerJobs extension to support newer powerJobs version.

#### Compile the Solution
##### Prerequisite
- Find and install the wix toolset found on the [Wix](https://wixtoolset.org/). This is necessary because the the installer project is made using the wix toolset. 
- Install or upgrade powerJobs Processor on your development machine.
- Clone this repository
##### Build Solution
- Open the Visual Studio solution which can be found under the [AutoCAD Electrical](/AutoCAD%20Electrical).
- In Visual Studio right-click on References and click “Add References”.
- Search for the assembly powerJobs.Common” in Assemblies-tab and add it to your project.
[see here](https://doc.coolorange.com/projects/coolorange-powerjobsprocessordocs/en/stable/_images/vs_add_reference.png)
- Replace the UpgradeCode with new GUID found in [ProductVariables.wxi](https://github.com/coolOrangeLabs/powerjobs-extensions/blob/5e620d5beabb785b12b513263fa3934d2e2c27ce/AutoCAD%20Electrical/Installer/Includes/ProductVariables.wxi#L3)
- Build the solution

If the build has been successful and you don't have any errors, then you can use the installer 'powerJobs.AcadElectricalPlugin.Setup_1.0.0.0_x64.msi'  which you will find in the VS Installer project's output folder.

Further information on how to create custom application for powerJobs can be found [on the powerJobs wiki](https://doc.coolorange.com/projects/coolorange-powerjobsprocessordocs/en/stable/jobprocessor/applications.html#custom-applications).

## At your own risk
The usage of these samples is at your own risk. There is no free support related to the samples. However, if you have questions to powerJobs, then visit http://www.coolorange.com/wiki or start a conversation in our support forum at http://support.coolorange.com/support/discussions

## Author
coolOrange s.r.l.  
Channel Readiness Team

![coolOrange](https://user-images.githubusercontent.com/36075173/46519882-4b518880-c87a-11e8-8dab-dffe826a9630.png)
