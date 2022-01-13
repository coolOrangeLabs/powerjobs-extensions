[![Windows](https://img.shields.io/badge/Platform-Windows-lightgray.svg)](https://www.microsoft.com/en-us/windows/)
[![PowerShell](https://img.shields.io/badge/PowerShell-5-blue.svg)](https://microsoft.com/PowerShell/)
[![Vault](https://img.shields.io/badge/Autodesk%20Vault-2021-yellow.svg)](https://www.autodesk.com/products/vault/)
[![powerJobs](https://img.shields.io/badge/coolOrange%20powerJobs-21-orange.svg)](https://www.coolorange.com/en-eu/enhance.html#powerJobs)

## Disclaimer
THE SAMPLE CODE ON THIS REPOSITORY IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.
THE USAGE OF THIS SAMPLE IS AT YOUR OWN RISK AND **THERE IS NO SUPPORT** RELATED TO IT.

# MsOffice-PowerJobs-Extension
Custom powerJobs extension to export PDFs from Microsoft Office Applications - Word, Excel, PowerPoint.
### Getting Started
This is a custom powerJobs application which supports exporting PDFs from Microsoft Office Applications - Word, Excel, PowerPoint. 
#### Prerequisite
Following applications are required for running this job:
-   Microsoft Office (Either Word or Excel or PowerPoint or all of the above depending on the requirements)
-   powerJobs 2022 stream (v22.0.20 and above)

#### Installation
Use the installer found in the release page and follow the instructions. After the installation is complete, you will find 
- new job "Sample.MsOffice.CreatePDF.ps1" added to the Jobs folder. 
- new module 'MsOfficeLoad.psm1' added to the modules folder 
- new custom powerJobs application 'MsOfficeApplication.dll' in the subfolder 'MsOfficeSupport' under the modules folder.

#### Setting and running the job
Rename the job Sample.MsOffice.CreatePDF.ps1 to Custom.CreatePDF.ps1 so your PDF job can be identified as your own job and will not be overwritten by future updates. You can modify this copied job as you please.

Depending on the needs, the job can be configured to run automatically during lifecycle state change or in some other ways.
Further information on how to setup and run the job on lifecycle state change can be found on [powerJob's Getting Started documentation](https://doc.coolorange.com/projects/coolorange-powerjobsprocessordocs/en/stable/getting_started.html#how-to-embed-the-job-in-a-status-change)

### Upgrade the powerJobs extension
The powerJobs extension will work only with powerJobs **2022 streams** (equal or greater than 22.0.20). For newer powerJobs' streams you will need to build and compile this solution. Follow the steps below to upgrade the powerJobs extension to support newer powerJobs version.

#### Compile the solution
- install or upgrade powerJobs Processor on your development machine. 
- Open the Visual Studio solution which can be found under [MsOffice-PowerJobs-Extension](MsOffice/Source/MsOffice-PowerJobs-Extension).
- In Visual Studio right-click on References and click “Add References”.
- Search for the assembly powerJobs.Common” in Assemblies-tab and add it to your project.
[see here](https://doc.coolorange.com/projects/coolorange-powerjobsprocessordocs/en/stable/_images/vs_add_reference.png)
- build the solution

If the build has been successful and you don't have any errors, you will find the new assembly in powerJobs Module folder - [MsOfficeSupport](MsOffice/Source/powerJobs/Modules/MsOfficeSupport/)

More information on how to create custom application for powerJobs can be found on the [powerJobs documentation](https://doc.coolorange.com/projects/coolorange-powerjobsprocessordocs/en/stable/jobprocessor/applications.html#custom-applications).

## At your own risk
The usage of these samples is at your own risk. There is no free support related to the samples. However, if you have questions to powerJobs, then visit http://www.coolorange.com/wiki or start a conversation in our support forum at http://support.coolorange.com/support/discussions

## Author
coolOrange s.r.l.  
Channel Readiness Team

![coolOrange](https://user-images.githubusercontent.com/36075173/46519882-4b518880-c87a-11e8-8dab-dffe826a9630.png)
