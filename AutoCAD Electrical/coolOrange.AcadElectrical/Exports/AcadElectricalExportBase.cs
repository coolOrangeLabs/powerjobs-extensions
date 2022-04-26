using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading;
using coolOrange.AutoCADElectrical.Helpers;
using IniParser;
using IniParser.Model;
using log4net;
using powerJobs.Common;
using powerJobs.Common.Applications;

namespace coolOrange.AutoCADElectrical.Exports
{
    public class AcadElectricalExportBase : DocumentExportBase
    {
        static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        protected dynamic AcadActiveDocument { get; set; }

        public AcadElectricalExportBase(IDocument sourceDocument, ExportSettings settings) : base(sourceDocument, settings)
        {
        }

        public override void Execute()
        {
            bool finished = false;
            Proxy.Instance.Collection.Add(() =>
            {
                try
                {
                    // Check if config file is specified and exists
                    if (!Settings.Options.ContainsKey("ConfigFile"))
                        throw new ApplicationException("Option 'ConfigFile' is not specified!");

                    var configFile = Settings.Options["ConfigFile"] as IFileInfo;
                    if (configFile == null || !configFile.Exists)
                        throw new ApplicationException($"Configuration file '{configFile}' does not exist!");

                    // activate AutoCAD electrical project file
                    dynamic acad = ((Application) SourceDocument.Application).AcadApplication;
                    AcadActiveDocument = acad.ActiveDocument;
                    ActivateProjectFile(SourceDocument.OpenSettings.File);

                    // creating DSD file
                    var dsdFilename = CreateDsdFile(SourceDocument.OpenSettings.File, configFile);

                    // publishing PDF file
                    PublishFile(dsdFilename);
                }
                catch (Exception ex)
                {
                    Log.Error($"Failed to export : {ex.Message}", ex);
                    throw;
                }
                finally
                {
                    finished = true;
                }
            });
            Proxy.Instance.Run();
            while (!finished)
            {
                Thread.Sleep(1000);
            }
        }

        public override string Name { get; }

        public override HashSet<string> SupportedDocumentTypes { get; }

        protected void ActivateProjectFile(IFileInfo wdpFile)
        {
            try
            {
                Log.Info($"Activating AutoCAD Electrical project file '{wdpFile.FullName}' ...");
                var wdpFilename = wdpFile.FullName.Replace('\\', '/');
                AcadDocHelper.SendCommandWait(AcadActiveDocument, $"(c:wd_makeproj_current \"{wdpFilename}\"){Environment.NewLine}");
                Log.Info($"Successfully activated AutoCAD Electrical project file!");
            }
            catch (Exception ex)
            {
                var msg = $"ActivateProjectFile(): Error activating AutoCAD Electrical project file '{wdpFile.FullName}': {ex.Message}";
                Log.Error(msg, ex);
                throw;
            }
        }

        protected void PublishFile(string dsdFilename)
        {
            try
            {
                // prepare for publish command
                Log.Debug("Setting variable BACKGROUNDPLOT = 0 ...");
                AcadDocHelper.SendCommandWait(AcadActiveDocument, "_backgroundplot 0 ");

                Log.Debug("Setting variable FILEDIA = 0 ...");
                AcadDocHelper.SendCommandWait(AcadActiveDocument, "_filedia 0 ");

                // call publish command
                Log.Info("Starting publish command ...");

                AcadDocHelper.SendCommandWait(AcadActiveDocument, $"_-publish {dsdFilename}{Environment.NewLine}");
                Log.Info($"Successfully created {Name} file!");
            }
            catch (Exception ex)
            {
                var msg = $"PublishFile(): Error publishing {Name} file: {ex.Message}";
                Log.Error(msg, ex);
                throw;
            }
        }

        protected string CreateDsdFile(IFileInfo wdpFile, IFileInfo dsdTemplateFile)
        {
            try
            {
                Log.Info($"Creating DSD file for publishing to {Name} ...");
                var dwgFiles = CollectDwgFilesFromProject(wdpFile);

                // creating DSD file
                var parser = new FileIniDataParser();
                var dsdTemplateIni = parser.ReadFile(dsdTemplateFile.FullName);
                var dwgTemplateSection = dsdTemplateIni.Sections.GetSectionData("DWF6Sheet:DwgTemplate");

                var dsdFileData = new IniData();
                foreach (var section in dsdTemplateIni.Sections)
                {
                    if (section.SectionName == "DWF6Sheet:DwgTemplate")
                    {
                        foreach (var dwgFile in dwgFiles)
                        {
                            var dwgFilename = dwgFile.Value;
                            var dwgSection = new SectionData($"DWF6Sheet:{dwgFile.Key}".Replace('/', '_'));
                            foreach (var key in dwgTemplateSection.Keys)
                                dwgSection.Keys.AddKey(key.KeyName, key.KeyName == "DWG" ? dwgFilename : key.Value);
                            dsdFileData.Sections.Add(dwgSection);
                        }
                    }
                    else
                    {
                        var clonedSection = section.Clone() as SectionData;
                        if (clonedSection != null)
                            dsdFileData.Sections.Add(clonedSection);
                    }
                }

                // set target path and required AcePublish settings
                var aceTempPath = Path.Combine(Path.GetTempPath(), $"AcePub_{Path.GetFileNameWithoutExtension(wdpFile.Name)}");
                var mdbName = Path.GetFileNameWithoutExtension(wdpFile.Name) + ".mdb";
                var originMdbDir = Properties.Settings.Default.OverrideAcadElectricalUserPath;
                if (string.IsNullOrEmpty(originMdbDir))
                {
                    var printerConfigPath = new System.IO.DirectoryInfo(AcadAppHelper.GetPrinterConfigPath(((Application)SourceDocument.Application).AcadApplication));
                    originMdbDir = Path.Combine(printerConfigPath.Parent.FullName, "support", "user");
                }

                string targetType;
                switch (Name)
                {
                    case "DWFx": targetType = "4"; break;
                    case "PDF": targetType = "6"; break;
                    case "DWF": targetType = "1"; break;
                    default: targetType = "1"; break;
                }

                dsdFileData["Target"]["Type"] = targetType;
                dsdFileData["Target"]["DWF"] = Settings.DestinationFile.FullName;
                dsdFileData["AcePublish"]["TempFolder"] = aceTempPath + "\\";
                dsdFileData["AcePublish"]["OriginWdpPath"] = wdpFile.FullName;
                dsdFileData["AcePublish"]["OriginMdbPath"] = Path.Combine(originMdbDir, mdbName);
                dsdFileData["AcePublish"]["TempWdpPath"] = Path.Combine(aceTempPath, wdpFile.Name);
                dsdFileData["AcePublish"]["TempMdbPath"] = Path.Combine(aceTempPath, mdbName);

                var dsdFilename = Path.Combine(wdpFile.Directory.FullName, Path.GetFileNameWithoutExtension(wdpFile.Name) + ".dsd");
                parser.WriteFile(dsdFilename, dsdFileData, Encoding.UTF8);
                Log.Info("Successfully created ");
                return dsdFilename;
            }
            catch (Exception ex)
            {
                var msg = $"CreateDsdFile(): Error creating DSD file: {ex.Message}";
                Log.Error(msg, ex);
                throw;
            }
        }

        private Dictionary<string, string> CollectDwgFilesFromProject(IFileInfo wdpFile)
        {
            Log.Info("Parsing WDP file to collect DWG files ...");
            var dwgFiles = new Dictionary<string, string>();

            try
            {
                using (var wdpFileContent = new StringReader(wdpFile.ReadAllText()))
                {
                    string line;
                    string dwgDescription = null;
                    while (null != (line = wdpFileContent.ReadLine()))
                    {
                        // Ignore lines that start with: '*', '+', '?' and five '='
                        // Sample:
                        // *[2]SAMPLE
                        // +[24]U:\DATA\acad setup\default_cat.mdb
                        // ?[72]W1,T1,TYPE,T2,W2,REF,SH,SHDWGNAM,FILENAME,FULLFILENAME
                        // =====SUB=SCHEMATIC[Empty]
                        //
                        // Lines starting with exact three '=' characters contain the DWG description for the following DWG file.
                        // Note:
                        //   - There could be more than one description. Only the first one is used.
                        //   - It could be that there is no description
                        // Sample:
                        // === 3.15 / 4.4MVA PACKAGE SUBSTATION
                        // === c / w 6300 AMPERE LV SWITCHBOARD
                        // === REF: LL LV SWB01
                        // 28871 - E01 - 1revB.dwg
                        //
                        if (line.StartsWith("*") || line.StartsWith("+") || line.StartsWith("?") || line.StartsWith("====="))
                            continue;

                        if (string.IsNullOrEmpty(dwgDescription) && line.StartsWith("==="))
                            dwgDescription = line.Substring(3);

                        if (line.EndsWith(".dwg"))
                        {
                            var dwgFullFilename = Path.Combine(wdpFile.Directory.FullName, line);
                            var key = string.IsNullOrEmpty(dwgDescription)
                                ? Path.GetFileNameWithoutExtension(dwgFullFilename)
                                : $"{Path.GetFileNameWithoutExtension(dwgFullFilename)}-{dwgDescription}";
                            dwgFiles.Add(key, dwgFullFilename);
                            dwgDescription = null;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException("Error parsing WDP file to collect DWG files!", ex);
            }
            Log.Info("Successfully collected DWG files ...");
            return dwgFiles;
        }
    }
}
