using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Rhino;
using Rhino.Commands;

namespace HOK.DivaManager
{
    [System.Runtime.InteropServices.Guid("a05a8872-4d49-4270-8851-ce954a8cd457")]
    public class AutoLocation : Command
    {
        private string epwFileName = "";
        private string tempDirectory = @"C:\DIVA\Temp";
        private string resultDirectory = "";
        private string materialFile = @"C:\DIVA\Daylight\material.rad";

        static AutoLocation _instance;
        public AutoLocation()
        {
            _instance = this;
        }

        ///<summary>The only instance of the AutoLocation command.</summary>
        public static AutoLocation Instance
        {
            get { return _instance; }
        }

        public override string EnglishName
        {
            get { return "AutoLocation"; }
        }

        protected override Result RunCommand(RhinoDoc doc, RunMode mode)
        {
            // TODO: complete command.

            bool layerAdded = AddLayer(doc);
            if (layerAdded) { RhinoApp.WriteLine("Layers are successfully added."); }

            bool directoryCreated = CreateDirectories(doc);
            if (directoryCreated) { RhinoApp.Write("Directories are successfully created."); }

            bool materialCopied = CopyMaterialFile();
            if (materialCopied) { RhinoApp.Write("material.rad has been copied."); }

            bool epw2weaRun = SetWeatherFile(doc, @"C:\DIVA\WeatherData\USA_GA_Atlanta-Hartsfield-Jackson.Intl.AP.722190_TMY3.epw");
            return Result.Success;
        }

        private bool AddLayer(RhinoDoc doc)
        {
            bool result = false;
            try
            {
                Dictionary<string, Color> analysisDictionary = new Dictionary<string, Color>();
                analysisDictionary.Add("Analysis Grid", Color.FromArgb(255, 255, 255));
                analysisDictionary.Add("Node Values", Color.FromArgb(190, 190, 190));
                analysisDictionary.Add("Nodes", Color.FromArgb(0, 0, 0));
                analysisDictionary.Add("Radiance Legend", Color.FromArgb(0, 0, 0));
                analysisDictionary.Add("Node Surfaces", Color.FromArgb(255, 175, 0));

                Dictionary<string, Color> thermalDictionary = new Dictionary<string, Color>();
                thermalDictionary.Add("ep_adiabatic", Color.FromArgb(0, 0, 0));
                thermalDictionary.Add("ep_ceiling", Color.FromArgb(144, 144, 144));
                thermalDictionary.Add("ep_floor", Color.FromArgb(254, 254, 254));
                thermalDictionary.Add("ep_shading", Color.FromArgb(25, 25, 25));
                thermalDictionary.Add("ep_wall", Color.FromArgb(200, 200, 200));
                thermalDictionary.Add("ep_window", Color.FromArgb(63, 191, 191));

                int analysisIndex = doc.Layers.Find("DIVA Analysis Surfaces", true);
                if (analysisIndex < 0)
                {
                    analysisIndex=doc.Layers.Add("DIVA Analysis Surfaces", Color.FromArgb(0, 175, 200));
                }

                Rhino.DocObjects.Layer analysisLayer=doc.Layers[analysisIndex];
                Guid parentGuid=analysisLayer.Id;
                foreach(string layerName in analysisDictionary.Keys)
                {
                    if (Rhino.DocObjects.Layer.IsValidName(layerName) && doc.Layers.Find(layerName, true) < 0)
                    {
                        Rhino.DocObjects.Layer childLayer = new Rhino.DocObjects.Layer();
                        childLayer.ParentLayerId = parentGuid;
                        childLayer.Name = layerName;
                        childLayer.Color = analysisDictionary[layerName];
                        int index = doc.Layers.Add(childLayer);
                    }
                }

                int thermalIndex = doc.Layers.Find("DIVA Thermal", true);
                if (thermalIndex < 0)
                {
                    thermalIndex = doc.Layers.Add("DIVA Thermal", Color.FromArgb(0, 0, 0));
                }

                Rhino.DocObjects.Layer thermalLayer = doc.Layers[thermalIndex];
                parentGuid = thermalLayer.Id;
                foreach (string layerName in thermalDictionary.Keys)
                {
                    if (Rhino.DocObjects.Layer.IsValidName(layerName) && doc.Layers.Find(layerName, true) < 0)
                    {
                        Rhino.DocObjects.Layer childLayer = new Rhino.DocObjects.Layer();
                        childLayer.ParentLayerId = parentGuid;
                        childLayer.Name = layerName;
                        childLayer.Color = thermalDictionary[layerName];
                        int index = doc.Layers.Add(childLayer);
                    }
                }
                result = true;
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine("Cannot add layers. " + ex.Message);
                result = false;
            }
            return result;
        }

        private bool CreateDirectories(RhinoDoc doc)
        {
            bool result = false;
            try
            {
                string docName = doc.Name;
                string folderName = docName.Replace(".3dm", "");
                tempDirectory = Path.Combine(tempDirectory, folderName);
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                resultDirectory = Path.Combine(Path.GetDirectoryName(doc.Path), folderName + " - DIVA");
                if (!Directory.Exists(resultDirectory))
                {
                    Directory.CreateDirectory(resultDirectory);
                }

                if (Directory.Exists(resultDirectory))
                {
                    string[] subFolders = { "Grid-Based", "Resources", "Thermal", "Visualizations"};
                    foreach (string folder in subFolders)
                    {
                        string subFolder = Path.Combine(resultDirectory, folder);
                        if (!Directory.Exists(subFolder))
                        {
                            Directory.CreateDirectory(subFolder);
                        }
                    }
                    result=true;
                }
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine("Cannot create directories." + ex.Message);
            }
            return result;
        }

        private bool CopyMaterialFile()
        {
            bool result = false;
            try
            {
                if (File.Exists(materialFile))
                {
                    string destination = Path.Combine(resultDirectory, "Resources");
                    if (Directory.Exists(destination))
                    {
                        destination = Path.Combine(destination, "material.rad");
                        try { File.Copy(materialFile, destination, true); result = true; }
                        catch { result = false; }
                    }
                }
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine("Cannot copy material.rad."+ex.Message);
            }
            return result;
        }

        private bool SetWeatherFile(RhinoDoc doc, string weatherFile)
        {
            bool result = false;
            try
            {
                doc.Strings.SetString("WeatherFile", weatherFile);
                string location = doc.Strings.GetValue("WeatherFile");

                string executable = @"C:\DIVA\DaysimBinaries\epw2wea.exe";
                string[] split = Path.GetFileNameWithoutExtension(weatherFile).Split('.');
                if (split.Length > 2)
                {
                    string weaFile = Path.Combine(Path.GetDirectoryName(weatherFile), split[0] + "." + split[1] + ".wea");
                    string arguments = weatherFile + " " + weaFile;
                    ExecutableManager.RunExcutable(executable, arguments);
                    result = true;
                }
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine(weatherFile + "Cannot set weather file.\n" + ex.Message);
            }
            return result;
        }
    }
}
