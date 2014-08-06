using System;
using System.IO;
using System.Threading;
using Rhino;
using Rhino.Commands;
using Rhino.FileIO;

namespace HOK.DivaSettings
{
    [System.Runtime.InteropServices.Guid("645bc02e-8a5a-42e9-8250-8a3274d413b0")]
    public class DivaDocumentUnits : Command
    {
        static DivaDocumentUnits _instance;
        private string[] optionValues = new string[] { "Inch", "Foot", "Millimeter", "Centimeter", "Meter" };

        public DivaDocumentUnits()
        {
            _instance = this;
        }

        ///<summary>The only instance of the DivaDocumentUnits command.</summary>
        public static DivaDocumentUnits Instance
        {
            get { return _instance; }
        }

        public override string EnglishName
        {
            get { return "DivaDocumentUnits"; }
        }

        protected override Result RunCommand(RhinoDoc doc, RunMode mode)
        {
            // TODO: complete command.

            if (SetModelUnit(doc))
            {
                if (SetPageUnit(doc))
                {
                   
                }
            }

            return Result.Success;
        }

        private bool SetModelUnit(RhinoDoc doc)
        {
            bool result = false;
            try
            {
                Rhino.Input.Custom.GetOption getOption = new Rhino.Input.Custom.GetOption();
                getOption.SetCommandPrompt("Set Model Units");

                foreach (string unit in optionValues)
                {
                    getOption.AddOption(unit);
                }

                getOption.Get();
                if (getOption.CommandResult() != Rhino.Commands.Result.Success)
                {
                    return false;
                }

                string selectedUnit = getOption.Option().EnglishName;
                switch (selectedUnit)
                {
                    case "Inch":
                        doc.AdjustModelUnitSystem(UnitSystem.Inches, true);
                        break;
                    case "Foot":
                        doc.AdjustModelUnitSystem(UnitSystem.Feet, true);
                        break;
                    case "Millimeter":
                        doc.AdjustModelUnitSystem(UnitSystem.Millimeters, true);
                        break;
                    case "Centimeter":
                        doc.AdjustModelUnitSystem(UnitSystem.Centimeters, true);
                        break;
                    case "Meter":
                        doc.AdjustModelUnitSystem(UnitSystem.Meters, true);
                        break;
                }

                doc.Views.Redraw();
                result = true;
            }
            catch (Exception ex)
            {
                string message = ex.Message;
            }
            return result;
        }

        private bool SetPageUnit(RhinoDoc doc)
        {
            bool result = false;
            try
            {
                Rhino.Input.Custom.GetOption getOption = new Rhino.Input.Custom.GetOption();
                getOption.SetCommandPrompt("Set Page Units");

                foreach (string unit in optionValues)
                {
                    getOption.AddOption(unit);
                }

                getOption.Get();
                if (getOption.CommandResult() != Rhino.Commands.Result.Success)
                {
                    return false;
                }

                string selectedUnit = getOption.Option().EnglishName;
                switch (selectedUnit)
                {
                    case "Inch":
                        doc.AdjustPageUnitSystem(UnitSystem.Inches, true);
                        break;
                    case "Foot":
                        doc.AdjustPageUnitSystem(UnitSystem.Feet, true);
                        break;
                    case "Millimeter":
                        doc.AdjustPageUnitSystem(UnitSystem.Millimeters, true);
                        break;
                    case "Centimeter":
                        doc.AdjustPageUnitSystem(UnitSystem.Centimeters, true);
                        break;
                    case "Meter":
                        doc.AdjustPageUnitSystem(UnitSystem.Meters, true);
                        break;
                }

                doc.Views.Redraw();
                result = true;
            }
            catch (Exception ex)
            {
                string message = ex.Message;
            }
            return result;
        }

        private bool ImportDwg(RhinoDoc doc, RunMode mode)
        {
            bool result = false;
            try
            {
                string fileName = string.Empty;
                if (mode == Rhino.Commands.RunMode.Interactive)
                {
                    fileName = Rhino.Input.RhinoGet.GetFileName(Rhino.Input.Custom.GetFileNameMode.Open, null, "Import", Rhino.RhinoApp.MainWindow());
                }
                else
                {
                    Rhino.Input.RhinoGet.GetString("DWG file to import", false, ref fileName);
                }

                fileName = fileName.Trim();
                if (string.IsNullOrEmpty(fileName))
                {
                    return false;
                }

                string rhinoFileName = fileName.Replace(".dwg", ".3dm");
                if (File.Exists(fileName) && Path.GetExtension(fileName).Contains("dwg"))
                {
                    /*
                    FileReadOptions readOptions = new FileReadOptions();
                    readOptions.BatchMode = true;
                    bool opened = Rhino.RhinoDoc.ReadFile(fileName, readOptions);
                    */

                    string script = string.Format("-Import \"{0}\" Enter", fileName);
                    bool runScript = Rhino.RhinoApp.RunScript(script, false);

                    if (runScript)
                    {
                        Rhino.FileIO.FileWriteOptions options = new Rhino.FileIO.FileWriteOptions();
                        options.SuppressDialogBoxes = true;
                        bool saved = RhinoDoc.ActiveDoc.WriteFile(rhinoFileName, options);
                    }
                }
                if (File.Exists(rhinoFileName))
                {
                    result = true;
                }
            }
            catch (Exception ex)
            {
                string message = ex.Message;
            }
            return result;
        }
    }
}
