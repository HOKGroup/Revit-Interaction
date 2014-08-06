using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace HOK.DivaAutomation.Exporter
{
    public class DwgExporter
    {
        public UIApplication m_app;
        public Document m_doc;
        private dynamic rhino = null;
        private string filePath = "";
        private ExportUnit exportUnit = ExportUnit.Default;

        private LightingTest selectedTest = LightingTest.Illuminance;
        private AnalysisBase selectedBase = AnalysisBase.None;
        private ElementId selectedView = ElementId.InvalidElementId;
        private List<Reference> selectedReference = new List<Reference>();

        private Dictionary<string, string> prefixes = new Dictionary<string, string>();
        private bool fileCreated = false;
        private string rhinoFileName = "";

        public bool FileCreated { get { return fileCreated; } set { fileCreated = value; } }
        public string RhinoFileName { get { return rhinoFileName; } set { rhinoFileName = value; } }

        public DwgExporter(UIApplication uiapp, LightingTest lightingTest, AnalysisBase analysisBase, ElementId viewId, List<Reference> references)
        {
            try
            {
                m_app = uiapp;
                m_doc = m_app.ActiveUIDocument.Document;

                selectedTest = lightingTest;
                selectedBase = analysisBase;
                selectedView = viewId;
                selectedReference = references;

                filePath = GetFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    MessageBox.Show("Please save the current Revit project before exporting.", "File Not Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    prefixes.Add("E-CAD", "E-CAD");
                    prefixes.Add("E-BIM", "E-BIM");
                    prefixes.Add("REVIT", "REVIT");

                    switch (selectedBase)
                    {
                        case AnalysisBase.Rooms:
                            CreateViewsFromRooms();
                            break;
                        case AnalysisBase.Areas:
                            CreateViewsFromAreas();
                            break;
                    }
                   
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to export to dwg.\n"+ex.Message, "DWG Export", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void CreateViewsFromRooms()
        {
            try
            {
               
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to create views from rooms.\n" + ex.Message, "Create Views From Rooms", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void CreateViewsFromAreas()
        {
            try
            {
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to create views from areas.\n"+ex.Message, "Create Views From Areas", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private ExportUnit GetExportUnit()
        {
            ExportUnit exportUnit = ExportUnit.Default;
            try
            {
                Units units = m_doc.GetUnits();
                FormatOptions formatOptions = units.GetFormatOptions(UnitType.UT_Length);
                DisplayUnitType dut = formatOptions.DisplayUnits;

                switch (dut)
                {
                    case DisplayUnitType.DUT_METERS:
                        exportUnit = ExportUnit.Meter;
                        break;
                    case DisplayUnitType.DUT_CENTIMETERS:
                        exportUnit = ExportUnit.Centimeter;
                        break;
                    case DisplayUnitType.DUT_DECIMAL_FEET:
                        exportUnit = ExportUnit.Foot;
                        break;
                    case DisplayUnitType.DUT_DECIMAL_INCHES:
                        exportUnit = ExportUnit.Inch;
                        break;
                    case DisplayUnitType.DUT_FEET_FRACTIONAL_INCHES:
                        exportUnit = ExportUnit.Foot;
                        break;
                    case DisplayUnitType.DUT_FRACTIONAL_INCHES:
                        exportUnit = ExportUnit.Inch;
                        break;
                    case DisplayUnitType.DUT_METERS_CENTIMETERS:
                        exportUnit = ExportUnit.Meter;
                        break;
                    case DisplayUnitType.DUT_MILLIMETERS:
                        exportUnit = ExportUnit.Millimeter;
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to get export units.\n"+ex.Message, "Get Export Unit", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            return exportUnit;
        }

        public bool ExportToDWG()
        {
            bool result = false;
            try
            {
                string inputFolder = GetInputDirectory();
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                bool exported = false;

                if (Directory.Exists(inputFolder))
                {
                    DWGExportOptions options = new DWGExportOptions();
                    options.ExportOfSolids = SolidGeometry.ACIS;
                    exportUnit = GetExportUnit();
                    options.TargetUnit = exportUnit;

                    ICollection<ElementId> views = new List<ElementId>();
                    views.Add(selectedView);

                    Guid transId = new Guid();
                    using (Transaction trans = new Transaction(m_doc))
                    {
                        try
                        {
                            transId = Guid.NewGuid();
                            trans.Start(transId.ToString());
                            exported = m_doc.Export(inputFolder, fileName, views, options);
                            trans.Commit();
                        }
                        catch (Exception ex)
                        {
                            trans.RollBack();
                            MessageBox.Show("DWG Export was failed.\n" + ex.Message, "DWG Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }

                string dwgFileName = Path.Combine(inputFolder, fileName + ".dwg");
                if (exported && File.Exists(dwgFileName))
                {
                    string unitString = "F";
                    switch (exportUnit)
                    {
                        case ExportUnit.Default:
                            unitString = "F";
                            break;
                        case ExportUnit.Inch:
                            unitString = "I";
                            break;
                        case ExportUnit.Foot:
                            unitString = "F";
                            break;
                        case ExportUnit.Millimeter:
                            unitString = "M";
                            break;
                        case ExportUnit.Centimeter:
                            unitString = "C";
                            break;
                        case ExportUnit.Meter:
                            unitString = "e";
                            break;
                    }

                    if (InitializeRhino())
                    {
                        string script = string.Format("-DivaDocumentUnits {0} {1}", unitString, unitString); //document units to foot
                        rhino.RunScript(script, false);

                        script = string.Format("-Import \"{0}\" Enter", dwgFileName); //model and layout unit to foot
                        rhino.RunScript(script, false);

                        rhinoFileName = dwgFileName.Replace(".dwg", ".3dm");
                        if (File.Exists(rhinoFileName))
                        {
                            File.Delete(rhinoFileName);
                        }

                        script = string.Format("-SaveAs \"{0}\" Enter", rhinoFileName);
                        rhino.RunScript(script, false);

                        if (File.Exists(rhinoFileName)) { result = true; }

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to export to dwg.\n"+ex.Message, "Export to DWG", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            return result;
        }

        private bool InitializeRhino()
        {
            bool initialized = false;
            try
            {
                string rhinoId = "Rhino5x64.Application";
                System.Type type = System.Type.GetTypeFromProgID(rhinoId);
                rhino = System.Activator.CreateInstance(type);
                if (null != rhino)
                {
                    rhino.Visible = 0;//non-visible 

                    const int bail_milliseconds = 15 * 1000;
                    int time_waiting = 0;
                    while (0 == rhino.IsInitialized())
                    {
                        Thread.Sleep(100);
                        time_waiting += 100;
                        if (time_waiting > bail_milliseconds)
                        {
                            MessageBox.Show("Rhino initialization failed");
                            return false;
                        }
                    }
                    initialized = true;
                }
            }
            catch
            {
                initialized = false;
            }
            return initialized;
        }

        private string GetInputDirectory()
        {
            string inputFolder = "";
            try
            {
                string fileLocation = GetProjectLocation(filePath);
                string projectNumber = "";
                string projectName = "";
                GetProjectInfo(filePath, out projectNumber, out projectName);

                string rootFolder=@"\\group\hok\"+fileLocation+@"\RESOURCES\BIM\DivaAutomation";

                if (Directory.Exists(rootFolder))
                {
                    string projectFolder = projectNumber + " - " + projectName;
                    projectFolder = Path.Combine(rootFolder, projectFolder);

                    if (!Directory.Exists(projectFolder))
                    {
                        Directory.CreateDirectory(projectFolder);
                    }

                    string testFolder = GetTestName();
                    testFolder = Path.Combine(projectFolder, testFolder);

                    if (!Directory.Exists(testFolder))
                    {
                        Directory.CreateDirectory(testFolder);
                    }

                    inputFolder = Path.Combine(testFolder, "Input");
                    if (!Directory.Exists(inputFolder))
                    {
                        Directory.CreateDirectory(inputFolder);
                    }
                }
                else
                {
                    MessageBox.Show("Please verify the file is located in the HOK folder structure.\n", "Directory Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to get file location.\n"+ex.Message, "Get File Location", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            return inputFolder;
        }

        private string GetFilePath()
        {
            string filePath = "";
            try
            {
                if (m_doc.IsWorkshared)
                {
                    ModelPath modelPath = m_doc.GetWorksharingCentralModelPath();
                    string centralPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath);
                    if (!string.IsNullOrEmpty(centralPath))
                    {
                        filePath = centralPath;
                    }
                    else
                    {
                        filePath = m_doc.PathName;
                    }
                }
                else
                {
                    filePath = m_doc.PathName;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to get file path.\n"+ex.Message, "Get File Path", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            return filePath;
        }

        private string GetTestName()
        {
            string testName = "";
            try
            {
                switch (selectedTest)
                {
                    case LightingTest.Illuminance:
                        testName = "Illuminance Values";
                        break;
                    case LightingTest.LEED:
                        testName = "LEED IEQ 8.1";
                        break;
                    case LightingTest.NECHPS:
                        testName = "NECHPS IEQ P2";
                        break;
                    case LightingTest.MACHPS:
                        testName = "MACHPS IEQ C2";
                        break;
                    case LightingTest.DaylightFactor:
                        testName = "Daylight Factor";
                        break;
                    case LightingTest.RadiationMap:
                        testName = "Radiation Map";
                        break;
                    case LightingTest.DaylightAutonomy:
                        testName = "Daylight Autonomy";
                        break;
                    case LightingTest.ContinuousDaylightAutonomy:
                        testName = "Continuous Daylight Autonomy";
                        break;
                    case LightingTest.DaylightAvailability:
                        testName = "Daylight Availability";
                        break;
                    case LightingTest.UDI:
                        testName = "Useful Daylight Illuminance (UDI)";
                        break;
                    case LightingTest.IlluminanceFromDA:
                        testName = "Illuminance from DA";
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to get lighting test name.\n"+ex.Message, "Get Test Name", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            return testName;
        }

        private string GetProjectLocation(string path)
        {
            string projectLocation = "";
            try
            {
                if (!string.IsNullOrEmpty(path))
                {
                    Regex regServer = new Regex(@"^\\\\group\\hok\\(.+?(?=\\))|^\\\\(.{2,3})-\d{2}svr(\.group\.hok\.com)?\\", RegexOptions.IgnoreCase);
                    Match regMatch = regServer.Match(path);
                    if (regMatch.Success)
                    {
                        if (string.IsNullOrEmpty(regMatch.Groups[1].Value))
                        {
                            projectLocation = regMatch.Groups[2].Value;
                        }
                        else
                        {
                            projectLocation = regMatch.Groups[1].Value;
                        }
                    }
                    if (string.IsNullOrEmpty(projectLocation))
                    {
                        try
                        {
                            ActiveDs.ADSystemInfo systemInfo = new ActiveDs.ADSystemInfo();
                            string siteName = systemInfo.SiteName;

                            if (!string.IsNullOrEmpty(siteName))
                            {
                                projectLocation = siteName;
                            }
                        }
                        catch
                        {
                            projectLocation = "";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to get project location.\n" + ex.Message, "Get Project Location", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            return projectLocation;
        }

        private void GetProjectInfo(string path, out string projectNumber, out string projectName)
        {
            projectNumber = "";
            projectName = "";
            try
            {
                if (!string.IsNullOrEmpty(path))
                {
                    try
                    {
                        string regPattern = @"\\([0-9]{2}[\.|\-][0-9]{4,5}[\.|\-][0-9]{2})(.*?)\\";
                        Regex regex = new Regex(regPattern, RegexOptions.IgnoreCase);
                        Match match = regex.Match(path);
                        if (match.Success)
                        {
                            projectNumber = match.Groups[1].Value;
                            projectName = match.Groups[2].Value;
                        }
                    }
                    catch { }

                    if (string.IsNullOrEmpty(projectNumber))
                    {
                        projectName = GetProjectName(path);
                        projectNumber = "00.00000.00";
                    }
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to get project info.\n" + ex.Message, "Get Project Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private string GetProjectName(string path)
        {
            string name = "";
            try
            {
                string[] paths = path.Split('\\');

                //Find E-BIM or E-CAD and get preceding values
                for (int i = 0; i < paths.Length; i++)
                {
                    if (prefixes.ContainsKey(paths[i]))
                    {
                        return paths[i - 1];
                    }
                }
            }
            catch
            {
                return "";
            }
            return name;
        }
    }

    
}
