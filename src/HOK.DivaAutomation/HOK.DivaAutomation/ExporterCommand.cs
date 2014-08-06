using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using HOK.DivaAutomation.Exporter;

namespace HOK.DivaAutomation
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]
    public class ExporterCommand:IExternalCommand
    {
        private UIApplication m_app;
        private Document m_doc;

        public Result Execute(ExternalCommandData commandData, ref string message, Autodesk.Revit.DB.ElementSet elements)
        {
            m_app = commandData.Application;
            m_doc = m_app.ActiveUIDocument.Document;

            DwgExporterWindow exporterWindow = new DwgExporterWindow(m_app);
            if (exporterWindow.ShowDialog() == true)
            {
                LightingTest selectedTest = exporterWindow.SelectedLightingTest;
                AnalysisBase analysisBase = exporterWindow.SelectedAnalysisBase;
                ElementId viewId = exporterWindow.SelectedViewId;

                exporterWindow.Close();

                List<Reference> selectedReference = new List<Reference>();
                UIDocument uidoc = new UIDocument(m_doc);
                switch (analysisBase)
                {
                    case AnalysisBase.Rooms:
                        ISelectionFilter roomFilter = new RoomSelectionFilter();
                        selectedReference = uidoc.Selection.PickObjects(ObjectType.Element, roomFilter, "Select multiple rooms to create 3d views").ToList();
                        break;
                    case AnalysisBase.Areas:
                        ISelectionFilter areaFilter = new AreaSelectionFilter();
                        selectedReference = uidoc.Selection.PickObjects(ObjectType.Element, areaFilter, "Select multiple areas to create 3d views").ToList();
                        break;
                }
                

                DwgExporter exporter = new DwgExporter(m_app, selectedTest, analysisBase, viewId, selectedReference);
                if (exporter.ExportToDWG())
                {
                    MessageBox.Show("Rhino file has been successfully created!!\n"+exporter.RhinoFileName, "Rhino File Creation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }

            return Result.Succeeded;
        }
    }

    public class RoomSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Autodesk.Revit.DB.Element elem)
        {
            if (null != elem.Category)
            {
                if (elem.Category.Name == "Rooms")
                {
                    return true;
                }
            }
            return false;
        }

        public bool AllowReference(Autodesk.Revit.DB.Reference reference, Autodesk.Revit.DB.XYZ position)
        {
            return true;
        }
    }

    public class AreaSelectionFilter : ISelectionFilter
    {

        public bool AllowElement(Autodesk.Revit.DB.Element elem)
        {
            if (null != elem.Category)
            {
                if (elem.Category.Name == "Areas")
                {
                    return true;
                }
            }
            return false;
        }

        public bool AllowReference(Autodesk.Revit.DB.Reference reference, Autodesk.Revit.DB.XYZ position)
        {
            return true;
        }
    }
}
