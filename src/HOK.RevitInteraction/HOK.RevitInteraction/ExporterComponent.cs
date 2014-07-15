using System;
using System.Collections.Generic;
using System.IO;
using Grasshopper.Kernel;
using Rhino;
using Rhino.Geometry;

namespace HOK.RevitInteraction
{
    public class ExporterComponent : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the ExporterComponent class.
        /// </summary>
        public ExporterComponent()
            : base("ExporterComponent", "Exporter",
                "This component will transfer data from Rhino to Revit",
                "HOK", "Revit Interaction")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("RunButton", "Run", "Connect to a button to trigger to transfer the result data to Revit", GH_ParamAccess.item, false);
            pManager.AddTextParameter("AnalysisType", "AType", "Set an anlaysls type from the list", GH_ParamAccess.item, "Daylight");
            pManager.AddTextParameter("FilePath", "Path", "Set the file path of the data result.", GH_ParamAccess.item, "");
            pManager[1].Optional = true;
            pManager[2].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Selected Revit", "Rvt", "Revit project information that will communicate with this component", GH_ParamAccess.item);
            pManager.AddTextParameter("Message", "Msg", "Rsult messages from the process", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool runCommand = false;
            string analysisType = "";
            string filePath = "";
            string revitDocName = "";
            string message = "";

            if (!DA.GetData(0, ref runCommand)) { return; }
            if (!DA.GetData(1, ref analysisType)) { return; }
            if (!DA.GetData(2, ref filePath)) { return; }

            if (runCommand)
            {
                List<Mesh> analysisGrid = FindAnalysisGrid();
                string docName = RhinoDoc.ActiveDoc.Name;
                string dirName = docName.Replace(".3dm", "");

                string divaTemp=@"C:\DIVA\Temp";
                string tempDirectory = Path.Combine(divaTemp, dirName);
                string resultDirectory = Path.Combine(Path.GetDirectoryName(RhinoDoc.ActiveDoc.Path), dirName + " - DIVA");

                string objFileName = Path.Combine(tempDirectory, dirName + "-AnalysisGrid.obj");
                string script = string.Format("-Export \"{0}\" Enter Enter Enter", objFileName);
                RhinoApp.RunScript(script, false);

                RegistryKeyManager.SetRegistryKeyValue("RhinoOutgoing", runCommand.ToString());
                RegistryKeyManager.SetRegistryKeyValue("RhinoOutgoingPath", filePath);
                RegistryKeyManager.SetRegistryKeyValue("DivaTempDirectory", tempDirectory);
                RegistryKeyManager.SetRegistryKeyValue("DivaResultDirectory", resultDirectory);
                RegistryKeyManager.SetRegistryKeyValue("RhinoOutgoingId", Guid.NewGuid().ToString());

                message = "Rhino is sending result data to Revit";

                revitDocName = RegistryKeyManager.GetRegistryKeyValue("RevitDocName");
                if (!string.IsNullOrEmpty(revitDocName))
                {
                    DA.SetData(0, revitDocName);
                    DA.SetData(1, message);
                }
            }
        }

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                //You can add image files to your project resources and access them like this:
                // return Resources.IconForThisComponent;
                return Properties.Resources.export;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("{27cda3da-4f66-4733-b5bc-1b98d95ba895}"); }
        }

        private List<Mesh> FindAnalysisGrid()
        {
            List<Mesh> anlysisGrid = new List<Mesh>();
            try
            {
                string layerName = "Analysis Grid";
                Rhino.RhinoDoc activeDoc = Rhino.RhinoDoc.ActiveDoc;
                activeDoc.Objects.UnselectAll();

                List<Guid> objIds = new List<Guid>();
                if (null != activeDoc)
                {
                    int layerIndex = activeDoc.Layers.Find(layerName, true);
                    if (layerIndex < 0)
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, layerName+" cannot be found.");
                        return null;
                    }
                    Rhino.DocObjects.Layer layer = activeDoc.Layers[layerIndex];
                    Rhino.DocObjects.RhinoObject[] objs = activeDoc.Objects.FindByLayer(layer);
                    if (objs != null)
                    {
                        foreach (Rhino.DocObjects.RhinoObject obj in objs)
                        {
                            if (obj.Geometry.ObjectType == Rhino.DocObjects.ObjectType.Mesh)
                            {
                                Mesh mesh = obj.Geometry as Mesh;
                                if (null != mesh)
                                {
                                    anlysisGrid.Add(mesh);
                                    Guid guid = obj.Id;
                                    objIds.Add(guid);
                                }
                            }
                        }
                        activeDoc.Objects.Select(objIds);
                    }
                }
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Analysis Grid cannot be exported.\n" + ex.Message);
            }
            return anlysisGrid;
        }
    }
}