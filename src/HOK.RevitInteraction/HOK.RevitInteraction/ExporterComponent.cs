using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
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
                RegistryKeyManager.SetRegistryKeyValue("RhinoOutgoing", runCommand.ToString());
                RegistryKeyManager.SetRegistryKeyValue("RhinoOutgoingPath", filePath);
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
    }
}