using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Rhino.Geometry;

namespace HOK.RevitInteraction
{
    public class ImporterComponent : GH_Component
    {
        private string rvtPath = "";
        private string dwgPath = "";
        private Guid transactionId = new Guid();

        public string RvtPath
        {
            get { return rvtPath; }
            set { rvtPath = value; }
        }

        public string DwgPath
        {
            get { return dwgPath; }
            set { dwgPath = value; }
        }

        public Guid TransactionId
        {
            get { return transactionId; }
            set { transactionId = value; }
        }
        /// <summary>
        /// Initializes a new instance of the ImporterComponent class.
        /// </summary>
        public ImporterComponent()
            : base("ImporterComponent", "Importer",
                "This component will trasfer data from Revit to Rhino",
                "HOK", "Revit Interaction")
        {
            rvtPath = GetRvtPath();
            dwgPath = GetDwgPath();
            transactionId = Guid.NewGuid();
        }

        private string GetDwgPath()
        {
            string path = "";
            //from registry keys
            return path;
        }

        private string GetRvtPath()
        {
            string path = "";
            return path;
        }
        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("Run Button", "Run", "Connect to a button to trigger to import a file from Revit.", GH_ParamAccess.item, false);
            pManager.AddTextParameter("dwgPath", "Path", "A file path that will be imported to the active Rhino session", GH_ParamAccess.item, dwgPath);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Message", "Msg", "Status message of importing files", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
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
                return Properties.Resources.import;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("{2f1d3b0d-9212-48fe-acc6-29761a8be274}"); }
        }
    }
}