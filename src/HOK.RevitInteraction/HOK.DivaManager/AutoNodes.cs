using System;
using System.Collections.Generic;
using Rhino;
using Rhino.Commands;
using Rhino.Geometry;

namespace HOK.DivaManager
{
    [System.Runtime.InteropServices.Guid("bdbfb799-55e9-44b7-9d67-2bac41856817")]
    public class AutoNodes : Command
    {
        static AutoNodes _instance;
        public AutoNodes()
        {
            _instance = this;
        }

        ///<summary>The only instance of the AutoNodes command.</summary>
        public static AutoNodes Instance
        {
            get { return _instance; }
        }

        public override string EnglishName
        {
            get { return "AutoNodes"; }
        }

        protected override Result RunCommand(RhinoDoc doc, RunMode mode)
        {
            // TODO: complete command.
            
            AddNodePoint(doc, 2.5, 3);

           

            return Result.Success;
        }

        private bool AddNodePoint(RhinoDoc doc, double offsetVal, double spacingVal)
        {
            bool result = false;
            try
            {
                int nodeLayerIndex = doc.Layers.Find("Nodes", true);
                if (nodeLayerIndex > -1)
                {
                    Rhino.DocObjects.RhinoObject[] rhobjs = doc.Objects.FindByLayer("A-FLOR");
                    for (int i = 0; i < rhobjs.Length; i++)
                    {
                        if (rhobjs[i] is Rhino.DocObjects.BrepObject)
                        {
                            Rhino.DocObjects.BrepObject brepObj = rhobjs[i] as Rhino.DocObjects.BrepObject;
                            Brep brep = brepObj.BrepGeometry;

                            foreach (BrepFace face in brep.Faces)
                            {
                                if (face.NormalAt(0, 0).Z == 1)
                                {
                                    Surface offsetSurface = face.Offset(offsetVal, doc.ModelAbsoluteTolerance);
                                    BoundingBox bb = offsetSurface.GetBoundingBox(true);
                                    int colNum = (int)Math.Floor((bb.Max.X - bb.Min.X) / spacingVal);
                                    int rowNum = (int)Math.Floor((bb.Max.Y - bb.Min.Y) / spacingVal);

                                    double remainderX = ((bb.Max.X - bb.Min.X) - (colNum * spacingVal)) / 2;
                                    double remainderY = ((bb.Max.Y - bb.Min.Y) - (rowNum * spacingVal)) / 2;

                                    int nodeNum = 0;
                                    List<Guid> ids = new List<Guid>();
                                    for (int col = 0; col < colNum; col++)
                                    {
                                        double xVal = bb.Min.X + remainderX + spacingVal*(col + 0.5);

                                        for (int row = 0; row < rowNum; row++)
                                        {
                                            double yVal = bb.Min.Y + remainderY + spacingVal * (row + 0.5);
                                            
                                            Point3d point = new Point3d(xVal, yVal, offsetVal);
                                            Rhino.DocObjects.ObjectAttributes attributes = new Rhino.DocObjects.ObjectAttributes();
                                            attributes.Name = "Node_" + nodeNum;
                                            attributes.LayerIndex = nodeLayerIndex;

                                            Guid guid = doc.Objects.AddPoint(point, attributes);
                                            ids.Add(guid);
                                            nodeNum++;
                                        }
                                    }
                                    int index = doc.Groups.Add(ids);
                                    break;
                                }
                            }
                            
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine("Cannot add node points. " + ex.Message);
            }
            return result;
        }

        private void GetNumberOfRowsColumns(Surface surface, double spacing, out int rowNum, out int colNum)
        {
            rowNum = 0;
            colNum = 0;
            try
            {
                BoundingBox bb = surface.GetBoundingBox(true);
                colNum = (int) Math.Floor((bb.Max.X - bb.Min.X) / spacing);
                rowNum = (int)Math.Floor((bb.Max.Y - bb.Min.Y) / spacing);
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine("Cannot get the number of rows and columns." + ex.Message);
            }
        }
        
    }
}
