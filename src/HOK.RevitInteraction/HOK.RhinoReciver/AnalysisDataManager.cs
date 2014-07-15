using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Analysis;
using Autodesk.Revit.UI;

namespace HOK.RhinoReciver
{
    public enum DataType
    {
        None,
        Daylight_Factor,//dat file
        Radiation_Map,//dat file
        Daylight_Autonomy//da file
    }

    public class AnalysisDataManager
    {
        private UIApplication m_app;
        private Document m_doc;
        private List<ObjMesh> objMeshes = new List<ObjMesh>();
        private string dataFilePath = "";
        private DataType analysisType = DataType.None;
        private List<double> values = new List<double>();
        private Dictionary<int/*index*/, DataContainer> dataDictionary = new Dictionary<int, DataContainer>();
        private List<Face> displayingFaces = new List<Face>();

        public Dictionary<int, DataContainer> DataDictionary { get { return dataDictionary; } set { dataDictionary = value; } }

        public AnalysisDataManager(UIApplication uiapp, List<ObjMesh> meshes, string filePath)
        {
            m_app = uiapp;
            m_doc = m_app.ActiveUIDocument.Document;
            objMeshes = meshes;
            dataFilePath = filePath;
        }

        public bool ReadResults()
        {
            bool result = false;
            try
            {
                if (File.Exists(dataFilePath))
                {
                    string extension = Path.GetExtension(dataFilePath);
                    if (extension == ".dat")
                    {
                        if (dataFilePath.EndsWith("df.dat"))
                        {
                            analysisType = DataType.Daylight_Factor;
                        }
                        else
                        {
                            analysisType = DataType.Radiation_Map;
                        }
                    }
                    else if (extension == ".DA")
                    {
                        analysisType = DataType.Daylight_Autonomy;
                    }

                    int index = 0;
                    using (StreamReader reader = new StreamReader(dataFilePath))
                    {
                        string line = string.Empty;
                        char[] delimiter = { ' ', '\t'};

                        while ((line = reader.ReadLine()) != null)
                        {
                            if (line.StartsWith("#")) { continue; }
                            string[] splitVals = line.Split(delimiter, StringSplitOptions.RemoveEmptyEntries);
                            DataContainer container = new DataContainer();
                            container.Index = index;

                            if (analysisType == DataType.Daylight_Factor || analysisType == DataType.Radiation_Map) //dat file
                            {
                                NodePoint node = new NodePoint(double.Parse(splitVals[0]), double.Parse(splitVals[1]), double.Parse(splitVals[2]));
                                double rVal = double.Parse(splitVals[6]);
                                double gVal = double.Parse(splitVals[7]);
                                double bVal = double.Parse(splitVals[8]);
                                double calVal = (rVal * 0.265 + gVal * 0.67 + bVal * 0.065) * 179;

                                container.Node = node;
                                container.ResultValue = calVal;
                            }
                            else if (analysisType == DataType.Daylight_Autonomy) //da file
                            {
                                NodePoint node = new NodePoint(double.Parse(splitVals[0]), double.Parse(splitVals[1]), double.Parse(splitVals[2]));
                                double daVal = double.Parse(splitVals[3]);

                                container.Node = node;
                                container.ResultValue = daVal;
                            }

                            if (dataDictionary.ContainsKey(index)) { dataDictionary.Remove(index); }
                            dataDictionary.Add(index, container);
                            index++;
                        }
                        reader.Close();
                    }
                    result = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to read result data.\n" + ex.Message, "Analysis Data Manager - Read Results", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            return result;
        }

        public bool CreateGeometry()
        {
            bool result = false;
            try
            {
                for (int i = 0; i < objMeshes.Count; i++)
                {
                    ObjMesh mesh = objMeshes[i];
                    
                    List<Curve> curveList = new List<Curve>();

                    ObjVertice startV = mesh.ObjVertex[mesh.ObjVertex.Count - 1];
                    ObjVertice endV = mesh.ObjVertex[0];

                    XYZ startPoint = new XYZ(startV.XValue, startV.YValue, startV.ZValue);
                    XYZ endPoint = new XYZ(endV.XValue, endV.YValue, endV.ZValue);

                    Line line = Line.CreateBound(startPoint, endPoint);
                    curveList.Add(line);

                    for (int j = 0; j < mesh.ObjVertex.Count-1; j++)
                    {
                        startV = mesh.ObjVertex[j];
                        endV = mesh.ObjVertex[j+1];

                        startPoint = new XYZ(startV.XValue, startV.YValue, startV.ZValue);
                        endPoint = new XYZ(endV.XValue, endV.YValue, endV.ZValue);

                        line = Line.CreateBound(startPoint, endPoint);
                        curveList.Add(line);
                    }
                    CurveLoop curveLoop = CurveLoop.Create(curveList);
                    List<CurveLoop> profile = new List<CurveLoop>();
                    profile.Add(curveLoop);
                    Solid extrusion = GeometryCreationUtilities.CreateExtrusionGeometry(profile, new XYZ(0, 0, 1), 1);

                    if (null != extrusion)
                    {
                        foreach (Face face in extrusion.Faces)
                        {
                            XYZ normal = face.ComputeNormal(new UV(0,0));
                            if (normal.Z > 0)
                            {
                                displayingFaces.Add(face); break;
                            }
                        }
                    }

                }
                result = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to create geometry for surfaces to be visulized with data.\n" + ex.Message, "Analysis Data Manager - Create Geometry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                result = false;
            }
            return result;
        }

        public bool VisualizeData()
        {
            bool result = false;
            try
            {
                string viewIdValue = RegistryKeyManager.GetRegistryKeyValue("RevitOutgoingViewId");
                ElementId viewId = new ElementId(int.Parse(viewIdValue));
                Autodesk.Revit.DB.View view = m_doc.GetElement(viewId) as Autodesk.Revit.DB.View;
                if (null != view)
                {
                    SpatialFieldManager sfm = SpatialFieldManager.GetSpatialFieldManager(view);
                    if (null == sfm)
                    {
                        sfm = SpatialFieldManager.CreateSpatialFieldManager(view, 1);
                    }
                    AnalysisResultSchema resultSchema = new AnalysisResultSchema(analysisType.ToString(), "Imported from DIVA");

                    int resultIndex = -1;

                    //check order
                    DataContainer tempContainer=dataDictionary[0];
                    XYZ node = new XYZ(tempContainer.Node.XValue, tempContainer.Node.YValue, tempContainer.Node.ZValue);
                    IntersectionResult intersection = displayingFaces[0].Project(node);
                    if (null == intersection) { displayingFaces.Reverse(); } //reverse the order of faces

                    foreach(int keyIndex in dataDictionary.Keys)
                    {
                        DataContainer container = dataDictionary[keyIndex];
                        Face face = displayingFaces[keyIndex];
                        List<double> dblList = new List<double>();
                        dblList.Add(container.ResultValue);

                        XYZ vectorZ = new XYZ(0, 0, -1);
                        Transform transform = Transform.CreateTranslation(vectorZ);
                        int index = sfm.AddSpatialFieldPrimitive(face, transform);
                        IList<UV> uvPts = new List<UV>();
                        IList<ValueAtPoint> valList = new List<ValueAtPoint>();

                        XYZ nodePoint = new XYZ(container.Node.XValue, container.Node.YValue, container.Node.ZValue);
                        IntersectionResult intersect = face.Project(nodePoint);
                        if (null != intersect)
                        {
                            UV nodeUV = intersect.UVPoint;
                            uvPts.Add(nodeUV);
                            valList.Add(new ValueAtPoint(dblList));

                            FieldDomainPointsByUV domainPoints = new FieldDomainPointsByUV(uvPts);
                            FieldValues values = new FieldValues(valList);

                            FieldValues vals = new FieldValues(valList);

                            if (resultIndex == -1) { resultIndex = sfm.RegisterResult(resultSchema); }
                            else { sfm.SetResultSchema(resultIndex, resultSchema); }

                            sfm.UpdateSpatialFieldPrimitive(index, domainPoints, values, resultIndex);
                        }
                        
                        result = true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to visualize the result data.\n"+ex.Message, "Analysis Data Manager - Visualize Data", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                result = false;
            }
            return result;
        }
    }

    public class DataContainer
    {
        private int index = -1;
        private ObjMesh mesh = null;
        private NodePoint node = null;
        private Face face = null;
        private double resultValue = 0;

        public int Index { get { return index; } set { index = value; } }
        public ObjMesh ImportedMesh { get { return mesh; } set { mesh = value; } }
        public NodePoint Node { get { return node; } set { node = value; } }
        public Face DisplayingFace { get { return face; } set { face = value; } }
        public double ResultValue { get { return resultValue; } set { resultValue = value; } }

        public DataContainer()
        {
        }
    }
}
