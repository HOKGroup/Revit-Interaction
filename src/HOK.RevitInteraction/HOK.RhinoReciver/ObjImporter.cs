using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HOK.RhinoReciver
{
    public static class ObjImporter
    {
        public static bool ReadObjFile(string objFile, out List<ObjMesh> meshes)
        {
            bool result = false;

            meshes = new List<ObjMesh>();
            try
            {
                using (StreamReader reader = new StreamReader(objFile))
                {
                    string line = string.Empty;
                    ObjMesh mesh = null;

                    while ((line = reader.ReadLine()) != null)
                    {
                        string[] split = line.Split(' ');
                        if (line.StartsWith("o "))
                        {
                            if (split.Length > 1)
                            {
                                if (null != mesh) { meshes.Add(mesh); }
                                string name = split[1];
                                mesh = new ObjMesh(name);
                            }
                        }
                        else if (line.StartsWith("v "))
                        {
                            if (split.Length > 3)
                            {
                                ObjVertice vertice = new ObjVertice(double.Parse(split[1]), double.Parse(split[2]), double.Parse(split[3]));
                                if (null != mesh)
                                {
                                    mesh.ObjVertex.Add(vertice);
                                }
                            }
                        }
                    }
                    meshes.Add(mesh);
                }
                result = true;
            }
            catch (Exception ex)
            {
                string message = ex.Message;
                result = false;
            }
            return result;
        }

    }

    public class ObjMesh
    {
        private string objName="";
        private List<ObjVertice> objVertex = new List<ObjVertice>();

        public string ObjName { get { return objName; } set { objName = value; } }
        public List<ObjVertice> ObjVertex { get { return objVertex; } set { objVertex = value; } }

        public ObjMesh(string name)
        {
            objName = name;
        }
       
    }

    public class ObjVertice
    {
        private double xval = 0;
        private double yval = 0;
        private double zval = 0;

        public double XValue { get { return xval; } set { xval = value; } }
        public double YValue { get { return yval; } set { yval = value; } }
        public double ZValue { get { return zval; } set { zval = value; } }

        public ObjVertice() { }

        public ObjVertice(double x, double y, double z)
        {
            xval = x;
            yval = y;
            zval = z;
        }
    }

    public class NodePoint
    {
         private double xval = 0;
        private double yval = 0;
        private double zval = 0;

        public double XValue { get { return xval; } set { xval = value; } }
        public double YValue { get { return yval; } set { yval = value; } }
        public double ZValue { get { return zval; } set { zval = value; } }

        public NodePoint(double x, double y, double z)
        {
            xval = x;
            yval = y;
            zval = z;
        }
    }
}
