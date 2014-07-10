using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HOK.RhinoReciver
{
    public class RhinoFileCreator
    {
        private string dwgFilePath = "";
        private string rhinoFileName = "";
        private string rhinoId = "Rhino5x64.Application";
        private dynamic rhino = null;

        public string DWGFileName { get { return dwgFilePath; } set { dwgFilePath = value; } }
        public string RhinoFileName { get { return rhinoFileName; } set { rhinoFileName = value; } }

        public RhinoFileCreator(string filePath)
        {
            dwgFilePath = filePath;
            rhinoFileName = dwgFilePath.Replace(".dwg", ".3dm");
        }

        public bool CreateRhinoFile()
        {
            bool created = false;
            try
            {
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

                    if (File.Exists(dwgFilePath))
                    {
                        string script = string.Format("-Import \"{0}\" Enter", dwgFilePath);
                        rhino.RunScript(script, false);

                        if (File.Exists(rhinoFileName))
                        {
                            File.Delete(rhinoFileName);
                        }

                        script = string.Format("-SaveAs \"{0}\" Enter", rhinoFileName);
                        rhino.RunScript(script, false);
                        if (File.Exists(rhinoFileName))
                        {
                            Marshal.FinalReleaseComObject(rhino);
                            rhino = null;
                            created = true;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Failed to create Rhino application.\nPlease make sure that you have Rhino 5 x64 installed on your machine.", "Rhino Application Initialization", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    created = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to start Rhino application.\n" + ex.Message, "Rhino Application Initialization", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                created = false;
            }
            return created;
        }

       
    }
}
