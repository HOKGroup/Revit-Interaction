using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            string rhinoId = "Rhino5x64.Application";
            System.Type type = System.Type.GetTypeFromProgID(rhinoId);
            dynamic rhino = System.Activator.CreateInstance(type);
            if (null != rhino)
            {
                rhino.Visible = 1;//non-visible 

                const int bail_milliseconds = 15 * 1000;
                int time_waiting = 0;
                bool initialized = true;
                while (0 == rhino.IsInitialized())
                {
                    Thread.Sleep(100);
                    time_waiting += 100;
                    if (time_waiting > bail_milliseconds)
                    {
                        //MessageBox.Show("Rhino initialization failed");
                        initialized = false;
                        break;
                    }
                }

                if (initialized)
                {
                    string dwgFile = @"\\Group\hok\ATL\RESOURCES\BIM\DivaAutomation\14.01007.00 -  DIVA Project\Illuminance Values\Input\DivaTest.dwg";
                    
                    string script = string.Format("-DivaDocumentUnits {0} {1}", "F", "F"); //document units to foot
                    rhino.RunScript(script, false);

                    script = string.Format("-Import \"{0}\" Enter", dwgFile); //model and layout unit to foot
                    rhino.RunScript(script, false);

                    string rhinoFileName = dwgFile.Replace(".dwg", ".3dm");
                    if (File.Exists(rhinoFileName))
                    {
                        File.Delete(rhinoFileName);
                    }

                    script = string.Format("-SaveAs \"{0}\" Enter", rhinoFileName);
                    rhino.RunScript(script, false);
                }
            }
        }

        
    }
}
