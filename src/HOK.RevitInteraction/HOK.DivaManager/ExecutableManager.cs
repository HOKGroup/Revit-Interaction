using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HOK.DivaManager
{
    public static class ExecutableManager
    {
        public static void RunExcutable(string exeFile, string arguments)
        {
            try
            {
                if (File.Exists(exeFile))
                {
                    System.Diagnostics.Process process = new System.Diagnostics.Process();
                    System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
                    //startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                    startInfo.FileName = exeFile;
                    startInfo.Arguments = arguments;
                    process.StartInfo = startInfo;
                    process.Start();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(exeFile+" Failed to run an executable.\n"+ex.Message, "Run Executable", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
