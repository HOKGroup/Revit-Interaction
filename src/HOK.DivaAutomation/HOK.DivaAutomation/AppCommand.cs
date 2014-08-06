using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;

namespace HOK.DivaAutomation
{
    public class AppCommand:IExternalApplication
    {

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            RibbonPanel rp = application.CreateRibbonPanel("DIVA Automation");
            string currentAssembly = System.Reflection.Assembly.GetAssembly(this.GetType()).Location;

            BitmapSource sendImage = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.send.GetHbitmap(), IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            BitmapSource importImage = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.import.GetHbitmap(), IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());

            PushButton sendButton = rp.AddItem(new PushButtonData("ExportFile", "Export File", currentAssembly, "HOK.DivaAutomation.ExporterCommand")) as PushButton;
            sendButton.LargeImage = sendImage;
            
            PushButton importButton = rp.AddItem(new PushButtonData("ImportData", "Import Data", currentAssembly, "HOK.DivaAutomation.ImporterCommand")) as PushButton;
            importButton.LargeImage = importImage;

            return Result.Succeeded;
        }
    }
}
