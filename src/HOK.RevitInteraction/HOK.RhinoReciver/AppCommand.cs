using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media.Imaging;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;

namespace HOK.RhinoReciver
{
    public class AppCommand:IExternalApplication
    {
        internal static AppCommand appCommand = null;
        private dynamic rhino = null;
        private bool rhinoInitialized=false;
        public PushButton toggleButton = null;

        private UIControlledApplication uicapp = null;
        private bool receiverActivated = false;
        private BitmapSource enableImage = null;
        private BitmapSource disableImage = null;
        private BitmapSource sendImage = null;

        public static AppCommand Instance
        {
            get { return appCommand; }
        }

        public dynamic RhinoInstance
        {
            get { return rhino; }
        }

        public bool ReceiverActivated
        {
            get { return receiverActivated; }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            DisposeRhino();
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            appCommand = this;
            uicapp = application;
            rhinoInitialized = InitializeRhino();

            enableImage = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.enable.GetHbitmap(), IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            disableImage = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.disable.GetHbitmap(), IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            sendImage = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.send.GetHbitmap(), IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions()); 

            RibbonPanel rp = application.CreateRibbonPanel(" Rhino ");
            string currentAssembly=System.Reflection.Assembly.GetAssembly(this.GetType()).Location;

            toggleButton = rp.AddItem(new PushButtonData("Activate Receiver", "Activate Receiver", currentAssembly, "HOK.RhinoReciver.Command")) as PushButton;
            toggleButton.LargeImage = enableImage;

            PushButton sendButton = rp.AddItem(new PushButtonData("Send", "Send", currentAssembly, "HOK.RhinoReciver.SendCommand")) as PushButton;
            sendButton.LargeImage = sendImage;

            return Result.Succeeded;
        }

        public bool InitializeRhino()
        {
            bool initialized = false;
            try
            {
                string rhinoId = "Rhino5x64.Application";
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
                    initialized = true;
                }
            }
            catch
            {
                initialized = false;
            }
            return initialized;
        }

        private void DisposeRhino()
        {
            if (rhinoInitialized && null != rhino)
            {
                Marshal.FinalReleaseComObject(rhino);
                rhino = null;
            }
        }

        public void Toggle()
        {
            string buttonText = toggleButton.ItemText;
            switch (buttonText)
            {
                case "Activate Receiver":
                    //activate
                    receiverActivated = true;
                    uicapp.Idling += OnIdling;

                    toggleButton.ItemText = "Disable Receiver";
                    toggleButton.LargeImage = disableImage;
                    toggleButton.Image = disableImage;
                    
                    break;
                case "Disable Receiver":
                    //disable
                    receiverActivated = false;
                    uicapp.Idling -= OnIdling;

                    toggleButton.ItemText = "Activate Receiver";
                    toggleButton.LargeImage = enableImage;
                    toggleButton.Image = enableImage;
                   
                    break;
            }
        }

        public void OnIdling(object sender, IdlingEventArgs e)
        {
           
            if (receiverActivated)
            {
                //UIApplication uiapp = sender as UIApplication;
                //Document doc = uiapp.ActiveUIDocument.Document;

                e.SetRaiseWithoutDelay();

                string value = RegistryKeyManager.GetRegistryKeyValue("RhinoOutgoing");
                if (!string.IsNullOrEmpty(value))
                {
                    if (value.ToLower() == "true")
                    {
                        TaskDialog dialog = new TaskDialog("Rhino Receiver");
                        dialog.MainInstruction = "Rhino Data";
                        dialog.MainContent = "Analysis data from Rhino is being sent to Revit.";
                        dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Visualize Analysis Results");
                        dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Cancel");

                        TaskDialogResult result = dialog.Show();
                        if (result == TaskDialogResult.CommandLink1)
                        {
                            //AVF
                        }
                        RegistryKeyManager.SetRegistryKeyValue("RhinoOutgoing", "False");
                    }
                }
            }
        }
    }
}
