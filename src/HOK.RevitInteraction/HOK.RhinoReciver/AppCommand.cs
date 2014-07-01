using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
        public PushButton toggleButton = null;

        private UIControlledApplication uicapp = null;
        private bool receiverActivated = false;
        private BitmapSource enableImage = null;
        private BitmapSource disableImage = null;

        public static AppCommand Instance
        {
            get { return appCommand; }
        }

        public bool ReceiverActivated
        {
            get { return receiverActivated; }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            appCommand = this;
            uicapp = application;

            enableImage = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.enable.GetHbitmap(), IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            disableImage = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.disable.GetHbitmap(), IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions()); 

            RibbonPanel rp = application.CreateRibbonPanel(" Rhino ");
            string currentAssembly=System.Reflection.Assembly.GetAssembly(this.GetType()).Location;

            PushButton sendButton = rp.AddItem(new PushButtonData("Activate Receiver", "Activate Receiver", currentAssembly, "HOK.RhinoReciver.Command")) as PushButton;
            sendButton.LargeImage = enableImage;
            toggleButton = sendButton;

            return Result.Succeeded;
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
