using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace HOK.RhinoReciver
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]
    public class SendCommand:IExternalCommand
    {
        private UIApplication m_app;
        private Document m_doc;

        public Result Execute(ExternalCommandData commandData, ref string message, Autodesk.Revit.DB.ElementSet elements)
        {
            try
            {
                m_app = commandData.Application;
                m_doc = m_app.ActiveUIDocument.Document;

                DialogResult dr = MessageBox.Show("Would you like to export the active view as DWG and open the file in Rhino?", "Export to DWG", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (dr == DialogResult.Yes)
                {
                    Autodesk.Revit.DB.View activeView = m_doc.ActiveView;
                    string revitPath = m_doc.PathName;

                    if (!string.IsNullOrEmpty(revitPath))
                    {
                        string directory = Path.GetDirectoryName(revitPath);
                        if (Directory.Exists(directory))
                        {
                            string fileName = Path.GetFileNameWithoutExtension(revitPath);

                            DWGExportOptions options = new DWGExportOptions();
                            ICollection<ElementId> views = new List<ElementId>();
                            views.Add(activeView.Id);

                            bool exported = false;
                            Guid transId = new Guid();
                            using (Transaction trans = new Transaction(m_doc))
                            {
                                try
                                {
                                    transId=Guid.NewGuid();

                                    trans.Start(transId.ToString());

                                    exported = m_doc.Export(directory, fileName, views, options);
                                    trans.Commit();

                                    
                                }
                                catch (Exception ex)
                                {
                                    trans.RollBack();
                                    MessageBox.Show("DWG Export was failed.\n" + ex.Message, "DWG Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                            }

                            if (exported)
                            {
                                string dwgFileName = Path.Combine(directory, fileName + ".dwg");
                                if (File.Exists(dwgFileName))
                                {
                                    RegistryKeyManager.SetRegistryKeyValue("RevitOutgoing", exported.ToString());
                                    RegistryKeyManager.SetRegistryKeyValue("RevitOutgoingViewId", activeView.Id.IntegerValue.ToString());
                                    RegistryKeyManager.SetRegistryKeyValue("RevitOutgoingPath", dwgFileName);
                                    RegistryKeyManager.SetRegistryKeyValue("RevitOutgoingId", transId.ToString());

                                    //Run Rhino Script
                                    dynamic rhino = AppCommand.Instance.RhinoInstance;
                                    if (File.Exists(dwgFileName))
                                    {
                                        string script = string.Format("-Import \"{0}\" Enter", dwgFileName);
                                        rhino.RunScript(script, false);

                                        string rhinoFileName = dwgFileName.Replace(".dwg", ".3dm");
                                        if (File.Exists(rhinoFileName))
                                        {
                                            File.Delete(rhinoFileName);
                                        }

                                        script = string.Format("-SaveAs \"{0}\" Enter", rhinoFileName);
                                        rhino.RunScript(script, false);


                                        if (File.Exists(rhinoFileName))
                                        {
                                            RegistryKeyManager.SetRegistryKeyValue("RhinoIncoming", true.ToString());
                                            RegistryKeyManager.SetRegistryKeyValue("RhinoIncomingPath", rhinoFileName);
                                            RegistryKeyManager.SetRegistryKeyValue("RhinoIncoming", transId.ToString());
                                            rhino.Visible = 1;
                                        }
                                    }
                                }
                            }

                        }
                    }
                }
                else if(dr==DialogResult.No)
                {
                    MessageBox.Show("Please open a view to be exported as DWG format.", "Active View", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to send data to Rhino.\n"+ex.Message, "Send to Rhino", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return Result.Succeeded;
        }
    }
}
