﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;

namespace HOK.RhinoReciver
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]
    class Command:IExternalCommand
    {
        private UIApplication m_app;
        private Document m_doc;
        
        public Result Execute(ExternalCommandData commandData, ref string message, Autodesk.Revit.DB.ElementSet elements)
        {
            m_app = commandData.Application;
            m_doc = m_app.ActiveUIDocument.Document;

            if (AppCommand.Instance.ReceiverActivated == false) //to be activated
            {
                RegistryKeyManager.SetRegistryKeyValue("RevitDocName", m_doc.PathName);
            }

            AppCommand.Instance.Toggle();
            return Result.Succeeded;
        }

       
    }
}
