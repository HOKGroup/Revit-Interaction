using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Permissions;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: RegistryPermissionAttribute(SecurityAction.RequestMinimum, All = "HKEY_CURRENT_USER")]

namespace HOK.RhinoReciver
{
    public static class RegistryKeyManager
    {
        private static string keyAddress = "Software\\Autodesk\\Revit\\Autodesk Revit 2015\\RevitInteraction";


        public static InteractionProperties GetRegistryKeyValues()
        {
            InteractionProperties ip = new InteractionProperties();
            try
            {
                RegistryKey rkey = Registry.CurrentUser.OpenSubKey(keyAddress, true);
                if (null != rkey)
                {
                    ip.RevitDocName = rkey.GetValue(ip.RevitDocNameKey).ToString();
                    ip.RhinoDocName = rkey.GetValue(ip.RhinoDocNameKey).ToString();

                    ip.RevitIncoming = Convert.ToBoolean(rkey.GetValue(ip.RevitIncomingKey).ToString());
                    ip.RevitIncomingPath = rkey.GetValue(ip.RevitIncomingPathKey).ToString();
                    ip.RevitIncomingId = Guid.Parse(rkey.GetValue(ip.RevitIncomingIdKey).ToString());

                    ip.RevitOutgoing = Convert.ToBoolean(rkey.GetValue(ip.RevitOutgoingKey).ToString());
                    ip.RevitOutgoingPath = rkey.GetValue(ip.RevitOutgoingPathKey).ToString();
                    ip.RevitOutgoingId = Guid.Parse(rkey.GetValue(ip.RevitOutgoingIdKey).ToString());

                    ip.RhinoIncoming = Convert.ToBoolean(rkey.GetValue(ip.RhinoIncomingKey).ToString());
                    ip.RhinoIncomingPath = rkey.GetValue(ip.RhinoIncomingPathKey).ToString();
                    ip.RhinoIncomingId = Guid.Parse(rkey.GetValue(ip.RhinoIncomingIdKey).ToString());

                    ip.RhinoOutgoing = Convert.ToBoolean(rkey.GetValue(ip.RhinoOutgoingKey).ToString());
                    ip.RhinoOutgoingPath = rkey.GetValue(ip.RhinoOutgoingPathKey).ToString();
                    ip.RhinoOutgoingId = Guid.Parse(rkey.GetValue(ip.RhinoOutgoinIdKey).ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to get registry key values.\n"+ex.Message, "Get Registry Key Values", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            return ip;
        }

        public static string GetRegistryKeyValue(string key)
        {
            string value = "";
            RegistryKey rkey = Registry.CurrentUser.OpenSubKey(keyAddress, true);
            if (null != rkey)
            {
                value = rkey.GetValue(key).ToString();
            }
            return value;
        }

        public static bool SetRegistryKeyValues(InteractionProperties ip)
        {
            bool result = false;
            try
            {
                RegistryKey rkey = Registry.CurrentUser.OpenSubKey(keyAddress, true);
                if (null == rkey)
                {
                    rkey = Registry.CurrentUser.CreateSubKey(keyAddress);
                }

                if (null != rkey)
                {
                    rkey.SetValue(ip.RevitDocNameKey, ip.RevitDocName);
                    rkey.SetValue(ip.RhinoDocNameKey, ip.RhinoDocName);

                    rkey.SetValue(ip.RevitIncomingKey, ip.RevitIncoming.ToString());
                    rkey.SetValue(ip.RevitIncomingPathKey, ip.RevitIncomingPath);
                    rkey.SetValue(ip.RevitIncomingIdKey, ip.RevitIncomingId.ToString());

                    rkey.SetValue(ip.RevitOutgoingKey, ip.RevitOutgoing.ToString());
                    rkey.SetValue(ip.RevitOutgoingPathKey, ip.RevitOutgoingPath);
                    rkey.SetValue(ip.RevitOutgoingIdKey, ip.RevitOutgoingId.ToString());

                    rkey.SetValue(ip.RhinoIncomingKey, ip.RhinoIncoming.ToString());
                    rkey.SetValue(ip.RhinoIncomingPathKey, ip.RhinoIncomingPath);
                    rkey.SetValue(ip.RhinoIncomingIdKey, ip.RhinoIncomingId.ToString());

                    rkey.SetValue(ip.RhinoOutgoingKey, ip.RhinoOutgoing.ToString());
                    rkey.SetValue(ip.RhinoOutgoingPathKey, ip.RhinoOutgoingPath);
                    rkey.SetValue(ip.RhinoOutgoinIdKey, ip.RhinoOutgoingId.ToString());

                    result = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to get registry key values.\n" + ex.Message, "Get Registry Key Values", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            return result;
        }

        public static void SetRegistryKeyValue(string key, string value)
        {
            RegistryKey rkey = Registry.CurrentUser.OpenSubKey(keyAddress, true);
            if (null == rkey)
            {
                rkey = Registry.CurrentUser.CreateSubKey(keyAddress);
            }

            if (null != rkey)
            {
                if (!string.IsNullOrEmpty(value))
                {
                    rkey.SetValue(key, value);
                }
            }
        }
    }

    public class InteractionProperties
    {
        private string revitDocName = "";
        private string rhinoDocName = "";

        private bool revitOutgoing = false;
        private string revitOutgoingPath = "";
        private Guid revitOutgoingId = new Guid();

        private bool revitIncoming = false;
        private string revitIncomingPath = "";
        private Guid revitIncomingId = new Guid();

        private bool rhinoOutgoing = false;
        private string rhinoOutgoingPath = "";
        private Guid rhinoOutgoingId = new Guid();

        private bool rhinoIncoming = false;
        private string rhinoIncomingPath = "";
        private Guid rhinoIncomingId = new Guid();

        public string RevitDocName { get { return revitDocName; } set { revitDocName = value; } }
        public string RhinoDocName { get { return rhinoDocName; } set { rhinoDocName = value; } }
        
        public string RevitDocNameKey = "RevitDocName";
        public string RhinoDocNameKey = "RhinoDocName";

        public bool RevitOutgoing { get { return revitOutgoing; } set { revitOutgoing = value; } }
        public string RevitOutgoingPath { get { return revitOutgoingPath; } set { revitOutgoingPath = value; } }
        public Guid RevitOutgoingId { get { return revitOutgoingId; } set { revitOutgoingId = value; } }
        
        public string RevitOutgoingKey="RevitOutgoing";
        public string RevitOutgoingPathKey = "RevitOutgoingPath";
        public string RevitOutgoingIdKey = "RevitOutgoingId";

        public bool RevitIncoming { get { return revitIncoming; } set { revitIncoming = value; } }
        public string RevitIncomingPath { get { return revitIncomingPath; } set { revitIncomingPath = value; } }
        public Guid RevitIncomingId { get { return revitIncomingId; } set { revitIncomingId = value; } }

        public string RevitIncomingKey = "RevitIncoming";
        public string RevitIncomingPathKey = "RevitIncomingPath";
        public string RevitIncomingIdKey = "RevitIncomingId";

        public bool RhinoOutgoing { get { return rhinoOutgoing; } set { rhinoOutgoing = value; } }
        public string RhinoOutgoingPath { get { return rhinoOutgoingPath; } set { rhinoOutgoingPath = value; } }
        public Guid RhinoOutgoingId { get { return rhinoOutgoingId; } set { rhinoOutgoingId = value; } }

        public string RhinoOutgoingKey = "RhinoOutgoing";
        public string RhinoOutgoingPathKey = "RhinoOutgoingPath";
        public string RhinoOutgoinIdKey = "RhinoOutgoingId";

        public bool RhinoIncoming { get { return rhinoIncoming; } set { rhinoIncoming = value; } }
        public string RhinoIncomingPath { get { return rhinoIncomingPath; } set { rhinoIncomingPath = value; } }
        public Guid RhinoIncomingId { get { return rhinoIncomingId; } set { rhinoIncomingId = value; } }

        public string RhinoIncomingKey = "RhinoIncoming";
        public string RhinoIncomingPathKey = "RhinoIncomingPath";
        public string RhinoIncomingIdKey = "RhinoIncomingId";

        public InteractionProperties()
        { }


    }
}
