using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Ionic.Zip;

namespace ClosedXML_Package
{
    class Program
    {
        static void Main(string[] args)
        {
#if DEBUG
            return;
#endif
            try
            {
                var targetPath = args[0]; 
                //if (!targetPath.EndsWith("Release")) return;

                var targetDll = String.Format(@"{0}\ClosedXML.dll", targetPath);
                var targetXml = String.Format(@"{0}\ClosedXML.XML", targetPath);

                if (!File.Exists(targetDll)) 
                {
                    MessageBox.Show("Missing: " + targetDll);
                    return;
                }
                if (!File.Exists(targetXml))
                {
                    MessageBox.Show("Missing: " + targetXml);
                    return;
                }

                Assembly assembly = Assembly.LoadFrom(targetDll);
                Version ver = assembly.GetName().Version;
                var targetZipPath = String.Format(@"{0}\ClosedXML_v{1}.zip", targetPath, ver);
                var targetZipInfo = new FileInfo(targetZipPath);
                if (targetZipInfo.Exists) targetZipInfo.Delete();

                using (ZipFile zip = new ZipFile())
                {
                    zip.AddFile(targetDll, "");
                    zip.AddFile(targetXml, "");
                    zip.Save(targetZipPath);
                }
            }
            catch (Exception ex)
            {
                var strMsg = "->" + ex.Message;
                if (ex.InnerException != null)
                    strMsg += Environment.NewLine + "=>" + ex.InnerException.Message;
                MessageBox.Show(strMsg);
            }

        }
    }
}
