using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddInLogging
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnCreateInfo_Click(object sender, RibbonControlEventArgs e)
        {
            Logger.Log.Info("An Info log is created");
        }

        private void btnCreateFatal_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var zero = 0;
                var divisionByZeroError = 13 / zero;
            }
            catch (Exception ex)
            {
                Logger.Log.Info("A Fatal Error log is created", ex);
            }
            
        }

        private void btnShowLog_Click(object sender, RibbonControlEventArgs e)
        {
            var folderName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "MyCompany"); 
            //See App.Config file for details (You can change MyCompany with your company name or app name etc.
            Process.Start("explorer.exe", folderName);
        }
    }
}
