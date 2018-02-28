using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddInForMacro
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnInjectMacro_Click(object sender, RibbonControlEventArgs e)
        {
            InjectMacro();
        }
        
        private void btnCreateTable_Click_1(object sender, RibbonControlEventArgs e)
        {
            RunMacro("CreateTable", "Table is created succesfully");
        }
        private void btnFormatTable_Click(object sender, RibbonControlEventArgs e)
        {
            RunMacro("FormatTable", "Table is formatted succesfully");
        }

        private void btnRunAll_Click(object sender, RibbonControlEventArgs e)
        {
            RunMacro("RunAll", "Table is formatted succesfully");
        }

        private void RunMacro(string MacroName, string SuccessMessage)
        {
            if (!isMacroInjected())
            {
                InjectMacro();
            }
            try
            {
                if (ExcelUtils.RunMacro(MacroName))
                {
                    MessageBox.Show(SuccessMessage);
                }
                else
                {
                    MessageBox.Show(string.Format("{0} could not be executed.", MacroName));
                }
            }
            catch
            {
                MessageBox.Show(string.Format("{0} could not be executed succesfully", MacroName));
            }
        }

        private void InjectMacro()
        {
            var assembly = Assembly.GetExecutingAssembly();
            StreamReader _textStreamReader;

            try
            {
                assembly = Assembly.GetExecutingAssembly();
                _textStreamReader = new StreamReader(assembly.GetManifestResourceStream("ExcelAddInForMacro.Resources.Macro.txt"));
                var macroText = _textStreamReader.ReadToEnd();
                var firstCodeModule = ExcelUtils.GetFirstCodeModule(Globals.ThisAddIn.Application.ActiveWorkbook);
                var newStandardModule = firstCodeModule != null ? firstCodeModule : Globals.ThisAddIn.Application.ActiveWorkbook.VBProject.VBComponents.Add(Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule);
                var codeModule = newStandardModule.CodeModule;
                if (codeModule.CountOfLines > 0)
                {
                    codeModule.DeleteLines(1, codeModule.CountOfLines - 1);
                }
                codeModule.AddFromString(macroText);
                //Globals.ThisAddIn.Application.ActiveWorkbook.Save();
            }
            catch (Exception ex)
            {

            }
        }

        private bool isMacroInjected()
        {
            try
            {
                var firstCodeModule = ExcelUtils.GetFirstCodeModule(Globals.ThisAddIn.Application.ActiveWorkbook);
                return firstCodeModule != null;
            }
            catch (Exception ex)
            {
                //Logger.Log.Fatal("Macro Injection Check could not be executed.", ex);
                return false;
            }
        }

        
    }
}
