using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using VBA = Microsoft.Vbe.Interop;

namespace ExcelAddInForMacro
{
    public class ExcelUtils
    {
       
        /// <summary>
        /// https://social.msdn.microsoft.com/Forums/office/en-US/7d8da16b-b04e-43ae-93d6-630d989fba9a/how-can-i-check-for-a-macros-existencepermission-from-my-addin?forum=exceldev
        /// </summary>
        /// <param name="wb"></param>
        /// <returns></returns>
        public static Hashtable GetMacros(Excel.Workbook wb)
        {
            Hashtable ht = new Hashtable();
            VBA.VBProject prj;
            VBA.CodeModule code;
            string composedFile;
            prj = wb.VBProject;

            foreach (VBA.VBComponent comp in prj.VBComponents)
            {
                // get the stand module of the project
                if (comp.Type == VBA.vbext_ComponentType.vbext_ct_StdModule)
                {
                    code = comp.CodeModule;
                    // Put the name of the code module at the top
                    composedFile = comp.Name + Environment.NewLine;

                    // Loop through the (1-indexed) lines
                    for (int i = 0; i < code.CountOfLines; i++)
                    {
                        composedFile +=
                            code.get_Lines(i + 1, 1) + Environment.NewLine;
                    }
                    ht.Add(comp.Name, composedFile);
                }
            }
            return ht;
        }
        public static VBA.VBComponent GetFirstCodeModule(Excel.Workbook wb)
        {

            VBA.VBProject prj;
            VBA.CodeModule code;
            prj = wb.VBProject;

            foreach (VBA.VBComponent comp in prj.VBComponents)
            {
                // get the stand module of the project
                if (comp.Type == VBA.vbext_ComponentType.vbext_ct_StdModule)
                {
                    return comp;
                }
            }
            return null;
        }
        public static bool RunMacro(string MacroName)
        {
            try
            {
                VBA.VBProject prj;
                VBA.CodeModule code;
                prj = Globals.ThisAddIn.Application.ActiveWorkbook.VBProject;

                foreach (VBA.VBComponent comp in prj.VBComponents)
                {
                    // get the stand module of the project
                    if (comp.Type == VBA.vbext_ComponentType.vbext_ct_StdModule)
                    {
                        code = comp.CodeModule;
                        if (code.ProcStartLine[MacroName, VBA.vbext_ProcKind.vbext_pk_Proc] > 0)
                        {
                            Globals.ThisAddIn.Application.Run(MacroName);
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public static Worksheet GetSheetByName(Workbook workbook, string sheetName)
        {
            foreach (var sheet in workbook.Sheets)
            {
                if (((Worksheet)sheet).Name.ToUpper() == sheetName.ToUpper())
                {
                    return (Worksheet)sheet;
                }
            }
            return null;
        }

        public static int GetColumnIndexOfTitle(Worksheet sheet, string columnName, int titleRow = 1)
        {
            try
            {
                for (int i = 1; i <= sheet.UsedRange.Columns.Count; i++)
                {
                    var cell = sheet.Cells[titleRow, i];
                    if (cell.Value.ToString().ToUpper() == columnName.ToUpper())
                    {
                        return i;
                    }
                }
                return -1;
            }
            catch (Exception ex)
            {
                return -1;

            }
        }
    }
}
