using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Collections;
using System.IO;
using ExcelCom = Microsoft.Office.Interop.Excel;
using Autodesk.Revit.DB;
using Microsoft.Win32;

namespace CalcTest.Tool
{
    class ExcelTool
    {
        //使用Com组件删除管道数据
        public void DeleteExcelDataByCom(string workBookFullName, string workSheetName, int startRow, int startCol, int elementId)
        {
            var excelApp = Marshal.GetActiveObject("Excel.Application") as ExcelCom.Application;
            if (excelApp != null)
            {
                bool workBookIsOpen = false;
                bool hasWorkSheet = false;

                //识别工作簿
                for (int i = 0; i < excelApp.Workbooks.Count; i++)
                {
                    var workBook = excelApp.Workbooks[i + 1] as ExcelCom.Workbook;
                    if (workBook != null && workBook.FullName == workBookFullName)
                    {
                        //识别工作表
                        for (int j = 0; j < excelApp.Worksheets.Count; j++)
                        {
                            var workSheet = excelApp.Worksheets[j + 1] as ExcelCom.Worksheet;
                            if (workSheet != null && workSheet.Name == workSheetName)
                            {
                                //获得最后一行
                                int lastRow = workSheet.UsedRange.Rows.Count + startRow - 1;
                                var range = workSheet.Range[workSheet.Cells[startRow + 1, startCol + 8], workSheet.Cells[lastRow, startCol + 8]] as ExcelCom.Range;
                                var rangeFind = range.Find(elementId);
                                if (rangeFind != null)
                                {
                                    (workSheet.Rows[rangeFind.Row] as ExcelCom.Range).Delete(ExcelCom.XlDirection.xlDown);
                                }
                                break;
                            }
                        }
                        break;
                    }

                    //如果没有识别到工作簿，应判断工作簿状态，分辨一下要不要创建Excel

                }
            }
            
        }

        //使用Com组件更新管道数据
        public void UpdaterExcelDataByCom(string workBookFullName,string workSheetName,int startRow,int startCol,ArrayList dataArrayList)
        {
            var excelApp = Marshal.GetActiveObject("Excel.Application") as ExcelCom.Application;
            if (excelApp != null)
            {
                bool workBookIsOpen = false;
                bool hasWorkSheet = false;

                //识别工作簿
                for (int i = 0; i < excelApp.Workbooks.Count; i++)
                {
                    var workBook = excelApp.Workbooks[i + 1] as ExcelCom.Workbook;
                    if (workBook != null && workBook.FullName == workBookFullName)
                    {
                        //识别工作表
                        for (int j = 0; j < excelApp.Worksheets.Count; j++)
                        {
                            var workSheet = excelApp.Worksheets[j + 1] as ExcelCom.Worksheet;
                            if (workSheet != null && workSheet.Name == workSheetName)
                            {
                                //获得最后一行
                                int lastRow = workSheet.UsedRange.Rows.Count + startRow - 1;
                                //查找要进行数据更新的行数
                                int updateRow = lastRow + 1;
                                int id = (int)dataArrayList[dataArrayList.Count - 1];
                                var range = workSheet.Range[workSheet.Cells[startRow + 1, startCol + 8], workSheet.Cells[lastRow, startCol + 8]] as ExcelCom.Range;
                                var rangeFind = range.Find(id);
                                if (rangeFind != null) { updateRow = rangeFind.Row; }
                                //更新数据
                                for (int k = 0; k < dataArrayList.Count; k++)
                                {
                                    workSheet.Cells[updateRow, startCol + k].Value = dataArrayList[k];
                                }
                                break;
                            }
                        }
                        break;
                    }

                    //如果没有识别到工作簿，应判断工作簿状态，分辨一下要不要创建Excel

                }

            }

        }

        //在打开的工作簿中全部更新一次数据
        public void UpdataInOpenWorkBook(Document doc, string workBookFullName, string workSheetName, int startRow, int startCol)
        {

            List<ArrayList> exportInformation = new List<ArrayList>();
            //获取项目中所有管道和管道附件
            FilteredElementCollector plumbingCollector = new FilteredElementCollector(doc);
            ElementIsElementTypeFilter filter1 = new ElementIsElementTypeFilter(true);

            List<ElementFilter> filterSet = new List<ElementFilter>();
            filterSet.Add(new ElementCategoryFilter(BuiltInCategory.OST_PipeCurves));
            filterSet.Add(new ElementCategoryFilter(BuiltInCategory.OST_PipeAccessory));
            LogicalOrFilter orFilter = new LogicalOrFilter(filterSet);

            plumbingCollector.WherePasses(new LogicalAndFilter(filter1, orFilter)).ToElements();


            int num = 0;

            //获得用于创建Excel的数据
            foreach (Element elem in plumbingCollector)
            {
                switch ((BuiltInCategory)elem.Category.Id.IntegerValue)
                {
                    case BuiltInCategory.OST_PipeCurves:
                        Autodesk.Revit.DB.Plumbing.Pipe p = elem as Autodesk.Revit.DB.Plumbing.Pipe;
                        //识别所有有缩写的管道
                        if (p.get_Parameter(BuiltInParameter.RBS_DUCT_PIPE_SYSTEM_ABBREVIATION_PARAM).AsString() != "")
                        {
                            num += 1;
                            exportInformation.Add(new Command.PipeCalculation(doc, p).PipeCalcInformation());
                        }
                        break;
                    case BuiltInCategory.OST_PipeAccessory:
                        FamilyInstance pa = elem as FamilyInstance;
                        //识别阀门
                        if (pa.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().Contains("阀"))
                        {
                            num += 1;
                            exportInformation.Add(new Command.PipeCalculation(doc, pa).PipeCalcInformation());
                        }
                        break;
                }


            }

            var excelApp = Marshal.GetActiveObject("Excel.Application") as ExcelCom.Application;
            if (excelApp != null)
            {
                bool workBookIsOpen = false;
                bool hasWorkSheet = false;
                //查找工作簿
                for (int i = 0; i < excelApp.Workbooks.Count; i++)
                {
                    var workBook = excelApp.Workbooks[i + 1] as ExcelCom.Workbook;
                    //Autodesk.Revit.UI.TaskDialog.Show("goodwish", workBook.FullName + "\n" + workBookFullName);
                    if (workBook != null && workBook.FullName == workBookFullName)
                    {
                        //Autodesk.Revit.UI.TaskDialog.Show("goodwish", "ishere");
                        workBookIsOpen = true;
                        //查找工作表
                        for (int j = 0; j < workBook.Worksheets.Count; j++)
                        {
                            var workSheet = excelApp.Worksheets[j + 1] as ExcelCom.Worksheet;
                            if (workSheet != null && workSheet.Name == workSheetName)
                            {
                                //获得最后一行
                                int lastRow = workSheet.UsedRange.Rows.Count + startRow - 1;
                                //清空数据
                                var range = workSheet.Range[workSheet.Cells[startRow + 1, startCol], workSheet.Cells[lastRow, startCol + 8]] as ExcelCom.Range;
                                range.Value2 = null;
                                //输入数据
                                for (int k = 0; k < exportInformation.Count; k++)
                                {
                                    var arrayList = exportInformation[k];
                                    for (int l = 0; l < arrayList.Count; l++)
                                    {
                                        workSheet.Cells[k + startRow + 1, l + startCol].Value = arrayList[l];
                                    }
                                }
                                break;
                            }
                        }
                        //工作簿中没有目标工作表
                        if (hasWorkSheet == false)
                        {
                            //to do
                        }
                        break;
                    }
                }
                //工作表未打开
                if (workBookIsOpen == false)
                {
                    //to do
                }
            }

        }

        //更新汇总表
        public void UpdataPivotTable(string workBookFullName,string workSheetName,int startRow,int startCol)
        {
            var excelApp = Marshal.GetActiveObject("Excel.Application") as ExcelCom.Application;
            if (excelApp != null)
            {
                bool workBookIsOpen = false;
                bool hasWorkSheet = false;
                //查找工作簿
                for (int i = 0; i < excelApp.Workbooks.Count; i++)
                {
                    var workBook = excelApp.Workbooks[i + 1] as ExcelCom.Workbook;
                    //Autodesk.Revit.UI.TaskDialog.Show("goodwish", workBook.FullName + "\n" + workBookFullName);
                    if (workBook != null && workBook.FullName == workBookFullName)
                    {
                        workBookIsOpen = true;
                        //查找工作表
                        for (int j = 0; j < workBook.Worksheets.Count; j++)
                        {
                            var workSheet = excelApp.Worksheets[j + 1] as ExcelCom.Worksheet;
                            if (workSheet != null && workSheet.Name == workSheetName)
                            {
                                //获得最后一行
                                int lastRow = workSheet.UsedRange.Rows.Count + startRow - 1;

                                //更新透视表
                                var _workSheet = workBook.Worksheets["汇总表"] as ExcelCom.Worksheet;
                                var _pt = _workSheet.PivotTables("汇总表") as ExcelCom.PivotTable;
                                //数据源 [工作表名称]!R[起始行]C[起始列]:R[总行数]R[总列数]
                                _pt.SourceData = workSheetName + "!R" + startRow.ToString() + "C" + startCol.ToString() + ":R" + lastRow.ToString() + "C" + (startCol + 8).ToString();
                                _pt.Update();

                                break;
                            }
                        }
                        //工作簿中没有目标工作表
                        if (hasWorkSheet == false)
                        {
                            //to do
                        }
                        break;
                    }
                }
                //工作表未打开
                if (workBookIsOpen == false)
                {
                    //to do
                }
            }
        }

        ////返回指定的工作簿运行的程序
        //public ExcelCom.Application GetWorkBookApplication(string workBookFullName)
        //{
        //    var excelapp=Marshal.GetActiveObject()
        //}


        //Excel版本
        public string ExcelVarsion
        {
            
            get {
                string varsions = "";
                if (Type.GetTypeFromProgID("Excel.Application") != null)
                {
                    RegistryKey rk = Registry.LocalMachine;
                    RegistryKey rk_2003 = rk.OpenSubKey(@"SOFTWARE\\Microsoft\\Office\\11.0\\Word\\InstallRoot\\");
                    if (rk_2003 != null)
                    {
                        string exePath = rk_2003.GetValue("Path").ToString() + "Excel.exe";
                        if (File.Exists(exePath)) varsions += "Excel2003;";
                    }
                    RegistryKey rk_2007= rk.OpenSubKey(@"SOFTWARE\\Microsoft\\Office\\12.0\\Word\\InstallRoot\\");
                    if (rk_2007 != null)
                    {
                        string exePath = rk_2007.GetValue("Path").ToString() + "Excel.exe";
                        if (File.Exists(exePath)) varsions += "Excel2007;";
                    }
                    RegistryKey rk_2013= rk.OpenSubKey(@"SOFTWARE\\Microsoft\\Office\\15.0\\Word\\InstallRoot\\");
                    if (rk_2013 != null)
                    {
                        string exePath = rk_2013.GetValue("Path").ToString() + "EXCEL.EXE";
                        if (File.Exists(exePath)) varsions += "Excel2013;";
                    }
                }
                return varsions; }
        }

        //Excel是否运行
        public bool isExcelRunning
        {
            get { return Process.GetProcessesByName("EXCEL").Length != 0; }
        }
    }
}
