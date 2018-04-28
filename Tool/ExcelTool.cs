using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Collections;
using ExcelCom = Microsoft.Office.Interop.Excel;
using Autodesk.Revit.DB;

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
                                int lastRow = workSheet.UsedRange.Rows.Count;
                                var range = workSheet.Range[workSheet.Cells[startRow + 1, startCol + 8], workSheet.Cells[lastRow + startRow, startCol + 8]] as ExcelCom.Range;
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
                                int lastRow = workSheet.UsedRange.Rows.Count;
                                //查找要进行数据更新的行数
                                int updateRow = lastRow + startRow;
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
            //获取项目中所有管道
            FilteredElementCollector pipeCollector = new FilteredElementCollector(doc);
            ElementCategoryFilter filter1 = new ElementCategoryFilter(BuiltInCategory.OST_PipeCurves);
            ElementIsElementTypeFilter filter2 = new ElementIsElementTypeFilter(true);
            pipeCollector.WherePasses(new LogicalAndFilter(filter1, filter2));
            //获得用于创建Excel的数据
            foreach (Autodesk.Revit.DB.Plumbing.Pipe p in pipeCollector)
            {
                if (p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString().Contains("P-"))
                {
                    exportInformation.Add(new Command.PipeCalculation(doc, p).PipeCalcInformation());
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
                                int lastRow = workSheet.UsedRange.Rows.Count;
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

        ////返回指定的工作簿运行的程序
        //public ExcelCom.Application GetWorkBookApplication(string workBookFullName)
        //{
        //    var excelapp=Marshal.GetActiveObject()
        //}


        //Excel版本
        public int[] ExcelNumber
        {
            get { return new int[] { 2007 }; }
        }
        //Excel是否运行
        public bool isExcelRunning
        {
            get { return Process.GetProcessesByName("EXCEL").Length != 0; }
        }
    }
}
