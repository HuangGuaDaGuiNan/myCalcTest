using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Plumbing;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections;
using ExcelCom = Microsoft.Office.Interop.Excel;

namespace CalcTest.Command
{
    [Transaction(TransactionMode.Manual)]
    class CreateExcelFile : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            List<ArrayList> exportInformation = new List<ArrayList>();
            //获取项目中所有管道
            FilteredElementCollector pipeCollector = new FilteredElementCollector(doc);
            ElementCategoryFilter filter1 = new ElementCategoryFilter(BuiltInCategory.OST_PipeCurves);
            ElementIsElementTypeFilter filter2 = new ElementIsElementTypeFilter(true);
            pipeCollector.WherePasses(new LogicalAndFilter(filter1, filter2));


            int num = 0;

            //获得用于创建Excel的数据
            foreach (Pipe p in pipeCollector)
            {
                if (p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString().Contains("P-"))
                {
                    num += 1;
                    exportInformation.Add(new PipeCalculation(doc, p).PipeCalcInformation());
                }
            }

            TaskDialog.Show("test", "识别" + num.ToString() + "根管道");



            //Excel名称及路径
            string fileName = "pipeCalc";
            string filePath = @"C:\Users\Administrator\Desktop\test\" + fileName + ".xlsx";
            //Sheet1名称
            string sheet1Name = "给排水工程量";
            //第一个单元格
            int startCol = 2;
            int startRow = 2;

            //判断是否打开了Excel文件
            var excelTool = new Tool.ExcelTool();
            if (excelTool.ExcelNumber.Contains(2007) && excelTool.isExcelRunning)
            {
                //捕获程序
                var excelApp = Marshal.GetActiveObject("Excel.Application") as ExcelCom.Application;



                //识别工作簿
                for(int i = 0; i < excelApp.Workbooks.Count; i++)
                {
                    var workBook = excelApp.Workbooks[i + 1] as ExcelCom.Workbook;
                    if (workBook != null && workBook.FullName == filePath)
                    {
                        //识别工作表
                        for (int j = 0; j < excelApp.Worksheets.Count; j++)
                        {
                            var workSheet = excelApp.Worksheets[j + 1] as ExcelCom.Worksheet;
                            if (workSheet != null && workSheet.Name == sheet1Name)
                            {
                                ////获得最后一列
                                //int row = workSheet.Cells.SpecialCells(ExcelCom.XlCellType.xlCellTypeLastCell).Row;
                                ////添加数据
                                //workSheet.Range["A" + (row + 1).ToString()].Value = "test";

                                excelTool.UpdataInOpenWorkBook(doc, filePath, sheet1Name, startRow, startCol);
                                //var _pt = workBook.PivotTables[0] as ExcelCom.PivotTables;
                                //_pt.

                                break;
                            }
                        }
                    }

                    //如果没有识别到工作簿，应判断工作簿状态，分辨一下要不要创建Excel

                }
            }
            else
            {
                //创建Excel
                if (File.Exists(filePath)) File.Delete(filePath);
                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePath)))
                {
                    //创建明细表
                    var st = excelPackage.Workbook.Worksheets.Add(sheet1Name);
                    //表头
                    st.Cells[startRow, startCol].Value = "区域";
                    st.Cells[startRow, startCol+1].Value = "系统";
                    st.Cells[startRow, startCol+2].Value = "项目名称";
                    st.Cells[startRow, startCol+3].Value = "材质";
                    st.Cells[startRow, startCol+4].Value = "规格";
                    st.Cells[startRow, startCol+5].Value = "连接方式";
                    st.Cells[startRow, startCol+6].Value = "单位";
                    st.Cells[startRow, startCol+7].Value = "工程量";
                    st.Cells[startRow, startCol+8].Value = "ID";
                    for (int i = 0; i < exportInformation.Count; i++)
                    {
                        var arrayList = exportInformation[i];
                        for (int j = 0; j < arrayList.Count; j++)
                        {
                            st.Cells[i + startRow + 1, j + startCol].Value = arrayList[j];
                        }
                    }

                    //创建汇总表
                    var st1 = excelPackage.Workbook.Worksheets.Add("汇总表");
                    //数据源
                    var dataRange = st.Cells[st.Dimension.Address];
                    //创建透视表
                    var pt = st1.PivotTables.Add(st1.Cells["A1"], dataRange, "汇总表");
                    //添加字段
                    pt.RowFields.Add(pt.Fields["区域"]);
                    pt.RowFields.Add(pt.Fields["系统"]);
                    pt.RowFields.Add(pt.Fields["项目名称"]);
                    pt.RowFields.Add(pt.Fields["材质"]);
                    pt.RowFields.Add(pt.Fields["规格"]);
                    pt.RowFields.Add(pt.Fields["连接方式"]);
                    pt.RowFields.Add(pt.Fields["单位"]);
                    pt.DataFields.Add(pt.Fields["工程量"]);
                    //关闭行列汇总计算
                    pt.RowGrandTotals = false;
                    pt.ColumGrandTotals = false;
                    //关闭+/-符号
                    pt.ShowDrill = false;

                    //关闭
                    //关闭分类汇总
                    foreach (var field in pt.RowFields)
                    {
                        field.SubTotalFunctions = eSubTotalFunctions.None;
                    }
                    //表格形式
                    foreach (var field in pt.Fields)
                    {
                        field.Compact = false;
                        field.Outline = false;
                        field.ShowAll = false;
                        field.SubtotalTop = false;
                    }


                    excelPackage.Save();
                }
            }




            return Result.Succeeded;
        }
    }
}
