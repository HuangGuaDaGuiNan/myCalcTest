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

            ////获取路径
            //string fullName = null;
            //if (fullName == null)
            //{
            //    FolderBrowserDialog fbd = new FolderBrowserDialog();
            //    if (fbd.ShowDialog() == DialogResult.OK)
            //    {
            //        fullName = fbd.SelectedPath + "\\blumbingCalc.xlsx";
            //    }
            //}

            string fullName = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\test\PlumbingCalc.xlsx";

            if (fullName != null)
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

                //获得用于创建Excel的数据
                foreach (Element elem in plumbingCollector)
                {
                    switch ((BuiltInCategory)elem.Category.Id.IntegerValue)
                    {
                        case BuiltInCategory.OST_PipeCurves:
                            Pipe p = elem as Pipe;
                            //识别所有有缩写的管道
                            if (p.get_Parameter(BuiltInParameter.RBS_DUCT_PIPE_SYSTEM_ABBREVIATION_PARAM).AsString() != "")
                            {
                                exportInformation.Add(new PipeCalculation(doc, p).PipeCalcInformation());
                            }
                            break;
                        case BuiltInCategory.OST_PipeAccessory:
                            FamilyInstance pa = elem as FamilyInstance;
                            //识别阀门
                            if (pa.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().Contains("阀") && pa.get_Parameter(BuiltInParameter.RBS_DUCT_PIPE_SYSTEM_ABBREVIATION_PARAM).AsString() != "")
                            {
                                exportInformation.Add(new PipeCalculation(doc, pa).PipeCalcInformation());
                            }
                            break;
                    }


                }

                ////Excel名称及路径
                //string fileName = "pipeCalc";
                //string filePath = @"C:\Users\Administrator\Desktop\test\" + fileName + ".xlsx";
                //Sheet1名称
                string sheetName = "给排水工程量";

                //判断是否打开了Excel文件
                var excelTool = new Tool.ExcelTool();
                if (excelTool.ExcelVarsion != "" && excelTool.isExcelRunning)
                {
                    //捕获程序
                    var excelApp = Marshal.GetActiveObject("Excel.Application") as ExcelCom.Application;
                    //记录状态
                    bool workBookIsOpen = false;
                    bool hasWorkSheet = false;
                    //识别工作簿
                    for (int i = 0; i < excelApp.Workbooks.Count; i++)
                    {
                        var workBook = excelApp.Workbooks[i + 1] as ExcelCom.Workbook;
                        if (workBook != null && workBook.FullName == fullName)
                        {
                            workBookIsOpen = true;
                            //识别工作表
                            for (int j = 0; j < excelApp.Worksheets.Count; j++)
                            {
                                var workSheet = excelApp.Worksheets[j + 1] as ExcelCom.Worksheet;
                                if (workSheet != null && workSheet.Name == sheetName)
                                {
                                    hasWorkSheet = true;
                                    ////获得最后一列
                                    //int row = workSheet.Cells.SpecialCells(ExcelCom.XlCellType.xlCellTypeLastCell).Row;
                                    ////添加数据
                                    //workSheet.Range["A" + (row + 1).ToString()].Value = "test";

                                    //刷新数据
                                    //表1
                                    excelTool.UpdataInOpenWorkBook(doc, fullName, sheetName, 2, 2);
                                    //汇总表
                                    excelTool.UpdataPivotTable(fullName, sheetName, 2, 2);
                                    break;
                                }
                            }
                            //当工作簿中无工作表
                            if (!hasWorkSheet)
                            {
                                //to do
                            }
                        }
                    }
                    //识别失败
                    if (!workBookIsOpen)
                    {
                        //创建Excel
                        if (File.Exists(fullName)) File.Delete(fullName);
                        creatExcelFile(fullName, sheetName, 2, 2, exportInformation);
                    }
                }
                else
                {
                    //创建Excel
                    if (File.Exists(fullName)) File.Delete(fullName);
                    creatExcelFile(fullName, sheetName, 2, 2, exportInformation);
                }
            }

            return Result.Succeeded;
        }

        //创建Excel
        void creatExcelFile(string excelFileFullName, string sheet1Name, int startRow, int startCol, List<ArrayList> data)
        {
            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(excelFileFullName)))
            {
                //创建明细表
                var st = excelPackage.Workbook.Worksheets.Add(sheet1Name);
                //表头
                st.Cells[startRow, startCol].Value = "区域";
                st.Cells[startRow, startCol + 1].Value = "系统";
                st.Cells[startRow, startCol + 2].Value = "项目名称";
                st.Cells[startRow, startCol + 3].Value = "材质";
                st.Cells[startRow, startCol + 4].Value = "规格";
                st.Cells[startRow, startCol + 5].Value = "连接方式";
                st.Cells[startRow, startCol + 6].Value = "单位";
                st.Cells[startRow, startCol + 7].Value = "工程量";
                st.Cells[startRow, startCol + 8].Value = "ID";
                for (int i = 0; i < data.Count; i++)
                {
                    var arrayList = data[i];
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

                //保存
                excelPackage.Save();
            }
        }

    }
}
