using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace CalcTest.Command
{
    [Transaction(TransactionMode.Manual)]
    class test_linkExcel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Excel名称及路径
            //string fileName = "pipeCalc";
            //string filePath = @"C:\Users\Administrator\Desktop\test\" + fileName + ".xlsx";
            string fullName = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\test\PlumbingCalc.xlsx";
            //Sheet1名称
            string sheet1Name = "给排水工程量";
            //第一个单元格
            int startCol = 2;
            int startRow = 2;

            //获得按钮组
            UIApplication uiapp = commandData.Application;

            foreach(var ribbonPanel in uiapp.GetRibbonPanels("Test"))
            {
                if(ribbonPanel.Name=="Excel Tool")
                {
                    foreach(var item in ribbonPanel.GetItems())
                    {
                        var rbg = item as RadioButtonGroup;
                        if (rbg != null&&rbg.Name== "Link_test")
                        {
                            //实例化更新器
                            Updater.PipeParameterUpdater ppu = new Updater.PipeParameterUpdater(commandData.Application.ActiveAddInId, fullName, sheet1Name);
                            //更新器ID
                            UpdaterId updaterId = ppu.GetUpdaterId();
                            if (UpdaterRegistry.IsUpdaterRegistered(updaterId))
                            {
                                if (UpdaterRegistry.IsUpdaterEnabled(updaterId))
                                {
                                    UpdaterRegistry.DisableUpdater(updaterId);
                                    //切换按钮
                                    foreach(ToggleButton tog in rbg.GetItems())
                                    {
                                        if(tog.Name== "link_disable")
                                        {
                                            rbg.Current = tog;
                                        }
                                    }
                                }
                                else
                                {
                                    Tool.ExcelTool excelTool = new Tool.ExcelTool();
                                    excelTool.UpdataInOpenWorkBook(commandData.Application.ActiveUIDocument.Document, fullName, sheet1Name, startRow, startCol);
                                    UpdaterRegistry.EnableUpdater(updaterId);
                                }
                            }
                            else
                            {
                                message = "未能启动更新器,请联系XXX";
                                return Result.Failed;
                            }
                            break;
                        }
                    }
                    break;
                }
            }


            return Result.Succeeded;
        }
    }
}
