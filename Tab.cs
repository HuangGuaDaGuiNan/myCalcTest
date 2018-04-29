using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using Microsoft.Win32;

namespace CalcTest
{
    [Transaction(TransactionMode.Manual)]
    class Tab : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            //检测是否安装Excel及其版本
            string excelNum = new Tool.ExcelTool().ExcelVarsion;

            //动态链接库路径
            string dllPath = typeof(Tab).Assembly.Location;

            //Tab名称
            string tabName = "Test";
            //创建Tab
            application.CreateRibbonTab(tabName);
            //创建Panel
            RibbonPanel mainPanel = application.CreateRibbonPanel(tabName, "Excel Tool");
            //创建按钮
            PushButtonData PBD_createExcel = new PushButtonData("CreateExcel", "导出数据", dllPath, "CalcTest.Command.CreateExcelFile");
            PushButton PB_createExcel = mainPanel.AddItem(PBD_createExcel) as PushButton;

            //PushButtonData PBD_linkExcel = new PushButtonData("LinkExcel", "创建实时连接", dllPath, "CalcTest.Command.LinkExcelFile");
            //PushButton PB_linkExcel = mainPanel.AddItem(PBD_linkExcel) as PushButton;
            //PB_linkExcel.AvailabilityClassName = "CalcTest.LinkExcelEnable";

            RadioButtonGroupData RBGD_linkExcel = new RadioButtonGroupData("Link_test");
            RadioButtonGroup RBG_linkExcel = mainPanel.AddItem(RBGD_linkExcel) as RadioButtonGroup;
            ToggleButton TB_linkExcel = RBG_linkExcel.AddItem(new ToggleButtonData("link_enable", "创建实时连接", dllPath, "CalcTest.Command.test_linkExcel"));
            ToggleButton TB_dislinkExcel = RBG_linkExcel.AddItem(new ToggleButtonData("link_disable", "关闭实时连接"));
            TB_dislinkExcel.Visible = false;
            RBG_linkExcel.Current = TB_dislinkExcel;
            TB_linkExcel.AvailabilityClassName = "CalcTest.LinkExcelEnable";


            //判断本机是否安装Excel
            if (excelNum != "")
            {
                //Excel名称及路径
                //string fileName = "pipeCalc";
                string fullName = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\test\PlumbingCalc.xlsx";
                //Sheet1名称
                string sheetName = "给排水工程量";

                //更新器
                Updater.PipeParameterUpdater ppu = new Updater.PipeParameterUpdater(application.ActiveAddInId, fullName, sheetName);
                UpdaterId updaterId = ppu.GetUpdaterId();
                //注册更新器
                if (!UpdaterRegistry.IsUpdaterRegistered(updaterId)) UpdaterRegistry.RegisterUpdater(ppu);
                //默认关闭
                UpdaterRegistry.DisableUpdater(updaterId);
                //过滤器
                //FilteredElementCollector plumbingCollector = new FilteredElementCollector(doc);
                ElementIsElementTypeFilter filter1 = new ElementIsElementTypeFilter(true);
                List<ElementFilter> filterSet = new List<ElementFilter>();
                filterSet.Add(new ElementCategoryFilter(BuiltInCategory.OST_PipeCurves));
                filterSet.Add(new ElementCategoryFilter(BuiltInCategory.OST_PipeAccessory));
                LogicalOrFilter orFilter = new LogicalOrFilter(filterSet);

                LogicalAndFilter andFilter = new LogicalAndFilter(filter1, orFilter);

                //订阅事件
                //标高参数改变
                UpdaterRegistry.AddTrigger(updaterId, andFilter, Element.GetChangeTypeParameter(new ElementId(BuiltInParameter.RBS_START_LEVEL_PARAM)));
                //系统类型参数改变
                UpdaterRegistry.AddTrigger(updaterId, andFilter, Element.GetChangeTypeParameter(new ElementId(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)));
                //尺寸参数改变
                UpdaterRegistry.AddTrigger(updaterId, andFilter, Element.GetChangeTypeParameter(new ElementId(BuiltInParameter.RBS_CALCULATED_SIZE)));
                //长度参数改变
                UpdaterRegistry.AddTrigger(updaterId, andFilter, Element.GetChangeTypeParameter(new ElementId(BuiltInParameter.CURVE_ELEM_LENGTH)));
                //元素添加
                UpdaterRegistry.AddTrigger(updaterId, andFilter, Element.GetChangeTypeElementAddition());
                //元素删除
                UpdaterRegistry.AddTrigger(updaterId, andFilter, Element.GetChangeTypeElementDeletion());
            }


            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
   
    //按钮可用性
    class LinkExcelEnable : IExternalCommandAvailability
    {
        public bool IsCommandAvailable(UIApplication applicationData, CategorySet selectedCategories)
        {
            UIDocument uidoc = applicationData.ActiveUIDocument;
            if (uidoc != null && !uidoc.Document.IsFamilyDocument)
            {
                Tool.ExcelTool excelTool = new Tool.ExcelTool();
                if (excelTool.ExcelVarsion != "")
                {
                    return true;
                }
            }
            return false;
        }
    }

}
