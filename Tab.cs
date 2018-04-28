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

            //检测Revit版本

            //检测是否安装Excel及其版本
            int[] excelNum = new Tool.ExcelTool().ExcelNumber;

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

            PushButtonData PBD_linkExcel = new PushButtonData("LinkExcel", "创建实时连接", dllPath, "CalcTest.Command.LinkExcelFile");
            PushButton PB_linkExcel = mainPanel.AddItem(PBD_linkExcel) as PushButton;

            //更新器
            Updater.PipeParameterUpdater ppu = new Updater.PipeParameterUpdater(application.ActiveAddInId);
            UpdaterId updaterId = ppu.GetUpdaterId();
            //注册更新器
            if (!UpdaterRegistry.IsUpdaterRegistered(updaterId)) UpdaterRegistry.RegisterUpdater(ppu);
            //默认关闭
            UpdaterRegistry.DisableUpdater(updaterId);
            //过滤器
            ElementCategoryFilter filter1 = new ElementCategoryFilter(BuiltInCategory.OST_PipeCurves);
            ElementIsElementTypeFilter filter2 = new ElementIsElementTypeFilter(true);
            LogicalAndFilter andFilter = new LogicalAndFilter(filter1, filter2);
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

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
    ////管道更新器
    //public class PipeParameterUpdater : IUpdater
    //{
    //    AddInId addinId;
    //    UpdaterId updaterId;
    //    public PipeParameterUpdater(AddInId id)
    //    {
    //        addinId = id;
    //        updaterId = new UpdaterId(addinId, new Guid("b0111042-f770-491b-b452-7353e49b2e35"));
    //    }
    //    public void Execute(UpdaterData data)
    //    {
    //        Tool.ExcelTool excelTool = new Tool.ExcelTool();
    //        if(excelTool.isExcelRunning)
    //        {
    //            var excelApp = Marshal.GetActiveObject("Excel.Application") as ExcelCom.Application;
    //            string fileName = "pipeCalc";
    //            string fullName = @"C:\Users\Administrator\Desktop\test\" + fileName + ".xlsx";
    //            string sheet1Name = "给排水工程量";
    //            int startCol = 2;
    //            int startRow = 2;

    //            Document doc = data.GetDocument();
    //            ICollection<ElementId> elementCollection = data.GetModifiedElementIds();
    //            foreach(ElementId id in elementCollection)
    //            {
    //                Autodesk.Revit.DB.Plumbing.Pipe pipe = doc.GetElement(id) as Autodesk.Revit.DB.Plumbing.Pipe;
    //                if (pipe != null)
    //                {
    //                    var dataArrayList = new Command.PipeCalculation(doc, pipe).PipeCalcInformation();
    //                    excelTool.UpdaterExcelDataByCom(excelApp, fullName, sheet1Name, startRow, startCol, dataArrayList);
    //                }
    //            }


    //        }
    //    }
    //    //加载错误提示
    //    public string GetAdditionalInformation()
    //    {
    //        return "管道更新器未成功加载";
    //    }
    //    //更新器优先级
    //    public ChangePriority GetChangePriority()
    //    {
    //        return ChangePriority.FloorsRoofsStructuralWalls;
    //    }
    //    //更新器全局标识
    //    public UpdaterId GetUpdaterId()
    //    {
    //        return updaterId;
    //    }
    //    //更新器名称
    //    public string GetUpdaterName()
    //    {
    //        return "管道更新器";
    //    }
    //}
}
