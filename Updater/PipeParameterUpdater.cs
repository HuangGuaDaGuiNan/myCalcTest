using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using System.Runtime.InteropServices;
using ExcelCom = Microsoft.Office.Interop.Excel;

namespace CalcTest.Updater
{
    class PipeParameterUpdater:IUpdater
    {
        AddInId addinId;
        UpdaterId updaterId;
        string fullName;
        string sheet1Name;
        public PipeParameterUpdater(AddInId id,string excelFileFullName,string sheetName)
        {
            addinId = id;
            updaterId = new UpdaterId(addinId, new Guid("b0111042-f770-491b-b452-7353e49b2e35"));
            fullName = excelFileFullName;
            sheet1Name = sheetName;
        }
        public void Execute(UpdaterData data)
        {
            Tool.ExcelTool excelTool = new Tool.ExcelTool();
            if (excelTool.isExcelRunning)
            {
                var excelApp = Marshal.GetActiveObject("Excel.Application") as ExcelCom.Application;

                //项目文档
                Document doc = data.GetDocument();
                //监视新增及修改的元素
                ICollection<ElementId> elementIdCollection_add = data.GetAddedElementIds();
                ICollection<ElementId> elementIdCollection_mod = data.GetModifiedElementIds();
                foreach (ElementId id in elementIdCollection_add.Union(elementIdCollection_mod))
                {
                    Element elem = doc.GetElement(id);
                    switch ((BuiltInCategory)elem.Category.Id.IntegerValue)
                    {
                        case BuiltInCategory.OST_PipeCurves:
                            var p = elem as Autodesk.Revit.DB.Plumbing.Pipe;
                            //识别所有有缩写的管道
                            if (p.get_Parameter(BuiltInParameter.RBS_DUCT_PIPE_SYSTEM_ABBREVIATION_PARAM).AsString() != "")
                            {
                                var dataArrayList = new Command.PipeCalculation(doc, p).PipeCalcInformation();
                                excelTool.UpdaterExcelDataByCom(fullName, sheet1Name, 2, 2, dataArrayList);
                            }
                            break;
                        case BuiltInCategory.OST_PipeAccessory:
                            FamilyInstance pa = elem as FamilyInstance;
                            //识别阀门
                            if (pa.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().Contains("阀") && pa.get_Parameter(BuiltInParameter.RBS_DUCT_PIPE_SYSTEM_ABBREVIATION_PARAM).AsString() != "")
                            {
                                var dataArrayList = new Command.PipeCalculation(doc, pa).PipeCalcInformation();
                                excelTool.UpdaterExcelDataByCom(fullName, sheet1Name, 2, 2, dataArrayList);
                            }
                            break;
                    }
                }
                //监视删除的元素
                ICollection<ElementId> elementIdCollection_del = data.GetDeletedElementIds();
                foreach (ElementId id in elementIdCollection_del)
                {
                    excelTool.DeleteExcelDataByCom(fullName, sheet1Name, 2, 2, id.IntegerValue);

                }
                excelTool.UpdataPivotTable(fullName, sheet1Name, 2, 2);

            }
        }
        //加载错误提示
        public string GetAdditionalInformation()
        {
            return "管道更新器未成功加载";
        }
        //更新器优先级
        public ChangePriority GetChangePriority()
        {
            return ChangePriority.FloorsRoofsStructuralWalls;
        }
        //更新器全局标识
        public UpdaterId GetUpdaterId()
        {
            return updaterId;
        }
        //更新器名称
        public string GetUpdaterName()
        {
            return "管道更新器";
        }
    }
}
