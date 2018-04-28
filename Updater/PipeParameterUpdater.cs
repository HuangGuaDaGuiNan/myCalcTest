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
        public PipeParameterUpdater(AddInId id)
        {
            addinId = id;
            updaterId = new UpdaterId(addinId, new Guid("b0111042-f770-491b-b452-7353e49b2e35"));
        }
        public void Execute(UpdaterData data)
        {
            Tool.ExcelTool excelTool = new Tool.ExcelTool();
            if (excelTool.isExcelRunning)
            {
                var excelApp = Marshal.GetActiveObject("Excel.Application") as ExcelCom.Application;
                string fileName = "pipeCalc";
                string fullName = @"C:\Users\Administrator\Desktop\test\" + fileName + ".xlsx";
                string sheet1Name = "给排水工程量";
                int startCol = 2;
                int startRow = 2;

                //项目文档
                Document doc = data.GetDocument();
                //监视新增及修改的元素
                ICollection<ElementId> elementIdCollection_add = data.GetAddedElementIds();
                ICollection<ElementId> elementIdCollection_mod = data.GetModifiedElementIds();
                foreach (ElementId id in elementIdCollection_add.Union(elementIdCollection_mod))
                {
                    Autodesk.Revit.DB.Plumbing.Pipe pipe = doc.GetElement(id) as Autodesk.Revit.DB.Plumbing.Pipe;
                    if (pipe != null)
                    {
                        var dataArrayList = new Command.PipeCalculation(doc, pipe).PipeCalcInformation();
                        excelTool.UpdaterExcelDataByCom(fullName, sheet1Name, startRow, startCol, dataArrayList);
                    }
                }
                //监视删除的元素
                ICollection<ElementId> elementIdCollection_del = data.GetDeletedElementIds();
                foreach (ElementId id in elementIdCollection_del)
                {
                    excelTool.DeleteExcelDataByCom(fullName, sheet1Name, startRow, startCol, id.IntegerValue);
                }

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
