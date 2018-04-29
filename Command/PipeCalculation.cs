using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Plumbing;
using System.Collections;

namespace CalcTest.Command
{
    class PipeCalculation
    {
        static Document _doc;
        static Element _element;
        public PipeCalculation(Document doc,Element element)
        {
            _doc = doc;
            _element = element;
        }
        public ArrayList PipeCalcInformation()
        {
            ArrayList information = new ArrayList();


            //按类别提取信息
            switch ((BuiltInCategory)_element.Category.Id.IntegerValue)
            {
                //管道
                case BuiltInCategory.OST_PipeCurves:
                    Pipe _pipe = _element as Pipe;
                    //管道系统
                    string plumbingSystemName = _pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString();
                    //管道直径
                    double pipeDn = UnitUtils.ConvertFromInternalUnits(_pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble(), DisplayUnitType.DUT_MILLIMETERS);
                    //获取管道计量信息
                    ////标高
                    //information.Add(_pipe.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsValueString());
                    //区域
                    information.Add(_pipe.ReferenceLevel.get_Parameter(BuiltInParameter.LEVEL_ELEV).AsDouble() < 0 ? "地下" : "地上");
                    //系统
                    information.Add(PlumbingSystemConvert(plumbingSystemName));
                    //项目名称
                    string itemName = _pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString().Split('-').Last() + "管";
                    information.Add(itemName);
                    //材质
                    string materialName = GetPipeMaterial(plumbingSystemName);
                    information.Add(materialName);
                    //规格
                    information.Add("DN" + _pipe.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString().Split(' ')[0]);
                    //连接方式
                    string connection = GetConnection(GetPipeMaterial(plumbingSystemName), pipeDn);
                    information.Add(connection);
                    //单位
                    information.Add("m");
                    //长度
                    information.Add(Math.Round(UnitUtils.ConvertFromInternalUnits(_pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble(), DisplayUnitType.DUT_METERS), 3));
                    break;
                //管道附件
                case BuiltInCategory.OST_PipeAccessory:
                    FamilyInstance _pipeAccessory = _element as FamilyInstance;
                    //管道系统
                    string plumbingAccessorySystemName = _pipeAccessory.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString();
                    //族名称
                    string pipeAccessoryName = _pipeAccessory.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString();
                    //区域
                    information.Add(_doc.GetElement(_pipeAccessory.LevelId).get_Parameter(BuiltInParameter.LEVEL_ELEV).AsDouble() < 0 ? "地下" : "地上");
                    //系统
                    information.Add(PlumbingSystemConvert(plumbingAccessorySystemName));
                    //项目名称
                    information.Add(pipeAccessoryName.Split('-').First());
                    //材质
                    information.Add("铸铁");
                    //规格
                    information.Add("DN" + _pipeAccessory.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsString().Split(' ')[0]);
                    //连接方式
                    information.Add(pipeAccessoryName.Split('-').Last());
                    //单位
                    information.Add("个");
                    //工程量
                    information.Add(1);
                    break;
            }


            //ID
            information.Add(_element.Id.IntegerValue);


            return information;


        }

        //系统判断
        string PlumbingSystemConvert(string systemName)
        {
            string name;
            switch (systemName)
            {
                case "P-给水":name = "给水系统";
                    break;
                case "P-热水给水":name = "给水系统";
                    break;
                case "P-热水回水":name = "给水系统";
                    break;
                case "P-废水":name = "排水系统";
                    break;
                case "P-污水":name = "排水系统";
                    break;
                case "P-通气":name = "排水系统";
                    break;
                case "P-喷淋":name = "消防系统";
                    break;
                case "P-消火栓":name = "消防系统";
                    break;
                case "M-冷冻水供水":name = "空调水系统";
                    break;
                case "M-冷冻水回水":name = "空调水系统";
                    break;
                default :name = "其他";
                    break;
            }
            return name;
        }

        //材质判断
        string GetPipeMaterial(string systemName)
        {
            string name;
            switch (systemName)
            {
                case "P-给水":
                    name = "钢塑复合管";
                    break;
                case "P-热水给水":
                    name = "钢塑复合管";
                    break;
                case "P-热水回水":
                    name = "钢塑复合管";
                    break;
                case "P-废水":
                    name = "PVC-U管";
                    break;
                case "P-污水":
                    name = "PVC-U管";
                    break;
                case "P-通气":
                    name = "PVC-U管";
                    break;
                case "P-喷淋":
                    name = "镀锌钢管";
                    break;
                case "P-消火栓":
                    name = "镀锌钢管";
                    break;
                case "M-冷冻水供水":
                    name = "无缝钢管";
                    break;
                case "M-冷冻水回水":
                    name = "无缝钢管";
                    break;
                default:
                    name = "未定义";
                    break;
            }
            return name;
        }
        //连接方式判断
        string GetConnection(string material,double Dn)
        {
            string connection;
            switch (material)
            {
                case "镀锌钢管":
                    connection = Dn > 65 ? "卡箍连接" : "螺纹连接";
                    break;
                case "钢塑复合管":
                    connection = Dn > 65 ? "卡箍连接" : "螺纹连接";
                    break;
                case "PVC-U管":
                    connection = "粘接";
                    break;
                case "无缝钢管":
                    connection = "焊接";
                    break;
                default:
                    connection = "未定义";
                    break;
            }
            return connection;
        }
    }


}
