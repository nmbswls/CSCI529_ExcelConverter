using UnityEngine;
using System.Collections;
using Excel;
using System.Collections.Generic;
using System.IO;
using System.Data;
using Newtonsoft.Json;
using System.Xml;
using System;
using UnityEditor;

public enum eExcelPropertyType
{
    String,
    Int,
    Float,
    Bool,
    Enum,
    Class,
}
public class ExcelProperty
{
    public string ExcelName;
    public string AttrName;
    public eExcelPropertyType Type;
    public string enumType;

    public ExcelProperty(string ExcelName, string AttrName, string typeString)
    {
        this.ExcelName = ExcelName;
        this.AttrName = AttrName;
        Enum.TryParse<eExcelPropertyType>( typeString, true,out this.Type);

    }

    public List<ExcelProperty> SubclassProperties;
    public int MaxListCount;
}



public class ExcelReader
{
    public static string bathPath = Application.dataPath + "/Files/";
    private Dictionary<string, List<ExcelProperty>> typeDict = new Dictionary<string, List<ExcelProperty>>();
    //private Dictionary<string, int> keywords = new Dictionary<string, int>();
    private Dictionary<string, Dictionary<string,int>> keywords = new Dictionary<string, Dictionary<string, int>>();

    [MenuItem("ss/sssss")]
    public static void TestSS()
    {
        ExcelReader er = new ExcelReader();
        er.LoadXmlStruct();
        er.LoadKeywords();

        List<JSONObject> data = er.LoadData();
        string outPath = bathPath + "/out/a0.txt";
        if (File.Exists(outPath))
        {
            File.Delete(outPath);
            File.WriteAllText(outPath, "");
        }
        JSONObject ret = new JSONObject();
        for (int i = 0; i < data.Count; i++)
        {

            ret.AddField(i+"",data[i]);
        }
        
        er.WriteTo(outPath, ret.ToString());

    }

    public void WriteTo(string path, string content)
    {
        File.AppendAllText(path, content);
    }

    public string[] keys;
    public List<JSONObject> LoadData()
    {

        FileStream fileStream = File.Open(bathPath + "t0.xlsx", FileMode.Open, FileAccess.Read, FileShare.Read);
        IExcelDataReader excelDataReader = ExcelReaderFactory.CreateOpenXmlReader(fileStream);

        DataSet result = excelDataReader.AsDataSet();

        int columns = result.Tables[0].Columns.Count;
        int rows = result.Tables[0].Rows.Count;

        List<string> excelDta = new List<string>();

        keys = new string[columns];
        for (int i = 0; i < columns; i++)
        {
            string key = result.Tables[0].Rows[0][i].ToString();
            if (key.Contains("_"))
            {
                key = key.Substring(0, key.IndexOf('_'));
            }
            keys[i] = key;
        }

        List<JSONObject> ret = new List<JSONObject>();


        List<ExcelProperty> properties = typeDict["CardAsset"];

        for (int i = 1; i < rows; i++)
        {
            //JSONObject jsonObject = new JSONObject();

            int idx = 0;
            DataRow aa = result.Tables[0].Rows[i];
            JSONObject oneRowObj = dfsReadExcel(properties, result.Tables[0].Rows[i], ref idx);

            Debug.Log(oneRowObj.ToString());
            ret.Add(oneRowObj);
            //for (int j = 0; j < columns; j++)
            //{
            //    string columnName = keys[i];


            //    ExcelProperty ep = typeDict[columnName];
            //    string attrName = ep.AttrName;

            //    // 获取表格中指定行指定列的数据 
            //    string value = result.Tables[0].Rows[i][j].ToString();

            //    if(ep.Type == eExcelPropertyType.Enum)
            //    {
            //        jsonObject.AddField(attrName, ObjEnumToIdx(value));
            //    }
            //    else if (ep.Type == eExcelPropertyType.List)
            //    {
            //        if (!jsonObject.HasField("value"))
            //        {
            //            jsonObject.AddField(value, new JSONObject(JSONObject.Type.ARRAY));
            //        }
            //        jsonObject.GetField(value).Add(value);
            //    }
            //    else
            //    {
            //        jsonObject.SetField(keys[i], value);
            //    }

            //}
        }
        return ret;
    }


    public JSONObject dfsReadExcel(List<ExcelProperty> properties, DataRow dataRow, ref int idx)
    {
        JSONObject jsonObject = new JSONObject();

        for (int i = 0; i < properties.Count; i++)
        {
            eExcelPropertyType type = properties[i].Type;
            if (type == eExcelPropertyType.Enum)
            {

                //jsonObject.AddField(properties[i].AttrName, dataRow[idx].ToString());
                jsonObject.AddField(properties[i].AttrName, ObjEnumToIdx(properties[i].enumType,dataRow[idx].ToString()));
                idx++;
            }
            else if (type == eExcelPropertyType.Class)
            {
                if (!jsonObject.HasField(properties[i].AttrName))
                {
                    jsonObject.AddField(properties[i].AttrName, new JSONObject(JSONObject.Type.ARRAY));
                }
                for (int j = 0; j < properties[i].MaxListCount; j++)
                {
                    JSONObject oneObj = dfsReadExcel(properties[i].SubclassProperties, dataRow, ref idx);
                    if (oneObj.ToString() == "{}")
                    {
                        Debug.Log("空"+ properties[i].AttrName + " "+i);
                    }
                    jsonObject.GetField(properties[i].AttrName).Add(oneObj);
                }
            }
            else if (type == eExcelPropertyType.Bool)
            {
                //jsonObject.AddField(properties[i].AttrName, dataRow[idx].ToString());
                string ret = "False";
                if(dataRow[idx].ToString() == "Y" || dataRow[idx].ToString() == "y")
                {
                    ret = "True";
                }
                jsonObject.AddField(properties[i].AttrName, ret);
                idx++;
            }
            else if(type==eExcelPropertyType.Int)
            {
                if (dataRow[idx].ToString() != "")
                {
                    jsonObject.AddField(properties[i].AttrName, int.Parse(dataRow[idx].ToString()));
                }
                idx++;
            }
            else if (type == eExcelPropertyType.Float)
            {
                if (dataRow[idx].ToString() != "")
                {
                    jsonObject.AddField(properties[i].AttrName, float.Parse(dataRow[idx].ToString()));
                }
                idx++;
            }
            else
            {
                if(dataRow[idx].ToString() != "")
                {
                    jsonObject.AddField(properties[i].AttrName, dataRow[idx].ToString());
                }
                idx++;
            }
        }
        return jsonObject;
    }


    public int ObjEnumToIdx(string enumName, string value)
    {
        if(value == null || value == "")
        {
            return 0;
        }
        if (keywords.ContainsKey(enumName))
        {
            Dictionary<string, int> subDIct = keywords[enumName];
            if (!subDIct.ContainsKey(value))
            {
                return 0;
            }
            return subDIct[value];
        }
        return 0;
    }

    //    <card>
    //      <name type="int"/>
    //      <child type="class">
    //          <
    //      </child>
    //    </card>


    public void LoadXmlStruct()
    {
        //也可以前面加上@，区别就是有@的话，双引号里面的内容不转义，比如" \" "
        //string filePath = Application.dataPath+@"/Resources/item.xml";
        string filePath = bathPath + "data.xml";
        if (File.Exists(filePath))
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(filePath);
            XmlNodeList rootNodes = xmlDoc.SelectSingleNode("root").ChildNodes;

            foreach (XmlNode node in rootNodes)
            {
                Debug.Log(node.Name);
                List<ExcelProperty> lep = dfsLoadNode(node);
                typeDict.Add(node.Name, lep);
            }
        }
    }


    public List<ExcelProperty> dfsLoadNode(XmlNode root)
    {
        List<ExcelProperty> lep = new List<ExcelProperty>();
        XmlNodeList nodes = root.ChildNodes;
        foreach (XmlElement ele in nodes)
        {
            string type = ele.GetAttribute("type");
            Debug.Log("type:" + type);
            if (type == "class")
            {
                List<ExcelProperty> ret = dfsLoadNode(ele);
                ExcelProperty np = new ExcelProperty(ele.InnerXml, ele.Name, type);
                np.SubclassProperties = ret;
                int maxCount = 1;
                int.TryParse(ele.GetAttribute("MaxCount"), out maxCount);
                np.MaxListCount = maxCount;
                lep.Add(np);
            }
            else if (type == "enum")
            {
                ExcelProperty ep = new ExcelProperty(ele.InnerXml, ele.Name, type);
                ep.enumType = ele.GetAttribute("etype");
                lep.Add(ep);
            }
            else
            {
                lep.Add(new ExcelProperty(ele.InnerXml, ele.Name, type));
            }

        }
        return lep;
    }

    public void LoadKeywords()
    {
        string filePath = bathPath + "enum.xml";
        if (File.Exists(filePath))
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(filePath);
            XmlNodeList rootNodes = xmlDoc.SelectSingleNode("root").ChildNodes;

            foreach (XmlNode node in rootNodes)
            {
                Debug.Log(node.Name);
                Dictionary<string, int> d0 = new Dictionary<string, int>();
                foreach (XmlElement value in node)
                {
                    d0.Add(value.Name,int.Parse(value.InnerText));
                }
                keywords.Add(node.Name,d0);
            }
        }


        
        
    }
}