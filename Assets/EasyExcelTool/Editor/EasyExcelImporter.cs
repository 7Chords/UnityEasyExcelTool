using Newtonsoft.Json;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using UnityEditor;
using UnityEngine;
//EXCEL资产信息类
class EasyExcelAssetInfo
{
    public Type AssetType { get; set; }
    public EasyExcelUsageAttribute attribute { get; set; }
    public string ExcelName
    {
        get
        {
            return string.IsNullOrEmpty(attribute.excelName) ? AssetType.Name : attribute.excelName;
        }
        private set { }
    }
}
public class EasyExcelImporter
{

    static List<EasyExcelAssetInfo> cachedAssetInfos = null;


    public static void ImportAllExcels()
    {
        bool hasImported = false;
        DirectoryInfo direction = new DirectoryInfo(Application.dataPath);//获取文件夹，exportPath是文件夹的路径
        FileInfo[] files = direction.GetFiles("*", SearchOption.AllDirectories);


        for(int i =0;i<files.Length; i++)
        {
            if (Path.GetExtension(files[i].FullName) == ".xls" || Path.GetExtension(files[i].FullName) == ".xlsx")
            {
                if (cachedAssetInfos == null)
                {
                    cachedAssetInfos = FindExcelAssetInfos();
                }
                string excelName = Path.GetFileNameWithoutExtension(files[i].FullName);

                if (excelName.StartsWith("~$"))
                {
                    continue;
                }

                EasyExcelAssetInfo info = cachedAssetInfos.Find(i => i.ExcelName == excelName);

                if (info == null)
                {
                    continue;
                }

                ImportExcel(files[i].FullName, info);

                hasImported = true;
            }
        }

        if (hasImported) //有导入新的excel数据就进行刷新
        {
            AssetDatabase.SaveAssets();
            AssetDatabase.Refresh();
        }
    }
    public static void ImportSingleExcel(string importedAsset,string createSOPath)
    {
        bool hasImported = false;

        //找到所有的excel文件
        if (Path.GetExtension(importedAsset) == ".xls" || Path.GetExtension(importedAsset) == ".xlsx")
        {
            if (cachedAssetInfos == null)
            {
                cachedAssetInfos = FindExcelAssetInfos();
            }
            string excelName = Path.GetFileNameWithoutExtension(importedAsset);

            if (excelName.StartsWith("~$"))
            {
                return;
            }


            EasyExcelAssetInfo info = cachedAssetInfos.Find(i => i.ExcelName == excelName);


            Debug.Log(cachedAssetInfos[0].ExcelName);
            Debug.Log(excelName);

            

            if (info == null)
            {
                return;
            }

            ImportExcel(importedAsset, info, createSOPath);

            Debug.Log("import");

            hasImported = true;

        }
        if (hasImported) //有导入新的excel数据就进行刷新
        {
            AssetDatabase.SaveAssets();
            AssetDatabase.Refresh();
        }
    }

    //找到应用程序中所有的带ExcelAssetAttribute脚本，并创建对象
    static List<EasyExcelAssetInfo> FindExcelAssetInfos()
    {
        List<EasyExcelAssetInfo> infoList = new List<EasyExcelAssetInfo>();

        foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())//获取已加载到此应用程序域的执行上下文中的程序集
        {
            foreach (var type in assembly.GetTypes())//获取此程序集中已定义的所有类型
            {
                var attributes = type.GetCustomAttributes(typeof(EasyExcelUsageAttribute), false);
                if (attributes.Length == 0) continue;
                var attribute = (EasyExcelUsageAttribute)attributes[0];
                var info = new EasyExcelAssetInfo()
                {
                    AssetType = type,
                    attribute = attribute
                };
                infoList.Add(info);
            }
        }
        return infoList;
    }
    static IWorkbook LoadBook(string excelPath)
    {
        using (FileStream stream = File.Open(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            if (Path.GetExtension(excelPath) == ".xls")
            {
                return new HSSFWorkbook(stream);
            }
            else
            {
                return new XSSFWorkbook(stream);
            }
        }
    }
    //加载或创建SO文件
    static UnityEngine.Object LoadOrCreateAsset(string assetPath, Type assetType)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(assetPath));

        // 加载现有资产
        var asset = AssetDatabase.LoadAssetAtPath(assetPath, assetType);

        // 如果不存在这个类型的 SO
        if (asset == null)
        {
            // 检查 assetType 是否为 ScriptableObject 的子类
            if (!typeof(ScriptableObject).IsAssignableFrom(assetType))
            {
                Debug.LogError($"Type {assetType} is not a ScriptableObject.");
                return null; // 或者你可以选择抛出异常
            }
            // 创建该类型的实例
            asset = ScriptableObject.CreateInstance(assetType);
            //更新特性？
            Debug.Log("create");
            cachedAssetInfos.Find(x => x.AssetType == assetType).attribute.assetPath = assetPath;
            // 确保 assetPath 以 .asset 结尾
            if (!assetPath.EndsWith(".asset"))
            {
                assetPath += ".asset";
            }

            AssetDatabase.CreateAsset(asset, assetPath);
            EditorUtility.SetDirty(asset); // 确保资产被标记为已修改
        }

        return asset;
    }
    static void ImportExcel(string excelPath, EasyExcelAssetInfo info,string createSOPath="")
    {
        string assetPath = "";
        if (!string.IsNullOrEmpty(createSOPath) )
        {
            assetPath = Path.Combine(createSOPath, info.AssetType.Name + ".asset");
        }
        else
        {
            string assetName = info.AssetType.Name + ".asset";

            if (string.IsNullOrEmpty(info.attribute.assetPath))
            {
                string basePath = Path.GetDirectoryName(excelPath);
                assetPath = Path.Combine(basePath, assetName);
            }
            else
            {
                assetPath = info.attribute.assetPath;
            }
        }

        UnityEngine.Object asset = LoadOrCreateAsset(assetPath, info.AssetType);

        Debug.Log(assetPath);

        IWorkbook book = LoadBook(excelPath);

        var assetFields = info.AssetType.GetFields();

        int sheetCount = 1;

        foreach (var assetField in assetFields)
        {
            if (assetField.FieldType.IsGenericType && assetField.FieldType.GetGenericTypeDefinition() == typeof(List<>))
            {
                Type elementType = assetField.FieldType.GetGenericArguments()[0];
                ISheet sheet = book.GetSheet(assetField.Name);
                if (sheet != null)
                {

                    var genericType = typeof(List<>).MakeGenericType(elementType);

                    var list = Activator.CreateInstance(genericType);

                    for (int i = 1; i <= sheet.LastRowNum; i++) // 从第一行开始，假设第一行是标题行，实际使用时可能需要根据情况调整
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue;
                        var obj = ScriptableObject.CreateInstance(elementType);
                        FieldInfo[] fields = elementType.GetFields(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                        for (int j = 0; j < fields.Length; j++)
                        {
                            ICell cell = row.GetCell(j);

                            string cellValue = cell?.ToString() ?? ""; // 处理空单元格

                            // 使用 JSON 反序列化对字段进行赋值
                            var deserializeMethod = typeof(JsonConvert).GetMethods()
                                .Where(m => m.Name == "DeserializeObject" && m.IsGenericMethodDefinition)
                                .FirstOrDefault();

                            var genericDeserialize = deserializeMethod.MakeGenericMethod(fields[j].FieldType);
                            var fieldValue = genericDeserialize.Invoke(null, new object[] { cellValue });
                            fields[j].SetValue(obj, fieldValue);

                        }
                        MethodInfo listAddMethod = list.GetType().GetMethod("Add", new Type[] { elementType });

                        listAddMethod.Invoke(list, new object[] { obj });

                    }

                    // 将列表赋值给 ScriptableObject 的字段
                    assetField.SetValue(asset, list);
                }

                sheetCount++;
            }
        }
      

        if (info.attribute.logOnImport)
        {
            Debug.Log(string.Format("Imported {0} sheets form {1}.", sheetCount, excelPath));
        }

        EditorUtility.SetDirty(asset);
    }
}
