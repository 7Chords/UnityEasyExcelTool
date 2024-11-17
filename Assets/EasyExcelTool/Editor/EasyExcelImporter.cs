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
public class EasyExcelImporter : AssetPostprocessor
{

    static List<EasyExcelAssetInfo> cachedAssetInfos = null;
    static void OnPostprocessAllAssets(string[] importedAssets, string[] deletedAssets, string[] movedAssets, string[] movedFromAssetPaths)
    {
        bool hasImported = false;

        foreach (string path in importedAssets)
        {
            //找到所有的excel文件
            if (Path.GetExtension(path) == ".xls" || Path.GetExtension(path) == ".xlsx")
            {
                if (cachedAssetInfos == null)
                {
                    cachedAssetInfos = FindExcelAssetInfos();
                }

                string excelName = Path.GetFileNameWithoutExtension(path);

                if (excelName.StartsWith("~$"))
                {
                    continue;
                }

                EasyExcelAssetInfo info = cachedAssetInfos.Find(i => i.ExcelName == excelName);

                if (info == null)
                {
                    continue;
                }

                ImportExcel(path, info);

                hasImported = true;
            }
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

        var asset = AssetDatabase.LoadAssetAtPath(assetPath, assetType);

        //不存在这个type的SO
        if (asset == null)
        {
            asset = ScriptableObject.CreateInstance(assetType.Name);
            //test
            assetPath = "Assets/Resources/Excels/TestListSO.asset";

            AssetDatabase.CreateAsset((ScriptableObject)asset, assetPath);
            //asset.hideFlags = HideFlags.NotEditable;//
        }

        return asset;
    }

    static void ImportExcel(string excelPath, EasyExcelAssetInfo info)
    {
        string assetPath = "";

        string assetName = info.AssetType.Name + ".asset";

        if (string.IsNullOrEmpty(info.attribute.assetPath))
        {
            string basePath = Path.GetDirectoryName(excelPath);
            assetPath = Path.Combine(basePath, assetName);
        }
        else
        {
            var path = Path.Combine("Assets", info.attribute.assetPath);
            assetPath = Path.Combine(path, assetName);
        }
        Debug.Log(assetPath);
        UnityEngine.Object asset = LoadOrCreateAsset(assetPath, info.AssetType);

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

                        var obj = Activator.CreateInstance(elementType);

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

                            //if (fields[j].FieldType.IsEnum)
                            //{
                            //    // 尝试将 cellValue 直接转换为枚举类型
                            //    if (Enum.TryParse(fields[j].FieldType, cellValue, out object enumValue))
                            //    {
                            //        fields[j].SetValue(obj, enumValue);
                            //    }
                            //    else
                            //    {
                            //        // 如果转换失败，您可以根据需要抛出异常或设置默认值
                            //        throw new ArgumentException($"Invalid enum value '{cellValue}' for field {fields[j].Name}");
                            //    }
                            //}
                            //else if(fields[j].FieldType==typeof(Int32))
                            //{
                            //    int temp = Convert.ToInt32(JsonConvert.DeserializeObject(cellValue));
                            //    fields[j].SetValue(obj, temp);
                            //}
                            //else
                            //{
                            //    // 使用 JSON 反序列化对字段进行赋值
                            //    var deserializeMethod = typeof(JsonConvert).GetMethods()
                            //        .Where(m => m.Name == "DeserializeObject" && m.IsGenericMethodDefinition)
                            //        .FirstOrDefault();

                            //    var genericDeserialize = deserializeMethod.MakeGenericMethod(fields[j].FieldType);
                            //    var fieldValue = genericDeserialize.Invoke(null, new object[] { cellValue });
                            //    fields[j].SetValue(obj, fieldValue);
                            //}

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
