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
//EXCEL�ʲ���Ϣ��
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
            //�ҵ����е�excel�ļ�
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

        if (hasImported) //�е����µ�excel���ݾͽ���ˢ��
        {
            AssetDatabase.SaveAssets();
            AssetDatabase.Refresh();
        }
    }

    //�ҵ�Ӧ�ó��������еĴ�ExcelAssetAttribute�ű�������������
    static List<EasyExcelAssetInfo> FindExcelAssetInfos()
    {
        List<EasyExcelAssetInfo> infoList = new List<EasyExcelAssetInfo>();

        foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())//��ȡ�Ѽ��ص���Ӧ�ó������ִ���������еĳ���
        {
            foreach (var type in assembly.GetTypes())//��ȡ�˳������Ѷ������������
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
    //���ػ򴴽�SO�ļ�
    static UnityEngine.Object LoadOrCreateAsset(string assetPath, Type assetType)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(assetPath));

        var asset = AssetDatabase.LoadAssetAtPath(assetPath, assetType);

        //���������type��SO
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

                    for (int i = 1; i <= sheet.LastRowNum; i++) // �ӵ�һ�п�ʼ�������һ���Ǳ����У�ʵ��ʹ��ʱ������Ҫ�����������
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue;

                        var obj = Activator.CreateInstance(elementType);

                        FieldInfo[] fields = elementType.GetFields(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                        for (int j = 0; j < fields.Length; j++)
                        {
                            ICell cell = row.GetCell(j);

                            string cellValue = cell?.ToString() ?? ""; // ����յ�Ԫ��

                            // ʹ�� JSON �����л����ֶν��и�ֵ
                            var deserializeMethod = typeof(JsonConvert).GetMethods()
                                .Where(m => m.Name == "DeserializeObject" && m.IsGenericMethodDefinition)
                                .FirstOrDefault();

                            var genericDeserialize = deserializeMethod.MakeGenericMethod(fields[j].FieldType);
                            var fieldValue = genericDeserialize.Invoke(null, new object[] { cellValue });
                            fields[j].SetValue(obj, fieldValue);

                            //if (fields[j].FieldType.IsEnum)
                            //{
                            //    // ���Խ� cellValue ֱ��ת��Ϊö������
                            //    if (Enum.TryParse(fields[j].FieldType, cellValue, out object enumValue))
                            //    {
                            //        fields[j].SetValue(obj, enumValue);
                            //    }
                            //    else
                            //    {
                            //        // ���ת��ʧ�ܣ������Ը�����Ҫ�׳��쳣������Ĭ��ֵ
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
                            //    // ʹ�� JSON �����л����ֶν��и�ֵ
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

                    // ���б�ֵ�� ScriptableObject ���ֶ�
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
