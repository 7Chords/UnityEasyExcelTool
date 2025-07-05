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
public class EasyExcelImporter
{

    static List<EasyExcelAssetInfo> cachedAssetInfos = null;


    public static void ImportAllExcels()
    {
        bool hasImported = false;
        DirectoryInfo direction = new DirectoryInfo(Application.dataPath);//��ȡ�ļ��У�exportPath���ļ��е�·��
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

        if (hasImported) //�е����µ�excel���ݾͽ���ˢ��
        {
            AssetDatabase.SaveAssets();
            AssetDatabase.Refresh();
        }
    }
    public static void ImportSingleExcel(string importedAsset,string createSOPath)
    {
        bool hasImported = false;

        //�ҵ����е�excel�ļ�
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

        // ���������ʲ�
        var asset = AssetDatabase.LoadAssetAtPath(assetPath, assetType);

        // ���������������͵� SO
        if (asset == null)
        {
            // ��� assetType �Ƿ�Ϊ ScriptableObject ������
            if (!typeof(ScriptableObject).IsAssignableFrom(assetType))
            {
                Debug.LogError($"Type {assetType} is not a ScriptableObject.");
                return null; // ���������ѡ���׳��쳣
            }
            // ���������͵�ʵ��
            asset = ScriptableObject.CreateInstance(assetType);
            //�������ԣ�
            Debug.Log("create");
            cachedAssetInfos.Find(x => x.AssetType == assetType).attribute.assetPath = assetPath;
            // ȷ�� assetPath �� .asset ��β
            if (!assetPath.EndsWith(".asset"))
            {
                assetPath += ".asset";
            }

            AssetDatabase.CreateAsset(asset, assetPath);
            EditorUtility.SetDirty(asset); // ȷ���ʲ������Ϊ���޸�
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

                    for (int i = 1; i <= sheet.LastRowNum; i++) // �ӵ�һ�п�ʼ�������һ���Ǳ����У�ʵ��ʹ��ʱ������Ҫ�����������
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue;
                        var obj = ScriptableObject.CreateInstance(elementType);
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
