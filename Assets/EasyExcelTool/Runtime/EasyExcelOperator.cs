using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Unity.Plastic.Newtonsoft.Json;
using UnityEngine;

public class EasyExcelOperator
{
    public static void CreateExcelByData(object dataList)
    {
        IWorkbook workbook = CreateWorkbookByData(dataList);

        using (FileStream fs = new FileStream(Application.streamingAssetsPath + "/Excels/TestExcel.csv", FileMode.Create))
        {
            workbook.Write(fs);
        }
    }
    private static IWorkbook CreateWorkbookByData(object dataLists)
    {
        HSSFWorkbook workbook = new HSSFWorkbook();

        if (dataLists == null)
        {
            return workbook;
        }

        Type dataType = dataLists.GetType();

        //�м���list
        FieldInfo[] fieldInfos = dataType.GetFields(BindingFlags.Public | BindingFlags.Instance);

        for (int i = 0; i < fieldInfos.Length; i++)
        {
            ISheet sheet = workbook.CreateSheet(fieldInfos[i].Name);
            //Ĭ�ϵ�һ���Ǳ�����
            IRow titleRow = sheet.CreateRow(0);

            if (fieldInfos[i].FieldType.IsGenericType && fieldInfos[i].FieldType.GetGenericTypeDefinition() == typeof(List<>))
            {
                IList objList = fieldInfos[i].GetValue(dataLists) as IList;

                Type elementType = fieldInfos[i].FieldType.GetGenericArguments()[0];

                FieldInfo[] infos = elementType.GetFields(BindingFlags.Public | BindingFlags.Instance);

                //���ݱ��������������� �����Ǳ�����
                for (int j = 0; j < infos.Length; j++)
                {
                    ICell cell = titleRow.CreateCell(j);

                    cell.SetCellValue(infos[j].Name);
                }
                //�ѵ�ǰ�б��������䵽���sheetҳ
                for (int k = 0; k < objList.Count; k++)
                {
                    SetRowOfWorkbook(ref workbook, objList[k], i, k + 1);
                }
            }
        }

        return workbook;
    }
    private static void SetRowOfWorkbook(ref HSSFWorkbook workbook, object singleData, int sheet, int row)
    {
        FieldInfo[] fieldInfos = singleData.GetType().GetFields(BindingFlags.Public | BindingFlags.Instance);

        IRow dataRow = workbook.GetSheetAt(sheet).CreateRow(row);

        for (int i = 0; i < fieldInfos.Length; i++)
        {
            if (fieldInfos[i].FieldType.IsValueType || fieldInfos[i].FieldType.IsClass)
            {
                ICell cell = dataRow.CreateCell(i);
                //��json���л����ݽ��д洢
                string jsonStr = JsonConvert.SerializeObject(fieldInfos[i].GetValue(singleData));

                cell.SetCellValue(jsonStr);
            }
        }
    }

    // Other methods for handling Excel to DataTable conversion can be added here

}
