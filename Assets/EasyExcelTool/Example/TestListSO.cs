using System;
using System.Collections.Generic;
using UnityEngine;
[CreateAssetMenu(fileName = "new TestListSO", menuName = "EasyExcelToolTest/TestListSO")]
[EasyExcelUsage(excelName = "TestListSO")]
[Serializable]
public class TestListSO : ScriptableObject
{

    public List<TestSO> soList_1;

    public List<TestSO> soList_2;

    //public List<Info> list_1;

    //public List<Info> list_2;


}
