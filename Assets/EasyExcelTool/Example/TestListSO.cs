using System.Collections;
using System.Collections.Generic;
using UnityEditor.VersionControl;
using UnityEngine;
[CreateAssetMenu(fileName = "new TestListSO", menuName = "EasyExcelToolTest/TestListSO")]
[EasyExcelUsage]
public class TestListSO : ScriptableObject
{

    public List<TestSO> soList_1;

    public List<TestSO> soList_2;

}
