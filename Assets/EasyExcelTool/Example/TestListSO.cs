using System.Collections;
using System.Collections.Generic;
using UnityEngine;
[CreateAssetMenu(fileName = "new TestListSO", menuName = "EasyExcelToolTest/TestListSO")]
public class TestListSO : ScriptableObject
{

    public List<TestSO> soList_1;

    public List<TestSO> soList_2;
}
