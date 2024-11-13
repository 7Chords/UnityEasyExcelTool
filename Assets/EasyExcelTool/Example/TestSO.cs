using System;
using System.Collections.Generic;
using UnityEngine;

public enum ColorType
{
    Yellow,
    Red,
    Blue,
    Black
}
[Serializable]
public struct TestInfo
{
    public int atk;

    public float speed;

    public string address;

}

[Serializable]
public class TestInfo_2
{
    public int hp;
    public float ap;
    public string weapon;
}

[CreateAssetMenu(fileName ="new TestSO",menuName ="EasyExcelToolTest/TestSO")]
public class TestSO : ScriptableObject
{
    public int age;

    public string myName;

    public ColorType colorType;

    public TestInfo info;

    public TestInfo_2 info_2;

    public List<int> list;

    public string[] strings;

}
