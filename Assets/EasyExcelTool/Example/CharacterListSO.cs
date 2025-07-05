using System;
using System.Collections.Generic;
using UnityEngine;
[CreateAssetMenu(fileName = "new CharacterListSO", menuName = "EasyExcelToolTest/CharacterListSO")]
[EasyExcelUsage(excelName = "CharacterListSO")]
public class CharacterListSO : ScriptableObject
{
    public List<CharacterSO> characterList_1;

    public List<CharacterSO> characterList_2;

}
