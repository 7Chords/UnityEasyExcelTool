using System;
using System.Collections.Generic;
using UnityEngine;

public enum CharacterClass
{
    fighter,
    warlord,
    archer,
    magician
}
[Serializable]
public struct WeaponInfo
{
    public string weaponName;

    public int atkDamage;

    public float criticalChance;

}

[Serializable]
public class CharacterBagSlot
{
    public string itemName;

    public int itemAmount;
}

[CreateAssetMenu(fileName = "new CharacterSO", menuName = "EasyExcelToolTest/CharacterSO")]
public class CharacterSO : ScriptableObject
{
    public string characterName;

    public CharacterClass characterClass;

    public WeaponInfo equipedWeapon;

    public string[] friendNames;

    public List<CharacterBagSlot> characterBag;

}
