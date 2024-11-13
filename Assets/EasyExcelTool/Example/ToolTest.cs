using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class ToolTest : MonoBehaviour
{
    // Start is called before the first frame update

    public ScriptableObject soData;
    void Start()
    {
        EasyExcelOperator.CreateExcelByData(soData);


    }

    // Update is called once per frame
    void Update()
    {
        
    }
}
