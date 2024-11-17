using System;


[AttributeUsage(AttributeTargets.Class, AllowMultiple = false/*, Inherited = false*/)]//ºÃ≥–≤‚ ‘
public class EasyExcelUsageAttribute : Attribute
{
    public string assetPath { get; set; }
    public string excelName { get; set; }
    public bool logOnImport { get; set; }
}
