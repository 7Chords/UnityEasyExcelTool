using UnityEditor;
using UnityEngine;
public class EasyExcelToolEditor : EditorWindow
{
    ScriptableObject sourceSO;

    string createExcelFolderPath;

    string selectedExcelPath;

    string createSOFolderPath;

    //利用构造函数来设置窗口名称
    EasyExcelToolEditor()
    {
        this.titleContent = new GUIContent("EasyExcelTool");
    }

    //添加菜单栏用于打开窗口
    [MenuItem("Tool/EasyExcelTool")]
    static void ShowEasyExcelToolWindow()
    {
        EditorWindow.GetWindow(typeof(EasyExcelToolEditor));
    }
    void OnGUI()
    {
        GUILayout.BeginVertical();

        //绘制标题
        GUILayout.Space(10);
        GUI.skin.label.fontSize = 24;
        GUI.skin.label.alignment = TextAnchor.MiddleCenter;
        GUILayout.Label("EasyExcelTool");

        GUILayout.Space(10);
        GUI.skin.label.fontSize = 15;
        GUILayout.Label("----------------通过SO生成Excel---------------");
        //绘制对象
        GUILayout.Space(10);
        sourceSO = (ScriptableObject)EditorGUILayout.ObjectField("SourceSO", sourceSO, typeof(ScriptableObject), true);



        //绘制对象
        GUILayout.Space(10);
        if (GUILayout.Button("Select Folder（The folder must exit in current project)"))
        {
            string folderPath = EditorUtility.OpenFolderPanel("Select Folder", "", "");
            if (!string.IsNullOrEmpty(folderPath))
            {
                createExcelFolderPath = folderPath;
            }
        }

        createExcelFolderPath = EditorGUILayout.TextField("ExcelSavePath", createExcelFolderPath);


        EditorGUILayout.Space();

        //创建Excel文件按钮
        if (GUILayout.Button("CreateExcel"))
        {
            EasyExcelOperator.CreateExcelByData(sourceSO, createExcelFolderPath + "/" + sourceSO.name + ".xls");
        }



        GUILayout.Space(10);
        GUILayout.Label("----------------通过Excel生成SO---------------");
        GUILayout.Space(10);
        GUI.skin.label.fontSize = 15;
        GUILayout.Label("----------------单个Excel操作---------------");
        GUI.skin.label.fontSize = 9;
        GUILayout.Label("(Excel文件尽量不要放在StreamingAssets文件夹中，统一导入时由于该文件夹无法写入，会出错)");

        if (GUILayout.Button("Select Excel File（The folder must exit in current project)"))
        {
            string[] fliter = { "Excel文件", "xls;*.xlsx" };
            string excelPath = EditorUtility.OpenFilePanelWithFilters("Select Excel", "", fliter);
            if (!string.IsNullOrEmpty(excelPath))
            {
                selectedExcelPath = excelPath;

            }
        }

        selectedExcelPath = EditorGUILayout.TextField("ExcelSelectedPath", selectedExcelPath);



        if (GUILayout.Button("Select SO Folder（The folder must exit in current project)"))
        {
            string folderPath = EditorUtility.OpenFolderPanel("Select Folder", "", "");
            if (!string.IsNullOrEmpty(folderPath))
            {
                createSOFolderPath = folderPath;

            }
        }
        createSOFolderPath = EditorGUILayout.TextField("SOFolderPath", createSOFolderPath);


        if (GUILayout.Button("CreateSO/RefreshSO"))
        {
            EasyExcelImporter.ImportSingleExcel(selectedExcelPath, "Assets/"+createSOFolderPath.Split("Assets/")[1]);
        }


        GUILayout.Space(10);
        GUI.skin.label.fontSize = 15;
        GUILayout.Label("----------------所有Excel操作---------------");

        if (GUILayout.Button("CreateAllSO/RefreshAllSO"))
        {
            EasyExcelImporter.ImportAllExcels();
        }

        GUILayout.EndVertical();
    }


}
