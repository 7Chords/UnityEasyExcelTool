using UnityEditor;
using UnityEngine;
using System.IO;
using UnityEditor.SceneManagement;
using UnityEngine.WSA;
using UnityEditor.VersionControl;


public class EasyExcelToolEditor : EditorWindow
{
    ScriptableObject sourceSO;

    string createExcelFolderPath;

    string selectedExcelPath;

    string createSOFolderPath;

    //���ù��캯�������ô�������
    EasyExcelToolEditor()
    {
        this.titleContent = new GUIContent("EasyExcelTool");
    }

    //��Ӳ˵������ڴ򿪴���
    [MenuItem("Tool/EasyExcelTool")]
    static void showWindow()
    {
        EditorWindow.GetWindow(typeof(EasyExcelToolEditor));
    }
    void OnGUI()
    {
        GUILayout.BeginVertical();

        //���Ʊ���
        GUILayout.Space(10);
        GUI.skin.label.fontSize = 24;
        GUI.skin.label.alignment = TextAnchor.MiddleCenter;
        GUILayout.Label("EasyExcelTool");

        GUILayout.Space(10);
        GUI.skin.label.fontSize = 15;
        GUILayout.Label("----------------ͨ��SO����Excel---------------");
        //���ƶ���
        GUILayout.Space(10);
        sourceSO = (ScriptableObject)EditorGUILayout.ObjectField("SourceSO", sourceSO, typeof(ScriptableObject), true);



        //���ƶ���
        GUILayout.Space(10);
        if (GUILayout.Button("Select Folder��The folder must exit in current project)"))
        {
            string folderPath = EditorUtility.OpenFolderPanel("Select Folder", "", "");
            if (!string.IsNullOrEmpty(folderPath))
            {
                createExcelFolderPath = folderPath;
               
            }
        }

        createExcelFolderPath = EditorGUILayout.TextField("ExcelSavePath", createExcelFolderPath);


        ////���������ı�����
        //GUILayout.Space(10);
        //GUILayout.BeginHorizontal();
        //GUILayout.Label("Description", GUILayout.MaxWidth(80));
        //description = EditorGUILayout.TextArea(description, GUILayout.MaxHeight(75));
        //GUILayout.EndHorizontal();

        EditorGUILayout.Space();

        //�����Ϊ"Save Bug"��ť�����ڵ���SaveBug()����
        if (GUILayout.Button("CreateExcel"))
        {
            EasyExcelOperator.CreateExcelByData(sourceSO, createExcelFolderPath+"/"+ sourceSO.name+".xls");
        }



        GUILayout.Space(10);
        GUILayout.Label("----------------ͨ��Excel����SO---------------");
        GUILayout.Space(10);
        if (GUILayout.Button("Select Excel File��The folder must exit in current project)"))
        {
            string[] fliter = { "Excel�ļ�", "xls;*.xlsx" };
            string excelPath = EditorUtility.OpenFilePanelWithFilters("Select Excel", "", fliter);
            if (!string.IsNullOrEmpty(excelPath))
            {
                selectedExcelPath = excelPath;

            }
        }

        selectedExcelPath = EditorGUILayout.TextField("ExcelSelectedPath", selectedExcelPath);



        if (GUILayout.Button("Select SO Folder��The folder must exit in current project)"))
        {
            string folderPath = EditorUtility.OpenFolderPanel("Select Folder", "", "");
            if (!string.IsNullOrEmpty(folderPath))
            {
                createSOFolderPath = folderPath;

            }
        }
        createSOFolderPath = EditorGUILayout.TextField("CreateSOFolderPath", createSOFolderPath);


        if (GUILayout.Button("CreateSO"))
        {
            //SaveBugWithScreenshot();
        }

        GUILayout.EndVertical();
    }


}
