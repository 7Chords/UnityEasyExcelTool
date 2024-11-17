using UnityEditor;
using UnityEngine;
using System.IO;
using UnityEditor.SceneManagement;
using UnityEngine.WSA;
using UnityEditor.VersionControl;


public class EasyExcelToolEditor : EditorWindow
{
    ScriptableObject sourceSO;

    string excelFolderStr;

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

        ////�����ı�
        //GUILayout.Space(10);
        //bugReporterName = EditorGUILayout.TextField("Bug Name", bugReporterName);

        ////���Ƶ�ǰ���ڱ༭�ĳ���
        //GUILayout.Space(10);
        //GUI.skin.label.fontSize = 12;
        //GUI.skin.label.alignment = TextAnchor.UpperLeft;
        //GUILayout.Label("Currently Scene:" + EditorSceneManager.GetActiveScene().name);

        ////���Ƶ�ǰʱ��
        //GUILayout.Space(10);
        //GUILayout.Label("Time:" + System.DateTime.Now);
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
                excelFolderStr = folderPath;
               
            }
        }

        excelFolderStr = EditorGUILayout.TextField("ExcelSavePath", excelFolderStr);


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
        }

        GUILayout.Space(10);
        GUILayout.Label("----------------ͨ��Excel����SO---------------");

        //�����Ϊ"Save Bug with Screenshot"��ť�����ڵ���SaveBugWithScreenshot() ����
        if (GUILayout.Button("CreateSO"))
        {
            //SaveBugWithScreenshot();
        }

        GUILayout.EndVertical();
    }


}
