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

    //利用构造函数来设置窗口名称
    EasyExcelToolEditor()
    {
        this.titleContent = new GUIContent("EasyExcelTool");
    }

    //添加菜单栏用于打开窗口
    [MenuItem("Tool/EasyExcelTool")]
    static void showWindow()
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

        ////绘制文本
        //GUILayout.Space(10);
        //bugReporterName = EditorGUILayout.TextField("Bug Name", bugReporterName);

        ////绘制当前正在编辑的场景
        //GUILayout.Space(10);
        //GUI.skin.label.fontSize = 12;
        //GUI.skin.label.alignment = TextAnchor.UpperLeft;
        //GUILayout.Label("Currently Scene:" + EditorSceneManager.GetActiveScene().name);

        ////绘制当前时间
        //GUILayout.Space(10);
        //GUILayout.Label("Time:" + System.DateTime.Now);
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
                excelFolderStr = folderPath;
               
            }
        }

        excelFolderStr = EditorGUILayout.TextField("ExcelSavePath", excelFolderStr);


        ////绘制描述文本区域
        //GUILayout.Space(10);
        //GUILayout.BeginHorizontal();
        //GUILayout.Label("Description", GUILayout.MaxWidth(80));
        //description = EditorGUILayout.TextArea(description, GUILayout.MaxHeight(75));
        //GUILayout.EndHorizontal();

        EditorGUILayout.Space();

        //添加名为"Save Bug"按钮，用于调用SaveBug()函数
        if (GUILayout.Button("CreateExcel"))
        {
        }

        GUILayout.Space(10);
        GUILayout.Label("----------------通过Excel生成SO---------------");

        //添加名为"Save Bug with Screenshot"按钮，用于调用SaveBugWithScreenshot() 函数
        if (GUILayout.Button("CreateSO"))
        {
            //SaveBugWithScreenshot();
        }

        GUILayout.EndVertical();
    }


}
