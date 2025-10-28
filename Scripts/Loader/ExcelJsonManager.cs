using System.Collections;
using System.Collections.Generic;
using System.IO;
using UnityEngine;

public class ExcelJsonManager : MonoBehaviour
{
    [Space, Header("Setting Folder Paths"), SerializeField]
    List<string> SettingFolderPaths;
    [Space, Header("Setting Json Path"), SerializeField]
    string SettingJsonPath;

    [Space, Header("Excel Folder Paths"), SerializeField]
    List<string> ExcelFolderPaths;

    [Space, Header("Excel Setting"), SerializeField]
    public ExcelSetting ExcelSetting;

    void Awake()
    {
        SettingFolderPaths.Add(Application.streamingAssetsPath);
        SettingFolderPaths.Add(Application.streamingAssetsPath + "/Setting Json");
        SettingFolderPaths.Add(Application.streamingAssetsPath + "/Setting Json/Excel");
        SettingJsonPath = Application.streamingAssetsPath + "/Setting Json/Excel/Excel Setting.json";

        ExcelFolderPaths.Add(Application.streamingAssetsPath + "/Data");
        ExcelFolderPaths.Add(Application.streamingAssetsPath + "/Data/Excel");

        CheckFolder();
        CheckJson();
        CheckExcelFile();
    }

    void CheckFolder()
    {
        foreach (string folder in SettingFolderPaths)
        {
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
        }

        foreach (string folder in ExcelFolderPaths)
        {
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
        }
    }

    void CheckJson()
    {
        if (!File.Exists(SettingJsonPath))
        {
            WriteJson(SettingJsonPath);
        }
        ReadJson(SettingJsonPath);
    }

    void CheckExcelFile()
    {
        if (!File.Exists(ExcelSetting.strExcelPath))
        {
            // Create an empty Excel file placeholder
            FileStream fileStream = File.Create(ExcelSetting.strExcelPath);
            fileStream.Close();
            Debug.Log("Excel Data.xlsx did not exist. Created an empty file at: " + ExcelSetting.strExcelPath);
        }
    }

    public void WriteJson(string path)
    {
        ExcelSetting ExcalSetting = new ExcelSetting()
        {
            strExcelPath = Application.streamingAssetsPath + "/Data/Excel/Excel Data.xlsx",
            strReportDataPath = ""
        };

        string settingdata = JsonUtility.ToJson(ExcalSetting);

        StreamWriter file = new StreamWriter(path);
        file.Write(settingdata);
        file.Close();
    }

    public void ReadJson(string path)
    {
        using (StreamReader streamreader = File.OpenText(path))
        {
            string settingdata = streamreader.ReadToEnd();
            streamreader.Close();

            ExcelSetting = JsonUtility.FromJson<ExcelSetting>(settingdata);
        }
    }
}

[System.Serializable]
public class ExcelSetting
{
    public string strExcelPath;
    public string strReportDataPath;
}