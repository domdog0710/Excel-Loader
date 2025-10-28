using FlexFramework.Excel;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using UnityEngine;
using UnityEngine.Networking;

public class ExcelLoader : MonoBehaviour
{
    [SerializeField, Space, Header("Excel Json Manager")]
    ExcelJsonManager ExcelJsonManager;

    [SerializeField, Space, Header("Excel Datas")]
    public List<ExcelData> ExcelDatas;

    [SerializeField, Space, Header("Excel Load Boolean")]
    bool bExcelLoad = false;

    public delegate void DownloadHandler(byte[] bytes);

    // Start is called before the first frame update
    void Start()
    {
        LoadExcel();
    }

    //���Jexcel
    void LoadExcel()
    {
        StartCoroutine(LoadFileAsync(ExcelJsonManager.ExcelSetting.strExcelPath, bytes =>
        {
            if (bytes != null && bytes.Length != 0)
            {
                ExcelDatas.Add(new ExcelData());

               var workbook = new WorkBook(bytes);
                //Debug.Log(book.Count);
                for (int i = 0; i < workbook.Count; i++)
                {
                    ExcelDatas[0].WorkBookDatas.Add(new WorkBookData());

                    GetRowData(workbook[i], i);
                    //Debug.Log("-------------------------");
                }
            }

            bExcelLoad = true;
        }
        ));
    }

    //Ū���C��M�C��
    void GetRowData(IEnumerable<Row> rows, int workbookid)
    {
        if (rows == null)
            return;

        var rowdata = rows.ToList();
        if (rowdata.Count == 0)
            return;

        if (ExcelDatas == null || ExcelDatas.Count == 0)
            return;

        var excel = ExcelDatas[0];
        if (excel.WorkBookDatas == null)
            return;

        if (workbookid < 0 || workbookid >= excel.WorkBookDatas.Count)
            return;

        var workbookdata = excel.WorkBookDatas[workbookid];
        if (workbookdata.RowDatas == null)
            workbookdata.RowDatas = new List<RowData>();

        for (int j = 0; j < rowdata.Count; j++) // ��
        {
            workbookdata.RowDatas.Add(new RowData());

            for (int i = 0; i < rowdata[j].Count; i++)//�C
            {
                string text = rowdata[j][i]?.Text ?? string.Empty;

                if (workbookdata.RowDatas[j].Datas == null)
                    workbookdata.RowDatas[j].Datas = new List<string>();

                ExcelDatas[0].WorkBookDatas[workbookid].RowDatas[j].Datas.Add(text);
            }
        }
    }


    //�D�P�B���J
    private IEnumerator LoadFileAsync(string path, DownloadHandler handler)
    {
        // streaming assets should be loaded via web request
        // on WebGL/Android platforms, this folder is in a compressed directory
        var url = Path.Combine(Application.streamingAssetsPath, path);
        using (var req = UnityWebRequest.Get(url))
        {
            yield return req.SendWebRequest();
            var bytes = req.downloadHandler.data;
            handler(bytes);
        }
    }

    [System.Serializable]
    public class ExcelData
    {
        public List<WorkBookData> WorkBookDatas = new List<WorkBookData>();
    }

    [System.Serializable]
    public class WorkBookData
    {
        public List<RowData> RowDatas = new List<RowData>();
    }

    [System.Serializable]
    public class RowData
    {
        public List<string> Datas;
    }
}