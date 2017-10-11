using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEditor;
using System.IO;
using XLSX;
public class ReadXlsxTest : MonoBehaviour {


    void OnGUI()
    {
        if (GUI.Button(new Rect(0, 0, 200, 100), "LoadXlsxOnLine"))
        {
            StartCoroutine(XLSXReader.LoadXlsxOnLine());
        }
        if (GUI.Button(new Rect(0, 100, 200, 100), "GetExcelSheetList"))
        {
            GetExcelSheetList();
        }
        if (GUI.Button(new Rect(0, 200, 200, 100), "GetHeadList"))
        {
            GetHeadList();
        }
        if (GUI.Button(new Rect(0, 300, 200, 100), "GetIDList"))
        {
            GetIDList();
        }
        if (GUI.Button(new Rect(0, 400, 200, 100), "GetRowColData"))
        {
            GetRowColData();
        }

    }

    static void GetRowColData()
    {
        Debug.Log("GetData :Data1, 10, data " + XLSXReader.GetRowColData("Data1", "10", "data"));
    }

    static void GetExcelSheetList()
    {
        List<string> temp = XLSXReader.GetExcelSheetList();
        for (int i = 0; i < temp.Count; i++)
        {
            Debug.Log(">>>GetExcelSheetList: [" + i + "]  " + temp[i]);
        }
    }

    static void GetHeadList()
    {
        List<string> temp = XLSXReader.GetHeadList("Data1");
        for (int i = 0; i < temp.Count; i++)
        {
            Debug.Log(">>>GetHeadList: [" + i + "]  " + temp[i]);
        }
    }

    static void GetIDList()
    {
        List<string> temp = XLSXReader.GetIDList("Data1");
        for (int i = 0; i < temp.Count; i++)
        {
            Debug.Log(">>>GetFirstRowList: [" + i + "]  " + temp[i]);
        }
    }
}
