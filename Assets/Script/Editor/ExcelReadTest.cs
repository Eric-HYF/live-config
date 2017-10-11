using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEditor;
using System.IO;
using XLSX;


public class ExcelReadTest 
{


    const string TEST_DATA_PATH = "File/Data.xlsx";

    [MenuItem("Tools/Read Xlsx Assets")]
    static void ReadTestExcel()
    {
        string path = Path.GetFullPath( Path.Combine(Application.dataPath , TEST_DATA_PATH));
        Debug.Log(path);


        if(File.Exists(path))
        {
            Debug.Log("Excel is existed");

            XLSX.ExcelFile ef = XLSXReader.ReadXLSX(path);

            if(ef != null)
            {
                Debug.Log(">>> Find Excel File: "+ ef);
                Debug.Log(">>> Sheet Count: " + ef.SheetCount);
            }
            else
            {
                Debug.Log("<color=red>Read Excel from "+path+" Failed!</color>");
            }
        }



    }

}
