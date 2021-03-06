﻿using System.Collections;
using System.Collections.Generic;
using System.IO;
using Excel;
using System.Data;
using UnityEngine;

namespace XLSX
{
    public class XLSXReader
    {    
         
        static Dictionary<string, List<List<string>>> m_excelData= new Dictionary<string, List<List<string>>>();

        //static Dictionary<string, Dictionary<string, List<string>>> XlsData=new Dictionary<string, Dictionary<string, List<string>>>();
      //  static List<string> HeadNameList = new List<string>();
        /// <summary>
        /// 获取所有的表名
        /// </summary>
        /// <returns></returns>
        public static List<string> GetExcelSheetList()
        {
            List<string> sheetList = new List<string>();
            foreach (var sheetdic in m_excelData)
            {
                sheetList.Add(sheetdic.Key);
            }
            return sheetList;
        }

        /// <summary>
        /// 获取第一行
        /// </summary>
        /// <returns></returns>
        public static List<string> GetHeadList(string sheetName)
        {
            List<List<string>> sheetData = m_excelData[sheetName];
           // List<string> idList = GetIDList(sheetName);
            List<string> headList = sheetData[0];
            return headList;
        }

        /// <summary>
        /// 获取第一列
        /// </summary>
        /// <returns></returns>
        public static List<string> GetIDList(string sheetName)
        {
            List<List<string>> sheetData = m_excelData[sheetName];
            List<string> idList = new List<string>();
            foreach (var sd in sheetData)
            {
                idList.Add(sd[0]);
            }
            return idList;
        }

        /// <summary>
        /// 获取某一行的数据
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="rowName"></param>
        /// <returns></returns>
        public static List<string> GetDataListByID(string sheetName,string idName) {
            try {
                List<List<string>> sheetData = m_excelData[sheetName];
                foreach (var sd in sheetData) {
                    if (sd[0].Equals(idName)) {
                        return sd;
                    }
                }
                return null;
            }
            catch (System.Exception e)
            {
                System.Console.WriteLine("sheetNameError or rowNameError");
                UnityEngine.Debug.Log("=== Error : sheetNameError or rowNameError" + "===");
                return null;
            }

        }
        /// <summary>
        /// 获取行列对应的数据值
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="rowName"></param>
        /// <param name="colName"></param>
        /// <returns></returns>
        public static string GetRowColData(string sheetName, string idName,string headName) {
            List<string> dataList = GetDataListByID(sheetName, idName);
            if (dataList != null)
            {
                List<string> headList = GetHeadList(sheetName);
                for (int i = 0; i < headList.Count; i++)
                {
                    if (headName.Equals(headList[i]))
                    {
                        return dataList[i];
                    }
                }
                return "";
            }
            return "";
           
        }


        public static void ReadXLSX(Stream stream, bool isFirstRowAsColumnNames) {
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            excelReader.IsFirstRowAsColumnNames = isFirstRowAsColumnNames;
            int index = 0;
            do
            {
                // 读取当前行

                // sheet name
                Log(">>> Table [{0}]", excelReader.Name);
                List<List<string>> sheetData = new List<List<string>>();
                // 读取列
                while (excelReader.Read())
                {
                    Log("-------------------------------");
                    Log("Count[{0}]", excelReader.FieldCount);


                    List<string> data = new List<string>();

                    for (int i = 0; i < excelReader.FieldCount; i++)
                    {
                        string value = excelReader.IsDBNull(i) ? "" : excelReader.GetString(i);
                        Log(value);
                        data.Add(value);
                        //if (index == 0)
                        //{
                        //    HeadNameList.Add(value);
                        //}
                    }
                    sheetData.Add(data);
                }
                m_excelData.Add(excelReader.Name, sheetData);
                index++;
            } while (excelReader.NextResult());
        }

   
        public static IEnumerator LoadXlsxOnLine()
        {
            WWW www = new WWW("http://localhost:81/Data.xlsx");
            yield return www;
            if (www.error != null)
            {
                Debug.Log(www.error);
                yield return null;
            }
            if (www.isDone)
            {
                byte[] bytes = www.bytes;
                Stream stream = new MemoryStream(bytes);
                ReadXLSX(stream, true);
            }
        }



        /// <summary>
        /// 读取EXCEL文件
        /// </summary>
        /// <param name="path"></param>
        public static ExcelFile ReadXLSX(string path, bool isFirstRowAsColumnNames = true)
        {
            Log("=== Read XLSX from : " + path + "===");

            ExcelFile file = null;


            try
            {
                FileStream stream = File.Open(path, FileMode.OpenOrCreate, FileAccess.ReadWrite);

                Log(">>>> stream: " + stream);


                ReadXLSX(stream, isFirstRowAsColumnNames);

               // DataSet result = excelReader.AsDataSet();

                //Log(">>>> result: " + result);

                //file = new ExcelFile(result.Tables);

                //System.Console.WriteLine("Get Excel File, Sheet Count: " + file.SheetCount);
            }
            catch (System.Exception e)
            {
                System.Console.WriteLine("" + e.Message);
                UnityEngine.Debug.Log("=== Error : " + e.Message + "===");
            }


            return file;
        }









        public static void Log(string format, params object[] args)
        {
            UnityEngine.Debug.LogFormat(format, args);
        }





    }





    public class ExcelFile
    {
        private Dictionary<string, WorkSheet> m_sheets;



        /// <summary>
        /// Excel文件
        /// </summary>
        /// <param name="tables"></param>
        public ExcelFile(DataTableCollection tables)
        {

            m_sheets = new Dictionary<string, WorkSheet>();

            int i = 0;
            int count = tables.Count;

            XLSXReader.Log(">>>[1] Tabel Count: " + tables.Count);

            while (i < count)
            {
                WorkSheet sheet = new WorkSheet(tables[i], i);
                m_sheets[sheet.name] = sheet;
                i++;
            }
        }

        /// <summary>
        /// 通过名称获取
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public WorkSheet GetSheet(string name)
        {
            if (m_sheets.ContainsKey(name))
            {
                return m_sheets[name];
            }
            return null;
        }

        /// <summary>
        /// 通过序号获取
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public WorkSheet GetSheet(int index)
        {
            foreach (var s in m_sheets.Values)
            {
                if (s.Index == index) return s;
            }
            return null;
        }



        /// <summary>
        /// 获取表内的数据
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="row"></param>
        /// <param name="cloum"></param>
        /// <returns></returns>
        public string GetItemFromSheet(string sheetName, int row, int cloum)
        {
            WorkSheet sheet = GetSheet(sheetName);
            if (sheet != null)
            {
                return GetItemFromSheet(sheet, row, cloum);
            }
            return null;
        }

        /// <summary>
        /// 获取表内的数据
        /// </summary>
        /// <param name="index"></param>
        /// <param name="row"></param>
        /// <param name="cloum"></param>
        /// <returns></returns>
        public string GetItemFromSheet(int index, int row, int cloum)
        {
            WorkSheet sheet = GetSheet(index);
            if (sheet != null)
            {
                return GetItemFromSheet(sheet, row, cloum);
            }
            return null;
        }



        public string GetItemFromSheet(WorkSheet sheet, int row, int cloum)
        {
            return sheet.GetItem(row, cloum);
        }

        /// <summary>
        /// 图标数量
        /// </summary>
        public int SheetCount
        {
            get { return m_sheets.Count; }
        }


        public string[] GetAllSheetNames()
        {
            List<string> keys = new List<string>();
            foreach (var kvp in m_sheets)
            {
                keys.Add(kvp.Key);
            }

            return keys.ToArray();
        }





    }








    /// <summary>
    /// 工作表
    /// </summary>
    public class WorkSheet
    {
        public string name;
        private int m_index;
        public int Index { get { return m_index; } set { m_index = value; } }
        private DataTable m_table;

        private List<string> m_headers;

        Dictionary<string, List<string>> m_data;



        public WorkSheet(DataTable table, int index = 0)
        {
            m_index = index;
            m_table = table;
            name = table.TableName;

            ParseDataTable(table);
        }



        public WorkSheet()
        {

        }




        void ParseDataTable(DataTable table)
        {
            m_headers = new List<string>();
            m_data = new Dictionary<string, List<string>>();


            int rowCount = table.Rows.Count;
            int colCount = table.Columns.Count;

            if (table.Rows.Count > 0)
            {
                // Make Header
                object[] objs = table.Rows[0].ItemArray;

                for (int i = 0; i < objs.Length; i++)
                {
                    string val = objs[i].ToString();
                    m_headers.Add(val);
                }
            }



        }



        /// <summary>
        /// 获取表内的元素
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="cloum">列</param>
        /// <returns></returns>
        public string GetItem(int row, int cloum)
        {
            return m_table.Rows[row][cloum].ToString();
        }












    }
}
