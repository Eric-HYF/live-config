using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using System;
using System.IO;

public class ExcelReader
{



    #region 静态方法


    /// <summary>
    /// 从文件系统加载Excel
    /// </summary>
    /// <param name="filePath"></param>
    /// <returns></returns>
    public static ExcelFile LoadFromFile(string filePath)
    {
        // TODO  



        return null;
    }


    /// <summary>
    /// 从网络加载Excel
    /// </summary>
    /// <param name="loadCompleteCallback"></param>
    public static void LoadFromNet(Action<ExcelFile> loadCompleteCallback)
    {

    }


    #endregion

}




/// <summary>
/// Excel 文件对象
/// </summary>
public class ExcelFile
{

    private Dictionary<string, WorkSheet> m_sheets;




    public ExcelFile()
    {

    }


    /// <summary>
    /// 读取文件流
    /// </summary>
    /// <param name="fileStream"></param>
    public void Read(FileStream fileStream)
    {

    }




}



/// <summary>
/// 工作表
/// </summary>
public class WorkSheet
{

    private string m_name;
    /// <summary>
    /// 工作表名
    /// </summary>
    public string Name { get { return m_name; } }

    private List<string> m_header;

    /// <summary>
    /// 表头
    /// </summary>
    public string[] Header { get { return m_header.ToArray(); } }


    /// <summary>
    /// 原始数据
    /// </summary>
    private List<List<string>> m_rawData;

    /// <summary>
    /// 元素列表
    /// </summary>
    private List<WorkSheetItem> m_items;


    /// <summary>
    /// 元素数量
    /// </summary>
    public int Count
    {
        get
        {
            if (m_items != null) return m_items.Count;
            return 0;
        }
    }



    public WorkSheet()
    {

    }

    /// <summary>
    /// 获取封装后的一行的元素对象
    /// </summary>
    /// <param name="index"></param>
    /// <returns></returns>
    public WorkSheetItem GetItem(int index)
    {
        if(m_items != null && index < m_items.Count)
        {
            return m_items[index];
        }
        return null;
    }

    /// <summary>
    /// 获取一行
    /// </summary>
    /// <param name="index"></param>
    /// <returns></returns>
    public string[] GetRow(int index)
    {
        if (m_rawData != null && index <  m_rawData.Count)
        {
            return m_rawData[index].ToArray();
        }
        return null;
    }


    /// <summary>
    /// 获取列标识
    /// </summary>
    /// <param name="key"></param>
    /// <returns></returns>
    public int GetColumIndex(string key)
    {
        if(m_header !=null && m_header.Contains(key))
        {
            return m_header.FindIndex(s => s.Equals(key));
        }
        return -1;
    }


    /// <summary>
    /// 获取列
    /// </summary>
    /// <param name="index"></param>
    /// <returns></returns>
    public string[] GetColum(int index)
    {
        if (m_header != null
            && m_rawData != null
            && index <m_header.Count)
        {

            List<string> col = new List<string>();
            for(int i = 0; i< m_rawData.Count; i++)
            {
                string val = m_rawData[i][index];
                col.Add(val);
            }
            return col.ToArray();
        }


        return null;
    }


    /// <summary>
    /// 获取列
    /// </summary>
    /// <param name="key">表头的Key</param>
    /// <returns></returns>
    public string[] GetColum(string key)
    {
        int index = GetColumIndex(key);
        if (index > -1) return GetColum(index);
        return null;
    }


    
    /// <summary>
    /// 按行列取值
    /// </summary>
    /// <param name="rowIndex"></param>
    /// <param name="colIndex"></param>
    /// <returns></returns>
    public string GetValue(int rowIndex, int colIndex)
    {
        if (rowIndex < Count && colIndex < m_header.Count)
        {
            return m_rawData[rowIndex][colIndex];
        }
        return "";
    }


    /// <summary>
    /// 通过键值取值
    /// </summary>
    /// <param name="key"></param>
    /// <param name="rowIndex"></param>
    /// <returns></returns>
    public string GetValue(string key, int rowIndex)
    {
        string[] col = GetColum(key);
        if(col != null && rowIndex < col.Length)
        {
            return col[rowIndex];
        }
        return "";
    }






}



public class WorkSheetItem 
{





}

