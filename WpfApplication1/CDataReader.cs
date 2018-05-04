using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
//using Microsoft.Office.Interop.Excel;
using Excel=  Microsoft.Office.Interop.Excel;

namespace WpfApplication1
{
    class CDataReader
    {
        String m_filename;
        List<CDataHKNJH> m_data{ get; set; }
        public  Boolean ReadDataFromFile()
        {

             //创建Application
            //Excel._Application spExcelApp = new Excel.ApplicationClass();





        
             return true;
        }
    }

    struct CDataHKNJH
    {
       //设备名称
        String m_DevName;
       //设备数量
        Int32 m_DevCount;
       //数量单位
        String m_DevUnit;
       //上次检修时间
        String m_LastCheckTime;
       //检修周期
        String m_CheckPeroid;
       //计划时间
        //Int32[] m_DevCount=new Int32[12];
       //工作地点
        String m_WorkSite;
       //
        String m_PlanTime;
    }



#region 操作Excle应用程序类
    /*
     //调用例子:
     Excel._Application   spExcelApp=new Excel.ApplicationClass();

     //模板文件
     Excel._Workbook spMouleBook = CExcelApp.OpenExcel(spExcelApp,"c:/Empty.xls");
     // Excel._Worksheet spMouleSheet=CExcelApp.GetSheetItem(spMouleBook,1);
    
     //输出文件
     Excel._Workbook  spOutBook =CExcelApp.AddSheetPage(spExcelApp);
     Excel._Worksheet spOutSheet =CExcelApp.GetSheetItem(spOutBook,1);
     Excel._Worksheet spOutSheet1 =CExcelApp.GetSheetItem(spOutBook,2);

     spOutSheet.Name="SheetName";

            
     //更新所有列宽
     //CExcelApp.SetColumnWidth(spMouleSheet,1,spOutSheet,1,11);

     //整行拷贝
     //Excel.Range spSou = CExcelApp.GetRangeByRow(spMouleSheet,1,1000);

     //Excel.Range spDes =CExcelApp.GetRangeByRow(spOutSheet,1);
     //CExcelApp.CopyRange(spSou,spDes);
     object spRgBeg = CExcelApp.GetRange(spOutSheet,1,1);
     object spRgEnd = CExcelApp.GetRange(spOutSheet,60000,18);
     spOutSheet.get_Range(spRgBeg,spRgEnd).set_Value(System.Reflection.Missing.Value,CExcelApp.GetRangeDataArr(GetJsTable()));

     object spRgBeg1 = CExcelApp.GetRange(spOutSheet1,1,1);
     object spRgEnd1 = CExcelApp.GetRange(spOutSheet1,60000,18);

     spOutSheet1.get_Range(spRgBeg1,spRgEnd1).set_Value(System.Reflection.Missing.Value,CExcelApp.GetRangeDataArr(GetJsTable()));

     //CExcelApp.Show(spExcelApp,true);  //是否显示Excel
     CExcelApp.SaveAs(spOutBook,"d:/z.xls");

     CExcelApp.CloseBook(spMouleBook,false);
     CExcelApp.CloseBook(spOutBook,false);

     CExcelApp.ReleaseSheet(spOutSheet);
     CExcelApp.ReleaseSheet(spOutSheet1);
     CExcelApp.ReleaseBook(spOutBook);
     CExcelApp.ReleaseBook(spMouleBook);
     CExcelApp.ReleaseExcelApp(spExcelApp);
     GC.Collect();
 */


    /// <summary>
    /// CExcelComm 的摘要说明。
    /// </summary>
    public class CExcelApp
    {
        /// <summary>
        /// 创建Excel相关对象
        /// </summary>
        static public Excel._Workbook OpenExcel(Excel._Application spExcelApp, string strFilePath)
        {
            try
            {
                if (strFilePath == "")
                    return null;
                Object vtMissing = System.Reflection.Missing.Value;
                if (spExcelApp == null)
                {
                    return null;
                }
                Excel._Workbook m_spWorkBook = spExcelApp.Workbooks.Open(strFilePath, vtMissing, vtMissing, vtMissing,
                    vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
                    vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, 0);

                return m_spWorkBook;

            }
            catch (Exception ex)
            {
                string strError = ex.ToString().Trim();
                return null;
            }
        }


        /// <summary>
        /// 是否显示Excel
        /// </summary>
        static public void Show(Excel._Application pApp, bool bShow)
        {
            if (pApp == null) return;
            pApp.Visible = bShow ? true : false;
        }


        /// <summary>
        /// 根据Sheet的Index得到WorkSheet对象
        /// </summary>
        static public Excel._Worksheet GetSheetItem(Excel._Workbook spWorkBook, int iIndex)
        {
            try
            {
                if (spWorkBook == null)
                {
                    return null;
                }
                return (Excel._Worksheet)spWorkBook.Sheets.get_Item(iIndex);
            }

            catch (Exception ex)
            {
                string strError = ex.ToString().Trim();
                return null;
            }

        }

        /// <summary>
        /// 得到打开Excel模板的Sheet个数
        /// </summary>
        static public int GetSheetCount(Excel._Workbook spWorkBook)
        {
            try
            {
                if (spWorkBook == null)
                {
                    return -1;
                }
                return spWorkBook.Sheets.Count;
            }

            catch (Exception ex)
            {
                string strError = ex.ToString().Trim();
                return -1;
            }
        }

        /// <summary>
        ///    从页中获得对应值
        /// </summary>
        static public bool GetValueFromSheet(Excel._Worksheet pSheet, long nRow, long nCol, ref string sValue)
        {
            return CExcelApp.GetCellValue(pSheet, nRow, nCol, ref sValue);
            //return true;
        }

        /// <summary>
        ///    设置值到指定页上
        /// </summary>
        static public bool SetValueToSheet(Excel._Worksheet pSheet, long nRow, long nCol, double dValue)
        {
            Excel.Range spRange = CExcelApp.GetRange(pSheet, nRow, nCol);
            if (spRange == null)
                return false;

            return CExcelApp.SetCellValue(spRange, dValue);
        }



        /// <summary>
        /// 根据Sheet和ExRange,获得指定项的值
        /// </summary>
        static public bool GetCellValue(Excel._Worksheet pSheet,   //指定页
            long nRow,                 //指定项
            long nCol,
            ref string strValue         //返回值
            )
        {
            Excel.Range spRange = GetRange(pSheet, nRow, nCol);
            if (spRange == null)
                return false;
            return GetCellValue(spRange, ref strValue);
        }

        /// <summary>
        /// 获得指定Range项得到相应的值
        /// </summary>
        static public bool GetCellValue(Excel.Range pRange,     //指定项
            ref string strValue       //返回值
            )
        {
            Object[,] saRet;
            string strRet = "";
            saRet = (System.Object[,])pRange.get_Value(System.Reflection.Missing.Value);
            try
            {
                long iRows;
                long iCols;
                iRows = saRet.GetUpperBound(0);
                iCols = saRet.GetUpperBound(1);

                for (long row = 1; row <= iRows; row++)
                {
                    for (long col = 1; col <= iCols; col++)
                    {
                        if (saRet[row, col] != "")
                            strRet += saRet[row, col] + ",";
                    }
                }
                string[] strRes = strRet.Split(',');
                for (int i = 0; i < strRes.Length - 1; i++)
                {
                    if (strRes[i] != "")
                        strValue += strRes[i] + ",";
                }
            }
            catch (Exception ex)
            {
                ex.ToString().Trim();
            }
            return true;
        }


        /// <summary>
        ///  根据Excle页和Range和值设置Cell字符串
        /// </summary>
        static public bool SetCellValue(Excel._Worksheet pSheet,        //指定页
            long nRow,       //行
            long nCol,         //列
            string szText        //更新值，为空时设置项目为空
            )
        {
            if (pSheet == null || Invalid(nRow, nCol))
                return false;

            Excel.Range spRange = (Excel.Range)pSheet.Cells.get_Item(nRow, nCol);
            if (spRange != null)
                return SetCellValue(spRange, szText);
            return false;
        }



        /// <summary>
        /// 根据Range和值设置Cell字符串
        /// </summary>
        static public bool SetCellValue(Excel.Range pRange,     //指定项
            string szText              //更新值，为空时设置项目为空
            )
        {
            if (szText != null)
                pRange.set_Value(System.Reflection.Missing.Value, szText);

            return true;
        }


        //设置指定范围的值
        /// <summary>
        /// 根据指定的页的起始项和结束项字符串
        /// </summary>
        static public bool SetRangeValue(Excel._Worksheet pSheet,        //指定页  //指定起始项，包括该项                                                          
            long nBegRow,
            long nBegCol,
            long nEndRow,
            long nEndCol,
            string szText       //更新值，为空时设置项目为空
            )
        {
            if (pSheet == null)
                return false;

            if (Invalid(nBegRow, nBegCol) || Invalid(nEndRow, nEndCol))
                return false;

            for (long i = nBegRow; i <= nEndRow; i++)
            {
                for (long j = nBegCol; j <= nEndCol; j++)
                {
                    SetCellValue(pSheet, i, j, szText);
                }
            }

            return true;
        }


        /// <summary>
        /// 根据Range和值设置Cell double值
        /// </summary>
        static public bool SetCellValue(Excel.Range pRange,             //指定项
            double dValue                    //更新值，为空时设置项目为空
            )
        {
            pRange.set_Value(System.Reflection.Missing.Value, dValue);
            return true;
        }


        /// <summary>
        /// 根据指定起始和终止范围，获得 Range 对象
        /// </summary>
        static public Excel.Range GetRange(Excel._Worksheet pSheet,     //指定页
            long nBegRow, //指定起始项，包括该项
            long nBegCol,
            long nEndRow,//指定结束项，包括该项
            long nEndCol
            )
        {
            if (pSheet == null || Invalid(nBegRow, nBegCol) || Invalid(nEndRow, nEndCol))
                return null;

            object spRgBeg = GetRange(pSheet, nBegRow, nBegCol);
            if (spRgBeg == null) return null;
            object spRgEnd = GetRange(pSheet, nEndRow, nEndCol);
            if (spRgEnd == null) return null;
            return pSheet.get_Range(spRgBeg, spRgEnd);
        }



        /// <summary>
        /// 根据指定行，获得 Range 对象
        /// </summary>
        static public Excel.Range GetRangeByRow(Excel._Worksheet pSheet,      //指定的页
            long nRow                       //指定起始行
            )
        {
            if (nRow < 1 || nRow > 0xFFFF)
                return null;

            return (Excel.Range)pSheet.Rows.get_Item(nRow, System.Reflection.Missing.Value);
        }

        /// <summary>
        /// 根据指定页和行，获得 Range 对象
        /// </summary>
        static public Excel.Range GetRangeByRow(Excel._Worksheet pSheet,      //指定的页
            long nRowBeg,                  //指定起始项，包括该项
            long nRowEnd                  //指定结束项，包括该项
            )
        {
            object spRgBeg = GetRangeByRow(pSheet, nRowBeg);
            object spRgEnd = GetRangeByRow(pSheet, nRowEnd);
            return pSheet.get_Range(spRgBeg, spRgEnd);
        }

        /// <summary>
        /// 根据指定列，获得 Range 对象
        /// </summary>
        static public Excel.Range GetRangeByCol(Excel._Worksheet pSheet,      //指定的页
            long nCol                    //指定起始列
            )
        {
            if (nCol < 1 || nCol > 0xFF)
                return null;

            return (Excel.Range)pSheet.Columns.get_Item(nCol, System.Reflection.Missing.Value);
        }

        /// <summary>
        /// 指定行及结束行
        /// </summary>

        static public Excel.Range GetRangeByCol(Excel._Worksheet pSheet,       //指定的页
            long nColBeg,                       //指定起始项，包括该项
            long nColEnd                       //指定结束项，包括该项
            )
        {
            object spRgBeg = GetRangeByCol(pSheet, nColBeg);
            object spRgEnd = GetRangeByCol(pSheet, nColEnd);

            return pSheet.get_Range(spRgBeg, spRgEnd);
        }


        /// <summary>
        /// 根据指定页和Cell,获得 Range 对象
        /// </summary>
        static public Excel.Range GetRange(Excel._Worksheet pSheet,
            long nRow, long nCol
            )
        {
            if (pSheet != null && Invalid(nRow, nCol) == false)
                return (Excel.Range)pSheet.Cells.get_Item(nRow, nCol);
            return null;
        }


        /// <summary>
        /// 拷贝范围的项到目标项
        /// </summary>
        static public bool CopyRange(Excel.Range pSou,        //数据源
            Excel.Range pDes        //目标
            )
        {

            Excel.Range spDes = pDes;
            object spDis = spDes;
            pSou.Copy(spDis);
            return true;

        }



        /// <summary>
        /// 设置指定列的宽度
        /// </summary>
        static public bool SetColumnWidth(Excel.Range pCol, ref double dWidth)
        {
            dWidth = (double)pCol.ColumnWidth;
            return true;
        }


        /// <summary>
        /// 更新指定列的宽度
        /// </summary>
        static public bool SetColumnWidth(Excel.Range pRange, double dWidth)
        {
            pRange.ColumnWidth = dWidth;
            return true;
        }


        /// <summary>
        /// 设置数据源和目标源的宽度相同
        /// </summary>
        static public bool SetColumnWidth(Excel.Range pSou, Excel.Range pDes)
        {
            pDes.ColumnWidth = pSou.ColumnWidth;
            return true;
        }

        /// <summary>
        /// 批量更新指定范围内的列与数据源同范围内的相同
        /// </summary>
        static public bool SetColumnWidth(Excel._Worksheet pSheetSou,    //数据源页
            long nColBegSou,                    //数据源的起始列，包括该列
            Excel._Worksheet pSheetDes,    //目标源页
            long nColBegDes,                    //目标源的起始列，包括该列
            long nCount                        //需要更新的列数
            )
        {
            Excel.Range spSou, spDes;
            long nSou = nColBegSou;
            long nDes = nColBegDes;
            for (long i = 0; i < nCount; i++, nSou++, nDes++)
            {
                spSou = GetRangeByCol(pSheetSou, nSou);
                spDes = GetRangeByCol(pSheetDes, nDes);
                SetColumnWidth(spSou, spDes);
            }
            return true;
        }


        /// <summary>
        /// 根据App对象,添加一个Excel Sheet对象
        /// </summary>
        static public Excel._Workbook AddSheetPage(Excel._Application spExcelApp)
        {
            if (spExcelApp == null) return null;
            return spExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
        }



        /// <summary>
        /// 核心方法,将一个dt转换成一个数组对象
        /// </summary>
        static public string[,] GetRangeDataArr(DataTable dt)
        {
            if (dt == null) return null;
            int iRow = dt.Rows.Count;
            if (iRow <= 0) return null;
            int iColumn = dt.Columns.Count;
            if (iColumn <= 0) return null;

            string[,] arrData = new string[iRow, iColumn];   //[row,col]
            for (int j = 0; j < iColumn; j++)
            {
                for (int i = 0; i < iRow; i++)
                {
                    arrData[i, j] = dt.Rows[i][j].ToString().Trim();

                }

            }
            return arrData;
        }




        /// <summary>
        /// 关闭指定薄
        /// </summary>
        static public bool SaveAs(Excel._Workbook pBook,
            string strFilePath
            )
        {
            pBook.SaveCopyAs(strFilePath);
            return true;
        }


        /// <summary>
        /// 关闭指定薄
        /// </summary>
        static public bool CloseBook(Excel._Workbook pBook,
            bool bSave
            )
        {
            bool bS = bSave ? true : false;
            pBook.Close(bS, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            return true;
        }

        /// <summary>
        /// 释放Excel Book对象
        /// </summary>
        static public bool ReleaseBook(Excel._Workbook pBook)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject((object)pBook);
            pBook = null;
            return true;
        }

        /// <summary>
        /// 释放Excel Sheet对象
        /// </summary>
        static public bool ReleaseSheet(Excel._Worksheet pWorkSheet)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject((object)pWorkSheet);
            pWorkSheet = null;
            return true;
        }


        /// <summary>
        /// 释放Excel App对象
        /// </summary>
        static public bool ReleaseExcelApp(Excel._Application pWorkApp)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject((object)pWorkApp);
            pWorkApp = null;
            return true;
        }


        //*****************************************************
        /// <summary>
        /// 是否无效 列范围(1,255) 行范围(1,~)
        /// </summary>
        static public bool Invalid(long nRow, long nCol)
        {
            return (nCol < 1 || nCol > 255) || (nRow < 1 || nRow > 65536);
        }

        /// <summary>
        /// 计算页面大小
        /// </summary>
        public long GetPageSize(long nEndRow, long nBegRow)
        {
            return nEndRow - nBegRow + 1;
        }



        /// <summary>
        /// 根据页面号，页面长度，偏移行号，返回新的坐标
        /// </summary>
        public bool MovePage(long nRow, long nPageSize, long nPageNo)
        {
            nRow += nPageSize * (nPageNo - 1);
            return true;
        }

    #endregion
    }

}
