using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Data;
//using Newtonsoft.Json;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;
using DataTable = System.Data.DataTable;

namespace AIC_Annual_Report
{
    /// <summary>
    /// Excel帮助类
    /// <para>用法1：使用using</para> 
    /// <para>用法2：不使用using，使用结束后显式调用Dispose()</para> 
    /// </summary>
    public class ExcelHelper : IDisposable
    {
        private Application appExcel = new Application();
        private Workbooks wbs = null;
        private Workbook wb = null;
        private Worksheet ws = null;

        private string filePath = "";
        private bool visible = false;
        private bool readOnly = false;
        private bool displayAlerts = false;
        private bool isSave = false;

        #region 构造函数
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="pFilePath"></param>
        /// <param name="pVisible"></param>
        /// <param name="pReadOnly"></param>
        /// <param name="pDisplayAlerts"></param>
        public ExcelHelper(string pFilePath, bool pVisible = false, bool pReadOnly = false, bool pIsSave = false, bool pDisplayAlerts = false)
        {
            filePath = pFilePath;
            visible = pVisible;
            displayAlerts = pDisplayAlerts;
            readOnly = pReadOnly;
            isSave = pIsSave;

            InitExcel();
        }

        public ExcelHelper()
        {
        }

        private void InitExcel()
        {
            appExcel.Visible = visible;
            appExcel.DisplayAlerts = displayAlerts;
            wbs = appExcel.Workbooks;
            if (File.Exists(filePath))
                wb = wbs.Open(filePath, ReadOnly: readOnly);
            else
                wb = wbs.Add();
        }
        #endregion

        #region 方法继承
        /// <summary>
        /// 实现IDisposable中的Dispose方法
        /// </summary>
        public void Dispose()
        {
            if (isSave)
            {
                Save();
            }

            if (wb != null)
            {
                wb.Close();
                Marshal.FinalReleaseComObject(wb);
            }

            if (ws != null)
                Marshal.FinalReleaseComObject(ws);

            if (wbs != null)
                Marshal.FinalReleaseComObject(wbs);

            if (appExcel != null)
            {
                appExcel.Quit();
                Marshal.FinalReleaseComObject(appExcel);
            }
        }
        #endregion

        #region 方法实现

        #region 替换方法
        /// <summary>
        /// 替换标签值
        /// </summary>
        /// <param name="pDicData"></param>
        /// <param name="pSheetName"></param>
        public void ReplaceData(Dictionary<string, string> pDicData, string pSheetName, int pRow = 0)
        {
            if (string.IsNullOrEmpty(pSheetName))
                return;

            List<string> sheetNameList = new List<string>() { pSheetName };
            ReplaceData(pDicData, sheetNameList, pRow);
        }

        /// <summary>
        /// 替换标签值
        /// </summary>
        /// <param name="FilePath"></param>
        /// <param name="DicData"></param>
        /// <param name="SheetNameList"></param>
        /// 


        public void ReplaceData(Dictionary<string, string> pDicData, List<string> pSheetNameList, int pRow = 0)
        {
            if (pSheetNameList == null || pSheetNameList.Count == 0)
                return;

            foreach (string sheetName in pSheetNameList)
            {
                ws = wb.Worksheets[sheetName];
                Range ReplaceArea = null;
                if (pRow == 0)
                    ReplaceArea = ws.Cells;
                else
                {
                    Range startRange = ws.Cells[1, 1];
                    Range endRange = ws.Cells[1, ws.UsedRange.SpecialCells(XlCellType.xlCellTypeLastCell).Column];
                    ReplaceArea = ws.Range[startRange, endRange];
                    ReplaceArea.Select();
                    Marshal.FinalReleaseComObject(startRange);
                    Marshal.FinalReleaseComObject(endRange);
                }

                foreach (var data in pDicData)
                {
                    object tag = data.Key;
                    object value = data.Value;
                    ReplaceArea.Replace(tag, value, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
                Marshal.FinalReleaseComObject(ReplaceArea);
                Marshal.FinalReleaseComObject(ws);
            }
        }

        #endregion

        #region 删除方法
        /// <summary>
        /// 清空单元格内容
        /// </summary>
        /// <param name="DelDic"></param>
        /// 
        public Worksheet CurrentSht(string pSheetName)
        {
            
            ws = wb.Worksheets[pSheetName];
            return ws;
        }
        public Boolean isExistsSht(string pSheetName)
        {
           Boolean blisExistsSht = false;


            try
            {
                ws = wb.Worksheets[pSheetName];
                blisExistsSht = true;
                return blisExistsSht;
            }
            catch
            {
                return blisExistsSht;

            }
            
        }

        public Boolean Exists(string strFindShtName)
        {
            Boolean found = false;
            foreach (Worksheet sheet in wb.Worksheets)
            {
                // Check the name of the current sheet
                if (sheet.Name == strFindShtName)
                {
                    found = true;
                    break; // Exit the loop now
                }
            }



            return found;
        }
        public Workbook Currentwb()
        {

            return wb;
        }
        public void ChangeShtName(string NewShtName, string OldShtName = "Sheet1")
        {
            ws = wb.Sheets[OldShtName];
            ws.Name = NewShtName;
        }

        public void SetColFormat(string ShtName, int intCol)
        {
            ws = wb.Sheets[ShtName];
            Range Rng = ws.Columns[intCol];
            Rng.NumberFormat = "@";


        }

        public void DeleteData(string pSheetName, int pRowNo, int pCellNo)
        {
            ws = wb.Sheets[pSheetName];
            Range rngCell = ws.Cells[pRowNo, pCellNo];
            rngCell.Clear();
            Marshal.FinalReleaseComObject(rngCell);
        }
        public void DeleteSpecifieSheet(string StrRegex)
        {

            foreach (Worksheet sht in wb.Sheets)

            {
                System.Text.RegularExpressions.Match m = System.Text.RegularExpressions.Regex.Match(sht.Name, StrRegex);//\.*Linksave\.*Templat

                if (m.Success)
                {
                    sht.Delete();
                }

            }



        }
        /// <summary>
        /// 选择性清空工作表内容
        /// </summary>
        /// <param name="FilePath"></param>
        /// <param name="DelDic"></param>

        public void Excel_FilterByOpenFile_Delete_Row(Worksheet oSht, string strHeaderText1, string strHeaderText2, Dictionary<string, string> dicVendorNbr)
        {
            oSht.Activate();
            Range xlRange = oSht.Range["A1:" + GetColumnName(oSht.UsedRange.SpecialCells(XlCellType.xlCellTypeLastCell).Column.ToString()) + oSht.UsedRange.SpecialCells(XlCellType.xlCellTypeLastCell).Row];
            var Data = xlRange.Value;

            //For this function title row should be row no "1"
            int i = 1;
            int intFilterClmn1 = 0;
            bool blnHeadingFound = false;
            for (int j = 1; j <= xlRange.SpecialCells(XlCellType.xlCellTypeLastCell).Column; j++)
            {
                if (Data[i, j] != null && Data[i, j].ToString() == strHeaderText1)
                {
                    intFilterClmn1 = j;
                    blnHeadingFound = true;
                    break;
                }
            }

            if (blnHeadingFound == false)
            {
                throw new Exception("Heading:" + strHeaderText1 + " could not found in the 1st row");
            }
            i = 1;
            int intFilterClmn2 = 0;
            blnHeadingFound = false;
            for (int j = 1; j <= xlRange.SpecialCells(XlCellType.xlCellTypeLastCell).Column; j++)
            {
                if (Data[i, j] != null && Data[i, j].ToString() == strHeaderText2)
                {
                    intFilterClmn2 = j;
                    blnHeadingFound = true;
                    break;
                }
            }

            if (blnHeadingFound == false)
            {
                throw new Exception("Heading:" + strHeaderText2 + " could not found in the 1st row");
            }



            Range rangeRow = null;
            for (i = xlRange.SpecialCells(XlCellType.xlCellTypeLastCell).Row; i > 1; i--)
            {


                if (Data[i, intFilterClmn1] != null)
                {
                    string key = Convert.ToString(Data[i, intFilterClmn1])+"<*>"+ Convert.ToString(Data[i, intFilterClmn2]);

                    if (dicVendorNbr.ContainsKey(key.Trim()) == false)

                    {
                        rangeRow = xlRange.Rows[i];
                        rangeRow.Delete(XlDeleteShiftDirection.xlShiftUp);
                        //Filter
                        //String[] FilterList = { "*" + key.Trim() + "*" };
                        //xlRange.AutoFilter(intFilterClmn, FilterList, XlAutoFilterOperator.xlFilterValues);
                        ////Delete the values
                        //xlRange.Offset[1, 0].SpecialCells(XlCellType.xlCellTypeVisible).EntireRow.Delete();
                        ////Turn off the filter
                        //oSht.AutoFilterMode = false;

                    }
                }
            }





        }

        public static string GetColumnName(string strColumnNumber)
        {
            int index = int.Parse(strColumnNumber);
            const string letters = "ZABCDEFGHIJKLMNOPQRSTUVWXY";

            int NextPos = (index / 26);
            int LastPos = (index % 26);
            if (LastPos == 0) NextPos--;

            if (index > 26)
                return GetColumnName(NextPos.ToString()) + letters[LastPos];
            else
                return letters[LastPos] + "";
        }

        public void DeleteData(string FilePath, Dictionary<string, string> DelDic)
        {
            Range range = null;
            foreach (var Dic in DelDic)
            {
                ws = wb.Worksheets[Dic.Key];
                ws.Select();

                //清除筛选条件
                if (ws.AutoFilter != null && ws.AutoFilter.FilterMode)
                    ws.AutoFilterMode = false;

                int BeginRow = int.Parse(Dic.Value.Split(',')[0]);
                int EndColumn = int.Parse(Dic.Value.Split(',')[1]);
                Range startRange = ws.Cells[BeginRow, 1];
                Range endRange = ws.Cells[ws.UsedRange.SpecialCells(XlCellType.xlCellTypeLastCell).Row, EndColumn];
                range = ws.Range[startRange, endRange];
                range.Select();
                range.ClearContents();

                Marshal.FinalReleaseComObject(startRange);
                Marshal.FinalReleaseComObject(endRange);
                Marshal.FinalReleaseComObject(range);
            }
        }

        /// <summary>
        /// 清空工作表内容
        /// </summary>
        /// <param name="DelDic"></param>
        public void DeleteAllData(List<string> pSheetNames)
        {
            foreach (var sheetName in pSheetNames)
            {
                ws = wb.Worksheets[sheetName];
                ws.Select();
                ws.Cells.Select();
                ws.Cells.ClearContents();
                Marshal.FinalReleaseComObject(ws);
            }
        }

        /// <summary>
        /// 删除列
        /// </summary>
        /// <param name="pSheetName"></param>
        /// <param name="pColumns"></param>
        public void DropColumns(string pSheetName, List<int> pColumns)
        {
            ws = wb.Sheets[pSheetName];
            ws.Select();

            pColumns = pColumns.OrderByDescending(m => m).ToList();
            foreach (int columnNo in pColumns)
            {
                Range rngColumn = ws.Columns[columnNo];
                rngColumn.Delete(XlDeleteShiftDirection.xlShiftToLeft);
                Marshal.FinalReleaseComObject(rngColumn);
            }

            Marshal.FinalReleaseComObject(ws);
        }

        /// <summary>
        /// 删除最后一行
        /// </summary>
        /// <param name="pSheetName"></param>
        public void DropLastRow(string pSheetName)
        {
            ws = wb.Sheets[pSheetName];
            ws.Select();

            int rowNum = ws.UsedRange.CurrentRegion.Rows.Count;
            Range rngRow = ws.Rows[rowNum];
            rngRow.Delete(XlDeleteShiftDirection.xlShiftUp);
            Marshal.FinalReleaseComObject(rngRow);
            Marshal.FinalReleaseComObject(ws);
        }
        /// <summary>
        /// 删除行
        /// </summary>
        /// <param name="pSheetName"></param>
        /// <param name="pColumns"></param>
        public void DropRows(string pSheetName, List<int> pRows)
        {
            ws = wb.Sheets[pSheetName];
            ws.Select();

            pRows = pRows.OrderByDescending(m => m).ToList();
            foreach (int RowNo in pRows)
            {
                Range rngRow = ws.Rows[RowNo];
                rngRow.Delete(XlDeleteShiftDirection.xlShiftUp);
                Marshal.FinalReleaseComObject(rngRow);
            }

            Marshal.FinalReleaseComObject(ws);
        }
        #endregion

        #region 设置工作表是否可见
        /// <summary>
        /// 设置工作表是否可见
        /// </summary>
        /// <param name="pSheetName"></param>
        /// <param name="pVisible">1显示 0隐藏</param>
        public void VisibleSheet(string pSheetName, bool pVisible)
        {
            if (string.IsNullOrEmpty(pSheetName))
                return;

            ws = wb.Worksheets[pSheetName];
            ws.Visible = pVisible ? XlSheetVisibility.xlSheetVisible : XlSheetVisibility.xlSheetHidden;
        }

        /// <summary>
        /// 设置工作表是否可见
        /// </summary>
        /// <param name="pShowSheets"></param>
        /// <param name="pVisible">1显示 0隐藏</param>
        public void VisibleSheet(List<string> pShowSheets, bool pVisible)
        {
            if (pShowSheets == null || pShowSheets.Count == 0)
                return;

            XlSheetVisibility visibility = pVisible ? XlSheetVisibility.xlSheetVisible : XlSheetVisibility.xlSheetHidden;
            foreach (string item in pShowSheets)
            {
                ws = wb.Worksheets[item];
                ws.Visible = visibility;
            }
        }
        #endregion

        #region 写入数据

        /// <summary>
        /// 将单个值写入工作表
        /// </summary>
        /// <param name="pValue">值</param>
        /// <param name="pSheetName">工作表名称</param>
        /// <param name="pRowNo">行号</param>
        /// <param name="pCellNo">列号</param>
        public void SetData(dynamic pValue, string pSheetName, int pRowNo, int pCellNo)
        {
            ws = wb.Sheets[pSheetName];
            Range rngCell = ws.Cells[pRowNo, pCellNo];
            rngCell.Value = pValue;
            Marshal.FinalReleaseComObject(rngCell);
        }

        /// <summary>
        /// 将多个值写入工作表
        /// </summary>
        /// <param name="pSescribes">规则[value,rowNo,cellNo]</param>
        /// <param name="pSheetName">工作表名称</param>
        public void SetData(List<DataSescribeModel> pSescribes, string pSheetName)
        {
            ws = wb.Sheets[pSheetName];
            foreach (DataSescribeModel item in pSescribes)
            {
                Range rngCell = ws.Cells[item.RowNo, item.ColNo];
                rngCell.Value = item.Value;
                Marshal.FinalReleaseComObject(rngCell);
            }
        }

        /// <summary>
        /// 将datatable数据写入空白工作表
        /// </summary>
        /// <param name="dtSource"></param>
        /// <param name="SetSht"></param>
        public void SetDataDataTableToSht(DataTable dtSource,string shtName)

        {
            int rowNum = dtSource.Rows.Count;//内存表的行数
            int columNum = dtSource.Columns.Count;//内存表的列数
            int columnIndex = -1;
            int rowIndex = 0;
            ws = wb.Sheets[shtName];


            object[,] arData = new object[rowNum+1 , columNum];

            foreach (System.Data.DataColumn dc in dtSource.Columns)
            {
                columnIndex++;
                arData[rowIndex, columnIndex] = dc.ColumnName;
            }
            //DataTable中的数据倒入array中
            for (int i = 0; i < rowNum; i++)//行
            {
                rowIndex++;
                columnIndex = -1;
                for (int j = 0; j < columNum; j++)//列
                {
                    columnIndex++;
                    arData[rowIndex, columnIndex] = dtSource.Rows[i][j];
                }
            }


           Range startRange = ws.Cells[1, 1];
           Range endRange = ws.Cells[rowNum+1, columNum];
           Range range = ws.Range[startRange, endRange];

           range.Value = arData;


        }

        /// <summary>
        /// 将datatable数据写入工作表
        /// </summary>
        /// <param name="pDataDT">datatable</param>
        /// <param name="pSheetName">工作表名称</param>
        /// <param name="pBeginRow">开始行数</param>
        /// <param name="pBeginCol"></param>
        /// <param name="pInsertRowNo">从哪一行开始插入</param>
        /// <param name="pInsertRows">插入行数</param>
        /// <param name="pFormulaCell"></param>
        /// <param name="pNumberFormatCell"></param>
        public void SetData(System.Data.DataTable pDataDT, string pSheetName, int pBeginRow = 0, int pBeginCol = 0, int pInsertRowNo = 0, int pInsertRows = 0, List<int> pFormulaCell = null, List<int> pNumberFormatCell = null)
        {
            ws = wb.Sheets[pSheetName];
            Range rngCell = null;
            Range rngInsertRows = null;
            Range startRange = null;
            Range endRange = null;
            Range range = null;
            try
            {
                int beginRowIndex = pBeginRow == 0 ? 2 : pBeginRow;
                int beginColIndex = pBeginCol == 0 ? 1 : pBeginCol;

                #region 插入行
                if (pInsertRowNo > 0)
                {
                    rngInsertRows = (Range)ws.Rows[pBeginRow + 1, Type.Missing];
                    foreach (var DataDR in pDataDT.Rows)
                        rngInsertRows.Insert(XlDirection.xlDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                    Marshal.FinalReleaseComObject(rngInsertRows);
                }
                #endregion

                #region 设置单元格格式
                if (pDataDT.Rows.Count > 0)
                {
                    if (pNumberFormatCell == null)
                    {
                        pNumberFormatCell = new List<int>();
                        for (int j = 0; j < pDataDT.Columns.Count; j++)
                            pNumberFormatCell.Add(j + 1);
                    }

                    System.Threading.Thread threadForCulture = new System.Threading.Thread(delegate () { });
                    string format = threadForCulture.CurrentCulture.DateTimeFormat.ShortDatePattern;
                    List<string> NumberTypes = new List<string>() { "Byte", "Int16", "Int32", "Boolean" };
                    foreach (int item in pNumberFormatCell)
                    {
                        string ColumnDataType = pDataDT.Columns[item - 1].DataType.Name;
                        startRange = ws.Cells[beginRowIndex, item];
                        endRange = ws.Cells[pDataDT.Rows.Count + beginRowIndex, item];
                        range = ws.Range[startRange, endRange];

                        Marshal.FinalReleaseComObject(startRange);
                        Marshal.FinalReleaseComObject(endRange);
                        if (NumberTypes.Contains(ColumnDataType))
                            range.NumberFormat = "0";
                        else if (ColumnDataType == "Double")
                            range.NumberFormat = "#,###0.000";
                        else if (ColumnDataType == "Decimal")
                            range.NumberFormat = "#,###0.000";
                        else if (ColumnDataType == "DateTime")
                            range.NumberFormat = format;
                        else
                            range.NumberFormat = "@";

                        Marshal.FinalReleaseComObject(startRange);
                        Marshal.FinalReleaseComObject(endRange);
                        Marshal.FinalReleaseComObject(range);
                    }
                }
                #endregion

                #region 列头数据
                if (pBeginRow == 0)
                {
                    for (int i = 0; i < pDataDT.Columns.Count; i++)
                    {
                        rngCell = ws.Cells[1, i + beginColIndex];
                        rngCell.Value = pDataDT.Columns[i].ColumnName;
                        
                        Marshal.FinalReleaseComObject(rngCell);
                    }
                }
                #endregion

                #region 数据
                if (pDataDT.Rows.Count > 0)
                {
                    int rowCount = pDataDT.Rows.Count;
                    int colCount = pDataDT.Columns.Count;
                    object[,] dataArray = new object[rowCount, colCount];
                    for (int i = 0; i < rowCount; i++)
                    {
                        for (int j = 0; j < colCount; j++)
                        {
                            dataArray[i, j] = pDataDT.Rows[i][j];
                        }
                    }

                    startRange = ws.Cells[beginRowIndex, beginColIndex];
                    endRange = ws.Cells[pDataDT.Rows.Count + beginRowIndex - 1, pDataDT.Columns.Count + beginColIndex - 1];
                    range = ws.Range[startRange, endRange];

                    range.Value = dataArray;

                    Marshal.FinalReleaseComObject(startRange);
                    Marshal.FinalReleaseComObject(endRange);
                    Marshal.FinalReleaseComObject(range);
                }
                #endregion

                #region 下拉公式
                if (pFormulaCell != null && pDataDT.Rows.Count > 1)
                {
                    foreach (int FormulaCellItem in pFormulaCell)
                    {
                        startRange = ws.Cells[beginRowIndex, FormulaCellItem];
                        endRange = ws.Cells[pDataDT.Rows.Count + beginRowIndex - 1, FormulaCellItem];
                        range = ws.Range[startRange, endRange];

                        startRange.AutoFill(range, XlAutoFillType.xlFillCopy);

                        Marshal.FinalReleaseComObject(startRange);
                        Marshal.FinalReleaseComObject(endRange);
                        Marshal.FinalReleaseComObject(range);
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                if (rngCell != null) Marshal.FinalReleaseComObject(rngCell);
                if (rngInsertRows != null) Marshal.FinalReleaseComObject(rngInsertRows);
                if (range != null) Marshal.FinalReleaseComObject(range);
                if (startRange != null) Marshal.FinalReleaseComObject(startRange);
                if (endRange != null) Marshal.FinalReleaseComObject(endRange);
                if (ws != null) Marshal.FinalReleaseComObject(ws);
            }
        }

        /// <summary>
        /// 将datatable数据写入工作表
        /// </summary>
        /// <param name="DataDTList"></param>
        public void SetData(List<DataDTToExcelModel> pDataDTList)
        {
            //写入数据
            foreach (var dataDT in pDataDTList)
            {
                if (dataDT.DataDT != null)
                {
                    ws = wb.Worksheets[dataDT.SheetName];
                    SetData(dataDT.DataDT, dataDT.SheetName, dataDT.BeginRow, dataDT.BeginCell, dataDT.InsertRowNo, dataDT.InsertRows, dataDT.FormulaCell, dataDT.NumberFormatCell);
                    Marshal.FinalReleaseComObject(ws);
                }
            }
        }

        public void AddColumns(string pSheetName, List<ColumnDataToExcel> pData)
        {
            pData = pData.OrderBy(m => m.ColumnNo).ToList();

            ws = wb.Worksheets[pSheetName];
            int EndRow = ws.UsedRange.SpecialCells(XlCellType.xlCellTypeLastCell).Row;

            Range RangeInsertColumn = null;
            Range RangeInsertValue = null;
            foreach (var data in pData)
            {
                object[,] InsertValue = new object[EndRow - 1, 1];
                for (int i = 0; i < EndRow - 1; i++)
                    InsertValue[i, 0] = data.ColumnValue;

                //插入列
                RangeInsertColumn = (Range)ws.Columns[data.ColumnNo, Type.Missing];
                RangeInsertColumn.Insert(XlDirection.xlToRight);
                Marshal.FinalReleaseComObject(RangeInsertColumn);
                //写入列名
                RangeInsertValue = ws.Cells[1, data.ColumnNo];
                RangeInsertValue.Value = data.ColumnName;
                Marshal.FinalReleaseComObject(RangeInsertValue);
                //写入列值
                Range Range0 = ws.Cells[2, data.ColumnNo];
                Range Range1 = ws.Cells[EndRow, data.ColumnNo];
                RangeInsertValue = ws.Range[Range0, Range1];
                RangeInsertValue.Value = InsertValue;
                RangeInsertValue.NumberFormat = "@";
                Marshal.FinalReleaseComObject(RangeInsertValue);
            }
            Marshal.FinalReleaseComObject(ws);
        }
        #endregion

        #region 获取数据
        public System.Data.DataTable GetData(string pSheetName)
        {
            int eachRow = 50000;
            System.Data.DataTable dt = null;
            ws = wb.Worksheets[pSheetName];
           
            Range rngBegin, rngEnd, rngCell;
            int colNum = ws.UsedRange.CurrentRegion.Columns.Count;
            int rowNum = ws.UsedRange.CurrentRegion.Rows.Count;

            int beginRow = 1;
            for (int endRow = 0; endRow < rowNum;)
            {
                endRow = (endRow + eachRow) > rowNum ? rowNum : (endRow + eachRow);

                rngBegin = ws.Cells[beginRow, 1];
                rngEnd = ws.Cells[endRow, colNum];
                rngCell = ws.Range[rngBegin, rngEnd];
                object[,] obj = (object[,])rngCell.Value;

                if (dt == null)
                    dt = ObjectHelper.ObjectToDataTable(obj);
                else
                {
                    System.Data.DataTable tempDT = dt.Clone();
                    tempDT = ObjectHelper.ObjectToDataTable(obj, false, tempDT);
                    foreach (DataRow item in tempDT.Rows)
                        dt.ImportRow(item);
                }
                Marshal.FinalReleaseComObject(rngBegin);
                Marshal.FinalReleaseComObject(rngEnd);
                Marshal.FinalReleaseComObject(rngCell);

                beginRow = beginRow + eachRow;
            }
            return dt;
        }



        /// <summary>
        /// 获取单个数据
        /// </summary>
        /// <param name="pSheetName"></param>
        /// <param name="pRowNo"></param>
        /// <param name="pColNo"></param>
        /// <returns></returns>
        public dynamic GetData(string pSheetName, int pRowNo, int pColNo)
        {
            List<DataSescribeModel> models = new List<DataSescribeModel>() {
                new DataSescribeModel (){ RowNo = pRowNo , ColNo = pColNo }
            };

            return GetData(pSheetName, models).FirstOrDefault().Value;
        }
        /// <summary>
        /// 获取单个数据
        /// </summary>
        /// <param name="pSheetName"></param>
        /// <param name="pModel"></param>
        /// <returns></returns>
        public DataSescribeModel GetData(string pSheetName, DataSescribeModel pModel)
        {
            List<DataSescribeModel> models = new List<DataSescribeModel>() { pModel };

            return GetData(pSheetName, models).FirstOrDefault();
        }
        /// <summary>
        /// 获取多个数据
        /// </summary>
        /// <param name="pSheetName"></param>
        /// <param name="pModels"></param>
        /// <returns></returns>
        public List<DataSescribeModel> GetData(string pSheetName, List<DataSescribeModel> pModels)
        {
            ws = wb.Worksheets[pSheetName];
            foreach (DataSescribeModel item in pModels)
            {
                if (item.RowEndNo == 0) item.RowEndNo = item.RowNo;
                if (item.ColEndNo == 0) item.ColEndNo = item.ColNo;

                Range rngBegin, rngEnd;
                rngBegin = ws.Cells[item.RowNo, item.ColNo];
                rngEnd = ws.Cells[item.RowEndNo, item.ColEndNo];

                Range rngCell = ws.Range[rngBegin, rngEnd];
                item.Value = rngCell.Value;
                Marshal.FinalReleaseComObject(rngCell);
            }

            return pModels;
        }
        #endregion

        #region 根据条件获取需要插入的行数
        public void SearchDifData(System.Data.DataTable pDataDT, string pSheetName, string searchName, List<string> copyArray, List<string> sumArray)
        {
            ws = wb.Sheets[pSheetName];
            Range rngCell = null;
            Range rngInsertRows = null;
            Range startRange = null;
            Range endRange = null;
            Range range = null;
            try
            {
                #region 获取对比的colume,设置需要复制、求和的colume

                int beginRowIndex = 2;
                int beginColIndex = 1;
                //需要对比的列
                int searchColume = 1;
                //需要求和的列
                List<int> sumColume = new List<int>();
                List<float> sumNumList = new List<float>();
                //需要复制的列
                List<int> copyColume = new List<int>();

                for (int i = 0; i < pDataDT.Columns.Count; i++)
                {
                    rngCell = ws.Cells[1, i + beginColIndex];
                    rngCell.Value = pDataDT.Columns[i].ColumnName;
                    if (rngCell.Value == searchName) searchColume = i;
                    foreach (string item in sumArray)
                    {
                        if (rngCell.Value == item)
                        {
                            sumColume.Add(i);
                            sumNumList.Add(0);
                        }
                    }
                    foreach (string item in copyArray)
                    {
                        if (rngCell.Value == item) copyColume.Add(i);
                    }
                    Marshal.FinalReleaseComObject(rngCell);
                }
                #endregion

                int rowCount = pDataDT.Rows.Count;
                int colCount = pDataDT.Columns.Count;
                object[,] dataArray = new object[rowCount, colCount];

                string pString = pDataDT.Rows[0][searchColume].ToString();
                int addNum = 2;

                for (int i = 0; i <= rowCount; i++)
                {
                    if (i == rowCount)
                    {
                        SetData(pString, pSheetName, i + addNum, searchColume + 1);
                        for (int j = 0; j < sumColume.Count; j++)
                        {
                            SetData(sumNumList[j], pSheetName, i + addNum, sumColume[j] + 1);
                        }
                        foreach (int copy in copyColume)
                        {
                            string copyString = pDataDT.Rows[i - 1][copy].ToString();
                            SetData(copyString, pSheetName, i + addNum, copy + 1);
                            DeleteData(pSheetName, i + addNum - 1, copy + 1);
                        }
                        return;
                    }
                    string current = pDataDT.Rows[i][searchColume].ToString();
                    if (pString != current)
                    {
                        InsetRow(pDataDT, pSheetName, i + addNum);
                        SetData(pString, pSheetName, i + addNum, searchColume + 1);
                        for (int j = 0; j < sumColume.Count; j++)
                        {
                            SetData(sumNumList[j], pSheetName, i + addNum, sumColume[j] + 1);
                            sumNumList[j] = Convert.ToSingle(pDataDT.Rows[i][sumColume[j]].ToString());
                        }
                        foreach (int copy in copyColume)
                        {
                            string copyString = pDataDT.Rows[i - 1][copy].ToString();
                            SetData(copyString, pSheetName, i + addNum, copy + 1);
                            DeleteData(pSheetName, i + addNum - 1, copy + 1);
                        }
                        addNum++;
                    }
                    else
                    {
                        if (i != 0)
                        {
                            foreach (int copy in copyColume)
                            {
                                DeleteData(pSheetName, i + addNum - 1, copy + 1);
                            }
                        }

                        for (int j = 0; j < sumColume.Count; j++)
                        {
                            float sumAdd = new int();
                            sumAdd = Convert.ToSingle(pDataDT.Rows[i][sumColume[j]].ToString());
                            sumNumList[j] += sumAdd;
                        }
                    }
                    pString = current;
                }

                //startRange = ws.Cells[beginRowIndex, beginColIndex];
                //endRange = ws.Cells[pDataDT.Rows.Count + beginRowIndex - 1, pDataDT.Columns.Count + beginColIndex - 1];
                //range = ws.Range[startRange, endRange];

                //range.Value = dataArray;

                //Marshal.FinalReleaseComObject(startRange);
                //Marshal.FinalReleaseComObject(endRange);
                //Marshal.FinalReleaseComObject(range);

            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                if (rngCell != null) Marshal.FinalReleaseComObject(rngCell);
                if (rngInsertRows != null) Marshal.FinalReleaseComObject(rngInsertRows);
                if (range != null) Marshal.FinalReleaseComObject(range);
                if (startRange != null) Marshal.FinalReleaseComObject(startRange);
                if (endRange != null) Marshal.FinalReleaseComObject(endRange);
                if (ws != null) Marshal.FinalReleaseComObject(ws);
            }
        }
        #endregion

        #region 插入行
        /// <summary>
        /// 
        /// </summary>
        /// <param name="pDataDT"></param>
        /// <param name="pSheetName"></param>
        /// <param name="pBeginRow"></param>
        public void InsetRow(System.Data.DataTable pDataDT, string pSheetName, int pBeginRow = 0)
        {
            ws = wb.Sheets[pSheetName];
            Range rngInsertRows = null;
            try
            {
                int beginRowIndex = pBeginRow == 0 ? 2 : pBeginRow;

                rngInsertRows = (Range)ws.Rows[beginRowIndex, Type.Missing];
                rngInsertRows.Insert(XlDirection.xlDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                Marshal.FinalReleaseComObject(rngInsertRows);
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                if (rngInsertRows != null) Marshal.FinalReleaseComObject(rngInsertRows);
                if (ws != null) Marshal.FinalReleaseComObject(ws);
            }
        }


        #endregion

        #region 保存
        private void Save()
        {
            wb.RefreshAll();
            if (File.Exists(filePath))
                wb.Save();
            else
                wb.SaveAs(filePath);
        }
        #endregion

        #endregion
    }

    #region Models
    public class DataSescribeModel
    {
        public string Name { get; set; }
        public dynamic Value { get; set; }
        public int RowNo { get; set; }
        public int ColNo { get; set; }
        public int RowEndNo { get; set; }
        public int ColEndNo { get; set; }
    }

    public class DataDTToExcelModel
    {
        public string SheetName;
        public System.Data.DataTable DataDT;
        public int BeginRow = 0;
        public int BeginCell = 0;
        public int InsertRowNo = 0;
        public int InsertRows = 0;
        public List<int> FormulaCell = null;
        public List<int> NumberFormatCell = new List<int>();
    }

    public class ColumnDataToExcel
    {
        public int ColumnNo = 0;
        public string ColumnName = "";
        public dynamic ColumnValue;
    }
    #endregion
}