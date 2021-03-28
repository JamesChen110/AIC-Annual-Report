using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelNS = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Data.OleDb;
using System.Data;
using System.Web;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace AIC_Annual_Report
{
    public class Excel
    {
        public Array Sheet_To_Array(ExcelNS.Worksheet sheet , string strFolderPath, string strWorkBookName)
        {
          
            int lastRow = sheet.Cells.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row;
            int lastColumn = sheet.Cells.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column;
            System.Array array = (System.Array)sheet.get_Range("").Cells.Value;
           

            return array;
        }


        public string saveAsExcel_FromActiveMHTFormat(string strFolderPath, string strWorkBookName)
        {
            //Gets Excel and gets Activeworkbook and worksheet
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel.Workbook oWB;
            oXL = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            oXL.Visible = true;
            oWB = (Microsoft.Office.Interop.Excel.Workbook)oXL.ActiveWorkbook;
            bool blnWBFound = false;

            //validate whether the active book path is correct or not
            if (oWB.Name != strWorkBookName + ".MHTML")
            {
                //Loop throught the workbooks and find out is it there or not, if not then reply as error
                foreach(var wb in oXL.Workbooks)
                {
                    if (oWB.Name != strWorkBookName + ".MHTML")
                    {
                        blnWBFound = true;
                        oWB = (Microsoft.Office.Interop.Excel.Workbook)wb;
                        break;
                    }
                }

                if (blnWBFound == false)
                {
                    return "Excel workbook" +  strWorkBookName + " is not found";
                }
            }

            //Delete if the file is already exist
            if (System.IO.File.Exists(strFolderPath + "\\" + strWorkBookName + ".xlsx") == true)
            {
                System.IO.File.Delete(strFolderPath + "\\" + strWorkBookName + ".xlsx");
            }                
            oWB.SaveAs(strFolderPath + "\\" + strWorkBookName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault);
            oWB.Close();
            return "Success";
        }                
        public string Excel_To_CSV_Conversion(string strFolderPath, string strWorkBookName, string strSheetName) {
            //Gets Excel and gets Activeworkbook and worksheet
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel.Workbook oWB;
            oXL = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            oXL.Visible = true;
            oWB = (Microsoft.Office.Interop.Excel.Workbook)oXL.ActiveWorkbook;
            bool blnWBFound = false;

            //validate whether the active book path is correct or not
            if (oWB.Name != strWorkBookName + ".xlsx") {
                //Loop throught the workbooks and find out is it there or not, if not then reply as error
                foreach (var wb in oXL.Workbooks) {
                    if (oWB.Name == strWorkBookName + ".xlsx") {
                        blnWBFound = true;
                        oWB = (Microsoft.Office.Interop.Excel.Workbook)wb;
                        break;
                    }
                }
                //Changed for testing
                if (blnWBFound == true) {
                }

                if (blnWBFound == false) {
                    //Open the new excel
                    //oWB = oXL.WorkbookOpen()
                    return "Excel workbook" + strWorkBookName + " is not found";
                }
            }

            //Delete if the file is already exist
            if (System.IO.File.Exists(strFolderPath + "\\" + strWorkBookName + ".csv") == true) {
                System.IO.File.Delete(strFolderPath + "\\" + strWorkBookName + ".csv");
            }

            //Save with the mentioned sheet
            bool blnSaved = false;

            //If it has only one sheet, then no need to check
            if (oWB.Worksheets.Count == 1) {
                oWB.SaveAs(strFolderPath + "\\" + strWorkBookName, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange);
                blnSaved = true;
            }
            else {
                foreach (Microsoft.Office.Interop.Excel.Worksheet sht in oWB.Worksheets) {
                    if (sht.Name.Trim().ToLower() == strSheetName.Trim().ToLower()) {
                        //Check whether this sheet is alreayd selected or not, if not then select it
                        if (oWB.ActiveSheet.name != sht.Name)
                                sht.Select();
                        oWB.Application.DisplayAlerts = false;
                        oWB.SaveAs(strFolderPath + "\\" + strWorkBookName, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange);
                        blnSaved = true;
                        oWB.Application.DisplayAlerts = true;
                    }
                }
            }      
            
            if (blnSaved == false) {
                return "Excel Sheet with the name:" + strSheetName + " is not found";
            }
            else {
                oWB.Close(false);
                return "Success";
            }
        }
        public void ExportToCSV(string excelFilePath, string csvOutputFile, int worksheetNumber = 1)
        {
            if (!File.Exists(excelFilePath)) throw new FileNotFoundException(excelFilePath);
            if (File.Exists(csvOutputFile)) throw new ArgumentException("File exists: " + csvOutputFile);

            //读取excel
            var cnnStr = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 8.0;IMEX=1;HDR=NO\"", excelFilePath);
            var cnn = new OleDbConnection(cnnStr);

            // get schema, then data
            var dt = new DataTable();
            try
            {
                cnn.Open();
                var schemaTable = cnn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (schemaTable.Rows.Count < worksheetNumber) throw new ArgumentException("The worksheet number provided cannot be found in the spreadsheet");
                string worksheet = schemaTable.Rows[worksheetNumber - 1]["table_name"].ToString().Replace("'", "");
                string sql = String.Format("select * from [{0}]", worksheet);
                var da = new OleDbDataAdapter(sql, cnn);
                da.Fill(dt);
            }
            catch (Exception e)
            {
                // ???
                throw e;
            }
            finally
            {
                // free resources
                cnn.Close();
            }

            StringBuilder sb = new StringBuilder();
            int i = 0;
            //for (i = 0; i <= dt.Columns.Count - 1; i++)
            //{
            //    if (i > 0) { sb.Append(","); }
            //    sb.Append(dt.Columns[i].ColumnName);
            //}
            //sb.Append("\n");
            foreach (DataRow dr in dt.Rows)
            {
                for (i = 0; i <= dt.Columns.Count - 1; i++)
                {
                    if (i > 0) { sb.Append(","); }
                    sb.Append(dr[i].ToString());
                }
                sb.Append("\n");
            }

           
            //写入csv
            byte[] buffer = System.Text.Encoding.UTF8.GetBytes(sb.ToString());
            byte[] outBuffer = new byte[buffer.Length + 3];
            outBuffer[0] = (byte)0xEF;
            outBuffer[1] = (byte)0xBB;
            outBuffer[2] = (byte)0xBF;
            Array.Copy(buffer, 0, outBuffer, 3, buffer.Length);

            File.WriteAllBytes(csvOutputFile, outBuffer);
        }
        public void CSV_To_Excel_Converstion(string strCSVFileFullpath) {
            ExcelNS.Application app = new ExcelNS.Application();
            app.DisplayAlerts = false;
            ExcelNS.Workbook wb = app.Workbooks.Open(strCSVFileFullpath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            wb.SaveAs(strCSVFileFullpath.Replace(".csv",".xlsx"), ExcelNS.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, Type.Missing, Type.Missing, ExcelNS.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);            
            wb.Close();
            app.DisplayAlerts = true;
            app.Quit();
        }
        public void Htm_To_Excel_Converstion(string strHtmFileFullpath) {
            ExcelNS.Application app = new ExcelNS.Application();
            app.DisplayAlerts = false;
            ExcelNS.Workbook wb = app.Workbooks.Open(strHtmFileFullpath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            wb.SaveAs(strHtmFileFullpath.Replace(".htm", ".xlsx"), ExcelNS.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, Type.Missing, Type.Missing, ExcelNS.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            wb.Close();
            app.DisplayAlerts = true;
            app.Quit();
        }
        public void Excel_Delete_BlankColumns(string excelFilePath, string strHeaderText,string strCheckMisAlignedData = "No") {
            if (!File.Exists(excelFilePath)) throw new FileNotFoundException(excelFilePath);
            ExcelNS.Application oXL = new ExcelNS.Application();
            ExcelNS.Workbook oWB;
            ExcelNS.Worksheet oSht;
            oWB = oXL.Workbooks.Open(excelFilePath, false, false);
            oXL.Visible = true;
            oSht = oWB.Sheets[1];
            ExcelNS.Range xlRange = oSht.Range["A1:" + GetColumnName(oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column.ToString()) + oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row];

            var Data = xlRange.Value;

            //find the title row
            for (int i = 1;i<= xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row;i++) {
                for (int j = 1; j <= xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column; j++) {
                    if (Data[i, j] != null &&  Data[i, j].ToString() == strHeaderText) {

                        if (i != 1) {
                            //Delete the previous rows
                            ExcelNS.Range deleteRows = (ExcelNS.Range)xlRange.Range["1:" + Convert.ToString(i - 1)];
                            deleteRows.EntireRow.Delete(Shift: ExcelNS.XlDeleteShiftDirection.xlShiftUp);
                        }                  

                        //xlRange.Rows[1, i - 1].EntireRow.Delete(Shift: ExcelNS.XlDeleteShiftDirection.xlShiftUp);
                        //change the i value to the highest to break the main loop
                        i = xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row + 100;
                        break;
                    }
                }
            }


            //Find the Empty columns and delete
            xlRange = oSht.Range["A1:" + GetColumnName(oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column.ToString()) + oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row];
            xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeConstants).EntireColumn.Hidden = true;
            try {
                if (xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible) != null && xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible).Count > 0) {
                    xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible).EntireColumn.Delete(ExcelNS.XlDeleteShiftDirection.xlShiftUp);
                }
            }
            catch (Exception e1) {

                e1 = e1;
            }
          
  
            xlRange.EntireColumn.Hidden = false;

            //Find the Empty columns and delete
            xlRange = oSht.Range["A1:" + GetColumnName(oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column.ToString()) + oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row];
            xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeConstants).EntireRow.Hidden = true;
            try {
                if (xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible) != null && xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible).Count > 0) {
                    xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible).EntireRow.Delete(ExcelNS.XlDeleteShiftDirection.xlShiftUp);
                }
            }
            catch (Exception e) {
                e = e;
            }                
            xlRange.EntireRow.Hidden = false;

            if (strCheckMisAlignedData.ToString().ToLower() == "yes") {
                Dictionary<int, bool> TitleClmns = new Dictionary<int, bool>();
                xlRange = oSht.Range["A1:" + GetColumnName(oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column.ToString()) + oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row];
                Data = xlRange.Value;
                int intFirstTitleIndex = 0;
                //Record all columns, whether its title or not
                for (int j = 1; j <= xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column; j++) {
                    if (Data[1, j] != null && Data[1, j].ToString() != "") {
                        if (intFirstTitleIndex == 0) intFirstTitleIndex = j;
                        TitleClmns.Add(j, true);
                    }
                    else {
                        TitleClmns.Add(j, false);
                    }                        
                }

                //Loop through rows and check for the mis-aligned data
                //find the title row
                bool blnMiddleMissingData,blnFirstDataMissing,blnPreviousMiddleMissingData =false;
                int intMiddleMissingIndex, intPreviousMiddleMissingIndex = 0;
                int intTotalColumns = xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column;
                int intTotalRows = xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row;
                for (int i = 2; i <= intTotalRows; i++) {
                    blnFirstDataMissing = false;
                    blnMiddleMissingData = false;
                    intMiddleMissingIndex = 0;
                    for (int j = 1; j <= intTotalColumns; j++) {
                        if (TitleClmns[j] == true) {
                            if (Data[i, j] == null || Data[i, j].ToString() == "") {
                                if (j == intFirstTitleIndex) {
                                    blnFirstDataMissing = true;
                                }
                                else if(blnFirstDataMissing == false) {
                                    blnMiddleMissingData = true;
                                    if (intMiddleMissingIndex == 0) intMiddleMissingIndex = j;
                                }
                            }
                            else if (blnMiddleMissingData == true) {
                                blnMiddleMissingData = false;
                                intMiddleMissingIndex = 0;
                            }
                        }

                        //Assumption is that the previous line is mis-aligned, 
                        //So cute paste this line data to the previous line from where missing data
                        if (blnFirstDataMissing == true && blnPreviousMiddleMissingData == true && (Data[i, j] != null && Data[i, j].ToString() != "")) {
                            //Calculate the total cells need to be cut & paste
                            int intTotalCells = intTotalColumns - intPreviousMiddleMissingIndex;
                            ExcelNS.Range rngCut = oSht.Range[GetColumnName(j.ToString()) + i + ":" + GetColumnName((j + intTotalCells).ToString()) + i];
                            ExcelNS.Range rngPaste = oSht.Range[GetColumnName(intPreviousMiddleMissingIndex.ToString()) + (i - 1).ToString() + ":" + GetColumnName((j + intTotalCells).ToString()) + (i - 1)];
                            //Cut
                            //rngCut.Copy();
                            //Paste
                            rngPaste.Insert(ExcelNS.XlInsertShiftDirection.xlShiftToRight, rngCut.Cut());
                            //rngPaste.PasteSpecial(ExcelNS.XlPasteType.xlPasteAll, ExcelNS.XlPasteSpecialOperation.xlPasteSpecialOperationNone);
                            break;
                        }else if (blnFirstDataMissing == false && j > 1 && (TitleClmns[j-1] == true && TitleClmns[j] == false) && (Data[i, j] != null && Data[i, j].ToString() != "")) {
                            //The data is misaligned - one extra cell to the right
                            //Copy the current cell data and add with previous cell data
                            //Delete the current cell with the option "shift to left"
                            //Then break the loop, no need the check the next columns
                            oSht.Range[GetColumnName((j-1).ToString()) + i].Value = Data[i, j-1].ToString() + Data[i, j].ToString();
                            oSht.Range[GetColumnName((j).ToString()) + i].Delete(ExcelNS.XlDeleteShiftDirection.xlShiftToLeft);
                            break;
                        }
                    }
                    //Mark if the data is missing from the middle
                    if (blnMiddleMissingData == true) {
                        blnPreviousMiddleMissingData = true;
                        intPreviousMiddleMissingIndex = intMiddleMissingIndex;
                    }
                }

                //Find the Empty columns and delete
                xlRange = oSht.Range["A1:" + GetColumnName(oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column.ToString()) + oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row];
                xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeConstants).EntireColumn.Hidden = true;
                try {
                    if (xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible) != null && xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible).Count > 0) {
                        xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible).EntireColumn.Delete(ExcelNS.XlDeleteShiftDirection.xlShiftUp);
                    }
                }
                catch (Exception e1) {

                    e1 = e1;
                }

                xlRange.EntireColumn.Hidden = false;
                //Find the Empty columns and delete
                xlRange = oSht.Range["A1:" + GetColumnName(oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column.ToString()) + oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row];
                xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeConstants).EntireRow.Hidden = true;
                try {
                    if (xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible) != null && xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible).Count > 0) {
                        xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible).EntireRow.Delete(ExcelNS.XlDeleteShiftDirection.xlShiftUp);
                    }
                }
                catch (Exception e) {
                    e = e;
                }
                xlRange.EntireRow.Hidden = false;

            }

            //Delete the header rows, if its in the data...
            //Delete duplicate headers
            oXL.DisplayAlerts = false;
            oWB.Save();
            oWB.Close();
            oXL.Quit();
        }
        public void Excel_Delete_DuplicateHeadings(string excelFilePath, string strHeaderText) {
            if (!File.Exists(excelFilePath)) throw new FileNotFoundException(excelFilePath);
            ExcelNS.Application oXL = new ExcelNS.Application();
            ExcelNS.Workbook oWB;
            ExcelNS.Worksheet oSht;
            oWB = oXL.Workbooks.Open(excelFilePath, false, false);
            oXL.Visible = true;
            oSht = oWB.Sheets[1];
            ExcelNS.Range xlRange = oSht.Range["A1:" + GetColumnName(oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column.ToString()) + oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row];
            var Data = xlRange.Value;

            //For this function title row should be row no "1"
            int i = 1;
            int intFilterClmn = 0;
            bool blnHeadingFound = false;
            for (int j = 1; j <= xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column; j++) {
                if (Data[i, j] != null && Data[i, j].ToString() == strHeaderText) {
                    intFilterClmn = j;
                    blnHeadingFound = true;
                    break;
                }
            }

            if (blnHeadingFound == false) {
                throw new Exception("Heading:" + strHeaderText +" could not found in the 1st row");
            }

            String[] FilterList = { strHeaderText};
            //Filter
            xlRange.AutoFilter(intFilterClmn, FilterList, ExcelNS.XlAutoFilterOperator.xlFilterValues);
            //Delete the values
            xlRange.Offset[1, 0].SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible).EntireRow.Delete();
            //Turn off the filter
            oSht.AutoFilterMode = false;

            //Delete the header rows, if its in the data...
            //Delete duplicate headers
            oXL.DisplayAlerts = false;
            oWB.Save();
            oWB.Close();
            oXL.Quit();
        }
        public void Excel_Remove_Duplicates(string excelFilePath, string strHeaderText) {
            if (!File.Exists(excelFilePath)) throw new FileNotFoundException(excelFilePath);
            ExcelNS.Application oXL = new ExcelNS.Application();
            ExcelNS.Workbook oWB;
            ExcelNS.Worksheet oSht;
            oWB = oXL.Workbooks.Open(excelFilePath, false, false);
            oXL.Visible = true;
            oSht = oWB.Sheets[1];
            ExcelNS.Range xlRange = oSht.Range["A1:" + GetColumnName(oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column.ToString()) + oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row];
            var Data = xlRange.Value;

            //For this function title row should be row no "1"
            int i = 1;
            int intFilterClmn = 0;
            bool blnHeadingFound = false;
            for (int j = 1; j <= xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column; j++) {
                if (Data[i, j] != null && Data[i, j].ToString() == strHeaderText) {
                    intFilterClmn = j;
                    blnHeadingFound = true;
                    break;
                }
            }

            if (blnHeadingFound == false) {
                throw new Exception("Heading:" + strHeaderText + " could not found in the 1st row");
            }

            //Remove duplicates
            xlRange.RemoveDuplicates(intFilterClmn);

            //Delete the header rows, if its in the data...
            //Delete duplicate headers
            oXL.DisplayAlerts = false;
            oWB.Save();
            oWB.Close();
            oXL.Quit();
            
        }
        public string Excel_Delete_Row(string excelFilePath, string strStartRowNo, string strEndRowNo, string strSheetNumber = "1") {
            //Gets Excel and gets Activeworkbook and worksheet
            if (!File.Exists(excelFilePath)) return "Excel file not found";
            ExcelNS.Application oXL = new ExcelNS.Application();
            ExcelNS.Workbook oWB;
            oWB = oXL.Workbooks.Open(excelFilePath, false, false);
            oXL.Visible = true;

            //Find the title
            ExcelNS._Worksheet wsSheet = oWB.Sheets[Convert.ToInt32(strSheetNumber)];
            ExcelNS.Range xlRange = wsSheet.Range[strStartRowNo + ":" + strEndRowNo];
            
            xlRange.Delete();

            
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(wsSheet);
            xlRange = null;
            wsSheet = null;

            oWB.Save();
            Marshal.ReleaseComObject(oWB);
            oWB = null;

            oXL.Quit();
            Marshal.ReleaseComObject(oXL);
            oXL = null;

            return "Success";
        }


        public string Excel_Copy_sheet(ExcelNS.Workbook oWB2, string excelFile1Path, string strCopySheetNumber, string strcopynewsheetname, string strPastebeforeSheetNumber="", string strPasteafterSheetNumber="" )
        {

            
            //Gets Excel and gets Activeworkbook and worksheet
            //if (!File.Exists(excelFile1Path)) return "Excel file not found";
            //if (!File.Exists(excelFile2Path)) return "Excel file not found";
            ExcelNS.Application oXL = oWB2.Application;

           
            ExcelNS.Workbook oWB;
            oXL.DisplayAlerts = false;
            //oWB = oXL.Workbooks.Open(excelFile1Path, false, false);
            oWB = oXL.Workbooks.Open(excelFile1Path, false, false);
            oXL.Visible = true;

            //Find the title
            ExcelNS.Worksheet wsCopySheet = oWB.Sheets[Convert.ToInt32(strCopySheetNumber)];
            





            if (strPasteafterSheetNumber == "")
            {
                oWB2.Activate();
                ExcelNS.Worksheet wsPasteSheet = oWB2.Sheets[Convert.ToInt32(strPastebeforeSheetNumber)];

                wsCopySheet.Copy(wsPasteSheet);
                
            }
            else
            {
                //oWB.ActiveSheet;
                //AcSht.Name = strcopynewsheetname;
                // ExcelNS.Worksheet wsPasteSheet = oWB2.Sheets[Convert.ToInt32(strPasteafterSheetNumber)];
                wsCopySheet.Copy(null, oWB2.Sheets[Convert.ToInt32(strPastebeforeSheetNumber)]);
                oWB2.Sheets[Convert.ToInt32(strPastebeforeSheetNumber)].name = strcopynewsheetname;
               

            }

            ExcelNS.Worksheet AcSht = oWB2.ActiveSheet;
            AcSht.Name = strcopynewsheetname;


            //oWB.ActiveSheet;
            //AcSht.Name = strcopynewsheetname;






            oWB.Close();
            oWB = null;
            // Marshal.ReleaseComObject(oWB);


            //oWB2 = null;
            //oXL.DisplayAlerts = true;
            //oXL.Quit();
            //Marshal.ReleaseComObject(oXL);
            //oXL = null;

            return "Success";
        }
        public string Excel_Copy_Data(string excelFile1Path, string excelFile2Path, string strCopySheetNumber, string strPasteSheetNumber, string strCopyRange, string strPasteRange) {
            //Gets Excel and gets Activeworkbook and worksheet
            if (!File.Exists(excelFile1Path)) return "Excel file not found";
            if (!File.Exists(excelFile2Path)) return "Excel file not found";
            ExcelNS.Application oXL = new ExcelNS.Application();
            ExcelNS.Workbook oWB, oWB2;
            oXL.DisplayAlerts = false;
            oWB = oXL.Workbooks.Open(excelFile1Path, false, false);
            oWB2 = oXL.Workbooks.Open(excelFile2Path, false, false);
            oXL.Visible = true;

            //Find the title
            ExcelNS._Worksheet wsCopySheet = oWB.Sheets[Convert.ToInt32(strCopySheetNumber)];
            ExcelNS._Worksheet wsPasteSheet = oWB2.Sheets[Convert.ToInt32(strPasteSheetNumber)];

            ExcelNS.Range xlCopyRange = wsCopySheet.Range[strCopyRange];
            wsCopySheet.Activate();
            xlCopyRange.Copy();
            ExcelNS.Range xlPasteRange = wsPasteSheet.Range[strPasteRange];
            wsPasteSheet.Activate();
            xlPasteRange.PasteSpecial(ExcelNS.XlPasteType.xlPasteAll);


            Marshal.ReleaseComObject(xlCopyRange);
            Marshal.ReleaseComObject(xlPasteRange);
            xlCopyRange = null;
            xlPasteRange = null;

            Marshal.ReleaseComObject(wsCopySheet);
            Marshal.ReleaseComObject(wsPasteSheet);
            wsCopySheet = null;
            wsPasteSheet = null;

            oWB.Close();
            oWB2.Save();
            oWB2.Close();
            Marshal.ReleaseComObject(oWB);
            Marshal.ReleaseComObject(oWB2);
            oWB = null;
            oWB2 = null;
            oXL.DisplayAlerts = true;
            oXL.Quit();
            Marshal.ReleaseComObject(oXL);
            oXL = null;

            return "Success";
        }
        public string Excel_Copy_Data_UsingSheetName(string excelFile1Path, string excelFile2Path, string strCopySheetName, string strPasteSheetName, string strCopyRange, string strPasteRange) {
            //Gets Excel and gets Activeworkbook and worksheet
            if (!File.Exists(excelFile1Path)) throw new Exception("Excel file not found");
            if (!File.Exists(excelFile2Path))  throw new Exception("Excel file not found");
            ExcelNS.Application oXL = new ExcelNS.Application();
            ExcelNS.Workbook oWB, oWB2;
            oXL.DisplayAlerts = false;
            oWB = oXL.Workbooks.Open(excelFile1Path, false, false);
            oWB2 = oXL.Workbooks.Open(excelFile2Path, false, false);
            oXL.Visible = true;

            ExcelNS._Worksheet wsCopySheet = null;
            ExcelNS._Worksheet wsPasteSheet = null;
            //Find the sheet
            foreach (var sheet in oWB.Sheets) {
                //Console.WriteLine(((ExcelNS._Worksheet)sheet).Name.ToLower());
                if (((ExcelNS._Worksheet)sheet).Name.ToLower().Trim() == strCopySheetName.ToLower().Trim()) {
                    wsCopySheet = (ExcelNS._Worksheet)sheet;
                    //Marshal.ReleaseComObject(sheet);

                    break;
                }
                Marshal.ReleaseComObject(sheet);
            }
            foreach (var sheet in oWB2.Sheets) {
                if (((ExcelNS._Worksheet)sheet).Name.ToLower().Trim() == strPasteSheetName.ToLower().Trim()) {
                    wsPasteSheet = (ExcelNS._Worksheet)sheet;
                    //Marshal.ReleaseComObject(sheet);
                    break;
                }
                Marshal.ReleaseComObject(sheet);
            }

            if (wsCopySheet == null) throw new System.Exception("Copy Sheet name is wrong/not exist");
            if (wsPasteSheet == null) throw new System.Exception("Paste Sheet name is wrong/not exist");

            ExcelNS.Range xlCopyRange = wsCopySheet.Range[strCopyRange];
            wsCopySheet.Activate();
            xlCopyRange.Copy();
            ExcelNS.Range xlPasteRange = wsPasteSheet.Range[strPasteRange];
            wsPasteSheet.Activate();
            xlPasteRange.PasteSpecial(ExcelNS.XlPasteType.xlPasteAll);

            Marshal.ReleaseComObject(xlCopyRange);
            Marshal.ReleaseComObject(xlPasteRange);
            xlCopyRange = null;
            xlPasteRange = null;

            Marshal.ReleaseComObject(wsCopySheet);
            Marshal.ReleaseComObject(wsPasteSheet);
            wsCopySheet = null;
            wsPasteSheet = null;

            oWB.Close();
            oWB2.Save();
            oWB2.Close();
            Marshal.ReleaseComObject(oWB);
            Marshal.ReleaseComObject(oWB2);
            oWB = null;
            oWB2 = null;
            oXL.DisplayAlerts = true;
            oXL.Quit();
            Marshal.ReleaseComObject(oXL);
            oXL = null;

            return "Success";
        }

        public string Excel_Copy_Data_UsingSheetName_AllAndValue(string excelFile1Path, string excelFile2Path, string strCopySheetName, string strPasteSheetName, string strCopyRange, string strPasteRange) {
            //Gets Excel and gets Activeworkbook and worksheet
            if (!File.Exists(excelFile1Path)) throw new Exception("Excel file not found");
            if (!File.Exists(excelFile2Path)) throw new Exception("Excel file not found");
            ExcelNS.Application oXL = new ExcelNS.Application();
            ExcelNS.Workbook oWB, oWB2;
            oXL.DisplayAlerts = false;
            oWB = oXL.Workbooks.Open(excelFile1Path, false, false);
            oWB2 = oXL.Workbooks.Open(excelFile2Path, false, false);
            oXL.Visible = true;

            ExcelNS._Worksheet wsCopySheet = null;
            ExcelNS._Worksheet wsPasteSheet = null;
            //Find the sheet
            foreach (var sheet in oWB.Sheets) {
                //Console.WriteLine(((ExcelNS._Worksheet)sheet).Name.ToLower());
                if (((ExcelNS._Worksheet)sheet).Name.ToLower().Trim() == strCopySheetName.ToLower().Trim()) {
                    wsCopySheet = (ExcelNS._Worksheet)sheet;
                    //Marshal.ReleaseComObject(sheet);

                    break;
                }
                Marshal.ReleaseComObject(sheet);
            }
            foreach (var sheet in oWB2.Sheets) {
                if (((ExcelNS._Worksheet)sheet).Name.ToLower().Trim() == strPasteSheetName.ToLower().Trim()) {
                    wsPasteSheet = (ExcelNS._Worksheet)sheet;
                    //Marshal.ReleaseComObject(sheet);
                    break;
                }
                Marshal.ReleaseComObject(sheet);
            }

            if (wsCopySheet == null) throw new System.Exception("Copy Sheet name is wrong/not exist");
            if (wsPasteSheet == null) throw new System.Exception("Paste Sheet name is wrong/not exist");
            
            ExcelNS.Range xlCopyRange = wsCopySheet.Range[strCopyRange];
            ExcelNS.Range xlPasteRange = wsPasteSheet.Range[strPasteRange];

            //First time paste all
            wsCopySheet.Activate();
            xlCopyRange.Copy();            
            wsPasteSheet.Activate();
            xlPasteRange.PasteSpecial(ExcelNS.XlPasteType.xlPasteAll);
            //Second time paste only values
            wsCopySheet.Activate();
            xlCopyRange.Copy();
            wsPasteSheet.Activate();
            xlPasteRange.PasteSpecial(ExcelNS.XlPasteType.xlPasteValues);

            Marshal.ReleaseComObject(xlCopyRange);
            Marshal.ReleaseComObject(xlPasteRange);
            xlCopyRange = null;
            xlPasteRange = null;

            Marshal.ReleaseComObject(wsCopySheet);
            Marshal.ReleaseComObject(wsPasteSheet);
            wsCopySheet = null;
            wsPasteSheet = null;

            oWB.Close();
            oWB2.Save();
            oWB2.Close();
            Marshal.ReleaseComObject(oWB);
            Marshal.ReleaseComObject(oWB2);
            oWB = null;
            oWB2 = null;
            oXL.DisplayAlerts = true;
            oXL.Quit();
            Marshal.ReleaseComObject(oXL);
            oXL = null;

            return "Success";
        }

        public void Excel_CreateNewExcel(string strExcelFileFullpath) {
            ExcelNS.Application app = new ExcelNS.Application();
            ExcelNS.Workbook wb = app.Workbooks.Add();
            wb.SaveAs(strExcelFileFullpath);
            wb.Close();            
            Marshal.FinalReleaseComObject(wb);
            wb = null;
            app.Quit();            
            Marshal.FinalReleaseComObject(app);
            app = null;
        }
        public void Excel_CreateNewSheet(ExcelNS.Workbook wb,string strNewSheetName)
        {
            wb.Sheets.Add(null,null,1);
        }




        public string Excel_GetHeader_Index(string strWorkBookName, string strHeaderText, string strSheetNumber = "1", string strHeaderRow_Number = "1") {
            //Gets Excel and gets Activeworkbook and worksheet
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel.Workbook oWB;
            oXL = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            //oXL.Visible = true;
            oWB = (Microsoft.Office.Interop.Excel.Workbook)oXL.ActiveWorkbook;
            bool blnWBFound = false;
            string strWBName,strHeaderIndex = "0";

            //Remove the extention if its already attached
            if (strWorkBookName.ToLower().Contains(".xlsx") == true) strWorkBookName = strWorkBookName.ToLower().Replace(".xlsx", "");
            if (strWorkBookName.ToLower().Contains(".xls") == true) strWorkBookName = strWorkBookName.ToLower().Replace(".xls", "");
            if (strWorkBookName.ToLower().Contains(".mhtml") == true) strWorkBookName = strWorkBookName.ToLower().Replace(".mhtml", "");

            //Loop throught the workbooks and find out is it there or not, if not then reply as error
            foreach (var wb in oXL.Workbooks) {
                Microsoft.Office.Interop.Excel.Workbook wbTem = (Microsoft.Office.Interop.Excel.Workbook)wb;
                strWBName = wbTem.Name.ToLower();
                if (strWBName == strWorkBookName + ".xls" || strWBName == strWorkBookName + ".xlsx" || strWBName == strWorkBookName + ".mhtml") {
                    blnWBFound = true;
                    oWB = wbTem;
                    break;
                }
            }

            if (blnWBFound == false) {
                return "Excel workbook" + strWorkBookName + " is not found";
            }
            oXL.Visible = true;

            //Find the title
            ExcelNS._Worksheet wsSheet = oWB.Sheets[Convert.ToInt32(strSheetNumber)];
            ExcelNS.Range xlRange = wsSheet.Range["A" + strHeaderRow_Number + ":" + GetColumnName(wsSheet.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column.ToString()) + strHeaderRow_Number];

            var Data = xlRange.Value;
            strHeaderText = strHeaderText.Trim().ToLower();
            //find the header column
            int i = 1;
            for (int j = 1; j <= oWB.Sheets[Convert.ToInt32(strSheetNumber)].UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column; j++) {
                if (Data[i, j] != null && Data[i, j].ToString().Trim().ToLower() == strHeaderText) {
                    strHeaderIndex = j.ToString() ;
                    break;
                }                
            }

            xlRange = null;
            wsSheet = null;

            return strHeaderIndex;
        }
        public string Excel_RunMacro(string strWorkBookName, string strMacroName) { //, string strArgument1,string strArgument2
            //Gets Excel and gets Activeworkbook and worksheet
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel.Workbook oWB;
            oXL = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            oXL.Visible = true;
            oWB = (Microsoft.Office.Interop.Excel.Workbook)oXL.ActiveWorkbook;
            bool blnWBFound = false;
            string strWBName, strHeaderIndex = "0";

            //Remove the extention if its already attached
            if (strWorkBookName.ToLower().Contains(".xlsm") == true) strWorkBookName = strWorkBookName.ToLower().Replace(".xlsm", "");
            if (strWorkBookName.ToLower().Contains(".xlsx") == true) strWorkBookName = strWorkBookName.ToLower().Replace(".xlsx", "");
            if (strWorkBookName.ToLower().Contains(".xls") == true) strWorkBookName = strWorkBookName.ToLower().Replace(".xls", "");
            if (strWorkBookName.ToLower().Contains(".mhtml") == true) strWorkBookName = strWorkBookName.ToLower().Replace(".mhtml", "");

            //Loop throught the workbooks and find out is it there or not, if not then reply as error
            foreach (var wb in oXL.Workbooks) {
                Microsoft.Office.Interop.Excel.Workbook wbTem = (Microsoft.Office.Interop.Excel.Workbook)wb;
                strWBName = wbTem.Name.ToLower();
                if (strWBName == strWorkBookName + ".xls" || strWBName == strWorkBookName + ".xlsx" || strWBName == strWorkBookName + ".mhtml" || strWBName == strWorkBookName + ".xlsm") {
                    blnWBFound = true;
                    oWB = wbTem;
                    break;
                }
            }

            if (blnWBFound == false) {
                return "Excel workbook" + strWorkBookName + " is not found";
            }

            //oWB.RunAutoMacros(ExcelNS.XlRunAutoMacro.)
            //oXL.Run(strMacroName, strArgument1, strArgument2);
            oXL.Run(strMacroName);
            return "Success";
        }
        ////Back up as on 2nd Feb,2019_8.40Pm
        //public void Excel_Delete_BlankColumns(string excelFilePath, string strHeaderText, string strCheckMisAlignedData = "No") {
        //    if (!File.Exists(excelFilePath)) throw new FileNotFoundException(excelFilePath);
        //    ExcelNS.Application oXL = new ExcelNS.Application();
        //    ExcelNS.Workbook oWB;
        //    ExcelNS.Worksheet oSht;
        //    oWB = oXL.Workbooks.Open(excelFilePath, false, false);
        //    oXL.Visible = true;
        //    oSht = oWB.Sheets[1];
        //    ExcelNS.Range xlRange = oSht.Range["A1:" + GetColumnName(oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column.ToString()) + oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row];

        //    var Data = xlRange.Value;

        //    //find the title row
        //    for (int i = 1; i <= xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row; i++) {
        //        for (int j = 1; j <= xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column; j++) {
        //            if (Data[i, j] != null && Data[i, j].ToString() == strHeaderText) {

        //                if (i != 1) {
        //                    //Delete the previous rows
        //                    ExcelNS.Range deleteRows = (ExcelNS.Range)xlRange.Range["1:" + Convert.ToString(i - 1)];
        //                    deleteRows.EntireRow.Delete(Shift: ExcelNS.XlDeleteShiftDirection.xlShiftUp);
        //                }

        //                //xlRange.Rows[1, i - 1].EntireRow.Delete(Shift: ExcelNS.XlDeleteShiftDirection.xlShiftUp);
        //                //change the i value to the highest to break the main loop
        //                i = xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row + 100;
        //                break;
        //            }
        //        }
        //    }


        //    //Find the Empty columns and delete
        //    xlRange = oSht.Range["A1:" + GetColumnName(oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column.ToString()) + oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row];
        //    xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeConstants).EntireColumn.Hidden = true;
        //    try {
        //        if (xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible) != null && xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible).Count > 0) {
        //            xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible).EntireColumn.Delete(ExcelNS.XlDeleteShiftDirection.xlShiftUp);
        //        }
        //    }
        //    catch (Exception e1) {

        //        e1 = e1;
        //    }


        //    xlRange.EntireColumn.Hidden = false;

        //    //Find the Empty columns and delete
        //    xlRange = oSht.Range["A1:" + GetColumnName(oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column.ToString()) + oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row];
        //    xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeConstants).EntireRow.Hidden = true;
        //    try {
        //        if (xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible) != null && xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible).Count > 0) {
        //            xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible).EntireRow.Delete(ExcelNS.XlDeleteShiftDirection.xlShiftUp);
        //        }
        //    }
        //    catch (Exception e) {
        //        e = e;
        //    }
        //    xlRange.EntireRow.Hidden = false;


        //    if (strCheckMisAlignedData.ToString().ToLower() == "yes") {
        //        Dictionary<int, bool> TitleClmns = new Dictionary<int, bool>();
        //        xlRange = oSht.Range["A1:" + GetColumnName(oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column.ToString()) + oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row];
        //        Data = xlRange.Value;
        //        int intFirstTitleIndex = 0;
        //        //Record all columns, whether its title or not
        //        for (int j = 1; j <= xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column; j++) {
        //            if (Data[1, j] != null && Data[1, j].ToString() != "") {
        //                if (intFirstTitleIndex == 0) intFirstTitleIndex = j;
        //                TitleClmns.Add(j, true);
        //            }
        //            else {
        //                TitleClmns.Add(j, false);
        //            }
        //        }

        //        //Loop through rows and check for the mis-aligned data
        //        //find the title row
        //        bool blnMiddleMissingData, blnFirstDataMissing, blnPreviousMiddleMissingData = false;
        //        int intMiddleMissingIndex, intPreviousMiddleMissingIndex = 0;
        //        int intTotalColumns = xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column;
        //        int intTotalRows = xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row;
        //        for (int i = 2; i <= intTotalRows; i++) {
        //            blnFirstDataMissing = false;
        //            blnMiddleMissingData = false;
        //            intMiddleMissingIndex = 0;
        //            for (int j = 1; j <= intTotalColumns; j++) {
        //                if (TitleClmns[j] == true) {
        //                    if (Data[i, j] == null || Data[i, j].ToString() == "") {
        //                        if (j == intFirstTitleIndex) {
        //                            blnFirstDataMissing = true;
        //                        }
        //                        else if (blnFirstDataMissing == false) {
        //                            blnMiddleMissingData = true;
        //                            if (intMiddleMissingIndex == 0) intMiddleMissingIndex = j;
        //                        }
        //                    }
        //                    else if (blnMiddleMissingData == true) {
        //                        blnMiddleMissingData = false;
        //                        intMiddleMissingIndex = 0;
        //                    }
        //                }

        //                //Assumption is that the previous line is mis-aligned, 
        //                //So cute paste this line data to the previous line from where missing data
        //                if (blnFirstDataMissing == true && blnPreviousMiddleMissingData == true && (Data[i, j] != null && Data[i, j].ToString() != "")) {
        //                    //Calculate the total cells need to be cut & paste
        //                    int intTotalCells = intTotalColumns - intPreviousMiddleMissingIndex;
        //                    ExcelNS.Range rngCut = oSht.Range[GetColumnName(j.ToString()) + i + ":" + GetColumnName((j + intTotalCells).ToString()) + i];
        //                    ExcelNS.Range rngPaste = oSht.Range[GetColumnName(intPreviousMiddleMissingIndex.ToString()) + (i - 1).ToString() + ":" + GetColumnName((j + intTotalCells).ToString()) + (i - 1)];
        //                    //Cut
        //                    //rngCut.Copy();
        //                    //Paste
        //                    rngPaste.Insert(ExcelNS.XlInsertShiftDirection.xlShiftToRight, rngCut.Cut());
        //                    //rngPaste.PasteSpecial(ExcelNS.XlPasteType.xlPasteAll, ExcelNS.XlPasteSpecialOperation.xlPasteSpecialOperationNone);
        //                    break;
        //                }
        //            }
        //            //Mark if the data is missing from the middle
        //            if (blnMiddleMissingData == true) {
        //                blnPreviousMiddleMissingData = true;
        //                intPreviousMiddleMissingIndex = intMiddleMissingIndex;
        //            }
        //        }

        //        //Find the Empty columns and delete
        //        xlRange = oSht.Range["A1:" + GetColumnName(oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column.ToString()) + oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row];
        //        xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeConstants).EntireColumn.Hidden = true;
        //        try {
        //            if (xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible) != null && xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible).Count > 0) {
        //                xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible).EntireColumn.Delete(ExcelNS.XlDeleteShiftDirection.xlShiftUp);
        //            }
        //        }
        //        catch (Exception e1) {

        //            e1 = e1;
        //        }

        //        xlRange.EntireColumn.Hidden = false;
        //        //Find the Empty columns and delete
        //        xlRange = oSht.Range["A1:" + GetColumnName(oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column.ToString()) + oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row];
        //        xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeConstants).EntireRow.Hidden = true;
        //        try {
        //            if (xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible) != null && xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible).Count > 0) {
        //                xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible).EntireRow.Delete(ExcelNS.XlDeleteShiftDirection.xlShiftUp);
        //            }
        //        }
        //        catch (Exception e) {
        //            e = e;
        //        }
        //        xlRange.EntireRow.Hidden = false;

        //    }
        //    oXL.DisplayAlerts = false;
        //    oWB.Save();
        //    oWB.Close();
        //    oXL.Quit();
        //}        
        public static string GetColumnName(string strColumnNumber) {
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
        public void Excel_DropLastRow(string pFilePath, string pSheetName)
        {
            using (ExcelHelper excel = new ExcelHelper(pFilePath, true, false, true))
            {
                excel.DropLastRow(pSheetName);
            }
        }
        public void XLS_To_XLSX_Converstion(string strHtmFileFullpath) {
            ExcelNS.Application app = new ExcelNS.Application();
            app.DisplayAlerts = false;
            ExcelNS.Workbook wb = app.Workbooks.Open(strHtmFileFullpath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            wb.SaveAs(strHtmFileFullpath.Replace(".xls", ".xlsx"), ExcelNS.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, Type.Missing, Type.Missing, ExcelNS.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            wb.Close();
            app.DisplayAlerts = true;
            app.Quit();
        }
        public void Excel_Walmart_Reconciliation(string pFilePath, string pSheetName)
        {
            using (ExcelHelper excel = new ExcelHelper(pFilePath, true, false, true))
            {

                System.Data.DataTable table = excel.GetData(pSheetName);

                List<string> copyList = new List<string>();
                copyList.Add("收货金额");
                copyList.Add("误差");

                List<string> sumList = new List<string>();
                sumList.Add("交货数量");
                sumList.Add("税");
                sumList.Add("净值");

                excel.SearchDifData(table, pSheetName, "采购单号", copyList, sumList);
            }
        }
        public void Excel_Filter_Delete_Row(string excelFilePath, string strHeaderText,string strFilterValue)
        {
            if (!File.Exists(excelFilePath)) throw new FileNotFoundException(excelFilePath);
            ExcelNS.Application oXL = new ExcelNS.Application();
            ExcelNS.Workbook oWB;
            ExcelNS.Worksheet oSht;
            oWB = oXL.Workbooks.Open(excelFilePath, false, false);
            oXL.Visible = true;
            oSht = oWB.Sheets[1];
            ExcelNS.Range xlRange = oSht.Range["A1:" + GetColumnName(oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column.ToString()) + oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row];
            var Data = xlRange.Value;

            //For this function title row should be row no "1"
            int i = 1;
            int intFilterClmn = 0;
            bool blnHeadingFound = false;
            for (int j = 1; j <= xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column; j++)
            {
                if (Data[i, j] != null && Data[i, j].ToString() == strHeaderText)
                {
                    intFilterClmn = j;
                    blnHeadingFound = true;
                    break;
                }
            }

            if (blnHeadingFound == false)
            {
                throw new Exception("Heading:" + strHeaderText + " could not found in the 1st row");
            }
            if (strFilterValue == "")
               {
                strFilterValue = null;
               }
            String[] FilterList = { strFilterValue };
            //Filter
            xlRange.AutoFilter(intFilterClmn, FilterList, ExcelNS.XlAutoFilterOperator.xlFilterValues);
            //Delete the values
            xlRange.Offset[1, 0].SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible).EntireRow.Delete();
            //Turn off the filter
            oSht.AutoFilterMode = false;

            //Delete the header rows, if its in the data...
            //Delete duplicate headers
            oXL.DisplayAlerts = false;
            oWB.Save();
            oWB.Close();
            oXL.Quit();
        }
        public void Excel_FilterByOpenFile_Delete_Row(ExcelNS.Worksheet oSht, string strHeaderText, string strFilterValue)
        {
            ExcelNS.Range xlRange = oSht.Range["A1:" + GetColumnName(oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column.ToString()) + oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row];
            var Data = xlRange.Value;

            //For this function title row should be row no "1"
            int i = 1;
            int intFilterClmn = 0;
            bool blnHeadingFound = false;
            for (int j = 1; j <= xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column; j++)
            {
                if (Data[i, j] != null && Data[i, j].ToString() == strHeaderText)
                {
                    intFilterClmn = j;
                    blnHeadingFound = true;
                    break;
                }
            }

            if (blnHeadingFound == false)
            {
                throw new Exception("Heading:" + strHeaderText + " could not found in the 1st row");
            }
            if (strFilterValue == "")
            {
                strFilterValue = null;
            }
            String[] FilterList = { strFilterValue };
            //Filter
            xlRange.AutoFilter(intFilterClmn, FilterList, ExcelNS.XlAutoFilterOperator.xlFilterValues);
            //Delete the values
            xlRange.Offset[1, 0].SpecialCells(ExcelNS.XlCellType.xlCellTypeVisible).EntireRow.Delete();
            //Turn off the filter
            oSht.AutoFilterMode = false;

            //Delete the header rows, if its in the data...
            //Delete duplicate headers
        }
        public void Excel_FilterbyList_Delete_Row(ExcelNS.Worksheet oSht ,string strHeaderText1, string strHeaderText2, Dictionary<string, string> dicFilter1, Dictionary<string, string> dicFilter2)
        {

            oSht.Activate();
            ExcelNS.Range xlRange = null;
            xlRange = oSht.Range["A10:" + GetColumnName(oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column.ToString()) + oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row];
            var Data = xlRange.Value;

            //For this function title row should be row no "1"
            int i = 1;
            int intFilterClmn1 = 0;
            int intFilterClmn2 = 0;
            bool blnHeadingFound1 = false, blnHeadingFound2= false;
            for (int j = 1; j <= xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column; j++)

            {
                if (Data[i, j] != null && Data[i, j].ToString() == strHeaderText1)
                {
                  

                    intFilterClmn1 = j;
                    blnHeadingFound1 = true;
                    break;
                }

            }
            for (int j = 1; j <= xlRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column; j++)

            {
                if (Data[i, j] != null && Data[i, j].ToString() == strHeaderText2)
                {


                    intFilterClmn2 = j;
                    blnHeadingFound2 = true;
                    break;
                }

            }
            if (blnHeadingFound1 == false || blnHeadingFound2 == false)
            {
                throw new Exception("GPPR Report Error");
            }

            //FilterList.Add("123");
            //
            //Filter
            //6938514100710
           
            int rowcount ;
            string strFilterValue1 , strFilterValue2;
            rowcount = oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row;
            for (int j= oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row; j>=11 ; j--)
            {

                rowcount = oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row;

                //xlRange = oSht.Range["A10:" + GetColumnName(oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Column.ToString()) + oSht.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row];
                //strFilterValue1  =Convert.ToString( xlRange.Cells[j, intFilterClmn1].Value);
                strFilterValue1 = Convert.ToString(oSht.Cells[j, intFilterClmn1].Value);
                // strFilterValue1 = strFilterValue1.Trim();


                strFilterValue2 = Convert.ToString(oSht.Cells[j, intFilterClmn2].value);
               // strFilterValue2 = strFilterValue2.Trim();
                if (dicFilter1.ContainsKey(strFilterValue1) == false && dicFilter2.ContainsKey(strFilterValue2) == false)
                {
                    xlRange = oSht.Range[j.ToString() + ":" + j.ToString()];

                    xlRange.Delete();

                }


            }


        
            //Delete the header rows, if its in the data...
       
        }



        public string Excel_Copy_Data_UsingSheetName_AllData(string excelFile1Path, string excelFile2Path, string strCopySheetName, string strPasteSheetName, string strCopyFromRow)
        {
            //Gets Excel and gets Activeworkbook and worksheet
            if (!File.Exists(excelFile1Path)) throw new Exception("Excel file not found");
            if (!File.Exists(excelFile2Path)) throw new Exception("Excel file not found");
            ExcelNS.Application oXL = new ExcelNS.Application();
            ExcelNS.Workbook oWB, oWB2;
            oXL.DisplayAlerts = false;
            oWB = oXL.Workbooks.Open(excelFile1Path, false, false);
            oWB2 = oXL.Workbooks.Open(excelFile2Path, false, false);
            oXL.Visible = true;

            ExcelNS._Worksheet wsCopySheet = null;
            ExcelNS._Worksheet wsPasteSheet = null;
            //Find the sheet
            foreach (var sheet in oWB.Sheets)
            {
                if (((ExcelNS._Worksheet)sheet).Name.ToLower().Trim() == strCopySheetName.ToLower().Trim())
                {
                    wsCopySheet = (ExcelNS._Worksheet)sheet;
                   // Marshal.ReleaseComObject(sheet);
                    break;
                }
                Marshal.ReleaseComObject(sheet);
            }
            foreach (var sheet in oWB2.Sheets)
            {
                if (((ExcelNS._Worksheet)sheet).Name.ToLower().Trim() == strPasteSheetName.ToLower().Trim())
                {
                    wsPasteSheet = (ExcelNS._Worksheet)sheet;
                    //Marshal.ReleaseComObject(sheet);
                    break;
                }
                Marshal.ReleaseComObject(sheet);
            }

            if (wsCopySheet == null) throw new Exception("Copy Sheet name is wrong/not exist");
            if (wsPasteSheet == null) throw new Exception("Paste Sheet name is wrong/not exist");

            ExcelNS.Range xlStartRange = wsCopySheet.Cells[Convert.ToInt32(strCopyFromRow),1];

            ExcelNS.Range xlEndRange = wsCopySheet.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell);
              
            ExcelNS.Range xlCopyRange = wsCopySheet.Range[xlStartRange, xlEndRange];
            wsCopySheet.Activate();
            xlCopyRange.Copy();
            ExcelNS.Range xlPasteRange = wsPasteSheet.Range["A" +  (wsPasteSheet.UsedRange.SpecialCells(ExcelNS.XlCellType.xlCellTypeLastCell).Row+1).ToString()];
            wsPasteSheet.Activate();
            xlPasteRange.PasteSpecial(ExcelNS.XlPasteType.xlPasteAll);


            Marshal.ReleaseComObject(xlCopyRange);
            Marshal.ReleaseComObject(xlPasteRange);
            xlCopyRange = null;
            xlPasteRange = null;

            Marshal.ReleaseComObject(wsCopySheet);
            Marshal.ReleaseComObject(wsPasteSheet);
            wsCopySheet = null;
            wsPasteSheet = null;

            oWB.Close();
            oWB2.Save();
            oWB2.Close();
            Marshal.ReleaseComObject(oWB);
            Marshal.ReleaseComObject(oWB2);
            oWB = null;
            oWB2 = null;
            oXL.DisplayAlerts = true;
            oXL.Quit();
            Marshal.ReleaseComObject(oXL);
            oXL = null;

            return "Success";
        }
    }
}
