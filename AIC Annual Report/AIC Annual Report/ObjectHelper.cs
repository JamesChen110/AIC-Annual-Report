using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Threading.Tasks;
using System.Diagnostics;

namespace AIC_Annual_Report
{
    public class ObjectHelper
    {
        /// <summary>
        /// Object转DataTable
        /// </summary>
        /// <param name="pObj"></param>
        /// <param name="pIncludeTitle"></param>
        /// <returns></returns>
        public static DataTable ObjectToDataTable(object[,] pObj, bool pIncludeTitle = true, DataTable pDataTable = null)
        {
            DataTable dt = new DataTable();

            int oneDimensional = pObj.GetLength(0);
            int twoDimensional = pObj.GetLength(1);

            //列名
            if (pDataTable != null)
            {
                dt = pDataTable.Clone();
            }
            else
            {
                for (int i = 1; i <= twoDimensional; i++)
                {
                    object objItem = pObj[1, i];
                    string columnName = (objItem == null || objItem.ToString() == "") || !pIncludeTitle ? "Column" + i : objItem.ToString();
                    dt.Columns.Add(columnName);
                }
            }

            //数据
            for (int i = (pIncludeTitle ? 2 : 1); i <= oneDimensional; i++)
            {
                dt.Rows.Add();
                for (int j = 1; j <= twoDimensional; j++)
                    dt.Rows[dt.Rows.Count - 1][j - 1] = pObj[i, j];
            }

            return dt;
        }
        /// <summary>
        /// string转decimal
        /// </summary>
        /// <param name="Str"></param>
        /// <returns></returns>
        public static decimal ConvertToDecimal(string Str)
        {
            decimal result = 0;
            if (Str.ToUpper().Contains("E"))
            {
                result = Convert.ToDecimal(Decimal.Parse(Str, System.Globalization.NumberStyles.Float));
            }
            return result;
        }

        public static List<T> DataColumnToList<T>(DataTable pData, string pColName)
        {
            return pData.AsEnumerable().Select(m => m.Field<T>(pColName)).ToList();
        }

        public static string ConvertTime(string pTime)
        {
            try
            {
                int H = (int)Math.Truncate(decimal.Parse(pTime) * 24);
                int mm = (int)Math.Truncate((decimal.Parse(pTime) * 24 * 60) % 60);
                int ss = Convert.ToInt16((decimal.Parse(pTime) * 24 * 60 * 60) % 60 % 60);

                if (ss == 60)
                {
                    mm = mm + 1;
                    ss = 0;
                }
                if (mm == 60)
                {
                    H = H + 1;
                    mm = 0;
                }
                return H + ":" + mm + ":" + ss;
            }
            catch (Exception ex)
            {

                throw;
            }
        }


    }
}
