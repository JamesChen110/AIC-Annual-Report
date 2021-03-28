using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using System.Diagnostics;
using System.Data.SqlClient;
using System.Data.Odbc;
using System.Globalization;

namespace AIC_Annual_Report
{
    class Connection

    {

        OdbcConnection Odbcconn = null;
        OleDbConnection OleDbConn = null;

        /// <summary>
        /// 连接Hive,返回Dt
        /// </summary>
        /// <param name="StrOfferID"></param>
        /// <param name="StrValidFrom"></param>
        /// <param name="StrValidTo"></param>
        /// <returns></returns>
        public DataTable GetDataFromHive(string StrOfferID, string StrValidFrom, string StrValidTo)
        {

            int RetryNum = 0;
            //OdbcConnection conn = null;
            DataTable dt = null;
            do
            {
                try
                {
                    RetryNum += 1;

                    System.Data.Odbc.OdbcConnectionStringBuilder conbuilder = new OdbcConnectionStringBuilder();
                    conbuilder.Driver = "Hortonworks Hive ODBC Driver";
                    conbuilder.Dsn = "HADOOP_CNPROD1_KNOX";
                    Odbcconn = new OdbcConnection(conbuilder.ConnectionString);
                    Odbcconn.ConnectionTimeout = 300;
                    Odbcconn.Open();
                    break;
                }

                catch (Exception ex)

                {
                    if (RetryNum == 3)
                    {
                        throw new Exception(@"Failed to connect to data source: " + Environment.CommandLine + ex.ToString());
                        
                    }

                }

            }
            while (RetryNum < 3);

            RetryNum = 0;
            do
            {
                try
                {
                    RetryNum += 1;
                    String sql;

                    //sql = "SELECT promo_pkg_id,promo_pkg_start_dt, promo_pkg_end_dt,dept_nbr,catg_nbr,fineline, CONCAT(lpad(CAST(vendor_nbr as varchar(6)),6, '0'), lpad(CAST(dept_nbr as varchar(6)),2, '0'), seq_nbr) AS vendor_nbr_9,vendor_name,store_nbr,  upc_nbr,item_nbr, cn_desc, tax_rate, -round(SUM(upc_discount),2)  AS upc_discount,-round(SUM(upc_discount*(tax_rate+1)),2)  AS upc_discount_tax, round(SUM(upc_sales),2) AS upc_sales,round(SUM(upc_sales*(tax_rate+1)),2) AS upc_sales_tax,SUM(upc_qty) AS upc_qty FROM   cn_ods_markdown.o_bas_transaction_smart_summary_d where promo_pkg_id=" + StrOfferID + "  and visit_date between '" + StrValidFrom + "' and '" + StrValidTo + "' GROUP BY promo_pkg_id, promo_pkg_start_dt, promo_pkg_end_dt,dept_nbr, catg_nbr, fineline, vendor_nbr, vendor_name, store_nbr, upc_nbr, item_nbr, cn_desc, tax_rate, seq_nbr";
                    //sql = "SELECT promo_pkg_id as `Offer ID`,promo_pkg_start_dt as `Promotion Begin Date` , promo_pkg_end_dt as `Promotion End Date`,dept_nbr as `Dept`,catg_nbr as `Category Nbr`,fineline as `Fineline`, CONCAT(substr(vendor_nbr , 6), substr(dept_nbr, 2), seq_nbr) AS `Vendor #(9位)`,vendor_name as `Vendor Name` ,store_nbr as `Store`, upc_nbr as `UPC`, item_nbr as `Item#`, cn_desc as `Item Desc.1`, tax_rate as `Tax%`, -round(SUM(upc_discount),2)  AS `Linksave MD$(未税)`,-round(SUM(upc_discount*(tax_rate+1)),2)  AS `Linksave MD$(含税)`, round(SUM(upc_sales),2) AS `Net Sales(含税）`,round(SUM(upc_sales*(tax_rate+1)),2) AS `Net Sales(未税）`, SUM(upc_qty) AS `POS Qty` FROM cn_ods_markdown.o_bas_transaction_smart_summary_d where promo_pkg_id=" + StrOfferID + "  and visit_date between '" + StrValidFrom + "' and '" + StrValidTo + "' GROUP BY visit_date, promo_pkg_id, store_nbr, promo_pkg_start_dt, promo_pkg_end_dt, dept_nbr, seq_nbr, catg_nbr, fineline, vendor_nbr, vendor_name, upc_nbr, item_nbr, cn_desc, tax_rate";
                    sql = "SELECT promo_pkg_id, promo_pkg_start_dt, promo_pkg_end_dt, dept_nbr, lpad(CAST(catg_nbr as varchar(2)), 2, '0') AS catg_nbr, fineline, vendor_nbr_9, vendor_name, store_nbr, upc_nbr, item_nbr, cn_desc, tax_rate, SUM(upc_discount) AS upc_discount, round(SUM(upc_discount) * (1 + tax_rate), 2) as upc_discount_with_tax,SUM(upc_sales) AS upc_sales, round(SUM(upc_sales) * (1 + tax_rate), 2) AS upc_sales_with_tax, SUM(upc_qty) AS upc_qty FROM cn_ods_markdown.o_bas_item_performance_report_view_d where promo_pkg_id=" + StrOfferID + "  and visit_date between '" + StrValidFrom + "' and '" + StrValidTo + "' GROUP BY promo_pkg_id, promo_pkg_start_dt,promo_pkg_end_dt, dept_nbr, catg_nbr, fineline, vendor_nbr_9, vendor_name, store_nbr, upc_nbr, item_nbr, cn_desc, tax_rate";
                    //sql = "SELECT promo_pkg_id as `Offer ID`,promo_pkg_start_dt as `Promotion Begin Date` , promo_pkg_end_dt as `Promotion End Date`,dept_nbr as `Dept`,catg_nbr as `Category Nbr`,fineline as `Fineline`, CONCAT(substr(vendor_nbr , 6), substr(dept_nbr, 2), seq_nbr) AS `Vendor #(9位)`,vendor_name as `Vendor Name`,store_nbr as `Store`,  upc_nbr as `UPC`, item_nbr as `Item#`, cn_desc as `Item Desc.1`, tax_rate as `Tax%`, round(SUM(upc_discount),2)  AS `Linksave MD$(未税)`,round(SUM(upc_discount*(tax_rate+1)),2)  AS `Linksave MD$(含税)`, round(SUM(upc_sales),2) AS `Net Sales(含税）`,round(SUM(upc_sales*(tax_rate+1)),2) AS `Net Sales(未税）`, SUM(upc_qty) AS `POS Qty` FROM cn_ods_markdown.o_bas_transaction_smart_summary_d where promo_pkg_id="+StrOfferID+"  and visit_date between '"+ StrValidFrom + "' and '" + StrValidTo + "' GROUP BY visit_date, promo_pkg_id, promo_pkg_start_dt, promo_pkg_end_dt, store_nbr, dept_nbr, seq_nbr, catg_nbr, fineline, vendor_nbr, vendor_name, upc_nbr, item_nbr, cn_desc, tax_rate";
                    
                    //sql = "SELECT promo_pkg_id,promo_pkg_start_dt, promo_pkg_end_dt,dept_nbr,catg_nbr,fineline, CONCAT(substr(vendor_nbr , 6), substr(dept_nbr, 2), seq_nbr) AS vendor_nbr_9,vendor_name,store_nbr,  upc_nbr, item_nbr, cn_desc, tax_rate, round(SUM(upc_discount),2)  AS upc_discount,round(SUM(upc_discount*(tax_rate+1)),2)  AS upc_discount_tax, round(SUM(upc_sales),2) AS upc_sales,round(SUM(upc_sales*(tax_rate+1)),2) AS upc_sales_tax, SUM(upc_qty) AS upc_qty FROM cn_ods_markdown.o_bas_transaction_smart_summary_d where promo_pkg_id="+StrOfferID+"  and visit_date between '"+ StrValidFrom + "' and '" + StrValidTo + "' GROUP BY visit_date, promo_pkg_id, promo_pkg_start_dt, promo_pkg_end_dt, store_nbr, dept_nbr, seq_nbr, catg_nbr, fineline, vendor_nbr, vendor_name, upc_nbr, item_nbr, cn_desc, tax_rate";

                    OdbcDataAdapter myda = new OdbcDataAdapter();
                    OdbcCommand myCom = new OdbcCommand(sql, Odbcconn);
                    myCom.CommandTimeout = 1800;
                    myda.SelectCommand = myCom;

                    dt= new DataTable(); //新建表对象
                    myda.Fill(dt); //用适配对象填充表对象
                                   //Debug.Print(dt.Rows[0][1].ToString());
                    Odbcconn.Close();
                    break;

                }

                catch (Exception ex)

                {
                    if (RetryNum == 3)
                    {
                        throw new Exception(@"SQL is Error: " + Environment.CommandLine + ex.ToString());

                    }

                }

            }
            while (RetryNum < 3);

            return dt;


        }            
    


        /// <summary>
        /// 读取Access中数据
        /// </summary>
        /// <param name="strDBFileFolderPath"></param>
        /// <param name="strDBFileName"></param>
        /// <param name="sql"></param>
        /// <returns></returns>



        public DataTable GetDataFromAccess(string strDBFileFolderPath, string sql = "select * from ams")
        {
            //OleDbConnection oleDb = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\c0c04nc\Documents\C# Access SQL\ams1.mdb");
           // strDBFileName= @"C:\Users\c0c04nc\Documents\C# Access SQL\ams.accdb";
            OleDbConnectionStringBuilder connStrBuild = new OleDbConnectionStringBuilder();
            connStrBuild.Provider = "Microsoft.ACE.OLEDB.12.0";
            connStrBuild.DataSource = strDBFileFolderPath; //+ strDBFileName;
            OleDbConnection conn = new OleDbConnection(connStrBuild.ConnectionString);

            conn.Open();
            OleDbDataAdapter dbDataAdapter = new OleDbDataAdapter(sql, conn); //创建适配对象
            DataTable dt = new DataTable(); //新建表对象
            dbDataAdapter.Fill(dt); //用适配对象填充表对象
            Debug.Print(dt.Rows[0][1].ToString());
            conn.Close();
            return dt;

            //OleDbConnection oleDb = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\c0c04nc\Documents\C# Access SQL\ams1.mdb; Persist Security Info = True");

            //获取表1的内容
            //oleDb.Open();
            //OleDbDataAdapter dbDataAdapter = new OleDbDataAdapter(sql, oleDb); //创建适配对象
            //DataTable dt = new DataTable(); //新建表对象
            //dbDataAdapter.Fill(dt); //用适配对象填充表对象
            //Debug.Print(dt.Rows[0][1].ToString());
            //oleDb.Close();
            //return dt;
        }
        /// <summary>
        /// 连接 SQL Server
        /// </summary>
        /// <returns></returns>
        public DataTable SQLConn()
        {
            
            SqlConnection Conn = new SqlConnection(@"Server=pcnnt57047sql\ni1;Database=MarkDown;User ID=svc_dax081;Password=4P%StKC&rJY%C6%r&JWK&tb1k");

            Conn.Open();
            string con, sql;
            con = "Server=.;Database=Exercise;Trusted_Connection=SSPI";
            sql = "SELECT promo_pkg_id,promo_pkg_start_dt, promo_pkg_end_dt,dept_nbr,catg_nbr,fineline, CONCAT(substr(vendor_nbr , 6), substr(dept_nbr, 2), seq_nbr) AS vendor_nbr_9,vendor_name,store_nbr,  upc_nbr, item_nbr, cn_desc, tax_rate, round(SUM(upc_discount),2)  AS upc_discount,round(SUM(upc_discount*(tax_rate+1)),2)  AS upc_discount_tax, round(SUM(upc_sales),2) AS upc_sales,round(SUM(upc_sales*(tax_rate+1)),2) AS upc_sales_tax, SUM(upc_qty) AS upc_qty FROM cn_ods_markdown.o_bas_transaction_smart_summary_d where promo_pkg_id=10655  and visit_date between '2020-01-22' and '2020-01-22' GROUP BY visit_date, promo_pkg_id, promo_pkg_start_dt, promo_pkg_end_dt, store_nbr, dept_nbr, seq_nbr, catg_nbr, fineline, vendor_nbr, vendor_name, upc_nbr, item_nbr, cn_desc, tax_rate";
            //SqlConnection mycon = new SqlConnection(con);
            Conn.Open();
            SqlDataAdapter myda = new SqlDataAdapter(sql, con);
            DataTable dt = new DataTable(); //新建表对象
            myda.Fill(dt); //用适配对象填充表对象
            Debug.Print(dt.Rows[0][1].ToString());
            Conn.Close();
            return dt;


        }
        /// <summary>
        /// 读取Excel中数据
        /// </summary>
        /// <param name="strExcelPath"></param>
        /// <param name="tableName"></param>
        /// <returns></returns>
        /// //static
        public DataTable GetExcelTableByOleDB(string strExcelPath, string tableName)
        {
            //try
            //{
                DataTable dtExcel = new DataTable();
                //数据表
                DataSet ds = new DataSet();
                //获取文件扩展名
                string strExtension = System.IO.Path.GetExtension(strExcelPath);
                string strFileName = System.IO.Path.GetFileName(strExcelPath);
                //Excel的连接
                
                switch (strExtension)
                {
                    case ".xls":
                    OleDbConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strExcelPath + ";" + "Extended Properties=\"Excel 8.0;HDR=yes;IMEX=1;\"");
                        break;
                    case ".xlsx":

                    OleDbConn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strExcelPath + ";" + "Extended Properties=\"Excel 8.0;HDR=yes;IMEX=1;\"");
                        //此连接可以操作.xls与.xlsx文件 (支持Excel2003 和 Excel2007 的连接字符串)  备注： "HDR=yes;"是说Excel文件的第一行是列名而不是数，"HDR=No;"正好与前面的相反。"IMEX=1 "如果列中的数据类型不一致，使用"IMEX=1"可必免数据类型冲突。 
                        break;
                    default:
                    OleDbConn = null;
                        break;
                }
                if (OleDbConn == null)
                {
                    return null;
                }
            OleDbConn.Open();
                //获取Excel中所有Sheet表的信息
                //System.Data.DataTable schemaTable = objConn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);
                //获取Excel的第一个Sheet表名
                // string tableName1 = schemaTable.Rows[0][2].ToString().Trim();
                string strSql = "select * from [" + tableName + "$]";
                //获取Excel指定Sheet表中的信息
                OleDbCommand objCmd = new OleDbCommand(strSql, OleDbConn);
                OleDbDataAdapter myData = new OleDbDataAdapter(strSql, OleDbConn);
                myData.Fill(ds, tableName);//填充数据
            OleDbConn.Close();
                //dtExcel即为excel文件中指定表中存储的信息
                dtExcel = ds.Tables[tableName];
                Debug.Print(dtExcel.Rows[0][0].ToString());
                return dtExcel;
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex.Message + "\r\n" + ex.StackTrace);
            //    return null;
            //}
        }




    }
}
