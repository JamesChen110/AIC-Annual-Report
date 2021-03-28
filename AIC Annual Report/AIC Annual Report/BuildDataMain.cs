using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.IO;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office;
using AIC_Annual_Report_BuildData;
using DataTable = System.Data.DataTable;
using System.Reflection;
using System.Windows.Forms;

namespace AIC_Annual_Report
{
    public class BuildDataMain
    {

        #region set variable
        
        string StrConfigFilePath;
        int MaxRetryNum = 1;

        string strEmailbody;
        ControlEmail controlemail = null;
        Dictionary<string, string> dicConfig = new Dictionary<string, string>();
        Dictionary<string, clsLoginInfo> dicCorporateLoginInfo = new Dictionary<string, clsLoginInfo>();
        /// <summary>
        /// 法人（公司信息--key:store and value:dic-->key:Icode(内码) and value :class(clsStandartTempInfo)
        /// </summary>
        Dictionary<string, Dictionary<string, clsStandartTempInfo>> dicCorporate = new Dictionary<string, Dictionary<string, clsStandartTempInfo>>();
        Dictionary<string, string> dicCorporateDefaultInfo = new Dictionary<string, string>();

        Dictionary<string, string> dicCorporateCalvalue = new Dictionary<string, string>();

        Dictionary<string, string> dicD0XXtoStoreNo = new Dictionary<string, string>();

        /// <summary>
        /// 3 法人字段映射-内码和沃尔玛报表
        /// </summary>
        Dictionary<string, clsIcodeVSSourceData> dicCorporateFieldMap = new Dictionary<string, clsIcodeVSSourceData>();
        /// <summary>
        /// 财务信息数据，key:store-item1(item2,3)
        /// </summary>
        Dictionary<string, string> dicFinancialInforSourceData = new Dictionary<string, string>();

        /// <summary>
        /// 关税 和 非关税信息
        /// </summary>
        Dictionary<string, string> dicTaxInforSourceData = new Dictionary<string, string>();



        /// <summary>
        /// 特种设备
        /// </summary>
        Dictionary<string, string> dicSpecialEquipmentInforSourceData = new Dictionary<string, string>();

        /// <summary>
        /// HR
        /// </summary>
        Dictionary<string, string> dicHRInforSourceData = new Dictionary<string, string>();

        //登录信息
        string strCorporateLoginInfoPath; //= @"C:\Users\c0c04nc\Documents\Project\合规年报\AIC Annual Report Build Data\Source\(工具结构化)法人公司登录信息汇总表new - 副本.xlsx";
        string strBranchLoginInfoPath;
        //登录信息


        string strStandardFieldtemplatePath; //= @"C:\Users\c0c04nc\Documents\Project\合规年报\AIC Annual Report Build Data\Source\年报字段.xlsx";
        string strSpecialEquipmentFilePath; //= @"C:\Users\c0c04nc\Documents\Project\合规年报\AIC Annual Report Build Data\Source\特种设备\特种设备汇总表(2020年合规工商年报模板)-仅辽宁1省需明细.xlsx";
        string strHROBasicInfoDetailFilePath; // = @"C:\Users\c0c04nc\Documents\Project\合规年报\AIC Annual Report Build Data\Source\HR\Associate Summary Info Report Wed Jan 06 2021.xlsx";
        string stroutputCorporateData; //= @"C:\Users\c0c04nc\Documents\Project\合规年报\AIC Annual Report Build Data\result\UploadCorporateData.xlsx";
        string strCompanySummaryFilePath; //= @"C:\Users\c0c04nc\Documents\Project\合规年报\AIC Annual Report Build Data\Source\（D0XX-4位店号）法人基表.xlsx";
        string strNonTariffsFilePath; //= @"C:\Users\c0c04nc\Documents\Project\合规年报\AIC Annual Report Build Data\Source\Tax\非关税\2020.11.28 2020年1~11月 megan cai 税金取数.xlsx";
        string strTariffsFilePath; //= @"C:\Users\c0c04nc\Documents\Project\合规年报\AIC Annual Report Build Data\Source\Tax\非关税\2020.11.28 2020年1~11月 megan cai 税金取数.xlsx";

        string strSocialSecurProFundInfoDetailFilePath; //= @"C:\Users\c0c04nc\Documents\Project\合规年报\AIC Annual Report Build Data\Source\HR\社保公积金缴纳明细表201901_C001_SI&PHF Detail Report_变式_Sample.XLSX";

        string strLevelOfEducationInfoDetailFilePath;//= @"C:\Users\c0c04nc\Documents\Project\合规年报\AIC Annual Report Build Data\Source\HR\教育及外语水平报表_Education and Language Proficiency Report_ Fri Dec 20 2019_Sample.xlsx";
        string strFinancialInforFilePath;// = @"C:\Users\c0c04nc\Documents\Project\合规年报\AIC Annual Report Build Data\Source\法人公司财务数据.xlsx";


        string strHRMappingFilePath;// = @"C:\Users\c0c04nc\Documents\Project\合规年报\AIC Annual Report Build Data\Source\配置表\HR\HR人事单位-门店匹配表.xlsx";
        string strBand12InfoDetailFilePath;// = @"C:\Users\c0c04nc\Documents\Project\合规年报\AIC Annual Report Build Data\Source\HR\HR\Band12员工信息-8028-手工.xlsx";
        string strHROtherFromStoreInfoDetailFilePath;// = @"C:\Users\c0c04nc\Documents\Project\合规年报\AIC Annual Report Build Data\Source\HR\HR\HR其他数据-门店-手工.xlsx";
        string strPayrollInfoDetailFilePath;//= @"C:\Users\c0c04nc\Documents\Project\合规年报\AIC Annual Report Build Data\Source\HR\HR\工资明细表-SAP-机器人.xlsx";

        string strBranchFinancialInforFilePath;
        string strHRInsuranceBaseFilePath;

        DataTable tblCorporateData = new DataTable();

        public string strDataType;
        public string strBuildResult;
        public UploadDataMain uploaddatadain = new UploadDataMain();

        //public form_Login from = new form_Login();

        //Main_testing
        public string strCurLog;

        #endregion
        public void Main_BuildData()
        {
            try
            {
                strBuildResult = "Running";
                //strBuildResult = "Succeed";
                strCurLog = "开始生成" + strDataType + "数据";
                StrConfigFilePath = @"\\cnnts8005fs\Private\Finance\GBS Innovations\Project\AIC Annual Report\AIC工商年报程序读取目录\Robot\Config\生成数据Config.xlsx";
                //"C:\Users\c0c04nc\Documents\Project\合规年报\AIC Annual Report Build Data\Config.xlsx"
                
                Config();

                strCorporateLoginInfoPath = dicConfig["strCorporateLoginInfoPath"];
                strBranchLoginInfoPath = dicConfig["strBranchLoginInfoPath"];
                strStandardFieldtemplatePath = dicConfig["strStandardFieldtemplatePath"];
                strSpecialEquipmentFilePath = dicConfig["strSpecialEquipmentFilePath"];
                strHROBasicInfoDetailFilePath = dicConfig["strHROBasicInfoDetailFilePath"];
                stroutputCorporateData = dicConfig["stroutputCorporateData"];
                strCompanySummaryFilePath = dicConfig["strCompanySummaryFilePath"];
                strNonTariffsFilePath = dicConfig["strNonTariffsFilePath"];
                strTariffsFilePath = dicConfig["strTariffsFilePath"];
                strSocialSecurProFundInfoDetailFilePath = dicConfig["strSocialSecurProFundInfoDetailFilePath"];
                strFinancialInforFilePath = dicConfig["strFinancialInforFilePath"];
                strLevelOfEducationInfoDetailFilePath = dicConfig["strLevelOfEducationInfoDetailFilePath"];
                strHRMappingFilePath = dicConfig["strHRMappingFilePath"];// = @"C:\Users\c0c04nc\Documents\Project\合规年报\AIC Annual Report Build Data\Source\配置表\HR\HR人事单位-门店匹配表.xlsx";
                strBand12InfoDetailFilePath = dicConfig["strBand12InfoDetailFilePath"];// = @"C:\Users\c0c04nc\Documents\Project\合规年报\AIC Annual Report Build Data\Source\HR\HR\Band12员工信息-8028-手工.xlsx";
                strHROtherFromStoreInfoDetailFilePath = dicConfig["strHROtherFromStoreInfoDetailFilePath"];// = @"C:\Users\c0c04nc\Documents\Project\合规年报\AIC Annual Report Build Data\Source\HR\HR\HR其他数据-门店-手工.xlsx";
                strPayrollInfoDetailFilePath = dicConfig["strPayrollInfoDetailFilePath"];//= @"C:\Users\c0c04nc\Documents\Project\合规年报\AIC Annual Report Build Data\Source\HR\HR\工资明细表-SAP-机器人.xlsx";
                strBranchFinancialInforFilePath = dicConfig["strBranchFinancialInforFilePath"];
                strHRInsuranceBaseFilePath = dicConfig["strHRInsuranceBaseFilePath"];

                //strDataType= dicConfig["strDataType"];

                CheckMoveDataSource();

                //分 法人/分支机构/全部
                if (strDataType == "法人")
                {

                    buildBasicData();
                    buildLoginInfo(strCorporateLoginInfoPath);
                    BuildCorporateHRData();
                    TaxInfor();
                    //读取财务信息
                    FinancialInfor(strFinancialInforFilePath);
                    SpecialEquipment();

                    outputCorporateData();
                    //法人

                }

                else if (strDataType == "分支机构")
                {

                    
                    // 1法人（公司信息）2法人字段映射-内码和沃尔玛报表 3法人-默认值字段表

                    buildBasicData();
                    buildLoginInfo(strBranchLoginInfoPath);
                    BuildCorporateHRData();
                    //读取财务信息
                    FinancialInfor(strBranchFinancialInforFilePath);
                    



                    TaxInfor();
                    SpecialEquipment();

                    outputCorporateData();
                    //分支
                }
                else
                {
                    //法人
                    strDataType = "法人";
                    buildBasicData();
                    buildLoginInfo(strCorporateLoginInfoPath);
                    BuildCorporateHRData();
                    TaxInfor();
                    //读取财务信息
                    FinancialInfor(strFinancialInforFilePath);
                    SpecialEquipment();

                    outputCorporateData();
                    //法人

                    //分支
                    strDataType = "分支机构";
                    // 1法人（公司信息）2法人字段映射-内码和沃尔玛报表 3法人-默认值字段表
                    buildBasicData();
                    buildLoginInfo(strBranchLoginInfoPath);
                    //读取财务信息
                    FinancialInfor(strBranchFinancialInforFilePath);
                    BuildCorporateHRData();



                    TaxInfor();
                    SpecialEquipment();

                    outputCorporateData();
                    //分支


                }



                strBuildResult = "Succeed";

                //MessageBox.Show("运行成功");

            }
            catch (Exception ex)
            {
                strCurLog = ex.Message;
                strBuildResult = "Fail";
                // MessageBox.Show("程序出错"+ex.ToString());
            }


        }

        /// <summary>
        /// 生成上传到法人年报网站数据
        /// </summary>
        public void CheckMoveDataSource()
        {
            Boolean blAllfilesExists = true;
            //AIC基本信息
            if (!File.Exists(strStandardFieldtemplatePath))
            {
                blAllfilesExists = false;
                strCurLog = strCurLog + Environment.NewLine + "这个文件不存在： " + strStandardFieldtemplatePath;
                //throw new Exception("这个文件不存在： " + strStandardFieldtemplatePath);
            }
            //AIC基本信息
            //登录信息表
            if (strDataType != "法人" && !File.Exists(strBranchLoginInfoPath))
            {
                blAllfilesExists = false;
                strCurLog = strCurLog + Environment.NewLine + "这个文件不存在： " + strBranchLoginInfoPath;
            }
            if (strDataType != "分支机构" && !File.Exists(strCorporateLoginInfoPath))
            {
                blAllfilesExists = false;
                strCurLog = strCurLog + Environment.NewLine + "这个文件不存在： " + strCorporateLoginInfoPath;

            }
            //登录信息表

            //HR
            if (!File.Exists(strHRMappingFilePath))
            {
                blAllfilesExists = false;
                strCurLog = strCurLog + Environment.NewLine + "这个文件不存在： " + strHRMappingFilePath;
            }

            if (!File.Exists(strPayrollInfoDetailFilePath))
            {

                blAllfilesExists = false;
                strCurLog = strCurLog + Environment.NewLine + "这个文件不存在： " + strPayrollInfoDetailFilePath;
                //throw new Exception("这个文件不存在： " + strPayrollInfoDetailFilePath);
            }

            if (!File.Exists(strHROBasicInfoDetailFilePath))
            {
                blAllfilesExists = false;
                strCurLog = strCurLog + Environment.NewLine + "这个文件不存在： " + strHROBasicInfoDetailFilePath;

                //throw new Exception("这个文件不存在： " + strHROBasicInfoDetailFilePath);
            }

            if (!File.Exists(strLevelOfEducationInfoDetailFilePath))
            {

                blAllfilesExists = false;
                strCurLog = strCurLog + Environment.NewLine + "这个文件不存在： " + strLevelOfEducationInfoDetailFilePath;
            }

            if (!File.Exists(strBand12InfoDetailFilePath))
            {
                blAllfilesExists = false;
                strCurLog = strCurLog + Environment.NewLine + "这个文件不存在： " + strBand12InfoDetailFilePath;
            }

            if (!File.Exists(strHROtherFromStoreInfoDetailFilePath))
            {
                blAllfilesExists = false;
                strCurLog = strCurLog + Environment.NewLine + "这个文件不存在： " + strHROtherFromStoreInfoDetailFilePath;
            }


            //HR


            //财务

            if (strDataType != "分支机构" && !File.Exists(strFinancialInforFilePath))
            {
                blAllfilesExists = false;
                strCurLog = strCurLog + Environment.NewLine + "这个文件不存在： " + strFinancialInforFilePath;
            }
            else if (strDataType != "法人" && !File.Exists(strBranchFinancialInforFilePath))
            {
                blAllfilesExists = false;
                strCurLog = strCurLog + Environment.NewLine + "这个文件不存在： " + strBranchFinancialInforFilePath;
            }
            //财务

            //税务
            if (!File.Exists(strNonTariffsFilePath))
            {
                blAllfilesExists = false;
                strCurLog = strCurLog + Environment.NewLine + "这个文件不存在： " + strNonTariffsFilePath;
                
            }

            if (strDataType !="分支机构" && !File.Exists(strTariffsFilePath))
            {
                blAllfilesExists = false;
                strCurLog = strCurLog + Environment.NewLine + "这个文件不存在： " + strTariffsFilePath;
            }

            //税务
            //特种设备
            if (!File.Exists(strSpecialEquipmentFilePath))
            {
                blAllfilesExists = false;
                strCurLog = strCurLog + Environment.NewLine + "这个文件不存在： " + strSpecialEquipmentFilePath;

            }
            if (!blAllfilesExists)
            {
                throw new Exception(strCurLog);
            }
        }
        private void BuildCorporateHRData()
        {
            dicHRInforSourceData = new Dictionary<string, string>();
            #region 0.HR人事范围-门店匹配表 or HR人事范围-公司代码匹配表
            //0.HR人事范围-门店匹配表 or HR人事范围-公司代码匹配表
            Dictionary<string, string> dicBasicTable;
            using (BuildDataExcelHelper excel = new BuildDataExcelHelper(strHRMappingFilePath, true, true, false))
            {
                DataTable tblBasicTable = excel.GetData("Sheet1");
                dicBasicTable = new Dictionary<string, string>();
                foreach (DataRow drItem in tblBasicTable.Rows)
                {

                    if (strDataType == "法人")
                    {
                        string strComCode = drItem["公司代码"].ToString().Trim();
                        if (dicD0XXtoStoreNo.ContainsKey(strComCode))
                        {
                            string strPersonnelArea = drItem["人事范围"].ToString().Trim();

                            if (strPersonnelArea != "")
                            {
                                string strStore = dicD0XXtoStoreNo[strComCode];
                                if (!dicBasicTable.ContainsKey(strPersonnelArea))
                                {
                                    dicBasicTable.Add(strPersonnelArea, strStore);
                                }
                            }
                        }
                    }
                    else
                    {
                        string strStore = drItem["店号"].ToString().Trim();
                        string strPersonnelArea = drItem["人事范围"].ToString().Trim();
                        if (strStore != "" && strPersonnelArea != "")
                        {
                            strStore = strStore.PadLeft(4, '0');

                            if (!dicBasicTable.ContainsKey(strPersonnelArea))
                            {
                                dicBasicTable.Add(strPersonnelArea, strStore);
                            }
                        }
                    }

                }

            }
            #endregion
            #region 2.社保公积金缴纳明细表
            //2.社保公积金缴纳明细表
            Dictionary<string, string> dicUserIDCount = new Dictionary<string, string>();
            foreach (FileInfo NextFile in new DirectoryInfo(strSocialSecurProFundInfoDetailFilePath).GetFiles())
            {
                if (NextFile.Extension.ToLower() == ".xlsx")
                {
                    using (BuildDataExcelHelper excel = new BuildDataExcelHelper(NextFile.FullName, true, true, false))
                    {
                        DataTable tblBasicTable = excel.GetData("Sheet1");
                        string strStore;
                        String strKey;
                        String strValue;
                        foreach (DataRow drItem in tblBasicTable.Rows)
                        {
                            string strTemp = drItem["人事范围"].ToString().Trim();

                            string strMonth= drItem["社保月份"].ToString().Trim();
                            string strUserId= drItem["员工号"].ToString().Trim();
                            if (strTemp != "" && dicBasicTable.ContainsKey(strTemp))
                            {
                                strStore = dicBasicTable[strTemp].PadLeft(4, '0');
                                //只取12月份的人数，因为为了和年底行业人数数目一致，而且这个表取出的数，如果这个员工在今年内都在这里工作，那么他在12月份有十二条数，所以要用字典去除
                                //这个仅用于人数统计，和金额统计无关
                                if (strMonth.Substring(4, 2) == "12" && ! dicUserIDCount.ContainsKey(strUserId))
                                {
                                    dicUserIDCount.Add(strUserId, "");
                                    //count_/314 养老保险金（公司）_金额
                                    strValue = drItem["/314 养老保险金（公司）_金额"].ToString().Trim();
                                    strKey = strStore + "-" + "count_/314 养老保险金（公司）_金额";
                                    if (strValue != "" && double.Parse(strValue) > 0)
                                    {
                                        strValue = "1";
                                    }
                                    else
                                    {
                                        strValue = "0";
                                    }
                                    SetDicValue(strValue, strKey, dicHRInforSourceData);
                                    //count_/324 失业保险金（公司）_金额
                                    strValue = drItem["/324 失业保险金（公司）_金额"].ToString().Trim();
                                    strKey = strStore + "-" + "count_/324 失业保险金（公司）_金额";
                                    if (strValue != "" && double.Parse(strValue) > 0)
                                    {
                                        strValue = "1";
                                    }
                                    else
                                    {
                                        strValue = "0";
                                    }
                                    SetDicValue(strValue, strKey, dicHRInforSourceData);

                                    //count_/334 医疗保险金（公司）_金额
                                    strValue = drItem["/334 医疗保险金（公司）_金额"].ToString().Trim();
                                    strKey = strStore + "-" + "count_/334 医疗保险金（公司）_金额";
                                    if (strValue != "" && double.Parse(strValue) > 0)
                                    {
                                        strValue = "1";
                                    }
                                    else
                                    {
                                        strValue = "0";
                                    }
                                    SetDicValue(strValue, strKey, dicHRInforSourceData);
                                    //count_/344 工伤保险金（公司）_金额
                                    strValue = drItem["/344 工伤保险金（公司）_金额"].ToString().Trim();
                                    strKey = strStore + "-" + "count_/344 工伤保险金（公司）_金额";
                                    if (strValue != "" && double.Parse(strValue) > 0)
                                    {

                                        strValue = "1";

                                    }
                                    else
                                    {
                                        strValue = "0";
                                    }
                                    SetDicValue(strValue, strKey, dicHRInforSourceData);
                                    //count_/354 生育保险金（公司）_金额
                                    strValue = drItem["/354 生育保险金（公司）_金额"].ToString().Trim();
                                    strKey = strStore + "-" + "count_/354 生育保险金（公司）_金额";
                                    if (strValue != "" && double.Parse(strValue) > 0)
                                    {

                                        strValue = "1";

                                    }
                                    else
                                    {
                                        strValue = "0";
                                    }
                                    SetDicValue(strValue, strKey, dicHRInforSourceData);


                                }



                                //sum_/314 养老保险金（公司）_金额
                                strValue = drItem["/314 养老保险金（公司）_金额"].ToString().Trim();
                                strKey = strStore + "-" + "sum_/314 养老保险金（公司）_金额";
                                if (strValue == "" )
                                {
                                    strValue = "0";
                                }
                                SetDicValue(strValue, strKey, dicHRInforSourceData);
                                //sum_/324 失业保险金（公司）_金额
                                strValue = drItem["/324 失业保险金（公司）_金额"].ToString().Trim();
                                strKey = strStore + "-" + "sum_/324 失业保险金（公司）_金额";
                                if (strValue == "")
                                {
                                    strValue = "0";
                                }
                                SetDicValue(strValue, strKey, dicHRInforSourceData);
                                //sum_/334 医疗保险金（公司）_金额
                                strValue = drItem["/334 医疗保险金（公司）_金额"].ToString().Trim();
                                strKey = strStore + "-" + "sum_/334 医疗保险金（公司）_金额";
                                if (strValue == "" )
                                {
                                    strValue = "0";
                                }
                                SetDicValue(strValue, strKey, dicHRInforSourceData);
                                //sum_/344 工伤保险金（公司）_金额
                                strValue = drItem["/344 工伤保险金（公司）_金额"].ToString().Trim();
                                strKey = strStore + "-" + "sum_/344 工伤保险金（公司）_金额";
                                if (strValue == "" )
                                {
                                    strValue = "0";
                                }
                                SetDicValue(strValue, strKey, dicHRInforSourceData);
                                //sum_/354 生育保险金（公司）_金额
                                strValue = drItem["/354 生育保险金（公司）_金额"].ToString().Trim();
                                strKey = strStore + "-" + "sum_/354 生育保险金（公司）_金额";
                                if (strValue == "" )
                                {
                                    strValue = "0";
                                }
                                SetDicValue(strValue, strKey, dicHRInforSourceData);
                            }
                        }




                    }

                }
                //else
                //{
                //    throw new Exception("这个文件不存在： " + NextFile.FullName);
                //}
            }
            //社保公积金缴纳明细表 
            #endregion
            #region 7.员工社保公积金计算基数报表-SAP-机器人
            // 7.员工社保公积金计算基数报表-SAP-机器人

            foreach (FileInfo NextFile in new DirectoryInfo(strHRInsuranceBaseFilePath).GetFiles())
            {

                if (NextFile.Extension.ToLower() == ".xlsx")
                {

                    using (BuildDataExcelHelper excel = new BuildDataExcelHelper(NextFile.FullName, true, true, false))
                    {
                        DataTable tblBasicTable = excel.GetData("Sheet1");

                        // string strStore;

                        String strKey = "";
                        string strInsuranceType;
                        String strValue = "";
                        //Personnel Area
                        foreach (DataRow drItem in tblBasicTable.Rows)
                        {
                            string strStore = drItem["店号"].ToString().Trim();
                            if (strStore != "")
                            {
                                strStore = strStore.PadLeft(4, '0');
                                // strInsuranceType = drItem["保险类型"].ToString().Trim();
                                strKey = strStore + "-" + "生育保险";
                                strValue = drItem["单位缴费基数：单位参加生育保险缴费基数（单位：万元）"].ToString().Trim();
                                SetDicValue(strValue, strKey, dicHRInforSourceData);


                                strKey = strStore + "-" + "工伤保险";
                                strValue = "0";
                                SetDicValue(strValue, strKey, dicHRInforSourceData);

                                strKey = strStore + "-" + "医疗保险";
                                strValue = drItem["单位缴费基数：单位参加职工基本医疗保险缴费基数（单位：万元）"].ToString().Trim();
                                SetDicValue(strValue, strKey, dicHRInforSourceData);
                                strKey = strStore + "-" + "失业保险";
                                strValue = drItem["单位缴费基数：单位参加失业保险缴费基数（单位：万元）"].ToString().Trim();
                                SetDicValue(strValue, strKey, dicHRInforSourceData);
                                strKey = strStore + "-" + "养老保险";
                                strValue = drItem["单位缴费基数：单位参加城镇职工基本养老保险缴费基数（单位：万元）"].ToString().Trim();
                                SetDicValue(strValue, strKey, dicHRInforSourceData);

                            }

                        }

                        //throw new Exception("这个文件不存在： " + NextFile.FullName);
                        #region robot source
                        //using (ExcelHelper excel = new ExcelHelper(NextFile.FullName, true, true, false))
                        //{
                        //    DataTable tblBasicTable = excel.GetData("Sheet1");
                        //    string strStore;

                        //    String strKey = "";
                        //    string strInsuranceType;
                        //    String strValue = "";
                        //    //Personnel Area
                        //    foreach (DataRow drItem in tblBasicTable.Rows)
                        //    {
                        //        string strTemp = drItem["合同所在单位人事范围"].ToString().Trim();
                        //        if (strTemp != "" && dicBasicTable.ContainsKey(strTemp))
                        //        {
                        //            strStore = dicBasicTable[strTemp].PadLeft(4, '0');
                        //            strInsuranceType = drItem["保险类型"].ToString().Trim();
                        //            switch (strInsuranceType)
                        //            {
                        //                case "生育保险":
                        //                    strKey = strStore + "-" + "生育保险";
                        //                    strValue = drItem["公司基数"].ToString().Trim();
                        //                    SetDicValue(strValue, strKey, dicHRInforSourceData);


                        //                    break;
                        //                case "工伤保险":
                        //                    strKey = strStore + "-" + "工伤保险";
                        //                    strValue = drItem["公司基数"].ToString().Trim();
                        //                    SetDicValue(strValue, strKey, dicHRInforSourceData);

                        //                    break;
                        //                case "医疗保险":
                        //                    strKey = strStore + "-" + "医疗保险";
                        //                    strValue = drItem["公司基数"].ToString().Trim();
                        //                    SetDicValue(strValue, strKey, dicHRInforSourceData);
                        //                    break;
                        //                case "失业保险":
                        //                    strKey = strStore + "-" + "失业保险";
                        //                    strValue = drItem["公司基数"].ToString().Trim();
                        //                    SetDicValue(strValue, strKey, dicHRInforSourceData);
                        //                    break;
                        //                case "养老保险":
                        //                    strKey = strStore + "-" + "养老保险";
                        //                    strValue = drItem["公司基数"].ToString().Trim();
                        //                    SetDicValue(strValue, strKey, dicHRInforSourceData);
                        //                    break;

                        //            }



                        //            //string strTemp = drItem["/150 PHF/SI base:Month Sa_Amount"].ToString().Trim();
                        //            //strPersonnelArea = drItem["Personnel Area"].ToString().Trim();//人事范围
                        //            //if (strTemp != "" && strPersonnelArea != "" && dicBasicTable.ContainsKey(strPersonnelArea))
                        //            //{

                        //            //    strKey = dicBasicTable[strPersonnelArea].PadLeft(4, '0') + "-" + "sum_/150 PHF/SI base:Month Sa_Amount";
                        //            //    strValue = strTemp;
                        //            //    SetDicValue(strValue, strKey, dicHRInforSourceData);
                        //            //}

                        //        }
                        //    }
                        //}
                        #endregion robt
                    }


                }



                // 7.员工社保公积金计算基数报表-SAP-机器人




            }
            #endregion
            #region 6.工资明细表-SAP-机器人(计算职工薪酬)
            // 6.工资明细表-SAP-机器人(计算职工薪酬)

            object[,] objCellsValue;
            using (BuildDataExcelHelper excel = new BuildDataExcelHelper(strPayrollInfoDetailFilePath, true, true, false))
            {

                Worksheet ws = excel.CurrentSht(strDataType);
                Range rngStart = ws.Cells[9, 1];
                objCellsValue = excel.Sheet_To_Array(ws, 1, 1, rngStart: rngStart);


            }

            for (int i = 1; i <= objCellsValue.GetLength(0); i++)
            {

                string strKey;
                if (objCellsValue[i, 1] != null && objCellsValue[i, 1].ToString() != "")
                {
                    //strKey = objCellsValue[i, 2].ToString().Trim();
                    string strValue;
                    double douSum;
                    for (int j = 2; j <= objCellsValue.GetLength(1); j++)
                    {
                        //列表头名称不能为“”
                        strKey = objCellsValue[i, 1].ToString().Trim();
                        strKey = strKey.Substring(strKey.Length - 4, 4);
                        if (strDataType == "法人")
                        {
                            //法人从essbase 导出数据没有store 需要用Dxxx -store no mapping表转换
                            if (dicD0XXtoStoreNo.ContainsKey(strKey))
                            {

                                strKey = dicD0XXtoStoreNo[strKey] + "-" + "本年职工薪酬";

                                if (objCellsValue[i, j] != null)
                                {
                                    strValue = objCellsValue[i, j].ToString().Trim();
                                }
                                else
                                {
                                    strValue = "";
                                }
                                SetDicValue(strValue, strKey, dicHRInforSourceData);


                            }
                        }
                        else
                        {

                            strKey = strKey + "-" + "本年职工薪酬";

                            if (objCellsValue[i, j] != null)
                            {
                                strValue = objCellsValue[i, j].ToString().Trim();
                            }
                            else
                            {
                                strValue = "";
                            }
                            SetDicValue(strValue, strKey, dicHRInforSourceData);
                        }

                    }


                }

            }


            #region 原sap取数计算逻辑

            //foreach (FileInfo NextFile in new DirectoryInfo(strPayrollInfoDetailFilePath).GetFiles())
            //{
            //    if (NextFile.Extension.ToLower() == ".xlsx")
            //    {
            //        using (ExcelHelper excel = new ExcelHelper(NextFile.FullName, true, true, false))
            //        {
            //            DataTable tblBasicTable = excel.GetData("Sheet1");
            //            string strStore;
            //            string strPersonnelArea;
            //            String strKey;
            //            String strValue;
            //            //Personnel Area
            //            foreach (DataRow drItem in tblBasicTable.Rows)
            //            {
            //                string strTemp = drItem["/150 PHF/SI base:Month Sa_Amount"].ToString().Trim();
            //                strPersonnelArea = drItem["Personnel Area"].ToString().Trim();//人事范围
            //                if (strTemp != "" && strPersonnelArea != "" && dicBasicTable.ContainsKey(strPersonnelArea))
            //                {

            //                    strKey = dicBasicTable[strPersonnelArea].PadLeft(4, '0') + "-" + "sum_/150 PHF/SI base:Month Sa_Amount";
            //                    strValue = strTemp;
            //                    SetDicValue(strValue, strKey, dicHRInforSourceData);
            //                }

            //            }
            //        }

            //    }
            //}
            //工资明细表-ESS-手工

            #endregion 原sap取数

            #endregion
            #region 1.HRO 基本信息写入字典
            //1.HRO 基本信息写入字典

            using (BuildDataExcelHelper excel = new BuildDataExcelHelper(strHROBasicInfoDetailFilePath, true, true, false))
            {
                DataTable tblBasicTable = excel.GetData("Sheet1");
                string strStore;
                String strKey;
                String strValue;
                foreach (DataRow drItem in tblBasicTable.Rows)
                {
                    string strTemp = drItem["Personnel Area"].ToString().Trim();


                    if (strTemp != "" && dicBasicTable.ContainsKey(strTemp))
                    {
                        strStore = dicBasicTable[strTemp].PadLeft(4, '0');
                        //从业人数
                        strKey = strStore + "-" + "从业人数";
                        strValue = "1";
                        SetDicValue(strValue, strKey, dicHRInforSourceData);

                        //(其中女性从业人数)
                        strKey = drItem["Gender"].ToString().Trim();
                            
                        if (strKey.ToLower() == "female")
                        {
                            strKey = strStore + "-" + "(其中女性从业人数)";
                            strValue = "1";
                        }
                        else
                        {
                            strKey = strStore + "-" + "(其中女性从业人数)";
                            strValue = "0";
                        }
                        SetDicValue(strValue, strKey, dicHRInforSourceData);
                        //中共党员人数（含预备党员）

                        strKey = drItem["Political Status Text"].ToString().Trim();
                            
                        if (strKey.ToLower() == "member of communist party of china")
                        {
                            strKey = strStore + "-" + "中共党员人数（含预备党员）";
                            strValue = "1";
                        }
                        else
                        {
                            strKey = strStore + "-" + "中共党员人数（含预备党员）";
                            strValue = "0";
                        }
                        SetDicValue(strValue, strKey, dicHRInforSourceData);
                        //外商投资基本信息-外籍职工

                        strKey = drItem["Nationality Text"].ToString().Trim();
                            
                        if (!strKey.ToLower().Contains("chinese") && !strKey.ToLower().Contains("china") && !strKey.ToLower().Contains("hong kong") && !strKey.ToLower().Contains("macau") && !strKey.ToLower().Contains("taiwan"))
                        {
                            strValue = "1";
                            strKey = strStore + "-" + "外商投资基本信息-外籍职工";
                        }
                        else
                        {
                            strValue = "0";
                            strKey = strStore + "-" + "外商投资基本信息-外籍职工";
                        }
                        SetDicValue(strValue, strKey, dicHRInforSourceData);

                        //从业人员信息-残疾人人数-雇工

                        //strKey = drItem["Disability"].ToString().Trim();
                        //if (strKey.ToLower() != "y")
                        //{
                        //    strKey = strStore + "-" + "从业人员信息-残疾人人数-雇工";
                        //    strValue = "1";
                        //    SetDicValue(strValue, strKey, dicHRInforSourceData);
                        //}


                        //从业人员信息-残疾人人数-雇工 从业人员中属于残疾人 企业已安置残疾人员数

                        strKey = drItem["Disability"].ToString().Trim();
                        if (strKey.ToLower() == "y")
                        {
                            strValue = "1";

                            strKey = strStore + "-" + "从业人员信息-残疾人人数-雇工";
                            SetDicValue(strValue, strKey, dicHRInforSourceData);
                            strKey = strStore + "-" + "从业人员中属于残疾人";
                            SetDicValue(strValue, strKey, dicHRInforSourceData);
                            strKey = strStore + "-" + "企业已安置残疾人员数";
                            SetDicValue(strValue, strKey, dicHRInforSourceData);

                        }
                        else
                        {

                            strValue = "0";

                            strKey = strStore + "-" + "从业人员信息-残疾人人数-雇工";
                            SetDicValue(strValue, strKey, dicHRInforSourceData);
                            strKey = strStore + "-" + "从业人员中属于残疾人";
                            SetDicValue(strValue, strKey, dicHRInforSourceData);
                            strKey = strStore + "-" + "企业已安置残疾人员数";
                            SetDicValue(strValue, strKey, dicHRInforSourceData);


                        }


                    }



                }




            }

            //HRO 基本信息写入字典
            #endregion
            #region 3.HRO 教育水平
            //3.HRO 教育水平
            using (BuildDataExcelHelper excel = new BuildDataExcelHelper(strLevelOfEducationInfoDetailFilePath, true, true, false))
            {
                DataTable tblBasicTable = excel.GetData("Sheet1");
                string strStore;
                String strKey;
                String strValue;
                foreach (DataRow drItem in tblBasicTable.Rows)
                {
                    string strTemp = drItem["人事范围代码"].ToString().Trim();
                    if (strTemp != "" && dicBasicTable.ContainsKey(strTemp))
                    {
                        strStore = dicBasicTable[strTemp].PadLeft(4, '0');


                        strKey = drItem["最高学历"].ToString().Trim();
                        if (strKey != "")
                        {
                            strKey = strStore + "-" + strKey;
                            strValue = "1";
                                
                        }
                        else
                        {

                            strKey = strStore + "-" + strKey;
                            strValue = "0";
                        }
                        SetDicValue(strValue, strKey, dicHRInforSourceData);
                    }
                }
            }
            //HRO 教育水平
            #endregion
            #region  4.Band12员工信息-8028-手工
            // 4.Band12员工信息-8028-手工
            using (BuildDataExcelHelper excel = new BuildDataExcelHelper(strBand12InfoDetailFilePath, true, true, false))
            {
                DataTable tblBasicTable = excel.GetData("Sheet1");
                string strStore;
                String strKey;
                String strValue;
                foreach (DataRow drItem in tblBasicTable.Rows)
                {
                    string strTemp = drItem["字段"].ToString().Trim();
                    strStore = "8028";
                    if (strTemp != "")
                    {
                        strKey = strStore + "-" + strTemp;
                        strValue = drItem["字段内容"].ToString().Trim();
                        if (strValue == "") strValue = "0";
                        SetDicValue(strValue, strKey, dicHRInforSourceData);
                    }

                }
            }

            //Band12员工信息-8028-手工
            #endregion
            #region 5.HR其他数据-门店-手工
            // 5.HR其他数据-门店-手工
            using (BuildDataExcelHelper excel = new BuildDataExcelHelper(strHROtherFromStoreInfoDetailFilePath, true, true, false))
            {

                DataTable tblBasicTable = excel.GetData(strDataType);
                string strStore;
                String strKey;
                String strValue;
                foreach (DataRow drItem in tblBasicTable.Rows)
                {
                    string strTemp = drItem["标准字段名(内码-数据库用）"].ToString().Trim();
                    strStore = drItem["店号"].ToString().Trim();//店号
                    if (strTemp != "" && strStore != "")
                    {
                        strKey = strStore.PadLeft(4, '0') + "-" + strTemp;
                        strValue = drItem["字段内容"].ToString().Trim();
                        SetDicValue(strValue, strKey, dicHRInforSourceData);
                    }

                }
            }
            //HR其他数据-门店-手工
            #endregion


        }
        private void SetDicValue(string p_strValue,string p_strKey,Dictionary<string,string> P_dic)
        {
            //douValue = double.Parse(objCellsValue[i, j].ToString());
            double douValue = 0;
            string strValue;
            if (double.TryParse(p_strValue, out douValue))
            {
                strValue = douValue.ToString();
            }
            else

            {
                strValue = p_strValue;
            }


            if (!P_dic.ContainsKey(p_strKey))
            {
                P_dic.Add(p_strKey, strValue);
            }
            else
            {
                double douTemp ;
                double douTemp1;
                if (double.TryParse(strValue, out douTemp) && double.TryParse(P_dic[p_strKey], out douTemp1))
                {
                    P_dic[p_strKey] = (douTemp + douTemp1).ToString();

                }
            }
            //if (p_strKey == "8028-1生育保险")
            //{
            //    Console.WriteLine(P_dic[p_strKey]);
            //}

        }

        private void SetdrCorporateData(string strStore,clsIcodeVSSourceData clsIcodeVSSourceDataItem, Dictionary<string, string> P_dic, DataRow drCorporateData,string strResult)
        {

            string strStoreKeyWord;
            double douSum = 0;
            string strTempValue = "";
            foreach (string strItem in clsIcodeVSSourceDataItem.lstKeyWords)
            {


                strStoreKeyWord = strStore + "-" + strItem;

                //if (strStoreKeyWord == "1021-单位缴费基数-单位参加生育保险缴费基数（单位万元）")
                //{
                //    strStoreKeyWord = "1021-单位缴费基数-单位参加生育保险缴费基数（单位万元）";


                //}

                //strConverUnit 不为空需要单位转换
                if (P_dic.ContainsKey(strStoreKeyWord))
                {

                    if (clsIcodeVSSourceDataItem.strConverUnit != "")
                    {
                        double douTemp;
                        if (double.TryParse(P_dic[strStoreKeyWord],out douTemp))
                        {
                            if (clsIcodeVSSourceDataItem.strIcode.Contains("人数"))
                            {
                                douSum = douSum +Math.Ceiling( double.Parse(P_dic[strStoreKeyWord]) * double.Parse(clsIcodeVSSourceDataItem.strConverUnit));
                            }
                            else
                            {
                                douSum = douSum + double.Parse(P_dic[strStoreKeyWord]) * double.Parse(clsIcodeVSSourceDataItem.strConverUnit);
                            }
                        }

                        

                    }
                    else
                    {
                        strTempValue = P_dic[strStoreKeyWord].ToString();
                    }

                    drCorporateData["数据收集结果"] = "";
                    drCorporateData["是否上传"] = "是";
                }
                else
                {
                    
                    drCorporateData["数据收集结果"] = strResult+ "数据源无法找到该字段信息";
                    drCorporateData["是否上传"] = "否";
                    //drCorporateData["处理方法"] = "在法人（公司信息）" + "表中添加该店号信息";

                }



            }


            if (clsIcodeVSSourceDataItem.strConverUnit != "")
            {

                drCorporateData["字段内容"] = douSum.ToString();
            }
            else
            {
                drCorporateData["字段内容"] = strTempValue;
            }

        }


        public void outputCorporateData()
        {
            //根据登录信息明细表的stroe 为基准，找到所有相关信息写入tblCorporateData table 表
            tblCorporateData = new DataTable();

            string[] lstHeaderName = { "店号", "分类", "字段", "字段内容","省份", "城市", "用户名", "密码", "联络员姓名", "联络员手机号码", "联络员身份证号码", "工具登录取值顺序","网址", "数据收集结果", "处理方法","是否上传", "上传状态", "上传结果" };
            DataRow drCorporateData = null;
            foreach (string item in lstHeaderName)
            {
                tblCorporateData.Columns.Add(item, typeof(string));

            }

            foreach (string strStore in dicCorporateLoginInfo.Keys)
            {

                clsLoginInfo clsLoginInfoItem = dicCorporateLoginInfo[strStore];

                if (dicCorporate.ContainsKey(strStore))
                {
                    clsStandartTempInfo clsStandartTempInfoItem = new clsStandartTempInfo();

                    foreach (var varItem in dicCorporate[strStore])
                    {
                        drCorporateData = tblCorporateData.NewRow();

                        drCorporateData["店号"] = strStore;

                        drCorporateData["用户名"] = clsLoginInfoItem.strloginAccount;
                        drCorporateData["密码"] = clsLoginInfoItem.strLoginPassword;
                        drCorporateData["联络员姓名"] = clsLoginInfoItem.strContactName;
                        drCorporateData["联络员手机号码"] = clsLoginInfoItem.strContactPhone;
                        drCorporateData["联络员身份证号码"] = clsLoginInfoItem.strContactID;
                        drCorporateData["工具登录取值顺序"] = clsLoginInfoItem.strLoginType;
                        drCorporateData["网址"] = clsLoginInfoItem.strURL;

                        drCorporateData["省份"] = clsLoginInfoItem.strProvince;
                        drCorporateData["城市"] = clsLoginInfoItem.strCity;



                        drCorporateData["数据收集结果"] = "";
                        drCorporateData["处理方法"] = "";
                        clsStandartTempInfoItem = varItem.Value;
                        drCorporateData["字段"] = varItem.Key;

                        drCorporateData["分类"] = clsStandartTempInfoItem.strCategory;
                        drCorporateData["是否上传"] = "是";




                        if (clsStandartTempInfoItem.strInternalCode == "外商投资经营情况-纳税总额-营业税(万元)")
                        {
                            drCorporateData["分类"] = clsStandartTempInfoItem.strCategory;

                        }


                        string strIcode = clsStandartTempInfoItem.strInternalCode;

                        string strStoreIcode = strStore + "-" + strIcode;

                        //if (strStoreIcode == "0111-生产经营情况信息-纳税总额(万元)")
                        //{
                        //    drCorporateData["分类"] = clsStandartTempInfoItem.strCategory;

                        //}





                        clsIcodeVSSourceData clsIcodeVSSourceDataItem = new clsIcodeVSSourceData();
                        if (dicCorporateFieldMap.ContainsKey(strIcode))
                        {
                            clsIcodeVSSourceDataItem = dicCorporateFieldMap[strIcode];




                        }
                        else
                        {

                            drCorporateData["数据收集结果"] = "3 法人字段映射-内码和沃尔玛报表" + "表没有该字段信息";
                            drCorporateData["处理方法"] = "在3 法人字段映射-内码和沃尔玛报表" + "表维护该字段信息";
                            drCorporateData["是否上传"] = "否";

                            tblCorporateData.Rows.Add(drCorporateData);
                            continue;
                            //break;

                        }



                        if (clsStandartTempInfoItem.strCategory == "公司信息")
                        {

                            //drCorporateData["字段"] = strStore;
                            drCorporateData["分类"] = clsStandartTempInfoItem.strCategory;
                            drCorporateData["字段内容"] = clsStandartTempInfoItem.strDefault;

                        }
                        else if (clsStandartTempInfoItem.strCategory == "财报信息")
                        {
                            SetdrCorporateData(strStore, clsIcodeVSSourceDataItem,dicFinancialInforSourceData, drCorporateData, "财报信息");


                        }
                        else if (clsStandartTempInfoItem.strCategory == "税务信息")
                        {
                           
                            
                            SetdrCorporateData(strStore, clsIcodeVSSourceDataItem,dicTaxInforSourceData, drCorporateData, "税务信息");

                            //SetdrCorporateData(strStore, clsIcodeVSSourceDataItem, dicTariffInforSourceData, drCorporateData, "税务信息");

                        }

                        else if (clsStandartTempInfoItem.strCategory == "特种设备信息")
                        {

                            SetdrCorporateData(strStore, clsIcodeVSSourceDataItem,dicSpecialEquipmentInforSourceData, drCorporateData, "特种设备信息");

                        }
                        else if (clsStandartTempInfoItem.strCategory == "HR信息")
                        {
                            //8028 是总部的法人代表，需要 添加特殊的数据
                            if (strStore == "8028")
                            {
                                if (clsIcodeVSSourceDataItem.strIcode == "外商投资基本情况-外籍职工")
                                {
                                    clsIcodeVSSourceDataItem.lstKeyWords.Add("外商投资基本情况-外籍职工Band12及以上");

                                }
                                if (clsIcodeVSSourceDataItem.strIcode == "外商投资基本情况-大专及以上学历")
                                {
                                    clsIcodeVSSourceDataItem.lstKeyWords.Add("外商投资基本情况-大专及以上学历Band12及以上");

                                }
                                if (clsIcodeVSSourceDataItem.strIcode == "外商投资基本情况-大学及以上学历")
                                {
                                    clsIcodeVSSourceDataItem.lstKeyWords.Add("外商投资基本情况-大专及以上学历Band12及以上");

                                }

                                if (clsIcodeVSSourceDataItem.strIcode == "外商投资基本情况-本年职工薪酬")
                                {
                                    clsIcodeVSSourceDataItem.lstKeyWords.Add("外商投资基本情况-本年职工薪酬Band12及以上");

                                }

                            }

                            SetdrCorporateData(strStore, clsIcodeVSSourceDataItem,dicHRInforSourceData, drCorporateData, "HR信息");




                        }
                        //如果 3.1 法人-默认值字段表 存在这个字段覆盖前面的字段内容值
                        if (dicCorporateDefaultInfo.ContainsKey(varItem.Key))
                        {
                            drCorporateData["字段内容"] = dicCorporateDefaultInfo[varItem.Key];
                            drCorporateData["数据收集结果"] = "";
                            drCorporateData["处理方法"] = "";
                            drCorporateData["是否上传"] = "是";
                        }

                        
                        if(!dicCorporateCalvalue.ContainsKey(strStoreIcode))
                        {
                            dicCorporateCalvalue.Add(strStoreIcode, drCorporateData["字段内容"].ToString());
                        }



                        //特殊化处理 item  盈亏情况
                        //1. 盈亏情况=计算值 - 根据“企业资产状况信息 - 利润总额”是否大于0，选择其一“盈利 / 亏损 / 收支平衡 / 与总公司合并报表”
                        //企业资产状况信息-利润总额（元）
                        if (strIcode == "盈亏情况")
                        {
                            string strTemp;
                            strTemp = strStore + "-" + "企业资产状况信息-利润总额（元）";
                            if (dicCorporateCalvalue.ContainsKey(strTemp))
                            {
                                double douTemp;
                                if (double.TryParse(dicCorporateCalvalue[strTemp], out douTemp))
                                {
                                    if (douTemp > 0)
                                    {

                                        drCorporateData["字段内容"] = "盈利";
                                    }
                                    else if (douTemp < 0)
                                    {
                                        drCorporateData["字段内容"] = "亏损";
                                    }
                                    else
                                    {

                                        drCorporateData["字段内容"] = "收支平衡";
                                    }

                                    drCorporateData["数据收集结果"] = "";
                                    drCorporateData["处理方法"] = "";
                                    drCorporateData["是否上传"] = "是";
                                }
                                else
                                {
                                    drCorporateData["数据收集结果"] = "无法根据 企业资产状况信息-利润总额（元） item 判断盈亏情况";
                                    drCorporateData["处理方法"] = "";
                                    drCorporateData["是否上传"] = "否";
                                }

                            }
                        
                        }



                        tblCorporateData.Rows.Add(drCorporateData);

                    }




                }
                else
                {
                    drCorporateData = tblCorporateData.NewRow();

                    drCorporateData["用户名"] = clsLoginInfoItem.strloginAccount;
                    drCorporateData["密码"] = clsLoginInfoItem.strLoginPassword;
                    drCorporateData["联络员姓名"] = clsLoginInfoItem.strContactName;
                    drCorporateData["联络员手机号码"] = clsLoginInfoItem.strContactPhone;
                    drCorporateData["联络员身份证号码"] = clsLoginInfoItem.strContactID;
                    drCorporateData["工具登录取值顺序"] = clsLoginInfoItem.strLoginType;
                    drCorporateData["网址"] = clsLoginInfoItem.strURL;


                    drCorporateData["店号"] = strStore;
                    drCorporateData["数据收集结果"] = strDataType + "（公司信息）" + "表没有该店号";
                    drCorporateData["处理方法"] = strDataType + "（公司信息）"  + "表中添加该店号信息";
                    drCorporateData["是否上传"] = "否";
                    tblCorporateData.Rows.Add(drCorporateData);
                }



            }
            //区别分支机构和法人
            stroutputCorporateData = dicConfig["stroutputCorporateData"];
            stroutputCorporateData = Path.Combine(stroutputCorporateData, strDataType+ " UploadCorporateData_"+DateTime.Now.ToString("yyyyMMddhhmmss")+".xlsx");
            if (File.Exists(stroutputCorporateData))
            {
                File.Delete(stroutputCorporateData);


            }
            Thread.Sleep(3000);

            //UploadCorporateData.xlsx
            using (BuildDataExcelHelper excel = new BuildDataExcelHelper(stroutputCorporateData,true,false, true))
            {
                excel.SetData(tblCorporateData, "Sheet1");

            }
        }

        public void buildLoginInfo(string strLoginInfoPath)
        {
            dicCorporateLoginInfo = new Dictionary<string, clsLoginInfo>();

            //if (!File.Exists(strLoginInfoPath))
            //{
            //    throw new Exception("这个文件不存在： " + strLoginInfoPath);
            //}

            using (BuildDataExcelHelper excel = new BuildDataExcelHelper(strLoginInfoPath, true, true, false))
            {

                DataTable tblLoginInfo = excel.GetData(strDataType + "公司年报登录信息", 3, 2, 2);

                string strStore = "";

                string strComCode = "";

                foreach (DataRow item in tblLoginInfo.Rows)
                {

                    strStore = item["Store NO.店号"].ToString().Trim();

                    strStore = strStore.PadLeft(4, '0');

                    //if (strStore == "0205")
                    //{
                    //    strStore = "0205";


                    //}


                    if (!dicCorporateLoginInfo.ContainsKey(strStore))
                    {
                        clsLoginInfo clsLoginInfoItem = new clsLoginInfo();

                        clsLoginInfoItem.strStore = strStore;
                        clsLoginInfoItem.strloginAccount = item["用户名"].ToString().Trim();
                        clsLoginInfoItem.strLoginPassword = item["密码"].ToString().Trim();
                        clsLoginInfoItem.strContactName = item["联络员姓名"].ToString().Trim();
                        clsLoginInfoItem.strContactPhone = item["联络员手机号码"].ToString().Trim();
                        clsLoginInfoItem.strContactID = item["联络员身份证号码"].ToString().Trim();


                        clsLoginInfoItem.strProvince = item["Province省份"].ToString().Trim();
                        clsLoginInfoItem.strCity = item["City城市"].ToString().Trim();


                        //工具登录取值顺序

                        clsLoginInfoItem.strLoginType= item["工具登录取值顺序"].ToString().Trim();

                        clsLoginInfoItem.strURL = item["网站入口"].ToString().Trim();

                        dicCorporateLoginInfo.Add(strStore, clsLoginInfoItem);

                    }

                    //法人登录信息表有 D0XX TO Store 对应表 作用于的法人非关税 数据源mapping
                    // dicD0XXtoStoreNo
                    if (strDataType == "法人")
                    {
                        strComCode= item["Company Code"].ToString().Trim();
                        if (!dicD0XXtoStoreNo.ContainsKey(strComCode) && strComCode !="")
                        {
                            dicD0XXtoStoreNo.Add(strComCode, strStore);
                        }

                    }

                }





            }

        }


        //写入法人（公司信息）
        /// <summary>
        /// 1（公司信息）2字段映射-内码和沃尔玛报表 3-默认值字段表
        /// </summary>
        public void buildBasicData()
        {
            dicCorporate = new Dictionary<string, Dictionary<string, clsStandartTempInfo>>();
            dicCorporateDefaultInfo = new Dictionary<string, string>();
            dicCorporateFieldMap = new Dictionary<string, clsIcodeVSSourceData>();
            if (!File.Exists(strStandardFieldtemplatePath))
            {
                throw new Exception("这个文件不存在： " + strStandardFieldtemplatePath);
            }

            using (BuildDataExcelHelper excel = new BuildDataExcelHelper(strStandardFieldtemplatePath, true, true, false))
            {


                DataTable tblTempInfo = excel.GetData(strDataType +"（公司信息）",intGetMaxRowColIndex:3, intGetMaxColRowIndex:2);
                //DataTable tblTempInfo = excel.GetData("Sheet1", intGetMaxRowColIndex: 3, intGetMaxColRowIndex: 2);
                string strStore = "";
                string InternalCode;
                string strValid;
                foreach (DataRow item in tblTempInfo.Rows)
                {

                    strStore = item["店号"].ToString().Trim();
                    strValid= item["有效性"].ToString().Trim();
                    if (strStore != "" && item["标准字段名(内码-数据库用）"].ToString().Trim() != ""&& strValid=="是")
                    {
                        InternalCode = item["标准字段名(内码-数据库用）"].ToString().Trim();
                        strStore = strStore.PadLeft(4, '0');
                        Dictionary<string, clsStandartTempInfo> dicInternalCode = null;
                        if (dicCorporate.ContainsKey(strStore))
                        {
                            dicInternalCode = dicCorporate[strStore];

                        }
                        else
                        {
                            dicInternalCode = new Dictionary<string, clsStandartTempInfo>();

                            dicCorporate.Add(strStore, dicInternalCode);
                        }

                        clsStandartTempInfo clsStandartTempInfoItem = new clsStandartTempInfo();

                        clsStandartTempInfoItem.strStore = strStore;
                        
                        clsStandartTempInfoItem.strCategory = item["分类"].ToString().Trim();


                        if (clsStandartTempInfoItem.strCategory == "HR信息")
                        {
                            clsStandartTempInfoItem.strCategory = clsStandartTempInfoItem.strCategory;

                        }


                        if (clsStandartTempInfoItem.strCategory == "公司信息")
                        {
                            clsStandartTempInfoItem.strDefault = item["字段内容"].ToString().Trim();
                        }
                        else 
                        {
                            clsStandartTempInfoItem.strDefault = "";
                        }
                        clsStandartTempInfoItem.strExternalCode = item["显示字段名(外码-网站显示）"].ToString().Trim();
                        clsStandartTempInfoItem.strInternalCode = InternalCode;

                        dicInternalCode[InternalCode] = clsStandartTempInfoItem;





                        dicCorporate[strStore]= dicInternalCode;

                    }
                }



                 tblTempInfo = excel.GetData(strDataType + "字段映射-内码和沃尔玛报表", 3, 2, 1);

                string strInternalCode = "";

                foreach (DataRow item in tblTempInfo.Rows)
                {

                    strInternalCode = item["标准字段名(内码-数据库用）"].ToString().Trim();
                    if (strInternalCode != "" && item["字段"].ToString().Trim() != "")
                    {


                        if (!dicCorporateFieldMap.ContainsKey(strInternalCode))
                        {
                            clsIcodeVSSourceData clsIcodeVSSourceDataItem = new clsIcodeVSSourceData();

                            //Unit conversion

                            clsIcodeVSSourceDataItem.strConverUnit = item["单位换算"].ToString().Trim();
                            //clsIcodeVSSourceDataItem.strCurrency = "元";

                            //if (strInternalCode.Contains("万元"))
                            //{
                            //    clsIcodeVSSourceDataItem.strCurrency = "万元";

                            //}

                            clsIcodeVSSourceDataItem.strCategory = item["分类"].ToString().Trim();
                            clsIcodeVSSourceDataItem.strIcode = strInternalCode;

                            clsIcodeVSSourceDataItem.lstKeyWords = item["字段"].ToString().Trim().Split('+').ToList();
                            dicCorporateFieldMap.Add(strInternalCode, clsIcodeVSSourceDataItem);
                        }

                    }
                }

                tblTempInfo = excel.GetData(strDataType + "默认值字段表", 3, 2, 1);

                strInternalCode = "";
                string strDefault = "";

                foreach (DataRow item in tblTempInfo.Rows)
                {
                    strInternalCode = item["标准字段名(内码-数据库用）"].ToString().Trim();
                    strDefault = item["默认值"].ToString().Trim(); //
                    if (!dicCorporateDefaultInfo.ContainsKey(strInternalCode) && strInternalCode != "")
                    {
                        dicCorporateDefaultInfo.Add(strInternalCode, strDefault);

                    }

                }





            }


        }

        //3.1 法人-默认值字段表 写入字典
        /// <summary>
        /// 默认值字段表写入字典
        /// </summary>


        /// <summary>
        /// 法人财报信息
        /// </summary>
        public void FinancialInfor(string p_strFinancialInforFilePath)
        {

            // tesyfgd(p_strFinancialInforFilePath);
            dicFinancialInforSourceData = new Dictionary<string, string>();



            if (strDataType == "法人")
            {
                #region 法人财务

                #region 审计手工报表格式改变，但保留格式读取方式

                //Range RngCellsValue;
                //object[,] objCellsValue = null;
                //using (ExcelHelper excel = new ExcelHelper(p_strFinancialInforFilePath, true, false, false))
                //{

                //    Worksheet ws = excel.CurrentSht("Finance report of Entities ");
                //    Range rngStart = ws.Cells[2, 1];


                //    objCellsValue = excel.Sheet_To_Array(ws, 2, 4, rngStart: rngStart);






                //    //由于excel数据有多级表头，且一级二级都是combine 单元格，所以要预先处理
                //    string strItemTemp;

                //    Dictionary<string, int> dicHeader = new Dictionary<string, int>();
                //    string strHeader;
                //    for (int j = 1; j <= objCellsValue.GetLength(1); j++) //int i = 1; i <= 3; i++
                //    {
                //        strHeader = "";
                //        for (int i = 2; i <= 4; i++)
                //        {

                //            strItemTemp = "";

                //            if (j == 25)
                //            {
                //                j = 25;
                //            }


                //            Range RngCur = ws.Cells[i, j];
                //            if (RngCur.MergeCells)
                //            {
                //                Range RngMerge = RngCur.MergeArea[1, 1];

                //                if (RngMerge.Value != null)
                //                {
                //                    strItemTemp = RngMerge.Value.ToString().Trim();
                //                }
                //            }
                //            else
                //            {
                //                if (RngCur.Value != null)
                //                {
                //                    strItemTemp = RngCur.Value.ToString().Trim();
                //                }

                //            }


                //            if (strItemTemp != "" && strItemTemp != null)
                //            {

                //                if (strHeader == "")
                //                {
                //                    strHeader = strItemTemp;
                //                }
                //                else
                //                {
                //                    strHeader = strHeader + "-" + strItemTemp;
                //                }
                //            }

                //        }


                //        //表头字段写入list
                //        if (strHeader != "" && !dicHeader.ContainsKey(strHeader))

                //        {
                //            dicHeader.Add(strHeader, j);
                //        }
                //    }


                //    //财务信息数据写入字典
                //    string strKey;

                //    for (int i = 4; i <= objCellsValue.GetLength(0); i++)
                //    {

                //        if (objCellsValue[i, 2] != null && objCellsValue[i, 2].ToString() != "")
                //        {
                //            string strStore = objCellsValue[i, 2].ToString().Trim().PadLeft(4, '0');
                //            string strValue;

                //            foreach (var item in dicHeader)
                //            {
                //                strKey = strStore + "-" + item.Key;
                //                if (objCellsValue[i, item.Value] != null)
                //                {
                //                    strValue = objCellsValue[i, item.Value].ToString().Trim();
                //                }
                //                else
                //                {
                //                    strValue = "";
                //                }
                //                SetDicValue(strValue, strKey, dicFinancialInforSourceData);
                //            }
                //        }

                //    }

                //}

                //法人
                #endregion
                //法人

                Range RngCellsValue;
                object[,] objCellsValue = null;
                using (BuildDataExcelHelper excel = new BuildDataExcelHelper(strFinancialInforFilePath, true, false, false))
                {

                    Worksheet ws = excel.CurrentSht("Sheet1");
                    Range rngStart = ws.Cells[2, 1];


                    objCellsValue = excel.Sheet_To_Array(ws, 1, 2, rngStart: rngStart);

                    //财务信息数据写入字典
                    string strKey;
                    
                    for (int i = 2; i <= objCellsValue.GetLength(0); i++)
                    {
                        if (objCellsValue[i, 1] != null && objCellsValue[i, 1].ToString() != "")
                        {
                            //strKey = objCellsValue[i, 1].ToString().Trim();
                            for (int j = 2; j <= objCellsValue.GetLength(1); j++)
                            {

                                if (objCellsValue[1, j] != null && objCellsValue[1, j].ToString() != "")
                                {
                                    string strComCode = objCellsValue[1, j].ToString().Trim();
                                    string strValue;

                                    if (objCellsValue[i, j] != null && objCellsValue[i, j].ToString() != "")
                                    {
                                        strValue = objCellsValue[i, j].ToString();
                                    }
                                    else
                                    {
                                        strValue = "0"; //
                                        
                                    }

                                    if (dicD0XXtoStoreNo.ContainsKey(strComCode))
                                    {
                                        strKey = dicD0XXtoStoreNo[strComCode] + "-" + objCellsValue[i, 1].ToString().Trim();
                                        SetDicValue(strValue, strKey, dicFinancialInforSourceData);

                                    }

                                }
                            }

                        }
                    }

                }



                #endregion 法人财务
            }
            else
            {


                #region 分支机构
                object[,] objCellsValue = null;
                using (BuildDataExcelHelper excel = new BuildDataExcelHelper(p_strFinancialInforFilePath, true, true, false))
                {
                    DataTable tblBasicTable = excel.GetData("工商年检",3,1,9);
                    Worksheet ws = excel.CurrentSht("工商年检");
                    Range rngStart = ws.Cells[9, 3];
                    objCellsValue = excel.Sheet_To_Array(ws, 3, 9, rngStart: rngStart);

                

                }
           
                for (int i = 1; i <= objCellsValue.GetLength(0); i++)
                {


                    //strKey = objCellsValue[i, 2].ToString().Trim();
                    string strValue;
                     //列表头名称不能为“”
                    if (objCellsValue[i, 1] != null && objCellsValue[i, 1].ToString() != "")
                    {
                        string  strKey = objCellsValue[i, 1].ToString().Trim().Substring(3,4);

                        strKey = strKey + "-net sales" ;

                        if (objCellsValue[i, 4] != null)
                        {
                            strValue = objCellsValue[i, 4].ToString().Trim();
                        }
                        else
                        {
                            strValue = "";
                        }

                        SetDicValue(strValue, strKey, dicFinancialInforSourceData);

                        strKey=objCellsValue[i, 1].ToString().Trim().Substring(3, 4);
                        strKey = strKey + "-net income";

                        if (objCellsValue[i, 5] != null)
                        {
                            strValue = objCellsValue[i, 5].ToString().Trim();
                        }
                        else
                        {
                            strValue = "";
                        }

                        SetDicValue(strValue, strKey, dicFinancialInforSourceData);




                    }



                }



                #endregion 分支机构
            }

        }

        /// <summary>
        /// 法人税务信息
        /// </summary>
        public void TaxInfor()
        {
            dicTaxInforSourceData = new Dictionary<string, string>();
            //dicTariffInforSourceData = new Dictionary<string, string>();
            Dictionary<string, int> dicHeader = new Dictionary<string, int>();
            string strHeader;
            string strItemTemp;
            string strLeftItemTemp;
            object[,] objCellsValue = null;
            Dictionary<string, string> dicBasicTable;
            //CompanySummary  法人基表 目的是使用Dxxx mapping store no

            //if (!File.Exists(strCompanySummaryFilePath))
            //{ 
            //    throw new Exception("这个文件不存在： " + strCompanySummaryFilePath);
            //}

            //using (ExcelHelper excel = new ExcelHelper(strCompanySummaryFilePath, true, true, false))
            //{
            //    DataTable tblBasicTable = excel.GetData("Sheet1");

            //    dicBasicTable = new Dictionary<string, string>();

            //    foreach (DataRow drItem in tblBasicTable.Rows)
            //    {
            //        string strStore = drItem["Store NO."].ToString().Trim();
            //        string strComCode = drItem["SAP-COM CODE"].ToString().Trim();
            //        if (strStore != "" && strComCode != "")
            //        {
            //            strStore = strStore.PadLeft(4, '0');

            //            if (!dicBasicTable.ContainsKey(strComCode))
            //            {
            //                dicBasicTable.Add(strComCode, strStore);
            //            }


            //        }



            //    }




            //}




            dicHeader = new Dictionary<string, int>();
            strHeader = "";
            objCellsValue = null;
            #region 非关税

            
            if (strDataType == "法人")
            {
                #region 法人 Tax



                if (!File.Exists(strNonTariffsFilePath))
                {

                    throw new Exception("这个文件不存在： " + strNonTariffsFilePath);
                }

                using (BuildDataExcelHelper excel = new BuildDataExcelHelper(strNonTariffsFilePath, true, true, false))
                {

                    Worksheet ws = excel.CurrentSht("法人");
                    Range rngStart = ws.Cells[8, 1];
                    objCellsValue = excel.Sheet_To_Array(ws, 2, 8, rngStart: rngStart);


                }


                //税务非关税信息数据写入字典
                string strKey;
                for (int i = 2; i <= objCellsValue.GetLength(0); i++)
                {


                    if (objCellsValue[i, 2] != null && objCellsValue[i, 2].ToString() != "")
                    {
                        //strKey = objCellsValue[i, 2].ToString().Trim();
                        string strValue;
                        for (int j = 2; j <= objCellsValue.GetLength(1); j++)
                        {
                            //列表头名称不能为“”
                            if (objCellsValue[1, j] != null && objCellsValue[1, j].ToString() != "")
                            {
                                strKey = objCellsValue[i, 2].ToString().Trim();

                                //法人从essbase 导出数据没有store 需要用Dxxx -store no mapping表转换
                                if (dicD0XXtoStoreNo.ContainsKey(strKey))
                                {

                                    strKey = dicD0XXtoStoreNo[strKey] + "-" + objCellsValue[1, j].ToString().Trim();

                                    if (objCellsValue[i, j] != null)
                                    {
                                        strValue = objCellsValue[i, j].ToString().Trim();
                                    }
                                    else
                                    {
                                        strValue = "";
                                    }
                                    SetDicValue(strValue, strKey, dicTaxInforSourceData);


                                }

                            }
                        }


                    }

                }

                #region 关税


                using (BuildDataExcelHelper excel = new BuildDataExcelHelper(strTariffsFilePath, true, true, false))
                {

                    DataTable tblTemp = excel.GetData("Sheet1");
                    double douTemp;
                    //string strKey;
                    foreach (DataRow row in tblTemp.Rows)
                    {
                        string strInvoicetype = row["Invoice type"].ToString().Trim();
                        string strValue = row["Amount"].ToString().Trim();

                        if (double.TryParse(strValue, out douTemp) && strInvoicetype.PadLeft(2, '0') == "03")
                        {
                            strKey = "8028-" + "外商投资经营情况-纳税总额-关税(万元)";
                            SetDicValue(strValue, strKey, dicTaxInforSourceData);
                        }
                    }

                }


                #endregion 关税



                #endregion 法人 Tax
            }
            else
            {

                #region 分支机构 Tax


                using (BuildDataExcelHelper excel = new BuildDataExcelHelper(strNonTariffsFilePath, true, true, false))
                {

                    Worksheet ws = excel.CurrentSht("门店");
                    Range rngStart = ws.Cells[8, 1];
                    objCellsValue = excel.Sheet_To_Array(ws, 2, 8, rngStart: rngStart);


                }


                //税务非关税信息数据写入字典
                string strKey;
                for (int i = 2; i <= objCellsValue.GetLength(0); i++)
                {


                    if (objCellsValue[i, 2] != null && objCellsValue[i, 2].ToString() != "")
                    {
                        //strKey = objCellsValue[i, 2].ToString().Trim();
                        string strValue;
                        for (int j = 3; j <= objCellsValue.GetLength(1); j++)
                        {
                            //列表头名称不能为“”
                            if (objCellsValue[1, j] != null && objCellsValue[1, j].ToString() != "")
                            {
                                strKey = objCellsValue[i, 2].ToString().Trim().Substring(3,4);


                                strKey = strKey + "-" + objCellsValue[1, j].ToString().Trim();

                                if (objCellsValue[i, j] != null)
                                {
                                    strValue = objCellsValue[i, j].ToString().Trim();
                                }
                                else
                                {
                                    strValue = "";
                                }
                                SetDicValue(strValue, strKey, dicTaxInforSourceData);



                            }
                        }


                    }

                }

                #endregion 分支机构 Tax
            }

            #endregion 非关税



        }



        /// <summary>
        /// Special equipment 特种设备数据源写入字典
        /// </summary>
        public void SpecialEquipment()
        {
            dicSpecialEquipmentInforSourceData = new Dictionary<string, string>();
            object[,] objCellsValue = null;

            using (BuildDataExcelHelper excel = new BuildDataExcelHelper(strSpecialEquipmentFilePath, true, true, false))
            {

                Worksheet ws = excel.CurrentSht("total");
                Range rngStart = ws.Cells[1, 1];
                objCellsValue = excel.Sheet_To_Array(ws, 1, 1, rngStart: rngStart,intmaxCol:15);


            }


            //特种设备信息数据写入字典
            string strKey;
            for (int i = 2; i <= objCellsValue.GetLength(0); i++)
            {


                if (objCellsValue[i, 1] != null && objCellsValue[i, 1].ToString() != "")
                {
                    //strKey = objCellsValue[i, 1].ToString().Trim().PadLeft(4,'0');
                    double douValue;
                    string strValue;
                    for (int j = 8; j <= objCellsValue.GetLength(1); j++)
                    {
                        //列表头名称不能为“”
                        if (objCellsValue[1, j] != null || objCellsValue[1, j].ToString() != "")
                        {
                            strKey = objCellsValue[i, 1].ToString().Trim().PadLeft(4, '0') + "-" + objCellsValue[1, j].ToString().Trim();

                            douValue = 0;

                            if (objCellsValue[i, j] != null)
                            {
                                strValue = objCellsValue[i, j].ToString().Trim();
                            }
                            else
                            {
                                strValue = "";
                            }
                            SetDicValue(strValue, strKey, dicSpecialEquipmentInforSourceData);
                        }
                    }


                }

            }



        }


        //每个Store的detail 为类的属性


        //法人或分支机构
        private class clsLoginInfo
        {
            public string strStore { get; set; }

            public string strloginAccount { get; set; }


            public string strLoginPassword { get; set; }
            public string strContactName{ get; set; }
            public string strContactPhone { get; set; }
            public string strContactID { get; set; }


            public string strLoginType { get; set; }
            public string strURL { get; set; }

            public string strProvince { get; set; }
            public string strCity { get; set; }


            //public string strStore { get; set; }
            //public string strStore { get; set; }
            //public string strStore { get; set; }
            //public string strStore { get; set; }
            //public string strStore { get; set; }
        }
        /// <summary>
        /// 3 法人字段映射-内码和沃尔玛报表 对应的 class
        /// </summary>
        private class clsIcodeVSSourceData
        {
            public string strCurrency { get; set; }

            public string strIcode { get; set; }


            public string strEcode { get; set; }
            public List<string> lstKeyWords { get; set; }
            public string strCategory { get; set; }
            public string strValue { get; set; }
            public string strConverUnit { get; set; }
            //public string strStore { get; set; }
            //public string strStore { get; set; }
            //public string strStore { get; set; }
            //public string strStore { get; set; }
        }


        private class clsStandartTempInfo
        {
            public string strStore { get; set; }

            public string strCategory { get; set; }


            public string strExternalCode { get; set; }
            public string strInternalCode { get; set; }
            public string strDefault { get; set; }
            public string strCalValue { get; set; }
            public string strContactID { get; set; }


            //Unit conversion
            
            //public string strStore { get; set; }
            //public string strStore { get; set; }
            //public string strStore { get; set; }
            //public string strStore { get; set; }



        }

        private void Config()
        {

            try
            {
                DataTable dblConfig = new DataTable();
                string strName;
                string strCode;
                string strDefalut;
                using (BuildDataExcelHelper excel = new BuildDataExcelHelper(StrConfigFilePath, true, false, true))
                {
                    dblConfig = excel.GetData("Asetting");
                    foreach (DataRow row in dblConfig.Rows)
                    {
                        strName = row["Name"].ToString().Trim();
                        strDefalut = row["Value"].ToString().Trim();

                        if (!dicConfig.ContainsKey(strName) && strName != "") dicConfig.Add(strName, strDefalut);
                    }

                }

            }

            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            
            }
        }


        private string RegularExpressions(object objInput, string strPattern)
        {
            string strResult = "";
            if (objInput == null) objInput = "";
            Match m = Regex.Match(objInput.ToString(), strPattern);//\.*Linksave\.*Templat  @"\d{1,}"
            if (m.Success)
            {
                strResult = m.Value.ToString();
            }

            return strResult;
        }





    }
}
