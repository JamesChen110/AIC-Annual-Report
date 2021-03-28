using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DBHelper;
using BrowserAutomation;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Threading;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Support.Extensions;
using OpenQA.Selenium.Interactions;

using System.IO;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using System.Diagnostics;
using System.Windows.Forms;
using System.Drawing;


using System.Collections.ObjectModel;
using Keys = OpenQA.Selenium.Keys;

namespace AIC_Annual_Report
{

    //class RuleException : Exception
    //{
    //    public RuleException(string message) : base(message)
    //    {
    //    }
    //}

    public class UploadDataMain
    {
        public clsBrowserAutomation clsBrowser = null;
        public IWebDriver chromeBrowser = null;

        public string strPasswordControlID;

        public string strUploadDataFilePath;
        public int intRowIndex;
        public string strCurLog = "";
        public Boolean blisSucceed = false;
        public string strCurStore;
        public string strCurCity;
        public string strCurProvince;
        public string strDataType; //法人或者分支机构

        public Dictionary<string, string> dicConfig = new Dictionary<string, string>();
        public Dictionary<string, string> dicConfigDefault = new Dictionary<string, string>();

        public Dictionary<string, string> dicConfigTextValue = new Dictionary<string, string>();
        public Dictionary<string, string> dicFinishedStores = new Dictionary<string, string>();

        public Dictionary<string, clsStoreInfo> dicStore = new Dictionary<string, clsStoreInfo>();

        //DicIwebItem
        public Dictionary<string, clsIwebItemInfo> DicIwebItems = new Dictionary<string, clsIwebItemInfo>();


        public List<string> litIwebItems = new List<string>();

        public DataTable tblConfig = new DataTable();

        string StrConfigFilePath =Path.Combine(Directory.GetCurrentDirectory(), "Config.xlsx");
        public DataTable tblUploadData = new DataTable();
        public DataTable tblSuccessfulStores = new DataTable();

        public List<string> lstStores= new List<string>();

        //全部的Log写入List
        List<string> litLog = new List<string>();

        //litTrackLog 不论是否成功 记录是否点击提交和后面的动作
        List<string> litTrackLog = null;
        int intRetry;



        string strLogPath = Path.Combine(new DirectoryInfo(string.Format(@"{0}..\..\..\..\", Environment.CurrentDirectory)).FullName, "log Screenshot");

        public void OpenWeb()
        {


            int intRetry = 0;
            do
            {
                try
                {


                    intRetry += 1;

                    //Close the previous  driver is it have opened            
                    foreach (System.Diagnostics.Process proTem in System.Diagnostics.Process.GetProcessesByName("chromedriver"))
                    {
                        proTem.CloseMainWindow();
                    }
                    //Close the chrome browser if its opened already
                    foreach (System.Diagnostics.Process proTem in System.Diagnostics.Process.GetProcessesByName("chrome"))
                    {
                        proTem.CloseMainWindow();
                    }


                    String strChromeAppPath = System.IO.File.ReadAllText(Path.Combine(Directory.GetCurrentDirectory(), "ChromeAppPath.txt"));

                    Process chrome = Process.Start(strChromeAppPath, "--remote-debugging-port=9222 --user-data-dir=\"C:\\data\\ChromeProfile\"");

                    //Wait for Chome window to open
                    chrome.WaitForInputIdle();
                    while (true)
                    {
                        if (chrome.MainWindowHandle != IntPtr.Zero)
                        {
                            //Fixed 2 seconds delay
                            System.Threading.Thread.Sleep(2000);
                            break;
                        }
                        System.Threading.Thread.Sleep(1000);
                    }
                    Thread.Sleep(2000);
                    ChromeOptions options = new ChromeOptions();
                    options.DebuggerAddress = "127.0.0.1:9222";
                    chromeBrowser = new ChromeDriver(options);
                    Console.WriteLine(chromeBrowser.Title);
                    clsBrowser = new clsBrowserAutomation();
                    //打开Chrome浏览器
                    //chromeBrowser = clsBrowser.openChrome();

                    //web窗口最大化
                    chromeBrowser.Manage().Window.Maximize();
                    chromeBrowser.Navigate().GoToUrl(dicStore[strCurStore].strURL);
                    //打开单一窗口网站

                    intRetry = 3;

                }
                catch (Exception ex)
                {
                    strCurLog = "无法正常打开网站!,请关闭chrome 浏览器 和 chrome drive 重试"+ ex;
                }
            }
            while (intRetry < 1);


        }

        public void Start_testing(string strTelCode)
        {



            //Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe", "--remote-debugging-port=9222 --user-data-dir=\"C:\\data\\ChromeProfile\"");

            //Process prc = new Process();
            //prc.StartInfo.FileName = "";

            //Step - 2
            //Connect to that browser

            //strCurProvince = dicStore[strCurStore].strProvince;
            //strCurCity = dicStore[strCurStore].strCity;


            if (chromeBrowser == null)
            {

                ChromeOptions options = new ChromeOptions();
                options.DebuggerAddress = "127.0.0.1:9222";
                chromeBrowser = new ChromeDriver(options);
                Console.WriteLine(chromeBrowser.Title);

                clsBrowser = new clsBrowserAutomation();
            }

            //IWebElement iwedf = chromeBrowser.FindElement(By.Id("formqynbwsjb")).FindElements(By.TagName("button"))[2];
            //iwedf.Click();
            //chromeBrowser.SwitchTo().Window(chromeBrowser.WindowHandles.Last());

            chromeBrowser.SwitchTo().Window(chromeBrowser.CurrentWindowHandle);

            StrConfigFilePath = Path.Combine(Directory.GetCurrentDirectory(), "Config.xlsx");



            intRetry = 0;
            string strItemName = "";
            do
            {
                try
                {
                    intRetry += 1;
                    SetDicIwebItemValue();
                    Config();
                    Boolean blNeedToSwitchTo = true;
                    Boolean blLoginSucceed = false;
                    Boolean blNeedUploadValue = false;
                    clsIwebItemInfo clsIwebItemInfoItem = null;
                    string strLogTitle = "";
                    foreach (string striwebItem in litIwebItems)
                    {
                        litLog = new List<string>();
                        try
                        {
                            blisSucceed = true;
                            clsIwebItemInfoItem = new clsIwebItemInfo();
                            strItemName = striwebItem;
                            blNeedUploadValue = false;
                            if (DicIwebItems.ContainsKey(striwebItem))
                            {
                                clsIwebItemInfoItem = DicIwebItems[striwebItem];
                                blNeedUploadValue = true;
                            }
                            else
                            {
                                clsIwebItemInfoItem.strItemUploadValue = "";

                            }


                            //if (strItemName == "社保信息")
                            //{
                            //    strItemName = "社保信息";
                            //}
                            //continue;
                            //IWebElement iwedf = chromeBrowser.FindElement(By.Id("zczk")).FindElements(By.ClassName("button"))[0];

                            clsIwebItemInfoItem.strItemConfigValue = "";

                            clsIwebItemInfoItem.tblItemConfig = tblConfig.Select("ItemName='" + striwebItem + "'", "ParentIndex").CopyToDataTable();


                            setclsIwebItemInfoItem(clsIwebItemInfoItem);
                            strItemName = clsIwebItemInfoItem.strModule + "_" + striwebItem;

                            strLogTitle = clsIwebItemInfoItem.strModule + "_" + clsIwebItemInfoItem.strItemName;
                            strCurLog = strLogTitle + " 正在开始...";

                            //System.Windows.Forms.SendKeys(Keys.);

                            //chromeBrowser.SwitchTo().Window(chromeBrowser.WindowHandles[1]);

                            //不需要重新输入登录信息
                            if (clsIwebItemInfoItem.strModule == "打开网站并获取验证码")
                            {
                                continue;
                            }


                            //仅testing
                            //if (clsIwebItemInfoItem.strItemName == "填写股东及出资信息保存按钮成功(Check)")
                            //{
                            //    continue;
                            //}


                            if (clsIwebItemInfoItem.strModule == "输入验证码")
                            {

                                clsIwebItemInfoItem.strItemConfigValue = strTelCode;

                            }

                            //if (chromeBrowser.Url == "https://amr.sz.gov.cn/annual/index")
                            //{
                            //    blLoginSucceed = true;
                            //}

                            if (clsIwebItemInfoItem.strModule == "选择年报")
                            {
                                //continue;

                                if (!blLoginSucceed)
                                {
                                    litLog.Add("登录失败");
                                    break;
                                }

                            }



                            //网站生成新的 tab
                            else if (clsIwebItemInfoItem.strModule == "基本信息")
                            {
                                if (blNeedToSwitchTo)
                                {
                                    chromeBrowser.SwitchTo().Window(chromeBrowser.WindowHandles[1]);
                                    blNeedToSwitchTo = false;
                                }

                            }
                            if (blLoginSucceed = true && striwebItem == "登录异常提示")
                            {
                                continue;
                            }

                            //仅测试使用
                            if (clsIwebItemInfoItem.strModule == "输入验证码" || clsIwebItemInfoItem.strModule == "打开网站并获取验证码" || clsIwebItemInfoItem.strModule == "点击登录")//点击登录
                            {
                                //continue;
                            }

                            //仅湖南四川使用


                            if (strCurProvince == "四川" || strCurProvince == "湖南" || strCurProvince == "云南" || strCurProvince == "福建" || strCurProvince == "河北")
                            {
                                if (striwebItem == "股东出资情况表-认缴出资时间" || striwebItem == "股东出资情况表-实缴出资时间")
                                {


                                    IWebElement iwebEle = IWebSearchIwebItem(chromeBrowser, clsIwebItemInfoItem, strItemName);
                                    SelectDate(clsIwebItemInfoItem.strItemUploadValue, iwebEle);
                                    strCurLog = strLogTitle + " is succeed !";
                                    //使用指定的function选择时间
                                    continue;
                                }

                            }

                            if (clsIwebItemInfoItem.strItemTagName == "checkbox")
                            {



                                IList<IWebElement> ilitiwebEle = IWebSearchIwebList(chromeBrowser, clsIwebItemInfoItem);
                                //ilitiwebEle = chromeBrowser.FindElements(By.Name ("INVNATURE"));


                                //选择config value or uploaddata value
                                string strValue = "";
                                if (clsIwebItemInfoItem.strIsValueFromData == "true")
                                {

                                    strValue = clsIwebItemInfoItem.strItemUploadValue;
                                }
                                else
                                {
                                    strValue = clsIwebItemInfoItem.strItemConfigValue;

                                }




                                foreach (string strItem in strValue.Split(','))
                                {
                                    if (clsIwebItemInfoItem.strcheckboxBy != null && clsIwebItemInfoItem.strcheckboxBy == "text")
                                    {
                                        clsBrowser.checkboxElement_SelectText(ilitiwebEle, strItem, clsIwebItemInfoItem.strcheckboxBy);


                                    }
                                    else
                                    {
                                        string strTemp = clsIwebItemInfoItem.strItemName + "-" + strItem;


                                        if (dicConfigTextValue.ContainsKey(strTemp))
                                        {


                                            // clsBrowser.checkboxElement_SelectText(ilitiwebEle, dicConfigTextValue[strItem]);
                                            // checkbox 部分需要用配置表Text_Value 转换找到对应的value
                                            //strcheckboxBy 为空 默认用iweb.text 查找
                                            //chromeBrowser.SwitchTo().DefaultContent();
                                            //chromeBrowser.SwitchTo().Frame("controlIframe");
                                            //IList<IWebElement> fdsf = chromeBrowser.FindElements(By.Name("rptAnnlForeInvestor.invNature"));


                                            clsBrowser.checkboxElement_SelectText(ilitiwebEle, dicConfigTextValue[strTemp], clsIwebItemInfoItem.strcheckboxBy);


                                            //IList<IWebElement> fdsf = chromeBrowser.FindElement(By.Id("basemodal")).FindElements(By.TagName("button"));

                                        }

                                        else
                                        {
                                            blisSucceed = false;
                                            strCurLog = strLogTitle + ": " + strTemp + " 在config文件没有对应网站 Value";

                                        }


                                    }
                                }



                            }
                            else if (clsIwebItemInfoItem.strItemTagName.ToLower() == "loopselect")

                            {

                                IList<IWebElement> ilitiwebEle = IWebSearchIwebList(chromeBrowser, clsIwebItemInfoItem);







                                //选择config value or uploaddata value
                                string strValue = "";
                                if (clsIwebItemInfoItem.strIsValueFromData == "true")
                                {

                                    strValue = clsIwebItemInfoItem.strItemUploadValue;
                                }
                                else
                                {
                                    strValue = clsIwebItemInfoItem.strItemConfigValue;

                                }



                                if (clsIwebItemInfoItem.strLoopSelectisFixedvalue.ToLower() != "true")
                                {

                                    if (dicConfigTextValue.ContainsKey(strValue))
                                    {
                                        strValue = dicConfigTextValue[strValue];
                                    }
                                }
                                clsBrowser.ClickElementByLoopFind(ilitiwebEle, strValue, clsIwebItemInfoItem.strLoopSelectBy);
                            }

                            //转到win窗口，关闭 window 弹窗
                            else if (clsIwebItemInfoItem.strItemTagName.ToLower() == "switchtoalert")

                            {
                                //Thread.Sleep(clsIwebItemInfoItem.intItemBeforeDelayTime * 1000);


                                Boolean blisRetry = true;

                                do
                                {
                                    try
                                    {
                                        WebDriverWait webWait = new WebDriverWait(chromeBrowser, new TimeSpan(0, 0, clsIwebItemInfoItem.intItemBeforeDelayTime));

                                        _ = webWait.Until(ExpectedConditions.AlertIsPresent());

                                        chromeBrowser.SwitchTo().Alert().Accept();

                                    }
                                    catch (WebDriverTimeoutException) //Alert window not found
                                    {
                                        blisRetry = false;
                                    }

                                }
                                while (blisRetry);
                                chromeBrowser.SwitchTo().DefaultContent();

                            }
                            else if (clsIwebItemInfoItem.strItemTagName.ToLower() == "switchto")
                            {
                                //chromeBrowser.SwitchTo().DefaultContent();
                                //chromeBrowser.SwitchTo().Frame("controlIframe");
                                //chromeBrowser.SwitchTo().Frame("dialog-popup-annl-worldhent");
                                IWebElement iwebEle = IWebSearchIwebItem(chromeBrowser, clsIwebItemInfoItem, strItemName);

                                chromeBrowser.SwitchTo().Frame(iwebEle);
                            }
                            else if (clsIwebItemInfoItem.strItemTagName.ToLower() == "switchtodefault")
                            {

                                chromeBrowser.SwitchTo().DefaultContent();
                            }

                            else

                            {

                                //chromeBrowser.SwitchTo().Window(chromeBrowser.WindowHandles[0]);
                                //多层级找到该元素
                                IWebElement iwebEle = IWebSearchIwebItem(chromeBrowser, clsIwebItemInfoItem, strItemName);
                                //iwebEle = chromeBrowser.FindElement(By.Id("capShouldType_C"));
                                //chromeBrowser.SwitchTo().DefaultContent();
                                //chromeBrowser.SwitchTo().Frame("editIframe");
                                // chromeBrowser.SwitchTo().Frame("editIframe");



                                //对该元素操作 ，input  click 
                                //iwebEle = chromeBrowser.FindElement(By.Id("capShouldType_C"));//.FindElement(By.Id("inv"));
                                //iwebEle = chromeBrowser.FindElement(By.Id("one")).FindElements(By.TagName("input"))[0];
                                //chromeBrowser.SwitchTo().Frame("controlIframe");
                                //iwebEle = chromeBrowser.FindElement(By.Id("infoDiv")).FindElements(By.TagName("table"))[0].FindElements(By.TagName("a"))[0];


                                //IList<IWebElement> ddilitiwebEle = chromeBrowser.FindElement(By.Id("CAPSOUL")).FindElements(By.TagName("option"));
                                //clsBrowser.ClickElementByLoopFind(ddilitiwebEle, "840", clsIwebItemInfoItem.strLoopSelectBy);
                                //iwebEle.Click();
                                //iwebEle.SendKeys("沃尔玛（WAL-MART STORES）");
                                //行动

                                //iwebEle.SendKeys("2015-01-24");
                                OperateIwebItem(clsIwebItemInfoItem.strAction, iwebEle, clsIwebItemInfoItem, strItemName);
                            }



                            if (strItemName.Contains("保存按钮"))
                            {
                                Console.WriteLine("");
                            }


                        }
                        catch (Exception ex)
                        {
                            blisSucceed = false;
                            strCurLog = ex.Message;
                        }
                        if (!blisSucceed)
                        {

                            litLog.Add(strCurProvince);
                            litLog.Add(strCurCity);
                            litLog.Add(strCurStore);
                            litLog.Add(striwebItem);
                            litLog.Add(strCurLog);
                            litLog.Add(DateTime.Now.ToString());

                            //WriterTXTLog(litLog);
                            //GetScreen();

                            clsIwebItemInfoItem.strUploadResult = strCurLog;
                        }
                        else
                        {
                            strCurLog = strLogTitle + " is succeed !";
                            clsIwebItemInfoItem.strUploadResult = "";
                            clsIwebItemInfoItem.strUploadStatus = "Succeed";
                        }

                        if (DicIwebItems.ContainsKey(striwebItem))
                        {

                            DicIwebItems[striwebItem] = clsIwebItemInfoItem;
                        }
                        //}
                        //while (1 == 1);

                        intRetry = 3;


                        //}
                        //catch (Exception ex)
                        //{
                        //    strCurLog = strItemName +"is Fail"+ Environment.NewLine+ex.Message;
                        //}
                    }
                    string strStore = "";
                    string strIwebItem = "";
                    foreach (DataRow dataRow in tblUploadData.Rows)
                    {
                        strStore = dataRow["店号"].ToString();
                        strIwebItem = dataRow["字段"].ToString();
                        if (strCurStore == strStore)

                        {
                            if (DicIwebItems.ContainsKey(strIwebItem))
                            {


                                if (DicIwebItems[strIwebItem].strUploadStatus != "Succeed")
                                {

                                    if (DicIwebItems[strIwebItem].strUploadResult != "" && DicIwebItems[strIwebItem].strUploadResult != null)
                                    {
                                        dataRow["上传结果"] = DicIwebItems[strIwebItem].strUploadResult;
                                    }
                                    else
                                    {
                                        dataRow["上传结果"] = "无上传要求，请检查确认";
                                    }

                                    dataRow["上传状态"] = "Fail";
                                }
                                else
                                {
                                    dataRow["上传状态"] = "Succeed";
                                    dataRow["上传结果"] = "";
                                }


                            }
                            else
                            {

                                dataRow["上传状态"] = "";
                                dataRow["上传结果"] = "无上传要求，请检查确认";

                            }




                        }

                    }
                    using (ExcelHelper excel = new ExcelHelper(strUploadDataFilePath, true, false, true))
                    {
                        Worksheet ws = excel.CurrentSht("Sheet1");
                        ws.Cells.ClearContents();
                        excel.SetData(tblUploadData, "Sheet1");

                        ws= excel.CurrentSht("Finished List");
                        ws.Activate();
                        int intStartRow = ws.Cells[ws.Rows.Count, 1].End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row + 1;
                        Range rng = ws.Cells[intStartRow, 1];
                        rng.Value = strCurStore;

                    }

                    strCurLog = "程序已经填写结束，请查看";

                }
                catch (Exception ex)
                {

                    strCurLog = "程序意外出错：" + ex.Message;

                }
            }
            while (intRetry < 1);
        }

        /// <summary>
        /// 在Config 对应省份的IwebItem 写入对应的upload值
        /// </summary>
        public void SetDicIwebItemValue()
        {


            DicIwebItems = new Dictionary<string, clsIwebItemInfo>();
            DataTable tblTemp = new DataTable();
            tblTemp = tblUploadData.Select("省份='" + strCurProvince + "' and 城市='" + strCurCity + "' and 店号='" + strCurStore + "'").CopyToDataTable();
            string strIwebItem = "";
            foreach (DataRow dataRow in tblTemp.Rows)

            {

                if (dataRow["是否上传"].ToString() == "是")
                {
                    strIwebItem = dataRow["字段"].ToString();

                    if (!DicIwebItems.ContainsKey(strIwebItem))
                    {
                        clsIwebItemInfo clsIwebItemInfo = new clsIwebItemInfo();
                        clsIwebItemInfo.strItemUploadValue = dataRow["字段内容"].ToString();
                        DicIwebItems[strIwebItem] = clsIwebItemInfo;
                    }

                }
            }





        }

        public void setclsIwebItemInfoItem(clsIwebItemInfo clsIwebItemInfoItem)
        {

            DataTable tblItemConfig = clsIwebItemInfoItem.tblItemConfig;
            foreach (DataRow dr in tblItemConfig.Rows)
            {
                clsIwebItemInfoItem.strAction = dr["Action"].ToString().Trim();
                clsIwebItemInfoItem.strItemTagName = dr["TagName"].ToString().Trim();
                clsIwebItemInfoItem.strModule = dr["Module"].ToString().Trim();
                clsIwebItemInfoItem.strItemName = dr["ItemName"].ToString().Trim();

                clsIwebItemInfoItem.strLoopSelectBy = dr["LoopSelectBy"].ToString().Trim();
                clsIwebItemInfoItem.strLoopSelectisFixedvalue = dr["LoopSelectisFixedvalue"].ToString().Trim();
                clsIwebItemInfoItem.strcheckboxBy = dr["CheckSelectBy"].ToString().Trim();
                clsIwebItemInfoItem.strAfterSendKey = dr["AfterSendKey"].ToString().Trim();
                clsIwebItemInfoItem.strItemConfigValue = dr["default"].ToString().Trim();
                clsIwebItemInfoItem.strIsValueFromData = dr["IsValueFromData"].ToString().Trim().ToLower();


                if (dr["BeforeDelayTime"].ToString().Trim() == "")
                {
                    clsIwebItemInfoItem.intItemBeforeDelayTime = 1;
                }
                else
                {
                    clsIwebItemInfoItem.intItemBeforeDelayTime = int.Parse(dr["BeforeDelayTime"].ToString().Trim());
                }
            }
        }

        public IList<IWebElement> IWebSearchIwebList(IWebDriver p_IWebDriver, clsIwebItemInfo clsIwebItemInfoItem)
        {

            IList<IWebElement> litiwebEle = null;
            DataTable tblItemConfig = clsIwebItemInfoItem.tblItemConfig;

            Thread.Sleep(clsIwebItemInfoItem.intItemBeforeDelayTime * 1000);

            IWebElement iwebEle = null;
            int intSelectIndex;
            foreach (DataRow dr in tblItemConfig.Rows)
            {


                string strSearchType = dr["SearchType"].ToString().Trim();

                string strfindElementType = dr["findElementType"].ToString().Trim().ToLower();
                string strParentID = dr["ParentID"].ToString().Trim();
                intSelectIndex = 0;
                if (dr["SelectIndex"].ToString().Trim() != "")
                {
                    intSelectIndex = Convert.ToInt32(dr["SelectIndex"].ToString().Trim());
                }




                if (strfindElementType == "checkbox" || strfindElementType == "loopselect")
                {
                    switch (strSearchType.ToLower())
                    {
                        case "id":
                            if (iwebEle == null)
                            {
                                litiwebEle = p_IWebDriver.FindElements(By.Id(strParentID));
                            }
                            else
                            {
                                litiwebEle = iwebEle.FindElements(By.Id(strParentID));
                            }

                            break;

                        case "tagname":

                            if (iwebEle == null)
                            {
                                litiwebEle = p_IWebDriver.FindElements(By.TagName(strParentID));
                            }
                            else
                            {
                                litiwebEle = iwebEle.FindElements(By.TagName(strParentID));
                            }
                            break;

                        case "name":

                            if (iwebEle == null)
                            {
                                litiwebEle = p_IWebDriver.FindElements(By.Name(strParentID));
                            }
                            else
                            {
                                litiwebEle = iwebEle.FindElements(By.Name(strParentID));
                            }


                            break;





                    }


                    break;
                }

                else if (strfindElementType == "findElement")
                {


                    switch (strSearchType.ToLower())
                    {
                        case "id":
                            if (iwebEle == null)
                            {
                                iwebEle = p_IWebDriver.FindElement(By.Id(strParentID));
                            }
                            else
                            {
                                iwebEle = iwebEle.FindElement(By.Id(strParentID));
                            }

                            break;

                        case "tagname":

                            if (iwebEle == null)
                            {
                                iwebEle = p_IWebDriver.FindElement(By.TagName(strParentID));
                            }
                            else
                            {
                                iwebEle = iwebEle.FindElement(By.TagName(strParentID));
                            }




                            break;

                        case "name":

                            if (iwebEle == null)
                            {
                                iwebEle = p_IWebDriver.FindElement(By.Name(strParentID));
                            }
                            else
                            {
                                iwebEle = iwebEle.FindElement(By.Name(strParentID));
                            }

                            break;
                        case "classname":

                            if (iwebEle == null)
                            {
                                iwebEle = p_IWebDriver.FindElement(By.ClassName(strParentID));
                            }
                            else
                            {
                                iwebEle = iwebEle.FindElement(By.ClassName(strParentID));
                            }

                            break;

                    }

                }
                else

                {
                    switch (strSearchType.ToLower())
                    {
                        case "id":
                            if (iwebEle == null)
                            {
                                iwebEle = p_IWebDriver.FindElements(By.Id(strParentID))[intSelectIndex];
                            }
                            else
                            {
                                iwebEle = iwebEle.FindElements(By.Id(strParentID))[intSelectIndex];
                            }

                            break;

                        case "tagname":

                            if (iwebEle == null)
                            {
                                iwebEle = p_IWebDriver.FindElements(By.TagName(strParentID))[intSelectIndex];
                            }
                            else
                            {
                                iwebEle = iwebEle.FindElements(By.TagName(strParentID))[intSelectIndex];
                            }




                            break;

                        case "name":

                            if (iwebEle == null)
                            {
                                iwebEle = p_IWebDriver.FindElements(By.Name(strParentID))[intSelectIndex];
                            }
                            else
                            {
                                iwebEle = iwebEle.FindElements(By.Name(strParentID))[intSelectIndex];
                            }

                            break;
                        case "classname":

                            if (iwebEle == null)
                            {
                                iwebEle = p_IWebDriver.FindElements(By.ClassName(strParentID))[intSelectIndex];
                            }
                            else
                            {
                                iwebEle = iwebEle.FindElements(By.ClassName(strParentID))[intSelectIndex];
                            }

                            break;


                    }

                }



            }
            return litiwebEle;
        }

        public IWebElement IWebSearchIwebItem(IWebDriver p_IWebDriver, clsIwebItemInfo clsIwebItemInfoItem, string strLogTitle)
        {

            DataTable tblItemConfig = clsIwebItemInfoItem.tblItemConfig;
            IWebElement iwebEle = null;


            int intSelectIndex;

            try
            {


                Thread.Sleep(clsIwebItemInfoItem.intItemBeforeDelayTime * 1000);
                intRowIndex = -1;
                foreach (DataRow dr in tblItemConfig.Rows)
                {

                    intRowIndex = intRowIndex + 1;
                    string strSearchType = dr["SearchType"].ToString().Trim();

                    string strfindElementType = dr["findElementType"].ToString().Trim();
                    string strParentID = dr["ParentID"].ToString().Trim();
                    intSelectIndex = 0;
                    if (dr["SelectIndex"].ToString().Trim() != "")
                    {
                        intSelectIndex = Convert.ToInt32(dr["SelectIndex"].ToString().Trim());
                    }



                    // if FindElement or  FindElements
                    //如果时 特别的 select  时，使用text 值 来选择对应的item
                    if (strfindElementType == "Ctrlselect")
                    {
                        break;
                    }

                    else if (strfindElementType == "FindElement")
                    {


                        switch (strSearchType.ToLower())
                        {
                            case "id":
                                if (iwebEle == null)
                                {

                                    iwebEle = p_IWebDriver.FindElement(By.Id(strParentID));


                                }
                                else
                                {
                                    iwebEle = iwebEle.FindElement(By.Id(strParentID));
                                }

                                break;

                            case "tagname":

                                if (iwebEle == null)
                                {
                                    iwebEle = p_IWebDriver.FindElement(By.TagName(strParentID));
                                }
                                else
                                {
                                    iwebEle = iwebEle.FindElement(By.TagName(strParentID));
                                }




                                break;

                            case "name":

                                if (iwebEle == null)
                                {
                                    iwebEle = p_IWebDriver.FindElement(By.Name(strParentID));
                                }
                                else
                                {
                                    iwebEle = iwebEle.FindElement(By.Name(strParentID));
                                }
                                break;

                            case "classname":

                                if (iwebEle == null)
                                {
                                    iwebEle = p_IWebDriver.FindElement(By.ClassName(strParentID));
                                }
                                else
                                {
                                    iwebEle = iwebEle.FindElement(By.ClassName(strParentID));
                                }
                                break;



                        }


                    }
                    else

                    {
                        switch (strSearchType.ToLower())
                        {
                            case "id":
                                if (iwebEle == null)
                                {
                                    iwebEle = p_IWebDriver.FindElements(By.Id(strParentID))[intSelectIndex];
                                }
                                else
                                {
                                    iwebEle = iwebEle.FindElements(By.Id(strParentID))[intSelectIndex];
                                }

                                break;

                            case "tagname":



                                if (iwebEle == null)
                                {
                                    iwebEle = p_IWebDriver.FindElements(By.TagName(strParentID))[intSelectIndex];
                                }
                                else
                                {
                                    iwebEle = iwebEle.FindElements(By.TagName(strParentID))[intSelectIndex];
                                }
                                break;

                            case "name":

                                if (iwebEle == null)
                                {
                                    iwebEle = p_IWebDriver.FindElements(By.Name(strParentID))[intSelectIndex];
                                }
                                else
                                {
                                    iwebEle = iwebEle.FindElements(By.Name(strParentID))[intSelectIndex];
                                }
                                break;

                            case "classname":

                                if (iwebEle == null)
                                {
                                    iwebEle = p_IWebDriver.FindElements(By.ClassName(strParentID))[intSelectIndex];
                                }
                                else
                                {
                                    iwebEle = iwebEle.FindElements(By.ClassName(strParentID))[intSelectIndex];
                                }
                                break;


                        }

                    }


                    if (tblItemConfig.Rows.Count == 1 || intRowIndex == tblItemConfig.Rows.Count - 2)
                    {
                        clsIwebItemInfoItem.IWebDirectParent = iwebEle;
                    }

                }
            }
            catch (Exception ex)
            {
                blisSucceed = false;
                if (strLogTitle.Contains("删除"))
                {

                    strCurLog = strLogTitle + "找不到默认跳过" + ex.Message;
                }
                else

                {
                    strCurLog = strLogTitle + "is Fail,无法找到该UI " + ex.Message;
                }




            }
            return iwebEle;

        }

        //实际要做的动作 input click ,select, etc
        public void OperateIwebItem(string strAction, IWebElement iwebEle, clsIwebItemInfo clsIwebItemInfoItem, string strLogTitle)
        {


            //选择config value or uploaddata value
            string strValue = "";
            if (clsIwebItemInfoItem.strIsValueFromData == "true")
            {

                strValue = clsIwebItemInfoItem.strItemUploadValue;
            }
            else
            {
                strValue = clsIwebItemInfoItem.strItemConfigValue;

            }




            //Thread.Sleep(clsIwebItemInfoItem.intItemBeforeDelayTime*1000);
            try
            {
                switch (strAction.ToLower())
                {



                    case "input":

                        //纳税总额  The total amount of tax paid
                        if (clsIwebItemInfoItem.strItemName.Contains("外商投资经营情况-纳税总额") ||
                            clsIwebItemInfoItem.strItemName.Contains("资产负债情况") ||
                            clsIwebItemInfoItem.strItemName.Contains("资产状况信息") ||//资产负债情况 资产状况信息
                            clsIwebItemInfoItem.strItemName.Contains("生产经营情况信息"))//生产经营情况信息

                        {
                            //需要特殊化处理 输入值会弹出提示窗口，需要点击确定


                            SetandCheckText_BySetTextofTax(iwebEle, strValue, strLogTitle, true, false, false, true, clsIwebItemInfoItem.strAfterSendKey);

                        }

                        else
                        {
                            if (clsIwebItemInfoItem.strItemName != null && clsIwebItemInfoItem.strItemTagName.ToLower() == "select")
                            {
                                SetandCheckText_BySetText(iwebEle, strValue, strLogTitle, false, false, false, false, clsIwebItemInfoItem.strAfterSendKey);
                            }

                            else
                            {
                                SetandCheckText_BySetText(iwebEle, strValue, strLogTitle, false, false, false, true, clsIwebItemInfoItem.strAfterSendKey);
                            }



                        }

                        //if (clsIwebItemInfoItem.strItemName != null && clsIwebItemInfoItem.strItemTagName.ToLower() == "select")
                        //{
                        //    SetandCheckText_BySetText(iwebEle, strValue, strLogTitle, false, false, false, false, clsIwebItemInfoItem.strAfterSendKey);
                        //}

                        //else
                        //{
                        //    SetandCheckText_BySetText(iwebEle, strValue, strLogTitle, true, false, false, true, clsIwebItemInfoItem.strAfterSendKey);
                        //}


                        //chromeBrowser.SwitchTo().Window(chromeBrowser.WindowHandles[1]);
                        //chromeBrowser.FindElement(By.Id("ifhasweb")).Clear();
                        //chromeBrowser.FindElement(By.Id("ifhasweb")).SendKeys("是");

                        //select element 也可以使用input action，


                        break;

                    case "click":

                        //iwebEle = chromeBrowser.FindElements(By.ClassName("btnAllA"))[0].FindElements(By.ClassName("button5"))[0];
                        //iwebEle.Click();
                        if (clsIwebItemInfoItem.strItemName.Contains("(Check)"))//clsIwebItemInfoItem.strItemName.Contains("保存按钮") || 
                        {
                            CatchAndClickElement(iwebEle, strLogTitle, false, true, clsIwebItemInfoItem.IWebDirectParent);

                        }
                        else
                        {
                            CatchAndClickElement(iwebEle, strLogTitle, false);

                        }

                        //点击保存后 关闭提示框
                        if (clsIwebItemInfoItem.strItemName.Contains("保存按钮"))//clsIwebItemInfoItem.strItemName.Contains("保存按钮") || 
                        {
                            Thread.Sleep(2000);

                            try
                            {
                                chromeBrowser.SwitchTo().DefaultContent();
                                Thread.Sleep(1000);

                                IList<IWebElement> ilstIWebElements = chromeBrowser.FindElements(By.ClassName("ui-button-text"));

                                for (int i = ilstIWebElements.Count - 1; i > -1; i--)
                                {
                                    ilstIWebElements[i].Click();
                                }



                            }

                            catch
                            {

                            }

                        }



                        break;

                    case "ctrlselect":


                        // string strValue = clsIwebItemInfoItem.strItemConfigValue;

                        foreach (string strItem in strValue.Split(','))
                        {


                            if (dicConfigTextValue.ContainsKey(strItem))
                            {
                                CatchAndCtrlSelctElement(iwebEle, dicConfigTextValue[strItem], strLogTitle, false, false, false);

                            }
                            else
                            {
                                blisSucceed = false;
                                strCurLog = strLogTitle + ": " + dicConfigTextValue[strItem] + " 在config文件没有对应网站 Value";

                            }
                        }



                        break;
                    case "gettext":


                        Boolean blNeedThrowException = false;

                        int intRetry = 0;

                        do
                        {
                            intRetry += 1;
                            try
                            {


                                if (clsIwebItemInfoItem.strItemName == "登录异常提示")
                                {
                                    //不一定存在（如果登录成功他是正常的）
                                    try
                                    {
                                        litLog.Add(strLogTitle + ": 这个节点发生错误。" + "<*>" + clsBrowser.GetText_byIWebElement(iwebEle));
                                    }
                                    catch
                                    {

                                        throw;
                                    }


                                }
                                else
                                {
                                    clsBrowser.GetText_byIWebElement(iwebEle);
                                }
                                intRetry = 3;
                            }
                            catch (Exception ex)
                            {
                                Thread.Sleep(2000);
                                if (intRetry == 3)
                                {
                                    if (blNeedThrowException)
                                    {
                                        // GetScreen(intRowIndex.ToString());
                                        throw new Exception(strLogTitle + ": 这个节点发生错误。" + "<*>" + ex);
                                    }
                                    else
                                    {
                                        // GetScreen(intRowIndex.ToString());
                                        litLog.Add(strLogTitle + ": 这个节点发生错误。" + "<*>" + ex);
                                    }
                                    // GetScreen(intRowIndex.ToString());

                                    //throw new Exception(strLogTitle + ": 这个节点发生错误。" + "<*>" + ex);
                                }
                            }


                        }

                        while (intRetry < 3);




                        break;

                }
            }
            catch (Exception ex)
            {
                blisSucceed = false;
                if (strLogTitle.Contains("删除"))
                {
                    strCurLog = strLogTitle + "找不到默认跳过" + ex.Message;
                }
                else

                {
                    strCurLog = strLogTitle + "is Fail,无法找到该UI " + ex.Message;
                }
            }



        }


        /// <summary>
        /// 因特殊无法使用配置表配置，所以该function的参数除日期外都是写死的
        /// 使用于股东及出资信息-认缴出资时间 股东及出资信息-实缴出资时间选择日期的Item
        /// </summary>
        /// <param name="strDate"></param>
        public void SelectDate(string strDate, IWebElement p_iweb)
        {
            List<string> lstYMDDate = new List<string>();
            lstYMDDate = strDate.Split('-').ToList();
            if (lstYMDDate.Count != 3)
            {
                blisSucceed = false;
                lstYMDDate = DateTime.Now.ToString("yyyy-MM-dd").Split('-').ToList();
                strCurLog = "日期格式有问题应为：yyyy-MM-dd,目前程序自动填写当前上传的日期，是不对的是需要手工更改";
            }
            chromeBrowser.SwitchTo().DefaultContent();
            chromeBrowser.SwitchTo().Frame(chromeBrowser.FindElement(By.Id("investorIframe")));


            p_iweb.Click();
            Thread.Sleep(1000);
            p_iweb.SendKeys(OpenQA.Selenium.Keys.Tab);
            Thread.Sleep(1000);

            //Month Input
            chromeBrowser.SwitchTo().DefaultContent();
            chromeBrowser.SwitchTo().Frame(chromeBrowser.FindElements(By.TagName("iframe"))[6]);
            IWebElement iwebMonth = chromeBrowser.FindElement(By.Id("dpTitle")).FindElements(By.ClassName("yminputfocus"))[0];//yminputfocus
            iwebMonth.SendKeys(lstYMDDate[1]);
            iwebMonth.SendKeys(OpenQA.Selenium.Keys.Tab);
            Thread.Sleep(1000);

            //YEar input
            iwebMonth = chromeBrowser.FindElement(By.Id("dpTitle")).FindElements(By.ClassName("yminputfocus"))[0];
            iwebMonth.SendKeys(lstYMDDate[0]);
            iwebMonth.SendKeys(OpenQA.Selenium.Keys.Tab);
            Thread.Sleep(1000);

            //Day Selection
            IList<IWebElement> ilstIwebEle = chromeBrowser.FindElements(By.ClassName("WdayTable"))[0].FindElements(By.TagName("td"));

            foreach (IWebElement webElement in ilstIwebEle)
            {
                if (webElement.Text.PadLeft(2, '0') == lstYMDDate[2].PadLeft(2, '0'))
                {
                    webElement.Click();
                    Thread.Sleep(2000);
                    break;
                }


            }
            chromeBrowser.SwitchTo().DefaultContent();
        }


        /// <summary>
        /// 加载 upload data 数据源.
        /// </summary>
        public void LoadingData()
        {





            //if (strDataType == "法人")
            //{
            //    strUploadDataFilePath = @"C:\Users\c0c04nc\Documents\Project\合规年报\AIC Annual Report Build Data\result\法人 UploadCorporateData.xlsx";
            //}
            //else
            //{
            //    strUploadDataFilePath = @"C:\Users\c0c04nc\Documents\Project\合规年报\AIC Annual Report Build Data\result\分支机构 UploadCorporateData.xlsx";
            //}
            try
            {
                //GetSuccessfulSht();
                GetUploadData();
                string strStore = "";
                dicFinishedStores = new Dictionary<string, string>();
                dicStore = new Dictionary<string, clsStoreInfo>();
                foreach (DataRow dataRow in tblSuccessfulStores.Rows)
                {
                    strStore = dataRow["店号"].ToString();


                    if (!dicFinishedStores.ContainsKey(strStore))
                    {
                        dicFinishedStores[strStore] = null;
                    }
                }



                lstStores = new List<string>();
                
                foreach (DataRow dataRow in tblUploadData.Rows)
                {
                    //filter 已经跑过的store
                    strStore = dataRow["店号"].ToString();
                    if (!dicFinishedStores.ContainsKey(strStore))
                    {
                        if (!lstStores.Contains(strStore))
                        {
                            lstStores.Add(strStore);
                        }

                        if (!dicStore.ContainsKey(strStore))
                        {
                            clsStoreInfo clsStoreInfoItem = new clsStoreInfo();
                            clsStoreInfoItem.strProvince = dataRow["省份"].ToString();
                            clsStoreInfoItem.strCity = dataRow["城市"].ToString();

                            clsStoreInfoItem.strUserName = dataRow["用户名"].ToString();
                            clsStoreInfoItem.strURL = dataRow["网址"].ToString();
                            clsStoreInfoItem.strPassword = dataRow["密码"].ToString();
                            clsStoreInfoItem.strIDNo = dataRow["联络员身份证号码"].ToString();

                            dicStore[strStore] = clsStoreInfoItem;

                        }
                    }



                }



            }
            catch (Exception ex)
            {
                strCurLog = "加载数据源异常：" + ex.Message;

            }






        }
        /// <summary>
        /// 将已经成功的Store no 写入dicSuccessfulStores
        /// </summary>
        public void GetSuccessfulSht()
        {







        }
           




        public class clsStoreInfo
        {


            public string strProvince { get; set; }

            public string strCity { get; set; }

            public string strUserName { get; set; }

            public string strURL { get; set; }

            public string strPassword { get; set; }

            public string strIDNo { get; set; }
        }
        public class clsIwebItemInfo
        {

            //网站上的item 显示字段
            public string strItemConfigValue { get; set; }
            public string strItemUploadValue { get; set; }
            public DataTable tblItemConfig { get; set; }

            public string strAction { get; set; }

            public string strItemName { get; set; }
            
            //直接上一级parent，目前是有一些保存按钮需要判断父级窗口是否正常关闭
            public IWebElement IWebDirectParent { get; set; }


            public int intItemBeforeDelayTime { get; set; }

            public string strItemTagName { get; set; }

            //网站上字段对应的模块
            public string strModule { get; set; }

            //仅使用于LoopSelect
            public string strLoopSelectBy { get; set; }

            //仅使用于LoopSelect
            
            public string strLoopSelectisFixedvalue { get; set; }

            //仅使用于checkbox SHIYNG
            public string strcheckboxBy { get; set; }

            public string strAfterSendKey { get; set; }

            public string strUploadResult { get; set; }

            public string strUploadStatus { get; set; }


            //IsValueFromData

            //判断数据是否来自UploadData
            public string strIsValueFromData { get; set; }
            ////item  对应的所有父级，包括它本身0
            //public string strParentID { get; set; }
            ////父级对应的顺序
            //public Int32 intParentIndex { get; set; }

            //public string strSearchType { get; set; }

            ////定义这个Item空间的类型
            //public string strTagName { get; set; }

            ////部分search 需要 对应的序号，比如 table 的指定行列，checkbox ,下拉列表
            //public Int32 intSelectIndex { get; set; }






        }








        //public void Start()
        //{








        //    litLog = new List<string>();

        //    StrConfigFilePath = @"C:\Users\c0c04nc\Documents\Project\合规年报\AIC Annual Report\Config.xlsx";
        //    Config();
        //    strLogPath = dicConfigDefault["strLogPath"];




        //    strUploadSinWinDataPath = dicConfigDefault["strUploadSinWinDataPath"];
        //    strUSBUploadSinWinDataPath = dicConfigDefault["strUSBUploadSinWinDataPath"];
        //    blNeedClickSubmitBtn = Boolean.Parse(dicConfigDefault["blNeedClickSubmitBtn"]);
        //    //如果没有找到文件则生成Log,如果存在另存到指定位置，上传结束后再移动到U盘
        //    if (File.Exists(strUSBUploadSinWinDataPath))
        //    {
        //        File.Copy(strUSBUploadSinWinDataPath, strUploadSinWinDataPath, true);
        //        Thread.Sleep(2000);
        //        //复制成功后直接删掉源文件，避免重复录入
        //        File.Delete(strUSBUploadSinWinDataPath);
        //    }
        //    else
        //    {
        //        List<object> litLog = new List<object>();
        //        litLog.Add("U盘路径不存在文件");
        //        litLog.Add(DateTime.Now);
        //        WriterTXTLog(litLog);

        //        return;
        //    }

        //    //使用chromeDriver 驱动打开chrome 


        //    intRetry = 0;
        //    do
        //    {
        //        try
        //        {

        //            intRetry += 1;
        //            clsBrowser = new clsBrowserAutomation();

        //            //打开Chrome浏览器
        //            chromeBrowser = clsBrowser.openChrome();
        //            //web窗口最大化
        //            chromeBrowser.Manage().Window.Maximize();
        //            //打开单一窗口网站

        //            if (litLog.Count != 0)
        //            {
        //                string strLog = string.Join(Environment.NewLine, litLog.ToArray());
        //                throw new Exception(strLog);
        //            }


        //            intRetry = 3;
        //        }
        //        catch (Exception ex)
        //        {
        //            Thread.Sleep(2000);
        //            if (intRetry == 3)
        //            {

        //                List<object> litLog = new List<object>();
        //                litLog.Add("chromeDriver异常");
        //                litLog.Add(DateTime.Now);
        //                WriterTXTLog(litLog);


        //                return;
        //            }
        //        }
        //    }
        //    while (intRetry < 3);

        //    //判断单一窗口是否在运行，
        //    Process[] ProcessItem = Process.GetProcessesByName("chromedriver");
        //    intRetry = 0;
        //    do
        //    {
        //        try
        //        {

        //            intRetry += 1;
        //            chromeBrowser.Navigate().GoToUrl(dicConfigDefault["MainUrl"].ToString());
        //            if (litLog.Count != 0)
        //            {
        //                string strLog = string.Join(Environment.NewLine, litLog.ToArray());
        //                throw new Exception(strLog);
        //            }

        //            intRetry = 3;
        //        }
        //        catch (Exception ex)
        //        {
        //            Thread.Sleep(2000);
        //            if (intRetry == 3)
        //            {

        //                List<object> litLog = new List<object>();
        //                litLog.Add("单一窗口遇到网路问题无法打开");
        //                litLog.Add(DateTime.Now);
        //                WriterTXTLog(litLog);

        //                chromeBrowser.Quit();
        //                return;
        //            }
        //        }
        //    }
        //    while (intRetry < 3);

        //     Boolean blNeedLogin = true;

        //    using (ExcelHelper excel = new ExcelHelper(strUploadSinWinDataPath, true, false,true))
        //    {
        //        dblUploadData = excel.GetData("Sheet1");
        //        intRowIndex=1;
        //        Worksheet ws = excel.CurrentSht("Sheet1");
        //        int intStatusCol = dblUploadData.Columns["Status"].Ordinal+1;
        //        int intResultCol = dblUploadData.Columns["Result"].Ordinal+1;
        //        int intRunTimeCol = dblUploadData.Columns["RunTime"].Ordinal+1;
        //        Range rng = null;
        //        foreach (DataRow dbrItem in dblUploadData.Rows)
        //        {
        //            litLog = new List<string>();
        //            litTrackLog = new List<string>();
        //            intRowIndex += 1;

        //            if (dbrItem["Status"].ToString() != "Succeed")
        //            {
        //                rng = ws.Cells[intRowIndex, intStatusCol];
        //                rng.Value = "";

        //                rng = ws.Cells[intRowIndex, intResultCol];
        //                rng.Value = "";
        //                try
        //                {
        //                    dbrUploadData = dbrItem;
        //                    //登录单一窗口网站
        //                    if (blNeedLogin)
        //                    {
        //                        //关闭通知窗口
        //                        CloseNotice();
        //                        Login();
        //                        Thread.Sleep(20000);
        //                        blNeedLogin = false;
        //                    }

        //                    //直接跳转到银行服务窗口
        //                    NavigateBankingServices();
        //                    //跳转到填写信息界面
        //                    GotoMainInterface();

        //                    if (litLog.Count != 0)
        //                    {
        //                        GetScreen(intRowIndex.ToString());
        //                        string strLog = string.Join(Environment.NewLine, litLog.ToArray());
        //                        if (litTrackLog.Count != 0)
        //                        {
        //                            strLog = strLog + Environment.NewLine + string.Join(Environment.NewLine, litTrackLog.ToArray());
        //                        }
        //                        throw new Exception(strLog);
        //                    }

        //                    rng = ws.Cells[intRowIndex, intStatusCol];
        //                    rng.Value = "Succeed";
        //                }
        //                catch (Exception ex)
        //                {
        //                    rng = ws.Cells[intRowIndex, intStatusCol];
        //                    rng.Value = "Fail";

        //                    rng = ws.Cells[intRowIndex, intResultCol];
        //                    rng.Value = ex.Message;



        //                    //dbrUploadData["Result"] = ex.Message;

        //                }

        //                rng = ws.Cells[intRowIndex, intRunTimeCol];
        //                rng.Value = DateTime.Now;
        //            }
        //        }
        //    }

        //    //关闭 driver
        //    chromeBrowser.Quit();

        //    //把文件更新到u盘
        //    try
        //    {
        //        File.Copy(strUploadSinWinDataPath, strUSBUploadSinWinDataPath, true);
        //    }
        //    catch
        //    {

        //        List<object> litLog = new List<object>();
        //        litLog.Add("文件更新到u盘失败");
        //        litLog.Add(DateTime.Now);
        //        WriterTXTLog(litLog);



        //    }

        //}
        private void Login()
        {

            //try
            //{
                chromeBrowser.SwitchTo().Frame("loginIframe1");


                CloseNotice();

                //CatchAndClickElement("点击卡介质",dicConfig["CardmediumControlID"]);
                //chromeBrowser.SwitchTo().DefaultContent();
                //chromeBrowser.SwitchTo().Frame("loginIframe2");
                
                //SetandCheckText_BySetText(dicConfig["PasswordControlID"], dbrUploadData["CardVerification"].ToString(),"输入登录密码",true,false);
                 
                //CatchAndClickElement("点击登录",dicConfig["loginbuttonControlID"]);
                //chromeBrowser.FindElement(By.Id(dicConfig["loginbuttonControlID"])).Click();
            //}
            //catch (Exception)
            
            //{ 
            
            
            
            //}
        }

        private void CloseNotice()
        {
           int intRetry = 0;
            do
            {
                try
                {
                    Thread.Sleep(3000);
                    intRetry += 1;
                    chromeBrowser.FindElement(By.ClassName(dicConfig["CloseNoticeControlClassName"])).Click();
                    
                    intRetry = 3;
                }
                catch (Exception ex)
                {
                    Thread.Sleep(2000);
                }
            }
            while (intRetry < 3);


        }

        private void NavigateBankingServices()
        {
           
            //等待登录时间
            Thread.Sleep(5000);
            CloseNotice();
            chromeBrowser.Navigate().GoToUrl(dicConfigDefault["BankingServicesUrl"]);


        }



            


        #region 信息录入的Function

        
        private void CatchAndCtrlSelctElement(IWebElement p_IWebElement, string strText, string strLogTitle = "", Boolean blnNeedCheckInputText = false, Boolean blnNeedCheckErrorText = true, Boolean blNeedSendCtrlKey = false, Boolean blNeedThrowException = false)
        {
            string strTemp = "";
            int intRetry = 0;
            if (blnNeedCheckInputText) strTemp = strText;
            do
            {
                try
                {
                    intRetry += 1;
                    clsBrowser.SelectElement_SelectText(p_IWebElement, strText, blNeedSendCtrlKey);
                    //if (clsBrowser.GetText_byIWebElement(p_IWebElement) == strTemp) intRetry = 3;

                    if (blnNeedCheckInputText)
                    {
                        if (clsBrowser.GetText_byIWebElement(p_IWebElement) == strTemp) intRetry = 3;
                    }
                    else intRetry = 3;


                }
                catch (Exception ex)
                {
                    Thread.Sleep(1000);
                    if (intRetry == 3)
                    {
                        if (blNeedThrowException)
                        {
                            GetScreen(intRowIndex.ToString());
                            throw new Exception(strLogTitle + ": 这个节点发生错误" + "<*>" + ex);
                        }
                        else
                        {
                            GetScreen(intRowIndex.ToString());
                            litLog.Add(strLogTitle + ": 这个节点发生错误" + "<*>" + ex);
                        }
                    }
                }
            }
            while (intRetry < 3);

            //查看填写信息是否规范
            //if (blnNeedCheckErrorText) GetErrorElementText(strLogTitle, strControlID + "-error");
        }

        private void SetandCheckText_BySetText(IWebElement p_IWebElement, string strText, string strLogTitle = "", Boolean blnNeedCheckInputText = false, Boolean blnNeedCheckErrorText = true, Boolean blNeedSendEnterKey = false, Boolean blNeedThrowException = false,string AfterSendKey="")
        {
            string strTemp = "";
            int intRetry = 0;
            if (blnNeedCheckInputText) strTemp = strText;
            do
            {
                try
                {
                    intRetry += 1;
                    clsBrowser.SetText_byIWebElement(p_IWebElement, strText, blNeedSendEnterKey, AfterSendKey, blnNeedCheckInputText);
                    intRetry = 3;

                    // if (clsBrowser.GetText_byIWebElement(p_IWebElement) == strTemp) intRetry = 3;

                    //if (blnNeedCheckInputText)
                    //{
                    //    if (clsBrowser.GetText_byIWebElement(p_IWebElement) == strTemp) intRetry = 3;
                    //}
                    //else
                    //{
                    //    intRetry = 3;
                    //}

                }
                catch (Exception ex)
                {
                   Thread.Sleep(1000);
                    if (intRetry == 3)
                    {
                        if (blNeedThrowException)
                        {
                            GetScreen(intRowIndex.ToString());
                            throw new Exception(strLogTitle + ": 这个节点发生错误" + "<*>" + ex);
                        }
                        else
                        {
                            GetScreen(intRowIndex.ToString());
                            litLog.Add(strLogTitle + ": 这个节点发生错误" + "<*>" + ex);
                        }
                    }
                }
            }
            while (intRetry <3);

        }
        /// <summary>
        /// 纳税总额需要特殊化处理，每次控制它都会弹出一个确定窗口，需要技术关闭它
        /// </summary>
        /// <param name="p_IWebElement"></param>
        /// <param name="strText"></param>
        /// <param name="strLogTitle"></param>
        /// <param name="blnNeedCheckInputText"></param>
        /// <param name="blnNeedCheckErrorText"></param>
        /// <param name="blNeedSendEnterKey"></param>
        /// <param name="blNeedThrowException"></param>
        /// <param name="AfterSendKey"></param>
        private void SetandCheckText_BySetTextofTax(IWebElement p_IWebElement, string strText, string strLogTitle = "", Boolean blnNeedCheckInputText = false, Boolean blnNeedCheckErrorText = true, Boolean blNeedSendEnterKey = false, Boolean blNeedThrowException = false, string AfterSendKey = "")
        {
            string strTemp = "";
            int intRetry = 0;
            if (blnNeedCheckInputText) strTemp = strText;
            do
            {
                try
                {
                    if (p_IWebElement == null)
                    {
                        intRetry = 3;
                        throw new Exception("无法找到该元素");



                    }
                    
                    intRetry += 1;

                    try
                    {
                        Thread.Sleep(1000);

                        IList<IWebElement> ilstIWebElements= chromeBrowser.FindElements(By.ClassName("ui-button-text"));
                        for (int i = ilstIWebElements.Count-1; i > -1; i--)
                        {
                            ilstIWebElements[i].Click();
                        }



                    }

                    catch
                    { }
                    try
                    {
                        p_IWebElement.Clear();

                    }
                    catch
                    { }
                    try
                    {
                        Thread.Sleep(1000);


                        IList<IWebElement> ilstIWebElements = chromeBrowser.FindElements(By.ClassName("ui-button-text"));
                        for (int i = ilstIWebElements.Count - 1; i > -1; i--)
                        {
                            ilstIWebElements[i].Click();
                        }




                        //foreach (IWebElement webElement1 in chromeBrowser.FindElements(By.ClassName("ui-dialog"))[0].FindElements(By.ClassName("ui-button-text")))
                        //{
                        //    //IWebElement webElement = webElement1.FindElements(By.TagName("button"))[0];

                        //    webElement1.Click();

                        //}

                        //IWebElement webElement = chromeBrowser.FindElements(By.ClassName("ui-dialog"))[0].FindElements(By.ClassName("ui-dialog-buttonset"))[0].FindElements(By.TagName("button"))[0];

                        //webElement.Click();


                    }
                    catch
                    { }




                    p_IWebElement.SendKeys(strText);


                    // if (clsBrowser.GetText_byIWebElement(p_IWebElement) == strTemp) intRetry = 3;

                    if (blnNeedCheckInputText)
                    {
                        if (clsBrowser.GetText_byIWebElement(p_IWebElement) == strTemp) intRetry = 3;
                    }
                    else
                    {
                        intRetry = 3;
                    }

                    p_IWebElement.SendKeys(Keys.Tab);

                    try
                    {
                        Thread.Sleep(1000);

                        IList<IWebElement> ilstIWebElements = chromeBrowser.FindElements(By.ClassName("ui-button-text"));
                        for (int i = ilstIWebElements.Count - 1; i > -1; i--)
                        {
                            ilstIWebElements[i].Click();
                        }

                    }
                    catch
                    { }




                }
                catch (Exception ex)
                {
                    Thread.Sleep(1000);
                    if (intRetry == 3)
                    {
                        if (blNeedThrowException)
                        {
                            GetScreen(intRowIndex.ToString());
                            throw new Exception(strLogTitle + ": 这个节点发生错误" + "<*>" + ex);
                        }
                        else
                        {
                            GetScreen(intRowIndex.ToString());
                            litLog.Add(strLogTitle + ": 这个节点发生错误" + "<*>" + ex);
                        }
                    }
                }
            }
            while (intRetry < 3);

        }
        private void CatchAndClickElement(IWebElement p_IWebElement, string strLogTitle = "", Boolean blNeedThrowException = false,Boolean blNeedCheckStatus=false, IWebElement p_IWebElementDirectParent=null)
        {

            int intRetry = 0;

            do
            {
                intRetry += 1;
                try
                {


                    if (p_IWebElement != null)
                    {
                        p_IWebElement.Click();
                    }
                    else
                    {
                        intRetry = 3;
                        throw (new Exception("该IWeb Element是null"));
                    }





                    //当时click保存按钮时，可能需要父级判断状态
                    if (blNeedCheckStatus)
                    {
                        try
                        {
                            int intdoIndex = 0;
                            do
                            {
                                Thread.Sleep(2000);
                                intdoIndex = intdoIndex + 1;
                            }
                            while (p_IWebElementDirectParent.Displayed && intdoIndex<3);
                        }
                        catch
                        { }


                    }
                    //p_IWebElement.FindElements(By.)


                    intRetry = 3;
                }
                catch (Exception ex)
                {
                    Thread.Sleep(2000);
                    if (intRetry == 3)
                    {
                        if(blNeedThrowException)
                        {
                           // GetScreen(intRowIndex.ToString());
                            throw new Exception(strLogTitle + ": 这个节点发生错误。" + "<*>" + ex);
                        }
                        else
                        {
                           // GetScreen(intRowIndex.ToString());
                            litLog.Add(strLogTitle + ": 这个节点发生错误。" + "<*>" + ex);
                        }
                       // GetScreen(intRowIndex.ToString());

                        //throw new Exception(strLogTitle + ": 这个节点发生错误。" + "<*>" + ex);
                    }
                }


            }

            while (intRetry<3);
        
        }

        private void SetandCheckText_BySetText_ByName(string strControlName, string strText, string strLogTitle = "",Boolean blnNeedCheckInputText=false, Boolean blnNeedCheckErrorText = true, Boolean blNeedSendEnterKey=false, Boolean blNeedThrowException = false)
        {
            string strTemp = "";
            int intRetry = 0;
            if (blnNeedCheckInputText) strTemp = strText;
            do
            {

                try
                {
                    intRetry += 1;
                    clsBrowser.SetText_ByName(chromeBrowser, strControlName, strText, blNeedSendEnterKey);
                    if(blnNeedCheckInputText)
                    {
                        if (clsBrowser.GetText_ByName(chromeBrowser, strControlName) == strTemp) intRetry = 3;
                    }
                    else intRetry = 3;
                }
                catch (Exception ex)
                {
                    Thread.Sleep(2000);
                    if (intRetry == 3)
                    {
                        if (blNeedThrowException)
                        {
                            GetScreen(intRowIndex.ToString());
                            throw new Exception(strLogTitle + ": 这个节点发生错误" + "<*>" + ex);
                        }
                        else
                        {
                            GetScreen(intRowIndex.ToString());
                            litLog.Add(strLogTitle + ": 这个节点发生错误" + "<*>" + ex);
                        }
                        
                        //
                    }
                }
            }
            while (intRetry < 3);
            //查看填写信息是否规范
            if (blnNeedCheckErrorText) GetErrorElementText(strLogTitle, "", strControlName + "-error");

        }
        private void SetandCheckText_BySelectItem(string strControlID, string strSelectText, string strLogTitle = "", Boolean blnNeedCheckInputText = false,Boolean blNeedThrowException=false)
        {
            SelectElement SelectElement=null;
            string strTemp = "";
            int intRetry = 0;
            if (blnNeedCheckInputText) strTemp = strSelectText;
            do
            {
                
                try
                {
                    intRetry += 1;
                    clsBrowser.SelectItem(chromeBrowser, strControlID, strSelectText);
                    SelectElement = new SelectElement(chromeBrowser.FindElement(By.Id(strControlID)));
                    if(SelectElement.SelectedOption.Text == strTemp) intRetry = 3;

                }
                catch (Exception ex)
                {
                    Thread.Sleep(2000);

                    if (intRetry == 3)
                    {
                        if (blNeedThrowException)
                        {
                            GetScreen(intRowIndex.ToString());
                            throw new Exception(strLogTitle + ": 这个节点发生错误" + "<*>" + ex);
                        }
                        else
                        {
                            GetScreen(intRowIndex.ToString());
                            litLog.Add(strLogTitle + ": 这个节点发生错误" + "<*>" + ex);
                        }
                    }
                }

            }
            while (intRetry < 3);//SelectElement.SelectedOption.Text != strTemp && intRetry < 3

        }

        private void GetErrorElementText(string strLogTitle = "", string strID = "", string strName = "", string strClassName = "", Boolean blNeedThrowException = false)
        {
            
            IWebElement iwebEle=null;
            Thread.Sleep(1000);
            int intRetry = 0;
            do
            {
                intRetry += 1;
                try
                {
                    if (strID != "") iwebEle = chromeBrowser.FindElement(By.Id(strID));
                    if (strName != "") iwebEle = chromeBrowser.FindElement(By.Name(strName));
                    if (strClassName != "") iwebEle = chromeBrowser.FindElement(By.ClassName(strClassName));
                    
                    string strValue = iwebEle.Text;

                    if (strValue == null) strValue = "";

                    intRetry = 3;
                    if (strValue != "")
                    {
                        
                        throw new ApplicationException(strLogTitle + ": " + strValue + " <*>");
                    }
                }
                catch (ApplicationException ex)
                {

                    if (intRetry == 3)
                    {
                        if (blNeedThrowException)
                        {
                            GetScreen(intRowIndex.ToString());
                            throw new Exception(ex.Message);
                        }
                        else
                        {

                            GetScreen(intRowIndex.ToString());
                            litLog.Add(ex.Message);
                        }
                    }

                    
                  // throw ex;
                    
                }
                catch (Exception ex)
                {
                    
                    Thread.Sleep(500);
                    
                }

            }

            while (intRetry < 3);
           
        }

        #endregion 信息录入 Function
        public void Config()
        {

            try
            {
                string strName;
                string strVilid;
                string strControlValue;
                string strDefalut;
                
                DataTable tblTemp = null;
                DataTable tblList = null;
                using (ExcelHelper excel = new ExcelHelper(StrConfigFilePath, true, true, true))
                {

                    tblTemp = excel.GetData("分配清单");
                    string strText;

                    try
                    {

                        try
                        {
                            tblList = tblTemp.Select("类型='" + strDataType + "' and 地区='" +strCurCity + "'").CopyToDataTable();
                        }
                        catch
                        {
                            tblList = tblTemp.Select("类型='" + strDataType + "' and 地区='" + strCurProvince + "'").CopyToDataTable();

                        }

                    }
                    catch
                    {
                        throw new Exception("没有类型" + strDataType + " 地区" + strCurProvince+"，所以找不到配置表");
                
                    }


                    tblConfig = excel.GetData(tblList.Rows[0]["配置表"].ToString());

                    foreach (DataRow row in tblConfig.Rows)
                    {
                        strName = row["ItemName"].ToString().Trim();
                        strVilid = row["vilid"].ToString().Trim();
                        if (strName != "" && !litIwebItems.Contains(strName) && strVilid.ToUpper() == "TRUE")
                        {

                            

                            //当 投资者信息-投资者性质 是 境内投资者才激活使用
                            if (strName == "投资者信息-反向投资股权投资额" ||
                                strName == "投资者信息-反向投资股权比例" ||
                                strName == "投资者信息-境外投资者是否为内地与香港、澳门《CEPA服务贸易协议》规定的香港/澳门服务提供者" ||
                                strName == "投资者信息-境外投资者是否为内地与香港、澳门《CEPA投资协议》规定的香港/澳门投资者" ||
                                strName == "投资者信息-境外投资者是否是华侨")
                            {
                                if (DicIwebItems.ContainsKey("投资者信息-投资者性质"))
                                {
                                    if (DicIwebItems["投资者信息-投资者性质"].strItemUploadValue == "境外投资者")
                                    {
                                        litIwebItems.Add(strName);
                                    } 
                                }
                            }
                            else
                            {
                                litIwebItems.Add(strName);
                            }



                        }
                        //if (!dicConfigDefault.ContainsKey(strName) && strName != "") dicConfigDefault.Add(strName, strDefalut);
                        //if (!dicConfig.ContainsKey(strControlID) && strControlID != "") dicConfig.Add(strControlID, strControlValue);
                    }

                    tblTemp = excel.GetData(tblList.Rows[0]["参数表"].ToString());
                    
                    foreach (DataRow row in tblTemp.Rows)
                    {
                        
                        strText= row["Text"].ToString().Trim();
                        strName = row["Value"].ToString().Trim();
                        if (!dicConfigTextValue.ContainsKey(strText) && strText != "") dicConfigTextValue.Add(strText, strName);
                       // if (!dicConfig.ContainsKey(strControlID) && strControlID != "") dicConfig.Add(strControlID, strControlValue);

                        //if (strName != "") dblUploadData.Columns.Add(strName, typeof(string));
                    }
                }

            }
            catch
            {
                blisSucceed = false;
                strCurLog = "config 文件无法找到，或者读取异常";

            }


        }

        public void GetUploadData()
        {

            try
            {
                
                using (ExcelHelper excel = new ExcelHelper(strUploadDataFilePath, true, false, true))
                {
                    tblUploadData = excel.GetData("Sheet1");
                    
                    //生成一个finished  list sht 并取数（已存在直接取数）
                    Workbook wb = excel.Currentwb();

                    Worksheet mySheet;



                    if (!excel.Exists("Finished List"))
                    {
                        object[,] arData = new object[1, 2];
                        mySheet = wb.Worksheets.Add(Type.Missing, wb.Worksheets[wb.Worksheets.Count], 1, XlSheetType.xlWorksheet);
                        mySheet.Name = "Finished List";

                        arData[0, 0] = "店号";
                        arData[0, 1] = "状态";

                        Range startRange = mySheet.Cells[1, 1];
                        Range endRange = mySheet.Cells[1, 2];
                        Range range = mySheet.Range[startRange, endRange];

                        range.Value = arData;

                    }
                    
                    tblSuccessfulStores = excel.GetData("Finished List");



                }

            }
            catch(Exception ex)
            {
                blisSucceed = false;
                strCurLog = "UploadData 文件无法找到，或者读取异常"+ex.Message;

            }


        }



        private void SetdbrDefalutData(DataRow dbrUpload)
        {
            foreach (var item in dicConfigDefault.Keys)
            {
                if (dbrUpload.Table.Columns.Contains(item))
                {
                    dbrUpload[item] = dicConfigDefault[item];
                }

            }
        
        
        
        }

        public string p_strLogMessages;

        private int _intLogMsgMaxLength;

        public object ExcelNS { get; private set; }

        private void AddLog(string in_strLogMessage)
        {
            string strQryInsert = "INSERT INTO Logging (LogAt,LogMessage) VALUES (@LogAt,@LogMessage)";
            int intInsertedRecords;
           
            //Add into the variable
            p_strLogMessages = in_strLogMessage + Environment.NewLine + p_strLogMessages;

            //Validate the log message lenght
            if (in_strLogMessage.Length > _intLogMsgMaxLength)
            {
                in_strLogMessage = in_strLogMessage.Substring(0, _intLogMsgMaxLength);
            }

            ////Set the parameters
            clsDBParameter[] parms = new clsDBParameter[2];
            //DateTime.Now.ToString("G")
            parms[0] = new clsDBParameter("@LogAt", DbType.DateTime, DateTime.Now.ToString("G"));
            parms[1] = new clsDBParameter("@LogMessage", DbType.String, in_strLogMessage);

            //Insert into database
            //intInsertedRecords = objOleDB.insertDataToDB(strQryInsert, parms);
        }


        private void WriterTXTLog(List<string> litNewLog)
        {

            // 从文件中读取并显示每行
            string strNewLog= string.Join(" ", litNewLog);

            List<string> litTXT = new List<string>();
            string line = "";
            using (StreamReader sr = new StreamReader(strLogPath))
            {
                while ((line = sr.ReadLine()) != null)
                {
                    litTXT.Add(line);
                }
            }
            litTXT.Add(strNewLog);
            using (StreamWriter sw = new StreamWriter(strLogPath))
            {
                foreach (string s in litTXT)
                {
                    sw.WriteLine(s);

                }
            }

        }

        public void GetScreen(string FileName="Cur")
        {
            //获取整个屏幕图像,不包括任务栏

            // Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            string strdf = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            //DirectoryInfo info = new DirectoryInfo(Environment.CurrentDirectory);
            DirectoryInfo strEnvPath = new DirectoryInfo(string.Format(@"{0}..\..\..\..\", Environment.CurrentDirectory));
            

            string strScreenShotPath =Path.Combine(strEnvPath.FullName, "log Screenshot");
            System.Drawing.Rectangle ScreenArea = Screen.PrimaryScreen.WorkingArea;
            Bitmap bmp = new Bitmap(ScreenArea.Width, ScreenArea.Height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.CopyFromScreen(0, 0, 0, 0, new Size(ScreenArea.Width, ScreenArea.Height));
                string dateTime = DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss");
                strScreenShotPath = Path.Combine(strScreenShotPath, FileName +" _ " + dateTime + ".jpg");
                bmp.Save(strScreenShotPath);//

            }
            //return;
        }
    }
}
