using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using System.Text.RegularExpressions;
using System.Net.Security;
namespace AIC_Annual_Report
{
    class ControlEmail
    {
        public ExchangeService exchange = new ExchangeService();
        public DirectoryInfo TheFolder = null;
        public string strNewSenderAddress;
        //创建一个 ExchangeService
        public ControlEmail(string EmailAddress)
            
        {
            int RetryNum = 0;
            do
            {
                try
                {
                    RetryNum += 1;

                        exchange.Credentials = (System.Net.NetworkCredential) System.Net.CredentialCache.DefaultCredentials;
                        exchange.AutodiscoverUrl(EmailAddress);
                }
                catch (Exception ex)

                {
                    if (RetryNum == 3)
                    {
                        throw new Exception(@"exchange Credentials is Error" + "<*>" + Environment.CommandLine + ex.ToString());
                    }
                }
            }
            while (RetryNum < 3);


        }




        public void SaveAttachments(string SavePath, string strApprovalEmailAdd, int intAddHours, string strKeySubject,string strSenderAddress="")//

        {
            int RetryNum = 0;

            do
            {
                try
                {
                    RetryNum += 1;
                    //ExchangeService exchange = new ExchangeService();
                    //引用window 用户的账号密码,同时需要转换credentials类型
                    //exchange.Credentials = (System.Net.NetworkCredential)System.Net.CredentialCache.DefaultCredentials;

                    //exchange.AutodiscoverUrl(EmailAddress);

                    Mailbox mbApproval;
                    Folder fldApproval;
                    //设置需要连接的邮箱
                    mbApproval = new Mailbox(strApprovalEmailAdd, "SMTP");
                    fldApproval = Folder.Bind(exchange, new FolderId(WellKnownFolderName.Inbox, mbApproval));
                    
                    //SearchFilter.IsGreaterThanOrEqualTo filter = new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd 00:00:00")).AddHours(-200));//IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, date);
                    FindItemsResults<Item> findResults = fldApproval.FindItems(SetFilter(intAddHours, strKeySubject, strSenderAddress), new ItemView(50));//(SetFilter()

                    Boolean blCurFolderExists;
                    //TheFolder = new DirectoryInfo(SavePath);

                    foreach (Item item in findResults)
                    {

                        EmailMessage message = EmailMessage.Bind(exchange, item.Id);
                        strNewSenderAddress = message.Sender.Address;
                        blCurFolderExists = true;
                        foreach (FileAttachment att in message.Attachments)
                        {

                            //找到附件是excel ,
                            if (RegularExpressions(Path.GetFileName(att.Name), @".*_R\.xls.*") !="")
                           
                            {
                                if (Directory.Exists(Path.Combine(SavePath, Path.GetFileNameWithoutExtension(att.Name.Replace("_R","")))) == false)
                                {
                                    TheFolder = Directory.CreateDirectory(Path.Combine(SavePath, Path.GetFileNameWithoutExtension(att.Name.Replace("_R", ""))));
                                    blCurFolderExists = false;

                                }

                                break;
                            }



                        }


                        if (blCurFolderExists == false)
                        {


                            foreach (FileAttachment att in message.Attachments)
                            {
                                if (Path.GetExtension(att.Name) == @".xls" || Path.GetExtension(att.Name) == @".xlsx")
                                {
                                    //删除不是_R的excel文件
                                    
                                    if (RegularExpressions(Path.GetFileName(att.Name), @".*_R\.xls.*") != "")
                                    {
                                        att.Load(Path.Combine(TheFolder.FullName, att.Name));
                                    }

                                }
                                else
                                {
                                    att.Load(Path.Combine(TheFolder.FullName, att.Name));
                                }
                                    
                            }

                        }

                    }

                    RetryNum = 3;
                }
                catch (Exception ex)

                {
                    if (RetryNum == 3)
                    {
                        throw new Exception(@"Getting attachments is Error" + "<*>" + Environment.CommandLine + ex.ToString());

                    }

                }

            }
            while (RetryNum < 3);



        }



        private static SearchFilter SetFilter(int intAddHours, string strKeySubject, string strSenderAddress)
        {


            //SearchFilter SearchFilter1 = new SearchFilter.ContainsSubstring(ItemSchema.Subject, "【NA Major Impact】 【P1】 System Down");
            //SearchFilter SearchFilter2 = new SearchFilter.ContainsSubstring(ItemSchema.Subject, "【NA Major Impact】 【P2】 System Slow");
            //SearchFilter SearchFilter3 = new SearchFilter.ContainsSubstring(ItemSchema.Subject, "【Brazil Major Impact】 【P1】 System Down");
            //SearchFilter SearchFilter4 = new SearchFilter.ContainsSubstring(ItemSchema.Subject, "【Brazil Major Impact】 【P2】 System Slow");

            //筛选今天的邮件
            SearchFilter start = new SearchFilter.IsGreaterThanOrEqualTo(EmailMessageSchema.DateTimeReceived, Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd 00:00:00")).AddHours(intAddHours));
            //SearchFilter end = new SearchFilter.IsLessThanOrEqualTo(EmailMessageSchema.DateTimeCreated, Convert.ToDateTime(DateTime.Now.ToString()));
            SearchFilter subject = new SearchFilter.ContainsSubstring(EmailMessageSchema.Subject, strKeySubject,ContainmentMode.Substring,ComparisonMode.IgnoreCase);
           // SearchFilter subject = new SearchFilter.ContainsSubstring
           SearchFilter SenderAddress = new SearchFilter.ContainsSubstring(EmailMessageSchema.From, strSenderAddress, ContainmentMode.Substring, ComparisonMode.IgnoreCase);

            //SearchFilter IsRead = new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false);

            //SearchFilter.SearchFilterCollection secondLevelSearchFilterCollection1 = new SearchFilter.SearchFilterCollection(LogicalOperator.And,
            //                                                                                                               start,
            //                                                                                                               end);

            //SearchFilter.SearchFilterCollection secondLevelSearchFilterCollection2 = new SearchFilter.SearchFilterCollection(LogicalOperator.Or,
            //                                                                                                             SearchFilter1, SearchFilter2, SearchFilter3, SearchFilter4);

            //SearchFilter.SearchFilterCollection firstLevelSearchFilterCollection = new SearchFilter.SearchFilterCollection(LogicalOperator.And,
            //                                                                                                                   secondLevelSearchFilterCollection1,
            //                                                                                                                   secondLevelSearchFilterCollection2);

           SearchFilter.SearchFilterCollection firstLevelSearchFilterCollection = new SearchFilter.SearchFilterCollection(LogicalOperator.And,start, subject, SenderAddress);


            return firstLevelSearchFilterCollection;
        }

        //    }

        //    catch (Exception ex)

        //    {
        //        if (RetryNum == 3)
        //        {
        //            throw new Exception(@"Failed to connect to data source: " + Environment.CommandLine + ex.ToString());

        //        }

        //    }

        //}

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



        //while (RetryNum < 3);

        public void SendEmail(string EmailAddress, string MessageSubject, string toRecipients,string ccRecipients,string strbody)
        {
           
           
            //引用window 用户的账号密码,同时需要转换credentials类型
            exchange.Credentials = (System.Net.NetworkCredential)System.Net.CredentialCache.DefaultCredentials;

            

            //设置需要连接的邮箱
           // exchange.AutodiscoverUrl(EmailAddress);

            EmailMessage message = new EmailMessage(exchange);



            message.Subject = MessageSubject;
           
            message.ToRecipients.Add(toRecipients);
            //message.CcRecipients.Add(ccRecipients);
            message.Body = strbody;
            message.Send();


        }



    }

}
