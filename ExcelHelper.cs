using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;

namespace ExcelDueDateNotify
{
    internal class ExcelHelper
    {
        private string _excelPath;
        private string _dueDateEvent_columnLetter;
        private string _dueDate_columnLetter;
        private string _emailTo;
        private string _emailFrom;
        private string _emailAccount;
        private string _emailPassword;
        private string _emailSmtp;
        private int _emailPort;

        #region 构造函数
        public ExcelHelper(ExcelHelperArgs helperArgs)
        {
            #region check null
            if (!File.Exists(helperArgs.FilePath))
            {
                throw new FileNotFoundException("!找不到Excel文件: " + helperArgs.FilePath);
            }
            #endregion

            _excelPath = helperArgs.FilePath;
            _dueDateEvent_columnLetter = helperArgs.DueDateEventColumnLetter;
            _dueDate_columnLetter = helperArgs.DueDateColumnLetter;
            _emailTo = helperArgs.EmailTo;
            _emailFrom = helperArgs.EmailFrom;
            _emailAccount = helperArgs.EmailAccount;
            _emailPassword = helperArgs.EmailPassword;
            _emailSmtp = helperArgs.EmailSmtp;
            _emailPort = int.Parse(helperArgs.EmailPort);
        }
        #endregion


        #region + private List<RowData> GetRowDatas()
        private List<RowData> GetRowDatas()
        {
            FileStream excelFileStream = null;
            XLWorkbook wb = null;

            try
            {
                excelFileStream = File.Open(_excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                wb = new XLWorkbook(excelFileStream);
                var sheet = wb.Worksheets.First();

                var rows = sheet.RowsUsed();

                if (rows.Count() <= 0)
                {
                    throw new ArgumentNullException($"\"{_excelPath}\" 获取行数 <= 0.");
                }

                var rowDatas = new List<RowData>();

                foreach (var row in rows)
                {
                    try
                    {
                        // 获取 这行 due date 单元格
                        var cell = row.Cell(_dueDate_columnLetter);

                        if (cell.TryGetValue(out DateTime due))
                        {
                            rowDatas.Add(new RowData
                            {
                                DueDate = due,
                                DaysOfDueDateSubtractToday = due.Date.Subtract(DateTime.Now.Date).Days,
                                DueDateEvent = row.Cell(_dueDateEvent_columnLetter)?.GetString(),
                                RowNumber = row.RowNumber()
                            });
                        }
                    }
                    catch (Exception)
                    {
                    }
                }

                return rowDatas;
            }
            catch (Exception ex)
            {
                throw new Exception("Fail:@GetRowDatas(): " + ex.Message);
            }
            finally
            {
                wb?.Dispose();
                excelFileStream?.Dispose();
            }
        }
        #endregion

        #region + private void SendEmail(RowData cd)
        private void SendEmail(RowData cd)
        {
            string fileName = Path.GetFileName(_excelPath);

            string subj = $"{fileName} , 第 {cd.RowNumber} 行到期时间 {cd.DueDate:yyyy-MM-dd}";
            MailMessage msg = new MailMessage();
            msg.To.Add(_emailTo);
            msg.From = new MailAddress(_emailFrom);
            msg.Subject = subj;
            msg.Body = cd.DueDateEvent;
            msg.IsBodyHtml = false;

            SmtpClient smtpClient = new SmtpClient();
            smtpClient.Timeout = 5000; // 5秒超时
            smtpClient.Host = _emailSmtp;
            smtpClient.Port = _emailPort;
            smtpClient.EnableSsl = true;
            smtpClient.Credentials = new NetworkCredential(_emailAccount, _emailPassword); //安福怀老院qq邮箱

            try
            {
                smtpClient.Send(msg);
                Console.WriteLine("send email: " + subj);
            }
            catch (Exception ex)
            {
                throw new Exception("邮件发送错误: " + ex.GetBaseException().Message);
            }
        }
        #endregion

        #region + public static void CalculateDueDate_And_SendEmail(string configPath)
        public static void CalculateDueDate_And_SendEmail(string configPath)
        {
            try
            {
                var helper = new ExcelHelper(ExcelHelperArgs.GetArgs(configPath));
                var rowData = helper.GetRowDatas();

                bool noExpired = true;

                foreach (var data in rowData)
                {
                    if (data.DaysOfDueDateSubtractToday == 1)
                    {
                        noExpired = false;

                        // send email
                        helper.SendEmail(data);
                    }
                }

                if (noExpired)
                {
                    Console.WriteLine("没有到期事件需提醒.");
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        #endregion
    }
}
