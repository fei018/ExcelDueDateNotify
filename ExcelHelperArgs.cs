using System;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelDueDateNotify
{
    internal class ExcelHelperArgs
    {
        public string FilePath { get; set; }

        public string DueDateColumnLetter { get; set; }

        public string DueDateEventColumnLetter { get; set; }

        public string EmailAccount { get; set; }

        public string EmailPassword { get; set; }

        public string EmailFrom { get; set; }

        public string EmailTo { get; set; }

        public string EmailSmtp { get; set; }

        public string EmailPort { get; set; }

        //===================

        public static ExcelHelperArgs GetArgs(string configPath)
        {
            ExcelHelperArgs inputArgs = new ExcelHelperArgs();

            string[] lines = File.ReadAllLines(configPath);

            if (lines.Length <= 0)
            {
                throw new Exception("配置文件里内容为空.");
            }

            #region check key value
            if (!lines.Any(l=>l.ToLower().Contains("excelfile=")))
            {
                throw new ArgumentNullException("excelFile=", "没有参数");
            }

            if (!lines.Any(l => l.ToLower().Contains("datecolumn=")))
            {
                throw new ArgumentNullException("dateColumn=", "没有参数");
            }

            if (!lines.Any(l => l.ToLower().Contains("eventcolumn=")))
            {
                throw new ArgumentNullException("eventColumn=", "没有参数");
            }

            if (!lines.Any(l => l.ToLower().Contains("emailsmtp=")))
            {
                throw new ArgumentNullException("emailSmtp=", "没有参数");
            }

            if (!lines.Any(l => l.ToLower().Contains("emailport=")))
            {
                throw new ArgumentNullException("emailPort=", "没有参数");
            }

            if (!lines.Any(l => l.ToLower().Contains("emailaccount=")))
            {
                throw new ArgumentNullException("emailAccount=", "没有参数");
            }

            if (!lines.Any(l => l.ToLower().Contains("emailpassword=")))
            {
                throw new ArgumentNullException("emailPassword=", "没有参数");
            }

            if (!lines.Any(l => l.ToLower().Contains("emailfrom=")))
            {
                throw new ArgumentNullException("emailFrom=", "没有参数");
            }

            if (!lines.Any(l => l.ToLower().Contains("emailto=")))
            {
                throw new ArgumentNullException("emailTo=", "没有参数");
            }
            #endregion

            #region get key value
            foreach (string line in lines)
            {
                string[] l = line.Trim()?.Split('=');

                if (l.Length != 2) continue;

                string key = l[0].Trim();
                string value = l[1].Trim();

                if (string.IsNullOrEmpty(value))
                {
                    throw new ArgumentNullException(key + "=", "值为空");
                }

                if (key.Equals("excelFile", StringComparison.OrdinalIgnoreCase))
                {
                    inputArgs.FilePath = value.Trim('"');
                }

                if (key.Equals("dateColumn", StringComparison.OrdinalIgnoreCase))
                {
                    inputArgs.DueDateColumnLetter = value;
                }

                if (key.Equals("eventColumn", StringComparison.OrdinalIgnoreCase))
                {
                    inputArgs.DueDateEventColumnLetter = value;
                }

                if (key.Equals("emailAccount", StringComparison.OrdinalIgnoreCase))
                {
                    inputArgs.EmailAccount = value;
                }

                if (key.Equals("emailPassword", StringComparison.OrdinalIgnoreCase))
                {
                    inputArgs.EmailPassword = value;
                }

                if (key.Equals("emailSmtp", StringComparison.OrdinalIgnoreCase))
                {
                    inputArgs.EmailSmtp = value;
                }

                if (key.Equals("emailPort", StringComparison.OrdinalIgnoreCase))
                {
                    inputArgs.EmailPort = value;
                }

                if (key.Equals("emailFrom", StringComparison.OrdinalIgnoreCase))
                {
                    inputArgs.EmailFrom = value;
                }

                if (key.Equals("emailTo", StringComparison.OrdinalIgnoreCase))
                {
                    inputArgs.EmailTo = value;
                }
            }
            #endregion

            return inputArgs;
        }

        public static bool CheckCmdArgs(string[] args, out string configPath)
        {
            configPath = null;

            StringBuilder sb = new StringBuilder();
            sb.AppendLine();
            sb.Append("命令格式: ExcelDueDateNotify.exe -c 配置文件路径");
            sb.AppendLine().AppendLine();
            sb.AppendLine("-c 配置文件路径, 路径有空格要用英文引号括起来");
            sb.AppendLine();
            sb.AppendLine("------配置文件内容格式-----");
            sb.AppendLine();

            string format = 
@"excelFile=
dateColumn=
eventColumn=
emailSmtp=
emailPort=
emailAccount=
emailPassword=
emailFrom=
emailTo=";

            sb.AppendLine(format);
            string error = sb.ToString();

            if (args.Length != 2 || args[0].ToLower() != "-c")
            {
                Console.WriteLine(error);
                return false;
            }

            configPath = args[1].Trim();

            if (string.IsNullOrEmpty(configPath))
            {
                Console.WriteLine(error);
                return false;
            }

            if (!File.Exists(configPath))
            {
                Console.WriteLine("配置文件不存在.");
                return false;
            }

            return true;
        }
    }
}
