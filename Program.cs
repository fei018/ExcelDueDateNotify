using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDueDateNotify
{
    internal class Program
    {
        static void Main(string[] args)
        {
			try
			{
				if (ExcelHelperArgs.CheckCmdArgs(args, out string config))
				{
					ExcelHelper.CalculateDueDate_And_SendEmail(config);
				}
			}
			catch (Exception ex)
			{
                Console.WriteLine(ex.Message);
            }
        }
    }
}
