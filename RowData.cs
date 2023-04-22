using System;

namespace ExcelDueDateNotify
{
    internal class RowData
    {
        /// <summary>
        /// 到期日期
        /// </summary>
        public DateTime DueDate { get; set; }

        public int RowNumber { get; set; }

        /// <summary>
        /// 单元格日期减去当天天数差
        /// </summary>
        public int DaysOfDueDateSubtractToday { get; set; }

        /// <summary>
        /// 到期事件内容
        /// </summary>
        public string DueDateEvent { get;set; }
    }
}
