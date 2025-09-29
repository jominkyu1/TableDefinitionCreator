using System;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using TableDefinitionCreator.enums;

namespace TableDefinitionCreator.utils
{
    internal static class ControlExtensions
    {
        /// <summary>
        /// 컨트롤의 더블버퍼를 활성화 합니다.
        /// (예외 무시)
        /// </summary>
        /// <param name="control"></param>
        /// <param name="enabled">활성화 여부</param>
        /// <param name="applyToChilds">자식 컨트롤에도 적용할지 여부</param>
        public static void SetDoubleBuffered(this Control control, bool enabled, bool applyToChilds)
        {
            try
            {
                const BindingFlags bindingFlags = BindingFlags.Instance | BindingFlags.NonPublic;
                PropertyInfo pi = control.GetType().GetProperty("DoubleBuffered", bindingFlags);
                pi?.SetValue(control, enabled, null);

                if (applyToChilds)
                {
                    foreach (Control child in control.Controls)
                        SetDoubleBuffered(child, enabled, true);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Control {control?.Name} SetDoubleBuffered Failed : {ex.Message}");
            }
        }

        /// <summary>
        /// DataGridView의 모든 컬럼의 정렬 모드를 일괄 변경합니다.
        /// </summary>
        /// <param name="dgv">정렬 모드를 변경할 DataGridView</param>
        /// <param name="sortMode">적용할 정렬 모드</param>
        /// <remarks>
        /// 지정된 DataGridView의 모든 컬럼에 동일한 정렬 모드가 적용됩니다.
        /// </remarks>
        public static void SwitchColumnSortable(this DataGridView dgv, DataGridViewColumnSortMode sortMode)
        {
            dgv.Columns.Cast<DataGridViewColumn>()
                .ToList()
                .ForEach(c => c.SortMode = sortMode);
        }

        /// <summary>
        /// DataTable의 내용을 클립보드로 복사합니다.
        /// </summary>
        /// <param name="dt">복사할 데이터가 포함된 DataTable</param>
        /// <param name="copyType">복사 형식</param>
        /// <param name="includeHeader">열 헤더를 포함할지 여부</param>
        /// <remarks>
        /// 데이터가 없는 경우(Rows.Count == 0) 복사를 수행하지 않습니다.
        /// </remarks>
        public static void CopyToClipBoard(this DataTable dt, CopyType copyType, bool includeHeader)
        {
            if (dt.Rows.Count == 0)
                return;

            string seperator;
            switch (copyType)
            {
                case CopyType.TSV:
                    seperator = "\t";
                    break;

                case CopyType.CSV:
                    seperator = ",";
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(copyType), copyType, null);
            }

            StringBuilder sb = new StringBuilder();
            if (includeHeader)
            {
                AppendHeader(dt, sb, seperator);
            }

            AppendRows(dt, sb, seperator);
            Clipboard.SetText(sb.ToString());
        }

        private static void AppendHeader(DataTable dt, StringBuilder sb, string seperator)
        {
            string arrHeaderNames = dt.Columns.Cast<DataColumn>()
                .Select(c => c.ColumnName)
                .Aggregate((a, b) => a + seperator + b);

            sb.AppendLine(arrHeaderNames);
        }

        private static void AppendRows(DataTable dt, StringBuilder sb, string seperator)
        {
            foreach (DataRow row in dt.Rows)
            {
                string arrRowValues = row.ItemArray
                    .Select(v => v == DBNull.Value ? "" : v.ToString())
                    .Aggregate((a, b) => a + seperator + b);

                sb.AppendLine(arrRowValues);
            }
        }
    }
}
