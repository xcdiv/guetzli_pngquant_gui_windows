using System;
using System.Collections.Generic;
using System.Text;
using System.Web.UI;
using System.Data;
using System.Web.UI.WebControls;
using System.Windows.Forms;

namespace DataMining.Common
{
    public class DataView2DataTable
    {
        public static string GetCellText(TableCell cell)
        {
            string text = cell.Text;
            if (!string.IsNullOrEmpty(text))
            {
                return text;
            }
            foreach (System.Web.UI.Control control in cell.Controls)
            {
                if (control != null && control is ITextControl)
                {
                    LiteralControl lc = control as LiteralControl;
                    if (lc != null)
                    {
                        continue;
                    }
                    ITextControl l = control as ITextControl;

                    text = l.Text.Replace(" ", "").Trim();
                    break;
                }
            }
            return text;
        }

        /// <summary>
        /// 从GridView的数据生成DataTable
        /// </summary>
        /// <param name="gv">GridView对象</param>
        public static DataTable GridView2DataTable(GridView gv)
        {
            DataTable table = new DataTable();
            int rowIndex = 0;
            List<string> cols = new List<string>();
            if (!gv.ShowHeader && gv.Columns.Count == 0)
            {
                return table;
            }
            GridViewRow headerRow = gv.HeaderRow;
            int columnCount = headerRow.Cells.Count;
            for (int i = 0; i < columnCount; i++)
            {
                string text = GetCellText(headerRow.Cells[i]);
                cols.Add(text);
            }
            foreach (GridViewRow r in gv.Rows)
            {
                if (r.RowType == DataControlRowType.DataRow)
                {
                    DataRow row = table.NewRow();
                    int j = 0;
                    for (int i = 0; i < columnCount; i++)
                    {
                        string text = GetCellText(r.Cells[i]);
                        if (!String.IsNullOrEmpty(text))
                        {
                            if (rowIndex == 0)
                            {
                                DataColumn dc = table.Columns.Add();
                                string columnName = cols[i];
                                if (String.IsNullOrEmpty(columnName))
                                {
                                    columnName = gv.Columns[i].HeaderText;
                                    if (string.IsNullOrEmpty(columnName))
                                    {
                                        continue;
                                    }
                                }
                                dc.ColumnName = columnName;
                                dc.DataType = typeof(string);
                            }
                            row[j] = text;
                        }
                        j++;
                    }
                    rowIndex++;
                    table.Rows.Add(row);
                }
            }
            return table;
        }


        public static DataTable DataGridView2DataTable(DataGridView dv)
        {
            DataTable dt = new DataTable();
            DataColumn dc;
            for (int i = 0; i < dv.Columns.Count; i++)
            {
                dc = new DataColumn();
                dc.ColumnName = dv.Columns[i].HeaderText.ToString();
                dt.Columns.Add(dc);
            }
            for (int j = 0; j < dv.Rows.Count - 1; j++)
            {
                DataRow dr = dt.NewRow();
                for (int x = 0; x < dv.Columns.Count; x++)
                {
                    dr[x] = dv.Rows[j].Cells[x].Value;
                }
                dt.Rows.Add(dr);
            }
            return dt;
        } 
    }
}
