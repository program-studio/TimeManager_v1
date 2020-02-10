using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ADGV;
using System.Data;

namespace TimeTimer
{
    public class ADGVManger
    {
        AdvancedDataGridView dgv = new AdvancedDataGridView();
        BindingSource bs = new BindingSource();
        ContextMenuStrip ms = new ContextMenuStrip();
        Label lbInfo = null;
        private string[] titleArr = null;
        public ADGVManger(Panel pn, Label info = null)
        {
            ///призначення усіх необхідні параметрів для відображення нашої таблиці, та всіх потрібних їй елементів(контекстне мею із функціями)
            dgv.Dock = DockStyle.Fill;
            dgv.DoubleBuffered(true);
            dgv.DataSource = bs;
            dgv.ReadOnly = true;
            dgv.AllowUserToDeleteRows = false;
            dgv.AllowUserToAddRows = false;
            dgv.AllowUserToOrderColumns = false;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgv.FilterStringChanged += Dgv_FilterStringChanged;
            dgv.SortStringChanged += Dgv_SortStringChanged;
            dgv.CurrentCellChanged += Dgv_CurrentCellChanged;
            ///додаємо менюшку для забезпечення вивантажень таблиць в ексель
            ToolStripMenuItem exportData = new ToolStripMenuItem("Export data to Excel");
            exportData.Click += ExportData_Click;
            ms.Items.Add(exportData);
            ///додаємо менюшку для забезпечення очищення фільтрів, щоб не перебирати кожну колонку окремо
            ToolStripMenuItem clearFilte = new ToolStripMenuItem("Clear filters");
            clearFilte.Click += ClearFilte_Click;
            ms.Items.Add(clearFilte);
            dgv.ContextMenuStrip = ms;
            pn.Controls.Add(dgv);
            lbInfo = info;
            if (lbInfo != null) lbInfo.Text = string.Empty;
        }
        public string[] TitleArr { set { titleArr = value; } }
        private void Dgv_CurrentCellChanged(object sender, EventArgs e)
        {
            UpdateLabelInfo();
        }
        private void ExportData_Click(object sender, EventArgs e)
        {
            ExcelApp excel = new ExcelApp();
            excel.ExportData(dgv, titleArr);
        }
        private void ClearFilte_Click(object sender, EventArgs e)
        {
            ClearFilter();
        }
        public void ClearFilter()
        {
            dgv.ClearFilter();
            dgv.ClearSort();
            bs.Filter = bs.Sort = string.Empty;
            UpdateLabelInfo();
        }
        private void Dgv_FilterStringChanged(object sender, EventArgs e)
        {
            bs.Filter = dgv.FilterString;
            UpdateLabelInfo();
        }
        private void Dgv_SortStringChanged(object sender, EventArgs e)
        {
            bs.Sort = dgv.SortString;
        }
        public AdvancedDataGridView DGV { get { return dgv; } }
        public void SetSourse(DataTable dt)
        {
            TitleArr = null;
            ClearFilter();
            bs.DataSource = dt;
            Application.DoEvents();
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                dgv.Columns[i].MinimumWidth = 110;
                dgv.Columns[i].ValueType = dt.Columns[i].DataType;
                if (!dt.Columns[i].DataType.ToString().Equals("System.String"))
                    dgv.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                if (dt.Columns[i].DataType.ToString().Equals("System.DateTime"))
                    dgv.Columns[i].DefaultCellStyle.Format = "dd.MM.yyyy HH:mm:ss";
                else if (dt.Columns[i].DataType.ToString().Equals("System.Int32"))
                    dgv.Columns[i].DefaultCellStyle.Format = "0,0";
                else if (dt.Columns[i].DataType.ToString().Equals("System.Double"))
                    dgv.Columns[i].DefaultCellStyle.Format = "0,0";
                UpdateLabelInfo();
            }
        }
        private void UpdateLabelInfo()
        {
            try
            {
                int x = dgv.DataSource != null && dgv.Rows.Count > 0 ? dgv.CurrentRow.Index : 0;
                if (lbInfo != null)
                    lbInfo.Text = x + 1 + " із " + dgv.RowCount.ToString() + " запис(ів)";
            }
            catch (Exception) { }
        }
    }
}
