using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace TimeTimer
{
    public class ExcelApp
    {
        private static Excel._Application app;
        public void ExportData(DataGridView dgv, string[] titleArr = null)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                app = new Excel.Application();
                Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                Excel._Worksheet worksheet = null;
                worksheet = workbook.Sheets[1];
                worksheet = workbook.ActiveSheet;
                int startRow = 0;
                if (titleArr != null)
                {
                    foreach (string str in titleArr)
                    {
                        string strTemp = str;
                        if (str.Contains("{DATE}"))
                            strTemp = str.Replace("{DATE}", Convert.ToDateTime(dgv.Rows[0].Cells["Дані_станом_на"].Value).ToShortDateString());
                        worksheet.Cells[startRow + 1, 1] = strTemp;
                        startRow++;
                    }
                }
                for (int i = 1; i < dgv.Columns.Count + 1; i++)
                    worksheet.Cells[startRow + 1, i] = dgv.Columns[i - 1].HeaderText;
                Excel.Range oRange = worksheet.Range[worksheet.Cells[startRow + 2, 1], worksheet.Cells[startRow + 1 + dgv.Rows.Count, dgv.Columns.Count]];
                object[,] arr = new object[dgv.Rows.Count, dgv.Columns.Count];
                for (int i = 0; i < dgv.Rows.Count; i++)
                {
                    for (int j = 0; j < dgv.Columns.Count; j++)
                    {
                        //if (dgv.Columns[j].ValueType.ToString().Equals("System.String"))
                        //    arr[i, j] = dgv.Rows[i].Cells[j].Value.ToString().Replace("\r\n", "");
                        //else
                        //    arr[i, j] = dgv.Rows[i].Cells[j].Value;
                        if (dgv.Columns[0].HeaderText == "Файл")
                        {
                            if (dgv.Rows[i].Cells[j].Value != null)
                                arr[i, j] = j == 0 ? "=HYPERLINK(\"" + dgv.Rows[i].Cells[j].Value.ToString() + "\",\"" + dgv.Rows[i].Cells[j].Value.ToString() + "\")" : dgv.Rows[i].Cells[j].Value;
                            else
                                arr[i, j] = "";
                        }
                        else
                            arr[i, j] = dgv.Rows[i].Cells[j].Value != null ? dgv.Rows[i].Cells[j].Value : "";
                    }
                }
                oRange.Value = arr;
                int lastCol = worksheet.Range["A" + startRow + 1].End[Excel.XlDirection.xlToRight].Column;
                worksheet.Range[worksheet.Cells[startRow + 1, 1], worksheet.Cells[startRow + 1, lastCol]].AutoFilter();
                worksheet.Range[worksheet.Cells[startRow + 1, 1], worksheet.Cells[startRow + 1, lastCol]].WrapText = true;
                worksheet.Range[worksheet.Cells[startRow + 1, 1], worksheet.Cells[startRow + 1, lastCol]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                worksheet.Range[worksheet.Cells[startRow + 1, 1], worksheet.Cells[startRow + 1, lastCol]].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                worksheet.Range["A3", "E3"].Interior.Color = Color.PaleGreen;  // колір клітинки
                worksheet.Range["A3", "E3"].Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;  // рамки в клітинці
               
                //worksheet.Columns["1:" + lastCol].AutoFit();
                oRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[startRow + 1 + dgv.Rows.Count, dgv.Columns.Count]];
                oRange.Columns.AutoFit();
                app.Visible = true;

                string fileName = dgv.Tag != null ? dgv.Tag.ToString() : "";
                if (fileName.Length > 0)
                {
                    DateTime curDate = Convert.ToDateTime(dgv.Rows[0].Cells["Дані_станом_на"].Value);
                    if (fileName.Contains("Актуалізація"))
                    {
                        string path = @"S:\DATA\SPV\03_Актуалізація\AutoCreate\" + curDate.ToString("yyyy.MM.dd") + @"\";
                        if (!Directory.Exists(path)) Directory.CreateDirectory(path);
                        workbook.SaveAs(path + fileName.Replace("{DATE}", curDate.ToShortDateString()) + ".xlsx");
                    }
                    else
                        workbook.SaveAs(fileName.Replace("{DATE}", curDate.ToShortDateString()) + ".xlsx");
                    MessageBox.Show("Файл збережено за посиланням: \n" + workbook.Path, "Вивантажено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                //MessageBox.Show("Дані вивантажено!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                if (app.Workbooks.Count == 1) app.Quit();
                app = null;
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }
        public static void ExportData(DataTable dt)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                app = new Excel.Application();
                Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                Excel._Worksheet worksheet = null;
                worksheet = workbook.Sheets[1];
                worksheet = workbook.ActiveSheet;
                int startRow = 0;
                for (int i = 1; i < dt.Columns.Count + 1; i++)
                    worksheet.Cells[startRow + 1, i] = dt.Columns[i - 1].ColumnName;
                Excel.Range oRange = worksheet.Range[worksheet.Cells[startRow + 2, 1], worksheet.Cells[startRow + 1 + dt.Rows.Count, dt.Columns.Count]];
                object[,] arr = new object[dt.Rows.Count, dt.Columns.Count];
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                        arr[i, j] = dt.Rows[i][j];
                }
                oRange.Value = arr;
                app.Visible = true;
                Cursor.Current = Cursors.Default;
                //MessageBox.Show("Дані вивантажено!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                if (app.Workbooks.Count == 1) app.Quit();
                app = null;
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void ExportCallQualityData(string formPath, string savePath, string fileFormat, DataTable dt, string contragent, string ocinyvach, string ocinyvanyi, string callType, string opDate, string criticalError, string filePath, string result)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                app = new Excel.Application();
                Excel._Workbook workbook = app.Workbooks.Open(formPath);
                //Excel._Worksheet worksheet = null;
                Excel._Worksheet worksheet = workbook.Sheets[1];
                worksheet = workbook.ActiveSheet;
                worksheet.Name = "Оцінка якості";

                worksheet.Cells.Replace("{OpDateTime}", opDate);
                worksheet.Cells.Replace("{Ochinuvanyy}", ocinyvanyi);
                worksheet.Cells.Replace("{CallType}", callType);
                worksheet.Cells.Replace("{Contragent}", contragent);
                worksheet.Cells.Replace("{FilePath}", filePath);
                worksheet.Cells.Replace("{Ocinyvach}", ocinyvach);
                worksheet.Cells.Replace("{CriticalError}", criticalError);
                worksheet.Cells.Replace("{Result}", result);
                int startRow = 14;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (i < dt.Rows.Count - 1)
                    {
                        worksheet.Rows[(startRow + i) + ":" + (startRow + i)].Copy();
                        worksheet.Rows[(startRow + i + 1) + ":" + (startRow + i + 1)].Insert();
                    }
                    worksheet.Range["B" + (startRow + i)].Value = dt.Rows[i][0].ToString();
                    worksheet.Range["D" + (startRow + i)].Value = dt.Rows[i][1].ToString();
                    worksheet.Range["E" + (startRow + i)].Value = dt.Rows[i][2].ToString();
                }
                worksheet.Rows[7 + ":" + (dt.Rows.Count - 1)].EntireRow().AutoFit();
                if (fileFormat.Contains("xlsx"))
                {
                    workbook.SaveAs(savePath);
                    app.Visible = true;
                }
                else if (fileFormat.Contains("pdf"))
                {
                    workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, savePath);
                    workbook.Close(false);
                    app.Quit();
                    System.Diagnostics.Process.Start(savePath);
                }
                Cursor.Current = Cursors.Default;
                //MessageBox.Show("Дані вивантажено!" + Environment.NewLine + savePath, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                if (app.Workbooks.Count == 1) app.Quit();
                app = null;
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
