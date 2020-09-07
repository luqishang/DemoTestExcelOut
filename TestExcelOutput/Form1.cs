using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using PA.Office.ExcelObjects;

namespace TestExcelOutput
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnTestExcelOutput_Click(object sender, EventArgs e)
        {         
            string templeteFileName = "..\\..\\excelTemplate\\検収記録_template_1.xlsx";
            //string outputExcelFileName = "..\\..\\excelTemplate\\PDFTemp\\検収記録.xlsx";
            //string outputExcelFileName = this.saveExcelFileDialog.FileName;
            string outputExcelFileName = GetOutPutPath("検収記録.xlsx");
            string pdfFileName = GetOutPutPath("検収記録.pdf");

            File.Copy(templeteFileName, outputExcelFileName, true);
            File.SetAttributes(outputExcelFileName, FileAttributes.Normal);

            //シート名
            string sheetName = "検収記録_例.店舗全体";

            //目標行の画線をコピーする
            ExcelFileSingleton excelSingleton = ExcelFileSingleton.GetInstance();

            try
            {
                excelSingleton.OpenExcel(outputExcelFileName);
                //新規の行(3行)を指定列(15:15)に挿入する
                excelSingleton.InsertRowOfSheet(sheetName, 15, 3);
                //新規の列(3列)を指定列(D:D)に挿入する
                excelSingleton.InsertColOfSheet(sheetName, "D", 3);

                //セルの値を設定する
                List<ExcelRowObject> rows = new List<ExcelRowObject>();
                ExcelRowObject headRow = new ExcelRowObject();
                ExcelCellObject cellD5 = new ExcelCellObject();
                cellD5.RowIndex = 5;
                cellD5.ColIndex = 4;
                cellD5.Value = "数量";

                ExcelCellObject cellE5 = new ExcelCellObject();
                cellE5.RowIndex = 5;
                cellE5.ColIndex = 5;
                cellE5.Value = "重量";

                ExcelCellObject cellF5 = new ExcelCellObject();
                cellF5.RowIndex = 5;
                cellF5.ColIndex = 6;
                cellF5.Value = "単価";

                headRow.Cells.Add(cellD5);
                headRow.Cells.Add(cellE5);
                headRow.Cells.Add(cellF5);
                rows.Add(headRow);
                excelSingleton.WriteRowsToSheet(sheetName, rows);

            }
            catch (Exception ex)
            {
                ex.ToString();
                throw ex;
            }
            finally
            {
                excelSingleton.CloseExcel();
            }

            //ExcelファイルをPDFに変換する
            ExcelSave excelSave = new ExcelSave();
            excelSave.SaveAsPdf(outputExcelFileName, pdfFileName);

            //PDFに変換後、TempのExcelファイルを削除する
            File.Delete(outputExcelFileName);

            MessageBox.Show("Output have been finished!!!");
        }

        private string GetOutPutPath(string fileName)
        {
            string path = "";
            string curPath = Directory.GetCurrentDirectory();
            int position = curPath.IndexOf("TestExcelOutput");
            path = curPath.Substring(0, position + 16);
            path += "excelTemplate\\PDFTemp\\" + fileName;

            return path;
        }

        /// <summary>
		/// EXCELのセルの内容を設定、取得します。
		/// </summary>
		/// <param name="year">年</param>
        /// <param name="month">月</param>
		/// <returns>Dictionary<key日, value曜日></returns>
		/// <remarks></remarks>
        public static Dictionary<int, string> Day(int year, int month)
        {
            //日付、曜日を格納するための辞書
            Dictionary<int, string> day = new Dictionary<int, string>();

            //月末が何日までか取得
            var firstDay = new DateTime(year, month, 1);
            var lastday = new DateTime(firstDay.Year, firstDay.Month, DateTime.DaysInMonth(firstDay.Year, firstDay.Month));

            //日付、曜日を格納
            for (int i = 1; i <= lastday.Day; i++)
            {
                //日付を定義
                var theDay = new DateTime(year, month, i);
                //曜日を取得
                string week = theDay.ToString("ddd");
                //辞書に追加
                day.Add(i, week);
            }
            return day;
        }
    }
}
