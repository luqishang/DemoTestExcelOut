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
using ExcelView;
using System.Reflection;
using OfficePositionAttributes;
using ExcelCommon;

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
            string templeteFileName = "..\\..\\excelTemplate\\検収記録_template.xlsx";
            //string outputExcelFileName = "..\\..\\excelTemplate\\PDFTemp\\検収記録.xlsx";
            //string outputExcelFileName = this.saveExcelFileDialog.FileName;
            string outputExcelFileName = GetOutPutPath("検収記録.xlsx");
            string pdfFileName = GetOutPutPath("検収記録.pdf");

            File.Copy(templeteFileName, outputExcelFileName, true);
            File.SetAttributes(outputExcelFileName, FileAttributes.Normal);

            //シート名
            string sheetName = "検収記録";

            //目標行の画線をコピーする
            ExcelFileSingleton excelSingleton = ExcelFileSingleton.GetInstance();

            try
            {
                excelSingleton.OpenExcel(outputExcelFileName);
                
                //新規の列(3列)を指定列(D:D)に挿入する
                excelSingleton.InsertColOfSheet(sheetName, "D", 3);

                // EXCELの固定セルの内容を設定する。
                SetHeadCellData2Excel(excelSingleton, sheetName);

                // EXCELの明細内容を設定する。
                SetDetailData2Excel(excelSingleton, sheetName);
                
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
            //File.Delete(outputExcelFileName);
            

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
		/// EXCELの固定セルの内容を設定する。
		/// </summary>
		/// <param name="excelSingleton"></param>
        /// <param name="sheetName"></param>
		/// <returns></returns>
		/// <remarks></remarks>
        private void SetHeadCellData2Excel(ExcelFileSingleton excelSingleton, string sheetName)
        {
            //今後、DBからデータを取得して、ExcelViewのobjectに格納する
            Dictionary<int, string> dictDay = ExcelComm.GetDayAndWeekName(DateTime.Now);
            InspectionRecordEM insRecordEV = new InspectionRecordEM();
            insRecordEV.Quantity = "数量";
            insRecordEV.Weight = "重量";
            insRecordEV.UnitPrice = "単価";
            //今後、DBからデータを取得して、ExcelViewのobjectに格納する

            //セルの値を設定する
            List<ExcelRowObject> rows = new List<ExcelRowObject>();
            ExcelRowObject headRow = new ExcelRowObject();

            var eVtype = typeof(InspectionRecordEM);
            foreach(PropertyInfo pf in eVtype.GetProperties())
            {
                string cellValue = (string)pf.GetValue(insRecordEV);
                var attribute = (ExcelCellPositionAttribute)pf.GetCustomAttributes(typeof(ExcelCellPositionAttribute), false).FirstOrDefault();

                ExcelCellObject cell = new ExcelCellObject();
                cell.RowIndex = attribute.Row;
                cell.ColIndex = attribute.Col;
                cell.Value = cellValue;

                headRow.Cells.Add(cell);
            }

            rows.Add(headRow);
            excelSingleton.WriteRowsToSheet(sheetName, rows);
        }

        /// <summary>
		/// EXCELの固定セルの内容を設定する。
		/// </summary>
		/// <param name="excelSingleton"></param>
        /// <param name="sheetName"></param>
		/// <returns></returns>
		/// <remarks></remarks>
        private void SetDetailData2Excel(ExcelFileSingleton excelSingleton, string sheetName)
        {
            //今後、DBからデータを取得して、ExcelViewのobjectに格納する
            int startRowIndex = 6;
            List<InspectionRecordDetailEM> details = new List<InspectionRecordDetailEM>();
            for(int i=0; i<30; i++)
            {
                InspectionRecordDetailEM detail = new InspectionRecordDetailEM();
                detail.GoodName = "品名_" + (i + 1).ToString();
                detail.ReceptionTime = DateTime.Now.ToString("h:mm:ss tt");
                detail.PackingRemark = "外箱に異常はない_" + (i + 1).ToString();
                detail.RowIndex = startRowIndex + i;

                details.Add(detail);
            }
            //今後、DBからデータを取得して、ExcelViewのobjectに格納する

            //新規の行を指定行(7)に挿入する
            excelSingleton.InsertRowOfSheet(sheetName, 7, 30);

            //セルの値を設定する
            List<ExcelRowObject> rows = new List<ExcelRowObject>();
            

            var eVtype = typeof(InspectionRecordDetailEM);

            foreach(InspectionRecordDetailEM detail in details)
            {
                ExcelRowObject row = new ExcelRowObject();
                foreach (PropertyInfo pf in eVtype.GetProperties())
                {   
                    var attribute = (ExcelColPositionAttribute)pf.GetCustomAttributes(typeof(ExcelColPositionAttribute), false).FirstOrDefault();
                    if (attribute is null)
                    {
                        continue;
                    }
                    string cellValue = (string)pf.GetValue(detail);

                    ExcelCellObject cell = new ExcelCellObject();
                    cell.RowIndex = detail.RowIndex;
                    cell.ColIndex = attribute.Col;
                    cell.Value = cellValue;

                    row.Cells.Add(cell);
                }
                rows.Add(row);
            }

            excelSingleton.WriteRowsToSheet(sheetName, rows);

        }


        
    }
}
