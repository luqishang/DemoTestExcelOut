using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCommon
{
    public class ExcelComm
    {
        /// <summary>
		/// EXCELのセルの内容を設定、取得します。
		/// </summary>
		/// <param name="year">年</param>
        /// <param name="month">月</param>
		/// <returns>Dictionary<key日, value曜日></returns>
		/// <remarks></remarks>
        public static Dictionary<int, string> GetDayAndWeekName(DateTime date)
        {
            //日付、曜日を格納するための辞書
            Dictionary<int, string> day = new Dictionary<int, string>();

            //DateTime型をint型に変換
            string yearString = date.ToString("yyyy");
            string monthString = date.ToString("MM");
            int year = int.Parse(yearString);
            int month = int.Parse(monthString);

            //月末が何日までか取得
            var lastday = DateTime.DaysInMonth(year, month);

            //日付、曜日を格納
            for (int i = 1; i <= lastday; i++)
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
