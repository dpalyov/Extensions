using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Extensions
{
    public static class Extensions
    {
        public static StringBuilder ParseHtml(this DataTable dt)
        {
            var html = new StringBuilder();
            //html.AppendLine($"<span style='font-weight:bold'>{appConfig.GetSection("StaticText").Value}</span>");
            html.AppendLine("<br/>");
            html.AppendLine("<table style='border:1px solid black'>");
            html.AppendLine("<thead style='background-color: black;color:white;'>");
            html.AppendLine("<tr>");

            foreach (DataColumn col in dt.Columns)
            {
                html.AppendLine($"<th style='border:1px solid black'>{col.ColumnName}</th>");
            }

            html.AppendLine("</tr>");
            html.AppendLine("<tbody >");

            foreach (DataRow row in dt.Rows)
            {
                html.AppendLine("<tr>");
                foreach (var col in row.ItemArray)
                {
                    html.AppendLine($"<td style='border:1px solid black; color: black'>{col}</td>");
                }
                html.AppendLine("</tr>");
            }
            html.AppendLine("</tbody>");
            html.AppendLine("</table>");

            return html;
        }

        public static DataTable ToDataTable<T>(this IEnumerable<T> dataset)
        {

            var table = new DataTable();

            var propsInfo = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            foreach (var prop in propsInfo)
            {
                table.Columns.Add(prop.Name);
            }

            foreach (var item in dataset)
            {
                var values = new object[propsInfo.Length];

                for (var i = 0; i < propsInfo.Length; i++)
                {
                    values[i] = propsInfo[i].GetValue(item, null);
                }
                table.Rows.Add(values);
            }
            return table;
        }

        public static string ExportXL<T>(this IEnumerable<T> dataset, string fileName, Dictionary<string, string> headers = null)
        {

            if (dataset.Count() <= 0)
            {
                throw new ArgumentException("Export contains no data");
            }

            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Tailoring");

                var dt = dataset.ToDataTable();

                var tbOffsetY = 1;
                var tbOffsetX = headers != null ? headers.Keys.Count + 3 : 1;

                var headerStyle = package.Workbook.Styles.CreateNamedStyle("Header");
                headerStyle.Style.Fill.PatternType = ExcelFillStyle.Solid;
                headerStyle.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(144, 189, 194));
                headerStyle.Style.Font.Bold = true;
                headerStyle.Style.Font.Color.SetColor(System.Drawing.Color.White);


                if (headers != null)
                {

                    for (var i = 0; i < headers.Keys.Count; i++)
                    {
                        ws.Cells[2 + i, tbOffsetY].Value = $"{headers.Keys.ElementAt(i)} - {headers.Values.ElementAt(i)}";
                        ws.Cells[2 + i, tbOffsetY].StyleName = "Header";

                    }
                }


                ws.Cells[tbOffsetX, tbOffsetY].LoadFromDataTable(dt, true);
                var xlTb = ws.Tables.Add(new ExcelAddressBase(tbOffsetX, tbOffsetY, tbOffsetX + dt.Rows.Count, tbOffsetY + dt.Columns.Count - 1), "Tailoring");
                xlTb.ShowHeader = true;
                xlTb.ShowFilter = true;
                xlTb.ShowRowStripes = true;
                xlTb.TableStyle = OfficeOpenXml.Table.TableStyles.Dark9;
                xlTb.HeaderRowCellStyle = "Header";


                for (int col = 1; col <= dt.Columns.Count; col++)
                {
                    ws.Column(col).AutoFit();
                }

                package.SaveAs(new FileInfo(fileName));
            }


            return fileName;
        }

        public static string ExportXL(this DataTable dt, string fileName, ExcelOptions xlOpts, Dictionary<string, string> headers = null)
        {

            if (dt.Rows.Count <= 0)
            {
                throw new ArgumentException("Export contains no data");
            }

            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add(xlOpts.WorksheetName);

                var tbOffsetY = xlOpts.TableOffsetY;
                var tbOffsetX = headers != null ? headers.Keys.Count + 3 : xlOpts.TableOffsetX;

                var headerStyle = package.Workbook.Styles.CreateNamedStyle("Header");
                headerStyle.Style.Fill.PatternType = ExcelFillStyle.Solid;
                headerStyle.Style.Fill.BackgroundColor.SetColor(xlOpts.HeaderBgColor);//System.Drawing.Color.FromArgb(144, 189, 194));
                headerStyle.Style.Font.Bold = xlOpts.HeaderBold;
                headerStyle.Style.Font.Color.SetColor(xlOpts.HeaderFontColor);//System.Drawing.Color.White);


                if (headers != null)
                {

                    for (var i = 0; i < headers.Keys.Count; i++)
                    {
                        ws.Cells[tbOffsetX + i, tbOffsetY].Value = $"{headers.Keys.ElementAt(i)} - {headers.Values.ElementAt(i)}";
                        ws.Cells[tbOffsetX + i, tbOffsetY].StyleName = "Header";

                    }
                }


                ws.Cells[tbOffsetX, tbOffsetY].LoadFromDataTable(dt, true);
                var xlTb = ws.Tables.Add(new ExcelAddressBase(tbOffsetX, tbOffsetY, tbOffsetX + dt.Rows.Count, tbOffsetY + dt.Columns.Count - 1), xlOpts.ExcelTableName);
                xlTb.ShowHeader = xlOpts.ShowHeader;
                xlTb.ShowFilter = xlOpts.ShowFilter;
                xlTb.ShowRowStripes = xlOpts.ShowStripes;
                xlTb.TableStyle = xlOpts.TableStyles;
                xlTb.HeaderRowCellStyle = "Header";


                for (int col = 1; col <= dt.Columns.Count; col++)
                {
                    ws.Column(col).AutoFit();
                }

                package.SaveAs(new FileInfo(fileName));
            }


            return fileName;
        }
    }
}
