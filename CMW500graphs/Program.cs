using System;
using System.IO;
using System.Xml;
using System.Text.RegularExpressions;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace CMW500graphs
{
    class Program
    {
        static void Main(string[] args)
        {
            string uut = args[0];
            string csvMetCalFileName = args[1];
            int maxFreq = Convert.ToInt32(args[2]);
            bool isFirstTest = Convert.ToBoolean(args[3]);

            string pattern = "M:/UserPrograms/";
            string replacement = "";
            Regex rgx = new Regex(pattern);
            string csvFileName = rgx.Replace(csvMetCalFileName, replacement);

            DirectoryInfo csvTempFolder = new DirectoryInfo(@"M:\UserPrograms\");
            FileInfo csvFullFileName = new FileInfo(csvTempFolder + csvFileName);

            string userDesktop = Environment.GetEnvironmentVariable("USERPROFILE") + @"\Desktop\";
			FileInfo book = new FileInfo(userDesktop + uut + ".xlsx");
            if (book.Exists && isFirstTest)
            {
                book.Delete();
                book = new FileInfo(userDesktop + uut + ".xlsx");
            }

            ExcelPackage package = new ExcelPackage(book);
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add(csvFileName);

            var csvText = sheet.Cells.LoadFromText(csvFullFileName);
            sheet.Cells["B1:E1"].Clear(); 
            sheet.Cells["B2"].Clear();
            sheet.Cells[maxFreq + 3, 2].Clear();
            sheet.Cells[maxFreq + 4, 2, maxFreq + 4, 5].Clear();
            sheet.Cells["A1"].Style.Font.Size = 22;
            sheet.Row(1).Merged = true;

            var chart = sheet.Drawings.AddChart("chart1", eChartType.Line);
            for (int col = 1; col <= 6; col++)
            {
                chart.Series.Add(csvText.Offset(1, col, maxFreq + 2, 1), csvText.Offset(1, 0, maxFreq + 2, 1));
            }

            chart.Title.Text = sheet.Cells[maxFreq + 4, 1].Value.ToString();
            chart.SetPosition(42, 350);
            chart.SetSize(800, 400);
            chart.DisplayBlanksAs = eDisplayBlanksAs.Gap;
            chart.Legend.Remove();
            RemoveGridlines(chart);

            chart.XAxis.CrossesAt = -1.4;
            chart.XAxis.MajorTickMark = eAxisTickMark.In;
            chart.XAxis.MinorTickMark = eAxisTickMark.None;
            chart.XAxis.MinValue = 0;
            chart.XAxis.MaxValue = maxFreq + 2;
            chart.XAxis.MajorUnit = 2;
            chart.XAxis.Title.Text = "Frequency (MHz)";
            chart.XAxis.Title.Font.Size = 12;

            chart.YAxis.MinValue = -1.4;
            chart.YAxis.MaxValue = 1.4;
            chart.YAxis.MajorUnit = 0.2;
            chart.YAxis.MinorTickMark = eAxisTickMark.None;
            chart.YAxis.Format = "0.0";
            chart.YAxis.Title.Text = "Error (dB)";
            chart.YAxis.Title.Font.Size = 11;

            var data = (ExcelLineChartSerie)chart.Series[0];
            data.Smooth = true;
            data.LineColor = Color.CornflowerBlue;

            var loLimit24 = (ExcelLineChartSerie)chart.Series[1];
            loLimit24.LineWidth = 1;
            loLimit24.LineColor = Color.Red;

            var loLimit12 = (ExcelLineChartSerie)chart.Series[2];
            loLimit12.LineWidth = 1;
            loLimit12.LineColor = Color.PaleVioletRed;

            var zero = (ExcelLineChartSerie)chart.Series[3];
            zero.LineWidth = 1;
            zero.LineColor = Color.Silver;

            var hiLimit12 = (ExcelLineChartSerie)chart.Series[4];
            hiLimit12.LineWidth = 1;
            hiLimit12.LineColor = Color.PaleVioletRed;

            var hiLimit24 = (ExcelLineChartSerie)chart.Series[5];
            hiLimit24.LineWidth = 1;
            hiLimit24.LineColor = Color.Red;

            package.Save();
            csvFullFileName.Delete();
        }

        static void RemoveGridlines(ExcelChart chart)
        {
            var chartXml = chart.ChartXml;
            var nsuri = chartXml.DocumentElement.NamespaceURI;
            var nsm = new XmlNamespaceManager(chartXml.NameTable);
            nsm.AddNamespace("c", nsuri);

            var catAxisNodes = chartXml.SelectNodes("c:chartSpace/c:chart/c:plotArea/c:catAx", nsm);
            if (catAxisNodes != null && catAxisNodes.Count > 0)
            {
                foreach (XmlNode catAxisNode in catAxisNodes)
                {
                    var major = catAxisNode.SelectSingleNode("c:majorGridlines", nsm);
                    if (major != null)
                        catAxisNode.RemoveChild(major);

                    var minor = catAxisNode.SelectSingleNode("c:minorGridlines", nsm);
                    if (minor != null)
                        catAxisNode.RemoveChild(minor);
                }
            }
            var valAxisNodes = chartXml.SelectNodes("c:chartSpace/c:chart/c:plotArea/c:valAx", nsm);
            if (valAxisNodes != null && valAxisNodes.Count > 0)
            {
                foreach (XmlNode valAxisNode in valAxisNodes)
                {
                    var major = valAxisNode.SelectSingleNode("c:majorGridlines", nsm);
                    if (major != null)
                        valAxisNode.RemoveChild(major);

                    var minor = valAxisNode.SelectSingleNode("c:minorGridlines", nsm);
                    if (minor != null)
                        valAxisNode.RemoveChild(minor);
                }
            }
        }
    }
}
