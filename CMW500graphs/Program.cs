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
            try
            {
                string temp = args[0];
            }
            catch
            {
                Console.WriteLine("This program was not meant to be run in standalone mode.");
                System.Threading.Thread.Sleep(3000);
                System.Environment.Exit(0);
            }

            string uut = args[0];
            string csvMetCalFileName = args[1];
            int maxFreq = Convert.ToInt32(args[2]);
            bool isFirstTest = Convert.ToBoolean(args[3]);

            string pattern = "M:/UserPrograms/CMW500/";
            string replacement = "";
            Regex rgx = new Regex(pattern);
            string csvFileName = rgx.Replace(csvMetCalFileName, replacement);

            DirectoryInfo csvTempFolder = new DirectoryInfo(@"M:\UserPrograms\CMW500\");
            FileInfo csvFullFileName = new FileInfo(csvTempFolder + csvFileName);

            string userDesktop = Environment.GetEnvironmentVariable("USERPROFILE") + @"\Desktop\";
            string bookName = userDesktop + uut + ".xlsx";

            FileInfo book = new FileInfo(bookName);
            if (IsFileinUse(book))
            {
                Console.WriteLine("Close the Workbook!");
                while (IsFileinUse(book))
                {
                    System.Threading.Thread.Sleep(500);
                }
            }
            if (book.Exists && isFirstTest)
            {
                book.Delete();
                book = new FileInfo(bookName);
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

            ExcelChart chart = (ExcelLineChart)sheet.Drawings.AddChart("chart1", eChartType.Line);
            for (int col = 1; col <= 6; col++)
            {
                chart.Series.Add(csvText.Offset(1, col, maxFreq + 2, 1), csvText.Offset(1, 0, maxFreq + 2, 1));
            }

            chart.Title.Text = sheet.Cells[maxFreq + 4, 1].Value.ToString();
            chart.SetPosition(42, 350);
            chart.SetSize(800, 300);
            chart.DisplayBlanksAs = eDisplayBlanksAs.Gap;
            chart.Legend.Remove();
            RemoveGridlines(chart);

            double yMax = Math.Ceiling((double)sheet.Cells["G2"].Value / 0.075) / 10;
            //double yMax = (double)sheet.Cells["G2"].Value / 0.075;

            //Console.WriteLine(yMax);
            //Console.ReadLine();

            chart.XAxis.CrossesAt = -yMax;
            chart.XAxis.MajorTickMark = eAxisTickMark.In;
            chart.XAxis.MinorTickMark = eAxisTickMark.None;
            chart.XAxis.MinValue = 0;
            chart.XAxis.MaxValue = maxFreq + 2;
            chart.XAxis.MajorUnit = 2;
            chart.XAxis.Title.Text = "Frequency (MHz)";
            chart.XAxis.Title.Font.Size = 12;

         //   chart.YAxis.MinValue = -yMax;// - 0.01;
         //   chart.YAxis.MaxValue = yMax + 0.05;
         //   chart.YAxis.MajorUnit = Math.Ceiling(yMax / 5) * 5;
            chart.YAxis.MinorTickMark = eAxisTickMark.None;
            chart.YAxis.Format = "0.0";
            chart.YAxis.Title.Text = "Error (dB)";
            chart.YAxis.Title.Font.Size = 12;
            chart.YAxis.CrossBetween = eCrossBetween.MidCat;

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
            #if !DEBUG
                csvFullFileName.Delete();
            #endif
            }

        static void RemoveGridlines(ExcelChart chart)
        {
            var chartXml = chart.ChartXml;
            var nsuri = chartXml.DocumentElement.NamespaceURI;
            var nsm = new XmlNamespaceManager(chartXml.NameTable);
            nsm.AddNamespace("c", nsuri);

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

        protected static bool IsFileinUse(FileInfo file)
        {
            FileStream stream = null;

            if (!file.Exists)
                return false;
            try
            {
                stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }
            return false;
        }
    }
}
