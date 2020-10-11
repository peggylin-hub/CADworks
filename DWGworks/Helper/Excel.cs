using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using System.IO;

using OfficeOpenXml;

namespace TunnalCal.Helper
{
    class Excel
    {
        //export data to excel
        public static void createCSV(List<string> data, string filepath)
        {
            StringBuilder output = new StringBuilder(1000);
            for (int r = 0; r < data.Count; r++)
            {
                output.Append(data[r]);
                output.Append(Environment.NewLine);
            }

            string text = output.ToString();
            DateTime today = DateTime.Now;
            string fileNameStr = $"_PointData_{today.ToString("yy_MM_dd-hh-mm")}.csv";
            string fileName = filepath + fileNameStr;
            System.IO.File.WriteAllText(fileName, text);
        }

        //export data to excel
        public static void createCSV(Dictionary<string, List<string>> data, string filepath)
        {
            StringBuilder output = new StringBuilder();
            foreach(KeyValuePair<string, List<string>> d in data)
            {
                string fn = d.Key;
                List<string> layers = d.Value;
                foreach(string ly in layers)
                {
                    output.AppendLine(fn + "," + ly);
                }
            }

            string text = output.ToString();
            string fileNameStr = $"_{DateTime.Now.ToString("yy_MM_dd-hh-mm")}.csv";
            string fileName = filepath + fileNameStr;
            System.IO.File.WriteAllText(fileName, text);
        }

        /// <summary>
        /// Create Point3d list by importing from excel
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="AssetExcelFile"></param>
        /// <param name="ExcelShtName"></param>
        /// <returns></returns>
        public static List<Point3d> getAllpoint(Document doc, string AssetExcelFile, string ExcelShtName)
        {
            List<Point3d> pts = new List<Point3d>();

            Dictionary<string, Point3d> data = new Dictionary<string, Point3d>();

            FileInfo fi = new FileInfo(AssetExcelFile);

            try
            {
                using (ExcelPackage pk = new ExcelPackage(fi))
                {
                    ExcelWorksheet ws = null; //pk.Workbook.Worksheets[1];

                    foreach (ExcelWorksheet sht in pk.Workbook.Worksheets)
                    {
                        if (sht.Name == ExcelShtName)
                            ws = sht;
                    }

                    int rowCount = ws.Dimension.End.Row;

                    for (int row = 2; row <= rowCount; row++)//read from the second row, instead of first row
                    {
                        double x = 0;
                        double y = 0;
                        double z = 0;

                        if (ws.Cells[row, 1].Value != null)
                            x = Convert.ToDouble(ws.Cells[row, 2].Value.ToString().Trim());
                        if (ws.Cells[row, 2].Value != null)
                            y = Convert.ToDouble(ws.Cells[row, 3].Value.ToString().Trim());
                        if (ws.Cells[row, 3].Value != null)
                            z = Convert.ToDouble(ws.Cells[row, 4].Value.ToString().Trim());

                        string text = $"{x},{y},{z}";
                        if (!data.ContainsKey(text))
                        {
                            Point3d pt = new Point3d(x, y, z);
                            data.Add(text, pt);
                            pts.Add(pt);
                        }
                    }
                    if (pts.Count() > 0)
                        return pts;
                    else
                        return null;
                }
            }
            catch 
            {
                return null;
            }

        }
    
        public static Dictionary<string, Dictionary<string, string>> getDwgLayerSetup(string filepath)
        {
            Dictionary<string, Dictionary<string,string>> data = new Dictionary<string, Dictionary<string, string>>();

            FileInfo fi = new FileInfo(filepath);
            try
            {
                using (ExcelPackage pk = new ExcelPackage(fi))
                {
                    ExcelWorksheet ws = pk.Workbook.Worksheets.First();
                    int rowCount = ws.Dimension.End.Row;

                    Dictionary<string, string> temp = new Dictionary<string, string>();
                    for (int row = 1; row <= rowCount; row++)//read from the second row, instead of first row
                    {
                        #region READING DATA FROM EXCEL
                        string fn = string.Empty;
                        string ly_old = string.Empty;
                        string ly_new = string.Empty;

                        if (ws.Cells[row, 1].Value != null)
                            fn = ws.Cells[row, 1].Value.ToString().Trim();
                        if (ws.Cells[row, 2].Value != null)
                            ly_old = ws.Cells[row, 2].Value.ToString().Trim();
                        if (ws.Cells[row, 3].Value != null)
                            ly_new = ws.Cells[row, 3].Value.ToString().Trim();
                        #endregion

                        if (data.ContainsKey(fn))
                        {
                            temp = data[fn];
                            if (!temp.ContainsKey(ly_old))
                            {
                                temp.Add(ly_old, ly_new);
                                data[fn] = temp;//update data
                            }
                        }
                        else
                        {
                            temp = new Dictionary<string, string>();
                            temp.Add(ly_old, ly_new);
                            data.Add(fn, temp);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                return null;
            }
            
            return data;
        }

        public static Dictionary<string, List<string>> getDwgBlockList(string filepath)
        {
            Dictionary<string, List<string>> data = new Dictionary<string, List<string>>();

            FileInfo fi = new FileInfo(filepath);
            try
            {
                using (ExcelPackage pk = new ExcelPackage(fi))
                {
                    ExcelWorksheet ws = pk.Workbook.Worksheets.First();
                    int rowCount = ws.Dimension.End.Row;

                    List<string> temp = new List<string>();
                    for (int row = 1; row <= rowCount; row++)//read from the second row, instead of first row
                    {
                        #region READING DATA FROM EXCEL
                        string fn = string.Empty;
                        string block = string.Empty;

                        if (ws.Cells[row, 1].Value != null)
                            fn = ws.Cells[row, 1].Value.ToString().Trim();
                        if (ws.Cells[row, 2].Value != null)
                            block = ws.Cells[row, 2].Value.ToString().Trim();
                        #endregion

                        if (data.ContainsKey(fn))
                        {
                            temp = data[fn];
                            if (!temp.Contains(block))
                            {
                                temp.Add(block);
                                data[fn] = temp;//update data
                            }
                        }
                        else
                        {
                            temp = new List<string>();
                            temp.Add(block);
                            data.Add(fn, temp);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                return null;
            }

            return data;
        }
    }
}
