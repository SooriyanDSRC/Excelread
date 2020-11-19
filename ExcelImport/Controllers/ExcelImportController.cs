using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Dynamic;
using System.Linq;

namespace ExcelImport.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelImportController : Controller
    {
        public Collection<object> GetExcelDetails()
        {
            Collection<object> clnPreviewer = new Collection<object>();
            string path = @"SampleSheets\Sample 4 Listing.xlsx";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = System.IO.File.OpenRead(path))
                {
                    pck.Load(stream);
                }

                int cntSheet = pck.Workbook.Worksheets.Count;

                DataSet dsWorkBook = new DataSet();

                DataSet dsExcelSplitter = new DataSet();

                for (int i = 0; i < cntSheet; i++)
                {
                    var sheetName = pck.Workbook.Worksheets[i].Name;

                    dsWorkBook.Tables.Add(sheetName);

                    var ws = pck.Workbook.Worksheets[i];

                    foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                    {
                        dsWorkBook.Tables[sheetName].Columns.Add(string.Format("Column {0}", firstRowCell.Start.Column));
                    }

                    var startRow = 1;

                    for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                    {
                        var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                        DataRow row = dsWorkBook.Tables[sheetName].Rows.Add();
                        foreach (var cell in wsRow)
                        {
                            row[cell.Start.Column - 1] = cell.Text;
                        }
                    }

                    var endRow = ws.Dimension.End.Row;
                    var endColumn = ws.Dimension.End.Column;

                    int celCount = 0;

                    foreach (DataRow dr in dsWorkBook.Tables[sheetName].Rows)
                    {
                        celCount = celCount + 1;
                        if (Convert.ToString(dr[0]).Contains("Joint Number") || Convert.ToString(dr[0]).Contains("Log Dist"))
                        {
                            break;
                        }
                    }

                    var headerStart = celCount == 0 ? 0 : celCount;

                    dsExcelSplitter = SplitExcelTables(dsWorkBook, i, sheetName, headerStart, endRow);
                    clnPreviewer.Add(ConvertDataSetToList(dsExcelSplitter));
                }
            }

            return clnPreviewer;
        }

        public static DataSet SplitExcelTables(DataSet dsWorkBook, int sheetIndex, string sheetName, int headerStart, int endRow)
        {
            DataSet dsExcelSplitter = new DataSet();
            dsExcelSplitter.Tables.Add("cmn_" + sheetIndex + "_" + sheetName);

            foreach (DataColumn dc in dsWorkBook.Tables[sheetName].Columns)
            {
                dsExcelSplitter.Tables["cmn_" + sheetIndex + "_" + sheetName].Columns.Add(dc.ColumnName);
            }

            for (int k = 0; k < headerStart - 1; k++)
            {
                DataRow dr = dsWorkBook.Tables[sheetName].Rows[k];
                dsExcelSplitter.Tables["cmn_" + sheetIndex + "_" + sheetName].ImportRow(dr);
            }

            dsExcelSplitter.Tables.Add("main_" + sheetIndex + "_" + sheetName);

            foreach (DataColumn dc in dsWorkBook.Tables[sheetName].Columns)
            {
                dsExcelSplitter.Tables["main_" + sheetIndex + "_" + sheetName].Columns.Add(dc.ColumnName);
            }

            for (int m = headerStart + 1; m < endRow; m++)
            {
                DataRow drm = dsWorkBook.Tables[sheetName].Rows[m];

                dsExcelSplitter.Tables["main_" + sheetIndex + "_" + sheetName].ImportRow(drm);
            }

            DataRow drr = dsWorkBook.Tables[sheetName].Rows[headerStart - 1];

            foreach (var colName in drr.ItemArray.Select((value, p) => new { p, value }))
            {
                dsExcelSplitter.Tables["main_" + sheetIndex + "_" + sheetName].Columns[colName.p].ColumnName =
                    Convert.ToString(colName.value) != string.Empty ? colName.value.ToString() : "Column" + colName.p;
            }

            return dsExcelSplitter;
        }

        public static Collection<object> ConvertDataSetToList(DataSet dsPreview)
        {
            Collection<object> clnPreview = new Collection<object>();

            foreach (DataTable dt in dsPreview.Tables)
            {
                List<dynamic> dynamicDt = dt.ToDynamic();
                clnPreview.Add(dynamicDt);
            }

            return clnPreview;
        }
    }

    public static class DynamicExtender
    {
        public static List<dynamic> ToDynamic(this DataTable dt)
        {
            var dynamicDt = new List<dynamic>();

            int i = 0;
            foreach (DataRow row in dt.Rows)
            {
                i = i + 1;

                if (i > 20)
                {
                    break;
                }

                dynamic dyn = new ExpandoObject();
                dynamicDt.Add(dyn);
                foreach (DataColumn column in dt.Columns)
                {
                    var dic = (IDictionary<string, object>)dyn;
                    dic[column.ColumnName] = row[column];
                }
            }

            return dynamicDt;
        }
    }
}
