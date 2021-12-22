using ClosedXML.Excel;
using FetchRetreive.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace FetchRetreive.Controllers
{
    public class FetchRecordsController : Controller
    {
        // GET: Default
        public ActionResult DashBoard()
        {
            Uploadfile uploadfile = new Uploadfile();
            return View(uploadfile);
        }
        [HttpPost]
        public ActionResult DashBoard(Uploadfile uploadfile)
        {
            string empname= string.Empty;
            string WorkLocation = string.Empty;
            string Clientname = string.Empty;
            string StartPeriod = string.Empty;
            string EndPeriod = string.Empty;
            if (ModelState.IsValid)
            {
                if (uploadfile.Excelfile.ContentLength > 0)
                {
                    if (uploadfile.Excelfile.FileName.EndsWith(".xlsx") || uploadfile.Excelfile.FileName.EndsWith(".xls"))
                    {
                        XLWorkbook workbook;
                        try
                        {
                            workbook = new XLWorkbook(uploadfile.Excelfile.InputStream);
                        }
                        catch (Exception ex)
                        {
                            ModelState.AddModelError("Excelfile", "Check your file " + ex.Message);
                            return View();
                        }
                        IXLWorksheet worksheet;
                        try
                        {
                            worksheet = workbook.Worksheet(1);
                        }
                        catch (Exception ex)
                        {
                            ModelState.AddModelError("Excelfile", "Sheet not found!" + ex.Message);
                            return View();
                        }

                        //var tablerecord = worksheet.Range("D14:S14").ToString().ToList();
                        int employeeworkingdays = 0;
                        int companyworkingdays = 0;
                        int leaves = 0;
                        //worksheet.Range["A2", "B2"].Value1 = "Range1";
                        for (var row1 = 14; row1 <= 17;)
                        {
                            for (var col = 4; col <= 19; col++)
                            {
                                string data = worksheet.Row(row1).Cell(col).Value.ToString();
                                companyworkingdays += (data == "WO" ? 0 : (data == "H" ? 0 : (data == "" ? 0 : 1)));
                                employeeworkingdays += (data == "WO" ? 0 : (data == "H" ? 0 : (data == "L" ? 0 : (data == "" ? 0 : 1))));
                                leaves += (data == "L" ? 1 : 0);
                            }
                            row1 = row1 + 3;
                        }

                        
                        foreach (var row in worksheet.RowsUsed())
                        {
                            //string str = row.Cell(2).Value.ToString();
                            //int empcode=(int)worksheet.Row(2).Cell(3).GetValue();
                            empname = worksheet.Row(9).Cell(4).Value.ToString();
                            WorkLocation = worksheet.Row(9).Cell(14).Value.ToString();
                            Clientname = worksheet.Row(8).Cell(14).Value.ToString();
                            StartPeriod = worksheet.Row(10).Cell(4).Value.ToString();
                            EndPeriod = worksheet.Row(10).Cell(14).Value.ToString();
                            break;
                            //string cellValue = worksheet.GetCellData("Resource Name", 2);
                            //string name = row.Cells("Resource Name",2).value.ToString();
                        }
                        int serial_number = 0;
                        ExporttoExcel(ref serial_number, empname, WorkLocation, Clientname, leaves, employeeworkingdays, companyworkingdays);
                    }
                    else
                    {
                        ModelState.AddModelError("Excelfile", "only .xlsx and .xls files are allowed");
                        return View();
                    }
                }
                else {
                    ModelState.AddModelError("Excelfile", "Not a valid file");
                    return View();
                }
            }
            return View();
           
        }
        private void ExporttoExcel(ref int serial_number, string name, string location, string cname, int leaves, int employeeworkingdays, int companyworkingdays, int empcode = 0, string Position = "", int bench = 0, int NonBilliable = 0, int Billiable = 0, int AddWorkingDays = 0, string WorkinLieu = "", string project = "", int BDM = 0, string manager = "", string projectOwner = "", string Geo = "")
        {
            try
            {
                DateTime now = DateTime.Now;
                string filename = "_dailyreport"+ now;
                serial_number = 1;

                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        var ws = wb.Worksheets.Add("Records");
                        
                        int cellrow = 1;
                        ws.Cell(cellrow, 1).InsertData(new List<string[]> { new string[] { "Sr.No" } });
                        ws.Cell(cellrow, 1).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(cellrow, 1).Style.Font.FontColor = XLColor.Red;

                        ws.Cell(cellrow, 2).InsertData(new List<string[]> { new string[] { "Resource Name" } } );
                        ws.Cell(cellrow, 2).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(cellrow, 2).Style.Font.FontColor = XLColor.Red;

                        ws.Cell(cellrow, 3).InsertData(new List<string[]> { new string[] { "Emp Code" } });
                        ws.Cell(cellrow, 3).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(cellrow, 3).Style.Font.FontColor = XLColor.Red;

                        ws.Cell(cellrow, 4).InsertData(new List<string[]> { new string[] { "Location" } });
                        ws.Cell(cellrow, 4).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(cellrow, 4).Style.Font.FontColor = XLColor.Red;

                        ws.Cell(cellrow, 5).InsertData(new List<string[]> { new string[] { "Position" } });
                        ws.Cell(cellrow, 5).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(cellrow, 5).Style.Font.FontColor = XLColor.Red;

                        ws.Cell(cellrow, 6).InsertData(new List<string[]> { new string[] { "Days Available" } });
                        ws.Cell(cellrow, 6).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(cellrow, 6).Style.Font.FontColor = XLColor.Red;

                        ws.Cell(cellrow, 7).InsertData(new List<string[]> { new string[] { "Leaves" } });
                        ws.Cell(cellrow, 7).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(cellrow, 7).Style.Font.FontColor = XLColor.Red;

                        ws.Cell(cellrow, 8).InsertData(new List<string[]> { new string[] { "Bench" } });
                        ws.Cell(cellrow, 8).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(cellrow, 8).Style.Font.FontColor = XLColor.Red;

                        ws.Cell(cellrow, 9).InsertData(new List<string[]> { new string[] { "Non-Billable" } });
                        ws.Cell(cellrow, 9).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(cellrow, 9).Style.Font.FontColor = XLColor.Red;

                        ws.Cell(cellrow, 10).InsertData(new List<string[]> { new string[] { "Billable" } });
                        ws.Cell(cellrow, 10).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(cellrow, 10).Style.Font.FontColor = XLColor.Red;

                        ws.Cell(cellrow, 11).InsertData(new List<string[]> { new string[] { "Additional Days Worked" } });
                        ws.Cell(cellrow, 11).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(cellrow, 11).Style.Font.FontColor = XLColor.Red;

                        ws.Cell(cellrow, 12).InsertData(new List<string[]> { new string[] { "Work In Lieu" } });
                        ws.Cell(cellrow, 12).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(cellrow, 12).Style.Font.FontColor = XLColor.Red;

                        ws.Cell(cellrow, 13).InsertData(new List<string[]> { new string[] { "Client Name" } });
                        ws.Cell(cellrow, 13).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(cellrow, 13).Style.Font.FontColor = XLColor.Red;

                        ws.Cell(cellrow, 14).InsertData(new List<string[]> { new string[] { "Project" } });
                        ws.Cell(cellrow, 14).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(cellrow, 14).Style.Font.FontColor = XLColor.Red;

                        ws.Cell(cellrow, 15).InsertData(new List<string[]> { new string[] { "BDM" } });
                        ws.Cell(cellrow, 15).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(cellrow, 15).Style.Font.FontColor = XLColor.Red;

                        ws.Cell(cellrow, 16).InsertData(new List<string[]> { new string[] { "Manager" } });
                        ws.Cell(cellrow, 16).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(cellrow, 16).Style.Font.FontColor = XLColor.Red;

                        ws.Cell(cellrow, 17).InsertData(new List<string[]> { new string[] { "Project Owner" } });
                        ws.Cell(cellrow, 17).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(cellrow, 17).Style.Font.FontColor = XLColor.Red;

                        ws.Cell(cellrow, 18).InsertData(new List<string[]> { new string[] { "Geo" } });
                        ws.Cell(cellrow, 18).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(cellrow, 18).Style.Font.FontColor = XLColor.Red;

                    //Records displays in the 
                        int datarow = 2;
                        ws.Cell(datarow, 1).InsertData(new List<string[]> { new string[] { Convert.ToString(serial_number) } });
                        ws.Cell(datarow, 1).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(datarow, 1).Style.Font.FontColor = XLColor.Black;

                        ws.Cell(datarow, 2).InsertData(new List<string[]> { new string[] { name } });
                        ws.Cell(datarow, 2).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(datarow, 2).Style.Font.FontColor = XLColor.Black;

                        ws.Cell(datarow, 3).InsertData(new List<string[]> { new string[] { Convert.ToString(empcode)} });
                        ws.Cell(datarow, 3).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(datarow, 3).Style.Font.FontColor = XLColor.Black;

                        ws.Cell(datarow, 4).InsertData(new List<string[]> { new string[] { location } });
                        ws.Cell(datarow, 4).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(datarow, 4).Style.Font.FontColor = XLColor.Black;

                        ws.Cell(datarow, 5).InsertData(new List<string[]> { new string[] { Position } });
                        ws.Cell(datarow, 5).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(datarow, 5).Style.Font.FontColor = XLColor.Black;

                        ws.Cell(datarow, 6).InsertData(new List<string[]> { new string[] { Convert.ToString(employeeworkingdays) } });
                        ws.Cell(datarow, 6).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(datarow, 6).Style.Font.FontColor = XLColor.Black;

                        ws.Cell(datarow, 7).InsertData(new List<string[]> { new string[] { Convert.ToString(leaves) } });
                        ws.Cell(datarow, 7).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(datarow, 7).Style.Font.FontColor = XLColor.Black;

                        ws.Cell(datarow, 8).InsertData(new List<string[]> { new string[] { Convert.ToString(bench) } });
                        ws.Cell(datarow, 8).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(datarow, 8).Style.Font.FontColor = XLColor.Black;

                        ws.Cell(datarow, 9).InsertData(new List<string[]> { new string[] { Convert.ToString(NonBilliable) } });
                        ws.Cell(datarow, 9).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(datarow, 9).Style.Font.FontColor = XLColor.Black;

                        ws.Cell(datarow, 10).InsertData(new List<string[]> { new string[] { Convert.ToString(Billiable) } });
                        ws.Cell(datarow, 10).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(datarow, 10).Style.Font.FontColor = XLColor.Black;

                        ws.Cell(datarow, 11).InsertData(new List<string[]> { new string[] { Convert.ToString(AddWorkingDays )} });
                        ws.Cell(datarow, 11).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(datarow, 11).Style.Font.FontColor = XLColor.Black;

                        ws.Cell(datarow, 12).InsertData(new List<string[]> { new string[] {Convert.ToString( WorkinLieu )} });
                        ws.Cell(datarow, 12).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(datarow, 12).Style.Font.FontColor = XLColor.Black;

                        ws.Cell(datarow, 13).InsertData(new List<string[]> { new string[] { cname } });
                        ws.Cell(datarow, 13).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(datarow, 13).Style.Font.FontColor = XLColor.Black;

                        ws.Cell(datarow, 14).InsertData(new List<string[]> { new string[] { project } });
                        ws.Cell(datarow, 14).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(datarow, 14).Style.Font.FontColor = XLColor.Black;

                        ws.Cell(datarow, 15).InsertData(new List<string[]> { new string[] { Convert.ToString(BDM )} });
                        ws.Cell(datarow, 15).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(datarow, 15).Style.Font.FontColor = XLColor.Black;

                        ws.Cell(datarow, 16).InsertData(new List<string[]> { new string[] { manager } });
                        ws.Cell(datarow, 16).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(datarow, 16).Style.Font.FontColor = XLColor.Black;

                        ws.Cell(datarow, 17).InsertData(new List<string[]> { new string[] { projectOwner } });
                        ws.Cell(datarow, 17).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(datarow, 17).Style.Font.FontColor = XLColor.Black;

                        ws.Cell(datarow, 18).InsertData(new List<string[]> { new string[] { Geo } });
                        ws.Cell(datarow, 18).Style.Fill.BackgroundColor = XLColor.FromArgb(238, 238, 238);
                        ws.Cell(datarow, 18).Style.Font.FontColor = XLColor.Black;

                       

                     ws.Columns().AdjustToContents();
                        Response.Clear();
                        Response.Buffer = true;
                        Response.Charset = "";
                        Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                        Response.AddHeader("content-disposition", "attachment;filename=\"" + filename + ".xlsx\"");
                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            wb.SaveAs(memoryStream);
                            memoryStream.WriteTo(Response.OutputStream);
                            Response.Flush();
                            Response.End();
                        }
                    }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}