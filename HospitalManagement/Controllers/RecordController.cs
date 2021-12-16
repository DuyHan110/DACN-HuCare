using BELibrary.Core.Entity;
using BELibrary.DbContext;
using System;
using System.IO;
using System.Linq;
using System.Web.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System.Collections.Generic;
using BELibrary.Entity;
using System.Drawing;
using HospitalManagement.Models;

namespace HospitalManagement.Controllers
{
    public class RecordController : BaseController
    {
        // private IHostingEnvironment Environment;
        // GET: Record
        public ActionResult Index()
        {
            return RedirectToAction("Index", "Account");
        }

        public ActionResult Attachment(Guid detailRecordId)
        {
            using (var workScope = new UnitOfWork(new HospitalManagementDbContext()))
            {
                //Check isRecord of current user
                var listData = workScope.AttachmentAssigns
                    .Include(x => x.Attachment).Where(x => x.DetailRecordId == detailRecordId).ToList();
                return View(listData);
            }
        }

        public ActionResult Prescription(Guid detailRecordId)
        {
            ViewBag.DetailRecordId = detailRecordId;
            using (var workScope = new UnitOfWork(new HospitalManagementDbContext()))
            {
                //Check isRecord of current user
                var listData = workScope.Prescriptions.Include(x => x.DetailPrescription.Medicine)
                    .Where(x => x.DetailRecordId == detailRecordId).ToList();
                return View(listData);
            }


            //string path = Path.Combine(this.Environment.WebRootPath, "HTML");
            //if (!Directory.Exists(path))
            //{
            //    Directory.CreateDirectory(path);
            //}

            //string input = Path.Combine(path, "html1.html");
            //string output = Path.Combine(path, "Grid.docx");
            //System.IO.File.WriteAllText(input, GridHtml);
            //DocumentCore documentCore = DocumentCore.Load(input);
            //documentCore.Save(output);
            //byte[] bytes = System.IO.File.ReadAllBytes(output);

            //Directory.Delete(path, true);

            //return File(bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Grid.docx");

        }




        private Stream CreateExcelFile(List<Prescription> listData, Stream stream = null)
        { 

                using (var excelPackage = new ExcelPackage(stream ?? new MemoryStream()))
                {
                    // Tạo author cho file Excel
                    excelPackage.Workbook.Properties.Author = "Hanker";
                    // Tạo title cho file Excel
                    excelPackage.Workbook.Properties.Title = "EPP test background";
                    // thêm tí comments vào làm màu 
                    excelPackage.Workbook.Properties.Comments = "This is my fucking generated Comments";
                    // Add Sheet vào file Excel
                    excelPackage.Workbook.Worksheets.Add("First Sheet");
                    // Lấy Sheet bạn vừa mới tạo ra để thao tác 
                    var workSheet = excelPackage.Workbook.Worksheets[1];
                // Đỗ data vào Excel file

                var listDetailPrescriptions = listData.Select(x => new DetailPrescriptionExport
                {
                    Amount = x.DetailPrescription.Amount, 
                    MedicineName = x.DetailPrescription.Medicine.Name,
                    Note = x.DetailPrescription.Note,
                    Unit = x.DetailPrescription.Unit,
                }).ToList();

                workSheet.Cells[8, 1].LoadFromCollection(listDetailPrescriptions, true, TableStyles.Dark9);

                    BindingFormatForExcelExist(workSheet, listDetailPrescriptions);


                    excelPackage.Save();
                    return excelPackage.Stream;
                }
             
        }
        private void BindingFormatForExcelExist(ExcelWorksheet worksheet, List<DetailPrescriptionExport> input)
        {

            // Set default width cho tất cả column
            worksheet.DefaultColWidth = 20;
            // Tự động xuống hàng khi text quá dài
            worksheet.Cells.Style.WrapText = true;
            // Tạo header

            int z = 8;

            worksheet.Cells[z, 1].Value = "Số Thứ Tự";
            worksheet.Cells[z, 2].Value = "Tên Thuốc ";
            worksheet.Cells[z, 3].Value = "Số lượng";
            worksheet.Cells[z, 4].Value = "Loại";
            worksheet.Cells[z, 5].Value = "Ghi Chú";
          

            // Lấy range vào tạo format cho range đó ở đây là từ A7 tới D7
            using (var range = worksheet.Cells["A8:E8"])
            {
                // Set PatternType
                range.Style.Fill.PatternType = ExcelFillStyle.DarkGray;

                range.Style.Font.Color.SetColor(Color.DarkViolet);
                // Set Màu cho Background
                range.Style.Fill.BackgroundColor.SetColor(Color.Aqua);
                // Canh giữa cho các text
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                // Set Font cho text  trong Range hiện tại
                range.Style.Font.SetFromFont(new Font("Arial", 10));
                // Set Border
                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                // Set màu ch Border
                range.Style.Border.Bottom.Color.SetColor(Color.Blue);
            }

         

            // Đỗ dữ liệu từ list vào 
            for (int i = 0; i < input.Count; i++)
            {
                
                var item = input[i];
                // worksheet.Cells[i + 1 + z, 1].Value = item.DetailPrescription.Id;
                worksheet.Cells[i + 1 + z, 1].Value = 1;
                worksheet.Cells[i + 1 + z, 2].Value = item.MedicineName;
                worksheet.Cells[i + 1 + z, 3].Value = item.Amount;
                worksheet.Cells[i + 1 + z, 4].Value = item.Unit;
                worksheet.Cells[i + 1 + z, 5].Value = item.Note;
                //Format lại color nế+6u như thỏa điều kiện
                //if (item.Revenue > 10000050)
                //{
                //    Ở đây chúng ta sẽ format lại theo dạng fromRow, fromCol, toRow, toCol
                //    using (var range = worksheet.Cells[i + 2, 1, i + 2, 6])
                //    {
                //        Format text đỏ và đậm
                //        range.Style.Font.Color.SetColor(Color.Red);
                //        range.Style.Font.Bold = true;
                //    }
                //}

            }
            // Format lại định dạng xuất ra ở cột Money 
            // fix lại width của column với minimum width là 15
            // worksheet.Cells[1 + z, 1, listItems.Count + 5 + z, 4].AutoFitColumns(15);

            // Thực hiện tính theo formula trong excel
            // Hàm Sum 
            //worksheet.Cells[listItems.Count + 3 + z, 3].Value = "Tổng SL nhập :";
            //worksheet.Cells[listItems.Count + 3 + z, 4].Formula = "SUM(C" + (z + 1) + ":C" + (listItems.Count + z + 1) + ")";
            //worksheet.Cells[listItems.Count + 4 + z, 3].Value = "Tổng SL bán:";
            //worksheet.Cells[listItems.Count + 4 + z, 4].Formula = "SUM(D" + (z + 1) + ":D" + (listItems.Count + z + 1) + ")";
            //worksheet.Cells[listItems.Count + 5 + z, 3].Value = "Tổng tồn kho:";
            //worksheet.Cells[listItems.Count + 5 + z, 4].Formula = "SUM(E" + (z + 1) + ":E" + (listItems.Count + z + 1) + ")";

            // Tổng doanh thu
            //worksheet.Cells[listItems.Count + 6 + z, 3].Value = "Tổng doanh thu: ";
            //worksheet.Cells[listItems.Count + 6 + z, 3].Style.Font.Color.SetColor(Color.Red);
            //worksheet.Cells[listItems.Count + 6 + z, 4].Style.Numberformat.Format = "#,##0";
            //worksheet.Cells[listItems.Count + 6 + z, 4].Formula = "SUM(F2:F" + (listItems.Count + 1) + ")";
            //worksheet.Cells[listItems.Count + 6 + z, 4].Style.Font.Color.SetColor(Color.Red);


            // Infor 
            worksheet.Cells[2, 2, 2, 4].Merge = true;
            var cellTitleInf = worksheet.Cells[2, 2, 2, 4];
            cellTitleInf.Value = "Đơn thuốc";
            cellTitleInf.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cellTitleInf.Style.Font.Color.SetColor(Color.Red);
            cellTitleInf.Style.Border.BorderAround(ExcelBorderStyle.Double);


            //worksheet.Cells[listItems.Count + 6 + 3, 3].Value = "Thống kê từ: ";
            //worksheet.Cells[listItems.Count + 6 + 3, 3].Style.Font.Color.SetColor(Color.Blue);
            //worksheet.Cells[listItems.Count + 6 + 3, 3].AutoFitColumns();
            //worksheet.Cells[listItems.Count + 6 + 3, 4].Value = input.StartTime+ " - "+ input.EndTime;
            //worksheet.Cells[listItems.Count + 6 + 3, 4].AutoFitColumns();
            //worksheet.Cells[listItems.Count + 6 + 3, 4].Style.Font.Color.SetColor(Color.Red);



            worksheet.Cells[3 + 1, 3].Value = "Người xuất: ";
            worksheet.Cells[3 + 1, 3].Style.Font.Color.SetColor(Color.Blue);
            worksheet.Cells[3 + 1, 4].Value = GetCurrentUser().FullName;
            worksheet.Cells[3 + 1, 4].AutoFitColumns();
            worksheet.Cells[3 + 1, 4].Style.Font.Color.SetColor(Color.Red);


            worksheet.Cells[4 + 1, 3].Value = "Ngày xuất: ";
            worksheet.Cells[4 + 1, 3].Style.Font.Color.SetColor(Color.Blue);
            worksheet.Cells[4 + 1, 4].Value = DateTime.Now.ToString("dd/MM/yyy HH:mm:ss");
            worksheet.Cells[4 + 1, 4].AutoFitColumns();
            worksheet.Cells[4 + 1, 4].Style.Font.Color.SetColor(Color.Red);



            worksheet.Cells[5 + 1, 3, 5 + 1, 4].Merge = true;
            var cellTimeInf = worksheet.Cells[5 + 1, 3, 5 + 1, 4];
            // cellTimeInf.Value = input.StartTime + " - " + input.EndTime;
            cellTimeInf.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cellTimeInf.Style.Font.Color.SetColor(Color.Red);
            cellTimeInf.Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }



        [HttpGet]
        public ActionResult Export(Guid detailRecordId)
        {
            using (var workScope = new UnitOfWork(new HospitalManagementDbContext()))
            {

                //Check isRecord of current user
                var listData = workScope.Prescriptions.Include(x => x.DetailPrescription.Medicine)
                .Where(x => x.DetailRecordId == detailRecordId).ToList();
                // Gọi lại hàm để tạo file excel
                var stream = CreateExcelFile(listData);
                // Tạo buffer memory strean để hứng file excel
                var buffer = stream as MemoryStream;
                // Đây là content Type dành cho file excel, còn rất nhiều content-type khác nhưng cái này mình thấy okay nhất
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                // Dòng này rất quan trọng, vì chạy trên firefox hay IE thì dòng này sẽ hiện Save As dialog cho người dùng chọn thư mục để lưu
                // File name của Excel này là ExcelDemo
                Response.AddHeader("Content-Disposition", "attachment; filename=ExcelDemo1.xlsx");
                // Lưu file excel của chúng ta như 1 mảng byte để trả về response
                Response.BinaryWrite(buffer.ToArray());
                // Send tất cả ouput bytes về phía clients
                Response.Flush();
                Response.End();
                // Redirect về luôn trang index :D
                return RedirectToAction("Index");
            }
        }
    }
}