using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.draw;
using System;
using System.IO;
using System.Text;
using System.Web.Mvc;

namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult TextDemo()
        {
            //获取字体
            string fontPath = @"C:\WINDOWS\Fonts\simsun.ttc,0";
            BaseFont baseFont = BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);

            Font titleFont = new Font(baseFont, 13, Font.NORMAL);
            Font contentFont = new Font(baseFont, 13, Font.BOLD);
            Font englishFont = new Font(baseFont, 9, Font.NORMAL);

            Document pdfDocument = new Document(PageSize.A4, 55, 55, 50, 0);

            using (MemoryStream ms = new MemoryStream())
            {
                PdfWriter writer = PdfWriter.GetInstance(pdfDocument, ms);
                //writer.PageEvent = new HeaderFooterHelper();
                pdfDocument.Open();

                LineSeparator line = new LineSeparator
                {
                    Offset = -3
                };

                PdfPTable table = new PdfPTable(4)
                {
                    WidthPercentage = 100
                };
                table.SetWidths(new[] { 2.5f, 4.5f, 3f, 4.5f });

                Paragraph title = new Paragraph("测试一", titleFont);
                //title.Add(GetChunk("测试一", contentFont, 20));
                table.AddCell(GetCell(title));
                Paragraph content = new Paragraph(GetContent("测试一"), contentFont);
                content.Add(line);
                table.AddCell(GetCell(content));

                title = new Paragraph("测试一", titleFont);
                table.AddCell(GetCell(title));
                content = new Paragraph(GetContent("测试一"), contentFont);
                content.Add(line);
                table.AddCell(GetCell(content));

                title = new Paragraph("test", englishFont);
                table.AddCell(GetCell(title, verticalAlignment: Element.ALIGN_TOP, height: 20));
                table.AddCell(GetCell(new Paragraph()));

                title = new Paragraph("test", englishFont);
                table.AddCell(GetCell(title, verticalAlignment: Element.ALIGN_TOP, height: 20));
                table.AddCell(GetCell(new Paragraph()));

                title = new Paragraph("测试一", titleFont);
                table.AddCell(GetCell(title));
                content = new Paragraph(GetContent("测试一"), contentFont);
                content.Add(line);
                table.AddCell(GetCell(content));

                title = new Paragraph("测试一", titleFont);
                table.AddCell(GetCell(title));
                content = new Paragraph(GetContent("测试一"), contentFont);
                content.Add(line);
                table.AddCell(GetCell(content));

                title = new Paragraph("test", englishFont);
                table.AddCell(GetCell(title, verticalAlignment: Element.ALIGN_TOP, height: 20));
                table.AddCell(GetCell(new Paragraph()));

                title = new Paragraph("test", englishFont);
                table.AddCell(GetCell(title, verticalAlignment: Element.ALIGN_TOP, height: 20));
                table.AddCell(GetCell(new Paragraph()));
                
                title = new Paragraph("开学时间", titleFont);
                table.AddCell(GetCell(title));
                content = new Paragraph("2018-12-14 09:00", contentFont);
                content.Add(line);
                table.AddCell(GetCell(content));

                title = new Paragraph("开学时间", titleFont);
                table.AddCell(GetCell(title));
                content = new Paragraph("2018-12-14 09:00", contentFont);
                content.Add(line);
                table.AddCell(GetCell(content));

                title = new Paragraph("Start Time", englishFont);
                table.AddCell(GetCell(title, verticalAlignment: Element.ALIGN_TOP, height: 20));
                table.AddCell(GetCell(new Paragraph()));

                title = new Paragraph("Start Time", englishFont);
                table.AddCell(GetCell(title, verticalAlignment: Element.ALIGN_TOP, height: 20));
                table.AddCell(GetCell(new Paragraph()));

                title = new Paragraph("姓名", titleFont);
                table.AddCell(GetCell(title));
                content = new Paragraph(GetContent("张三  李四", 48), contentFont);
                content.Add(line);
                table.AddCell(GetCell(content, colspan: 3));

                title = new Paragraph("Name", englishFont);
                table.AddCell(GetCell(title, verticalAlignment: Element.ALIGN_TOP, height: 20));
                table.AddCell(GetCell(new Paragraph(), colspan: 3));


                pdfDocument.Add(table);

                writer.Flush();
                writer.CloseStream = false;
                pdfDocument.Close();
                writer.Close();

                return File(ms.GetBuffer(), "application/pdf");
            }
        }

        private PdfPCell GetCell(Paragraph pg, int horizontalAlignment = Element.ALIGN_CENTER, int verticalAlignment = Element.ALIGN_MIDDLE, int height = 0, int colspan = 0)
        {
            PdfPCell cell = new PdfPCell(pg)
            {
                HorizontalAlignment = horizontalAlignment,
                VerticalAlignment = verticalAlignment,
                Border = Rectangle.NO_BORDER
            };

            if (height != 0)
            {
                cell.FixedHeight = height;
            }
            if (colspan > 0)
            {
                cell.Colspan = colspan;
            }

            return cell;
        }

        private Chunk GetChunk(string content, Font contentFont, int strLength = 22)
        {
            int len = strLength - Encoding.Default.GetByteCount(content);
            if (len > 0)
            {
                len = (int)Math.Round(len / 2.0, MidpointRounding.AwayFromZero);
            }

            Chunk chunk = new Chunk(GetInsertEmptyStrContent(content, len), contentFont);

            return chunk;
        }

        private string GetContent(string content, int strLength = 22)
        {
            int len = strLength - Encoding.Default.GetByteCount(content);
            if (len > 0)
            {
                len = (int)Math.Round(len / 2.0, MidpointRounding.AwayFromZero);
            }

            return GetInsertEmptyStrContent(content, len);
        }

        private string GetInsertEmptyStrContent(string content, int len)
        {
            if (len <= 0)
            {
                return content;
            }

            content = content.PadLeft(content.Length + len);

            return content;
        }

    }
}