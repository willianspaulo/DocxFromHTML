using System.IO;
using System.Web.Mvc;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;

namespace UsandoOPENXML.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public JsonResult GerarWord()
        {
            string Filepath = "C:\\Users\\willi_1jfalph\\Documents\\test.docx";

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(Filepath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text("Create text in body - CreateWordprocessingDocument <h2>Willians<h2> <table><tr><td>TESTE</td></tr></table>"));
            }

            return Json(true, JsonRequestBehavior.AllowGet);
        }

        public JsonResult CriarDocumento()
        {
            string Filepath = "C:\\Users\\willi_1jfalph\\Documents\\testDocHtml.docx";

            CriaDocumentoApartirHTML(Filepath, "<h2>Willians</h2>");

            return Json(true, JsonRequestBehavior.AllowGet);
        }

        public static void CriaDocumentoApartirHTML(string filename, string html)
        {
            if (System.IO.File.Exists(filename)) System.IO.File.Delete(filename);

            using (MemoryStream generatedDocument = new MemoryStream())
            {
                using (WordprocessingDocument package = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = package.MainDocumentPart;
                    if (mainPart == null)
                    {
                        mainPart = package.AddMainDocumentPart();
                        new Document(new Body()).Save(mainPart);
                    }

                    HtmlConverter converter = new HtmlConverter(mainPart);
                    converter.ParseHtml(html);

                    mainPart.Document.Save();
                }

                System.IO.File.WriteAllBytes(filename, generatedDocument.ToArray());
            }

            //System.Diagnostics.Process.Start(filename);
        }
    }
}