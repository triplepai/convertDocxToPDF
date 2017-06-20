using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Microsoft.Office;
using Microsoft.Office.Interop.Word;

using System.IO;

using iTextSharp.text.pdf;
using iTextSharp.text;

namespace ConvertDocToHTML.Controllers
{

    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            // Convert Input.docx into Output.doc
        //    Convert(Server.MapPath(Url.Content("~/Content/word2.docx")), Server.MapPath(Url.Content("~/Content/word3.doc")), WdSaveFormat.wdFormatDocument);
                                                              
                                                             
            Convert(Server.MapPath(Url.Content("~/Content/word2.docx")), Server.MapPath(Url.Content("~/Content/word3.pdf")), WdSaveFormat.wdFormatPDF);
                                                              
            // Convert Input.docx into Output.html            
         //   Convert(Server.MapPath(Url.Content("~/Content/word2.docx")), Server.MapPath(Url.Content("~/Content/word3.html")), WdSaveFormat.wdFormatHTML);

            

            //create pdfreader object to read sorce pdf
            PdfReader pdfReader = new PdfReader(System.IO.File.ReadAllBytes(Server.MapPath(Url.Content("~/Content/word3.pdf"))));
            //create stream of filestream or memorystream etc. to create output file
            FileStream stream = new FileStream(Server.MapPath(Url.Content("~/Content/word32.pdf")), FileMode.Create);
            //create pdfstamper object which is used to add addtional content to source pdf file
            PdfStamper pdfStamper = new PdfStamper(pdfReader, stream);
            //iterate through all pages in source pdf
            for (int pageIndex = 1; pageIndex <= pdfReader.NumberOfPages; pageIndex++)
            {
                //Rectangle class in iText represent geomatric representation... in this case, rectanle object would contain page geomatry
                iTextSharp.text.Rectangle pageRectangle = pdfReader.GetPageSizeWithRotation(pageIndex);
                //pdfcontentbyte object contains graphics and text content of page returned by pdfstamper
                PdfContentByte pdfData = pdfStamper.GetOverContent(pageIndex);
                //create fontsize for watermark
                pdfData.SetFontAndSize(BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 40);
                //create new graphics state and assign opacity
                PdfGState graphicsState = new PdfGState();
                graphicsState.FillOpacity = 0.4F;
                //set graphics state to pdfcontentbyte
                pdfData.SetGState(graphicsState);
                //set color of watermark
                pdfData.SetColorFill(BaseColor.RED);
                //indicates start of writing of text
                pdfData.BeginText();
                //show text as per position and rotation
                pdfData.ShowTextAligned(Element.ALIGN_CENTER, "Watermark", pageRectangle.Width / 2, pageRectangle.Height / 2, 45);
                //call endText to invalid font set
                pdfData.EndText();
            }
            
            //close stamper and output filestream
            pdfStamper.Close();
            stream.Close();

            // now delete the original file and rename the temp file to the original file
            //  System.IO.File.Delete(FileLocation);
            //  File.Move(FileLocation.Replace(".pdf", "[temp][file].pdf"), FileLocation);





            // ...and start a viewer.
            //  Process.Start(filename);
            return View();
        }

        public ActionResult About()
        {
     

            return View();
        }
        // Convert a Word 2008 .docx to Word 2003 .doc
        public static void Convert(string input, string output, WdSaveFormat format)
        {
            // Create an instance of Word.exe
          //  Word._Application oWord = new Word.Application();
            ApplicationClass oWord = new ApplicationClass();
            // Make this instance of word invisible (Can still see it in the taskmgr).
            oWord.Visible = false;

            // Interop requires objects.
            object oMissing = System.Reflection.Missing.Value;
            object isVisible = true;
            object readOnly = true;
            object oInput = input;
            object oOutput = output;
            object oFormat = format;

            // Load a document into our instance of word.exe
            _Document oDoc = oWord.Documents.Open(ref oInput, ref oMissing, ref readOnly, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref isVisible, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            // Make this document the active document.
           // oDoc.Activate();

            // Save this document in Word 2003 format.
            oDoc.SaveAs(ref oOutput, ref oFormat, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            // Always close Word.exe.
            oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
           
        }
        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}