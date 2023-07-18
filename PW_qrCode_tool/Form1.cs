using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using QRCoder;
using System.Text.RegularExpressions;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;
using ZXing;
using Bytescout.PDFRenderer;
using iText;
using iText.Kernel.Pdf;
using iText.Kernel.Utils;
using iText.Bouncycastleconnector;
using System.Collections.Generic;
using Stream = System.IO.Stream;
using System.Windows.Controls;
using PageRange = iText.Kernel.Utils.PageRange;
using Microsoft.Office.Interop.Word;
using iText.Kernel.Pdf.Xobject;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace PW_qrCode_tool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            tabControl1.SelectTab(tabPage1);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void OnTab1ChooseFile(object sender, EventArgs e)
        {
            System.Windows.Forms.OpenFileDialog fileDialog = new System.Windows.Forms.OpenFileDialog();
            fileDialog.Filter = "Plik Word (.docx ,.doc)|*.docx;*.doc";
            fileDialog.FileOk += delegate (object s, CancelEventArgs ev)
            {
                string ext = Path.GetExtension(fileDialog.FileName);
                if (ext != ".doc" && ext != ".docx")
                {
                    System.Windows.MessageBox.Show("Wybrany plik nie jest plikem Word!");
                    ev.Cancel = true;
                }
            };
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = fileDialog.FileName;
            }
            fileDialog.Dispose();
        }

        private void OnTab1ProcessFile(object sender, EventArgs e)
        {
            string path = textBox2.Text;
            string ext = Path.GetExtension(path);

            if (ext != ".doc" && ext != ".docx")
            {
                System.Windows.MessageBox.Show("Wybrany plik nie jest plikem Word!");
                return;
            }

            if (!File.Exists(path))
            {
                System.Windows.MessageBox.Show(String.Format("Plik {0} nie istnieje!", path));
            }
            else
            {
                ProcessWordFile(path);
            }
        }

        private void OnTab2ChooseFile(object sender, EventArgs e)
        {
            System.Windows.Forms.OpenFileDialog fileDialog = new System.Windows.Forms.OpenFileDialog();
            fileDialog.Filter = "Plik PDF (.pdf)|*.pdf";
            fileDialog.FileOk += delegate (object s, CancelEventArgs ev)
            {
                string ext = Path.GetExtension(fileDialog.FileName);
                if (ext != ".pdf")
                {
                    System.Windows.MessageBox.Show("Wybrany plik nie jest plikem PDF!");
                    ev.Cancel = true;
                }
            };
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = fileDialog.FileName;
            }
            fileDialog.Dispose();
        }

        private void OnTab2ProcessFile(object sender, EventArgs e)
        {
            string path = textBox1.Text;
            string ext = Path.GetExtension(path);

            if (ext != ".pdf")
            {
                System.Windows.MessageBox.Show("Wybrany plik nie jest plikem PDF!");
                return;
            }

            if (!File.Exists(path))
            {
                System.Windows.MessageBox.Show(String.Format("Plik {0} nie istnieje!", path));
            }
            else
            {
                ProcessPDFFile(path);
            }
        }

        protected void ProcessWordFile(string path)
        {
            if (path.Length == 0)
            {
                return;
            }

            string ext = Path.GetExtension(path);

            progressBar1.Minimum = 0;
            progressBar1.Maximum = 10000;
            progressBar1.Step = 1;
            progressBar1.Style = ProgressBarStyle.Continuous;

            label2.Text = "Przetwarzanie...";


            switch (ext)
            {
                case ".doc":
                    ProcessDoc(path);
                    break;
                case ".docx":
                    ProcessDocx(path);
                    break;
                default:
                    break;
            }
        }

        protected void ProcessDoc(string path)
        {
            Word._Application application = new Word.Application();
            object fileformat = Word.WdSaveFormat.wdFormatXMLDocument;

            object filename = path;
            object tempFileName = Path.GetFileName(path).ToLower().Replace(Path.GetExtension(path), "");
            string uuid = Guid.NewGuid().ToString();
            string newfilename = System.IO.Path.GetTempPath() + tempFileName + uuid + ".docx";
            Word._Document document = application.Documents.Open(filename);

            document.Convert();
            document.SaveAs(newfilename, fileformat);
            document.Close();

            document = null;

            application.Quit();
            application = null;
            ProcessDocx(newfilename);
            File.Delete(newfilename);
        }

        protected void ProcessDocx(string path)
        {

            // Odczyt i zmapowanie danych do wygenerowania QR Kodu ze stopki pliku docx
            int progressBarRange1 = progressBar1.Maximum / 2;
            int progressBarRange2 = progressBar1.Maximum;
            Word._Application application = new Word.Application();
            Word._Document documentWord = application.Documents.Open(path);
            var pages = documentWord.ComputeStatistics(Word.WdStatistic.wdStatisticPages, false);
            Word.Sections documentWordSections = documentWord.Sections;

            PageToCode[] pageToCodes = new PageToCode[documentWordSections.Count+1];
            int sectionNumber = 1;
            int startPage, endPage, currentPage, previousPage = 0;
            while (sectionNumber <= documentWordSections.Count)
            {
                Word.Section section = documentWordSections[sectionNumber];
                Word.HeaderFooter[] footers = { section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary], section.Footers[WdHeaderFooterIndex.wdHeaderFooterEvenPages], section.Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage]};
                bool flag = false;

                startPage = 1;
                currentPage = Convert.ToInt32(section.Range.Information[WdInformation.wdActiveEndPageNumber]);
                endPage = currentPage - previousPage;

                foreach (HeaderFooter f in footers)
                {
                    if (f != null)
                    {
                        flag = true;
                        string checkFooterText = f.Range.Text;
                        if (checkFooterText.Length > 0)
                        {
                            int foundIndex = checkFooterText.IndexOf("pernr") + 5;
                            string pernr = "";
                            try 
                            {
                                pernr = checkFooterText.Substring(foundIndex, 8);
                            } 
                            catch (Exception)
                            {
                                continue;
                            }
                            string pattern = @"^\d{8}$";
                            if (Regex.IsMatch(pernr, pattern))
                            {
                                pageToCodes[sectionNumber] = new PageToCode(sectionNumber, startPage, endPage, currentPage, previousPage, pernr);
                                break;
                            }
                        }
                    }
                }

                if (flag == false)
                {
                    System.Windows.MessageBox.Show("Wybrany plik nie posiada stopki!");
                    return;
                }
                previousPage = currentPage;
                sectionNumber += 1;
                progressBar1.Value += progressBarRange1/documentWordSections.Count;
            }

            progressBar1.Value = progressBarRange1;
            // Utworzenie tymczasowgo pliku pdf na podstawie wskazanego pliku docx
            object tempFileNameWithoutExtension = Path.GetFileNameWithoutExtension(path).ToLower();
            string uuid = Guid.NewGuid().ToString();
            string tempFilename = System.IO.Path.GetTempPath() + tempFileNameWithoutExtension + uuid + ".pdf";
            object fileFormat = WdSaveFormat.wdFormatPDF;
            documentWord.SaveAs(tempFilename, fileFormat);

            // Koniec przetwarzania wzorcowego pliku docx
            documentWord.Close();
            application.Quit();

            // Utworzenie docelowego pliku PDF przez dialog
            string newPdfFile = "";
            Stream myStream;
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();

            saveFileDialog1.Filter = "PDF Document (*.pdf)|*.pdf";
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (Path.GetExtension(saveFileDialog1.FileName) != ".pdf")
                {
                    System.Windows.MessageBox.Show("Plik musi mieć rozszerzenie .pdf!");
                }
                else
                {
                    if ((myStream = saveFileDialog1.OpenFile()) != null)
                    {
                        myStream.Close();
                        newPdfFile = saveFileDialog1.FileName;
                    }
                }

            }
            saveFileDialog1.Dispose();


            // Wstawianie do tymczasowego pliku pdf QR Kodów i zapisanie go pod nową nazwą
            if (newPdfFile.Length > 0)
            {
                var pdfDocument = new PdfDocument(new PdfReader(tempFilename), new PdfWriter(newPdfFile));
                iText.Layout.Document doc = new iText.Layout.Document(pdfDocument);
                foreach (PageToCode page in pageToCodes)
                {
                    if (page == null) {
                        continue;
                    }
                    startPage = page.previousPage + 1;
                    endPage = page.lastPage;
                    for (var i = page.pageNumberLow - 1; i < page.pageNumberHigh; i++)
                    {
                        string qrCode = page.pernr + "_" + (i + 1);
                        Bitmap qrCodeBitmap = GenerateQrCode(qrCode);
                        MemoryStream memoStream = new MemoryStream();
                        qrCodeBitmap.Save(memoStream, System.Drawing.Imaging.ImageFormat.Png);
                        PdfImageXObject xObject = new PdfImageXObject(iText.IO.Image.ImageDataFactory.CreatePng(memoStream.ToArray()));
                        
                        iText.Layout.Element.Image image = new iText.Layout.Element.Image(xObject, 100f);
                        image.SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.CENTER);
                        image.SetMarginTop(iText.Kernel.Geom.PageSize.A4.GetHeight() - 150);
                        doc.Add(image);
                    }
                    progressBar1.Value += (progressBarRange2 - progressBarRange1)/pageToCodes.Length;
                }
                pdfDocument.Close();
            }
                
            File.Delete(tempFilename);
            progressBar1.Value = progressBarRange2;
            label2.Text = "Ukończono!";
            System.Windows.MessageBox.Show("Wygenerowano plik z QR Kodami!");
        }

        //protected void ProcessDocx(string path)
        //{
        //    using (var document = DocX.Load(path))
        //    {
        //        Dictionary<int, string> sectionToPernr = new Dictionary<int, string>();
        //        int sectionNumber = 0;
        //        while (sectionNumber < document.Sections.Count)
        //        {
        //            Section section = document.Sections[sectionNumber];

        //            Footer[] footers = { section.Footers.First, section.Footers.Odd, section.Footers.Even };
        //            Footer footer = footers[0];
        //            bool flag = false;
        //            foreach (Footer f in footers)
        //            {
        //                if (f != null)
        //                {
        //                    footer = f;
        //                    flag = true;
        //                    break;
        //                }
        //            }

        //            if (flag == false)
        //            {
        //                System.Windows.MessageBox.Show("Wybrany plik nie posiada stopki!");
        //                return;
        //            }
        //            string checkFooterText = footer.Paragraphs.FirstOrDefault().Text;

        //            if (checkFooterText.Length == 0)
        //            {
        //                footer = section.Footers.Even;
        //                checkFooterText = footer.Paragraphs.FirstOrDefault().Text;
        //                if (checkFooterText.Length == 0)
        //                {
        //                    footer = section.Footers.Odd;
        //                }
        //            }
        //            if (footer != null)
        //            {
        //                // Odczytaj zawartość paragrafu w stopce
        //                Paragraph paragraph = footer.Paragraphs.FirstOrDefault();
        //                if (paragraph != null)
        //                {
        //                    string footerText = paragraph.Text;
        //                    if (footerText.Length > 0)
        //                    {
        //                        int foundIndex = footerText.IndexOf("pernr") + 5;
        //                        string pernr = footerText.Substring(foundIndex);
        //                        pernr = pernr.Trim('}');
        //                        string pattern = @"^\d{8}$";
        //                        if (Regex.IsMatch(pernr, pattern))
        //                        {
        //                            sectionToPernr.Add(sectionNumber,pernr.ToString());
                                    
                                    
        //                            //paragraph.RemoveText(0, footerText.Length, false, false);
        //                            //Bitmap qrCodeBitmap = GenerateQrCode(pernr);
        //                            //MemoryStream memoStream = new MemoryStream();
        //                            //qrCodeBitmap.Save(memoStream, System.Drawing.Imaging.ImageFormat.Png);
        //                            //Xceed.Document.NET.Image image = document.AddImage(memoStream);
        //                            //var picture = image.CreatePicture(100f, 100f);
        //                            //paragraph.Alignment = Alignment.center;
        //                            //paragraph.AppendPicture(picture);
        //                        }
        //                    }
        //                }
        //            }
        //            sectionNumber++;
        //        }
        //        Stream myStream;
        //        SaveFileDialog saveFileDialog1 = new SaveFileDialog();

        //        saveFileDialog1.Filter = "docx files (*.docx)|*.docx";
        //        saveFileDialog1.RestoreDirectory = true;

        //        if (saveFileDialog1.ShowDialog() == DialogResult.OK)
        //        {
        //            if (Path.GetExtension(saveFileDialog1.FileName) != ".docx")
        //            {
        //                System.Windows.MessageBox.Show("Plik musi mieć rozszerzenie .docx!");
        //            }
        //            else
        //            {
        //                if ((myStream = saveFileDialog1.OpenFile()) != null)
        //                {
        //                    myStream.Close();
        //                    document.SaveAs(saveFileDialog1.FileName);
        //                }
        //            }

        //        }
        //    }
        //}

        protected Bitmap GenerateQrCode(string code)
        {
            QRCodeGenerator qrGenerator = new QRCodeGenerator();
            QRCodeData qrCodeData = qrGenerator.CreateQrCode(code, QRCodeGenerator.ECCLevel.H);
            QRCode qrCode = new QRCode(qrCodeData);
            Bitmap qrCodeImage = qrCode.GetGraphic(20);
            return qrCodeImage;
        }

        protected void ProcessPDFFile(string path)
        {
            if (path.Length == 0)
            {
                return;
            }

            // Create an instance of Bytescout.PDFRenderer.RasterRenderer object and register it.
            RasterRenderer renderer = new RasterRenderer();
            //renderer.RegistrationName = "demo";
            //renderer.RegistrationKey = "demo";

            var reader = new BarcodeReader();
            

            // Load PDF document.
            renderer.LoadDocumentFromFile(path);
            //Page[] pdfPages = new Page[renderer.GetPageCount()];
            Dictionary <string, string> pernrPages = new Dictionary<string, string>();
            for (int i = 0; i < renderer.GetPageCount(); i++)
            {
                // Render first page of the document to BMP image file.
                
                System.Drawing.Image img = renderer.GetImage(i, 118);
                Bitmap btm = img as Bitmap;
                var pernr = reader.Decode(btm);
                if (pernr != null)
                {
                    if(pernrPages.ContainsKey(pernr.ToString()))
                    {
                        pernrPages[pernr.ToString()] = pernrPages[pernr.ToString()] + " ," + (i + 1).ToString();
                    }
                    else
                    {
                        pernrPages[pernr.ToString()] = (i + 1).ToString();
                    }

                    
                    //pdfPages[i] = new Page((i+1).ToString() , results.ToString());
                }
            }

            string outputFile = "";

            foreach (var page in pernrPages)
            {
                outputFile = Path.GetDirectoryName(path);
                outputFile = outputFile + "\\" + page.Key + ".pdf";
                ExtractPages(path, outputFile, page.Value);
            }
        }
        protected void ExtractPages(string sourcePDFpath, string outputFile, string pageRange)
        {
            var pdfDocument = new PdfDocument(new PdfReader(sourcePDFpath));
            var split = new ImprovedSplitter(pdfDocument, range => new PdfWriter(outputFile));
            var result = split.ExtractPageRange(new PageRange(pageRange));
            result.Close();
        }
    }

    class MySplitter : PdfSplitter
    {
        //string toFile=null;
        //public MySplitter(PdfDocument pdfDocument, string toFile) : base(pdfDocument)
        //{
        //    this.toFile = toFile;
        //}
        public MySplitter(PdfDocument pdfDocument) : base(pdfDocument)
        {
        }

        protected override PdfWriter GetNextPdfWriter(PageRange documentPageRange)
        {
            String toFile = @"C:\Users\marci\OneDrive\Pulpit\regulacje_pliki\Extracted.pdf";
            return new PdfWriter(toFile);
        }
    }

    class ImprovedSplitter : PdfSplitter
    {
        private Func<PageRange, PdfWriter> nextWriter;
        public ImprovedSplitter(PdfDocument pdfDocument, Func<PageRange, PdfWriter> nextWriter) : base(pdfDocument)
        {
            this.nextWriter = nextWriter;
        }

        protected override PdfWriter GetNextPdfWriter(PageRange documentPageRange)
        {
            return nextWriter.Invoke(documentPageRange);
        }
    }

    class PageToCode
    {
        public int sectionNumber { get; set; }
        public int pageNumberLow { get; set; }
        public int pageNumberHigh { get; set; }
        public int lastPage { get; set; }
        public int previousPage { get; set; }
        public string pernr { get; set; }

        public PageToCode(int sectionNumber, int pageNumberLow, int pageNumberHigh, int lastPage, int previousPage, string pernr)
        {
            this.sectionNumber = sectionNumber;
            this.pageNumberLow = pageNumberLow;
            this.pageNumberHigh = pageNumberHigh;
            this.lastPage = lastPage;
            this.previousPage = previousPage;
            this.pernr = pernr;
        }
    }


    class Page
    {
        public string pageNumber { get; set; }
        public string decodedQR { get; set; }

        public Page(string pageNumber, string decodedQR)
        {
            this.pageNumber = pageNumber;
            this.decodedQR = decodedQR;
        }
    }
}
