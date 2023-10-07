using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using QRCoder;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;
using ZXing;
using Bytescout.PDFRenderer;
using iText.Kernel.Pdf;
using iText.Kernel.Utils;
using System.Collections.Generic;
using Stream = System.IO.Stream;
using PageRange = iText.Kernel.Utils.PageRange;
using Microsoft.Office.Interop.Word;
using iText.Kernel.Pdf.Xobject;
using MessagingToolkit.QRCode.Codec;
using MessagingToolkit.QRCode.Codec.Data;
using System.Windows.Controls;
using System.Drawing.Drawing2D;
using ZXing.QrCode.Internal;

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
                progressBar2.Minimum = 0;
                progressBar2.Maximum = 10000;
                progressBar2.Step = 1;
                progressBar2.Style = ProgressBarStyle.Continuous;

                label4.Text = "Przetwarzanie...";
                label4.Refresh();

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
            label2.Refresh();

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
            progressBar1.Value = 0;
            _Application application = new Word.Application();
            _Document documentWord = application.Documents.Open(path);
            var pages = documentWord.ComputeStatistics(WdStatistic.wdStatisticPages, false);
            Sections documentWordSections = documentWord.Sections;

            PageToCode[] pageToCodes = new PageToCode[documentWordSections.Count+1];
            int sectionNumber = 1;
            int startPage, endPage, currentPage, previousPage = 0;
            while (sectionNumber <= documentWordSections.Count)
            {
                Section section = documentWordSections[sectionNumber];
                HeaderFooter[] footers = { section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary], section.Footers[WdHeaderFooterIndex.wdHeaderFooterEvenPages], section.Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage]};
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
                        image.SetMarginTop(iText.Kernel.Geom.PageSize.A4.GetHeight() - 155);
                        doc.Add(image);
                    }
                    progressBar1.Value += (progressBarRange2 - progressBarRange1)/pageToCodes.Length;
                }
                pdfDocument.Close();
            }             
            File.Delete(tempFilename);
            progressBar1.Value = progressBar1.Maximum;
            label2.Text = "Ukończono!";
            label2.Refresh();
            System.Windows.MessageBox.Show("Wygenerowano plik z QR Kodami!");
        }

        protected Bitmap GenerateQrCode(string code)
        {
            QRCodeGenerator qrGenerator = new QRCodeGenerator();
            QRCodeData qrCodeData = qrGenerator.CreateQrCode(code, QRCodeGenerator.ECCLevel.H);
            QRCoder.QRCode qrCode = new QRCoder.QRCode(qrCodeData);
            Bitmap qrCodeImage = qrCode.GetGraphic(20);

            Graphics g = Graphics.FromImage(qrCodeImage);
            g.DrawString(code, new System.Drawing.Font("Adagio Slab", 30, FontStyle.Bold), Brushes.Black, (qrCodeImage.Width/3)-10, 20);

            return qrCodeImage;
        }

        protected void ProcessPDFFile(string path)
        {
            if (path.Length == 0)
            {
                return;
            }
            int progressBarRange1 = progressBar2.Maximum / 2;
            int progressBarRange2 = progressBar2.Maximum;
            progressBar2.Value = 0;

            RasterRenderer renderer = new RasterRenderer();
            var reader = new BarcodeReader();
            
            renderer.LoadDocumentFromFile(path);
            Dictionary<String, Page> pernrPages = new Dictionary<String, Page>();
            List<int> errorPages = new List<int>();
            for (int i = 0; i < renderer.GetPageCount(); i++)
            {
                System.Drawing.Image img = renderer.GetImage(i, 118);
                Bitmap btm = img as Bitmap;
                var pernr = reader.Decode(btm);
                string decodedString = null;
                if (pernr == null)
                {
                    QRCodeDecoder decode = new QRCodeDecoder();
                    decodedString = decode.Decode(new QRCodeBitmapImage(btm));

                } else
                {
                    decodedString = pernr.ToString();
                }

                if (decodedString == null) {
                    errorPages.Add(i+1);
                    continue;
                }
                string[] decodedQr = decodedString.Split('_');
                string pernrString = decodedQr[0];
                int pageNumber = 0;
                if (int.TryParse(decodedQr[1], out _))
                {
                    pageNumber = Convert.ToInt32(decodedQr[1]);
                }
                else
                {
                    errorPages.Add(i + 1);
                    continue;
                }
                

                string pattern = @"^\d{8}$";
                if (pernrString != null && Regex.IsMatch(pernrString, pattern))
                {
                    Page pageMappingObj;
                    if (pernrPages.ContainsKey(pernrString))
                    {
                        pageMappingObj = pernrPages[pernrString];
                        pageMappingObj.AddPageMapping(pageNumber, (i+1));
                        pernrPages[pageMappingObj.pernr] = pageMappingObj;
                    }
                    else
                    {
                        pageMappingObj = new Page(pernrString);
                        pageMappingObj.AddPageMapping(pageNumber, (i+1));
                        pernrPages.Add(pageMappingObj.pernr, pageMappingObj);
                    }
                }
                else
                {
                    errorPages.Add(i + 1);
                    continue;
                }

                progressBar2.Value += progressBarRange1 / renderer.GetPageCount();
            }

            progressBar2.Value = progressBarRange1;
            string outputFolder = "";

            FolderBrowserDialog outputFolderDialog = new FolderBrowserDialog();
            outputFolderDialog.Description = "Wybierz folder w którym mają być zapisane przetworzone pliki";
            DialogResult result = outputFolderDialog.ShowDialog();
            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(outputFolderDialog.SelectedPath))
            {
                outputFolder = outputFolderDialog.SelectedPath;
            }
            else if(string.IsNullOrWhiteSpace(outputFolderDialog.SelectedPath))
            {
                System.Windows.MessageBox.Show("Ścieżka do wybranego folder nie istnieje!");
                label4.Text = string.Empty;
                label4.Refresh();
                progressBar2.Value = 0;
                return;
            }else
            {
                label4.Text = string.Empty;
                label4.Refresh();
                progressBar2.Value = 0;
                return;
            }
            
            foreach (var pernrPage in pernrPages)
            {
                string outputFile = outputFolder + "\\" + pernrPage.Key + ".pdf"; ;
                string pageRange = null;
                Dictionary<int, int> pagesMapping = pernrPage.Value.GetPageMapping();
                foreach (var pageMapping in pagesMapping)
                {
                    if(pageMapping.Key == 1)
                    {
                        pageRange = pageMapping.Value.ToString();
                    }
                    else
                    {
                        pageRange = pageRange + ',' + pageMapping.Value.ToString();
                    }
                    
                }
                if (pageRange.Length > 0)
                {
                    ExtractPages(path, outputFile, pageRange);
                }
                progressBar2.Value += (progressBarRange2-progressBarRange1)/pernrPages.Count();
            }

            progressBar2.Value = progressBar2.Maximum;
            label4.Text = "Ukończono!";
            label4.Refresh();
            if (errorPages.Count == 0)
            {
                System.Windows.MessageBox.Show("Wygenerowano pliki!");
            }
            else
            {
                string errorPagesMessage = "Nie udało się odczytać QR Kodu ze stron:";
                foreach(var errorPage in errorPages)
                {
                    errorPagesMessage = errorPagesMessage + System.Environment.NewLine + String.Format("{0}", errorPage);
                }
                System.Windows.MessageBox.Show(errorPagesMessage);
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
        public string pernr { get; set; }
        public Dictionary<int, int> pdfToDocPages { get; set; }

        public Page(string pernr)
        {
            this.pernr = pernr;
            this.pdfToDocPages = new Dictionary<int, int>();
        }

        public void AddPageMapping(int docPageNumber, int pdfPageNumber) 
        { 
            if (this.pdfToDocPages.ContainsKey(pdfPageNumber))
            {
                this.pdfToDocPages[docPageNumber] = pdfPageNumber;
            }
            else
            {
                this.pdfToDocPages.Add(docPageNumber, pdfPageNumber);
            }
            
            this.pdfToDocPages = this.pdfToDocPages.OrderBy(obj=>obj.Key).ToDictionary(obj=>obj.Key, obj=>obj.Value);
        }

        public Dictionary<int,int> GetPageMapping()
        {
            return this.pdfToDocPages;
        }
    }
}
