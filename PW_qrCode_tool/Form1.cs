using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Xceed.Words.NET;
using Xceed.Document.NET;
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
            using (var document = DocX.Load(path))
            {
                int page = 0;
                while (page < document.Sections.Count)
                {
                    Xceed.Document.NET.Section section = document.Sections[page];

                    Footer[] footers = { section.Footers.First, section.Footers.Odd, section.Footers.Even };
                    Footer footer = footers[0];
                    bool flag = false;
                    foreach (Footer f in footers)
                    {
                        if (f != null)
                        {
                            footer = f;
                            flag = true;
                            break;
                        }
                    }

                    if (flag == false)
                    {
                        System.Windows.MessageBox.Show("Wybrany plik nie posiada stopki!");
                        return;
                    }
                    string checkFooterText = footer.Paragraphs.FirstOrDefault().Text;

                    if (checkFooterText.Length == 0)
                    {
                        footer = section.Footers.Even;
                        checkFooterText = footer.Paragraphs.FirstOrDefault().Text;
                        if (checkFooterText.Length == 0)
                        {
                            footer = section.Footers.Odd;
                        }
                    }
                    if (footer != null)
                    {
                        // Odczytaj zawartość paragrafu w stopce
                        Xceed.Document.NET.Paragraph paragraph = footer.Paragraphs.FirstOrDefault();
                        if (paragraph != null)
                        {
                            string footerText = paragraph.Text;
                            if (footerText.Length > 0)
                            {
                                int foundIndex = footerText.IndexOf("pernr") + 5;
                                string pernr = footerText.Substring(foundIndex);
                                pernr = pernr.Trim('}');
                                string pattern = @"^\d{8}$";
                                if (Regex.IsMatch(pernr, pattern))
                                {
                                    paragraph.RemoveText(0, footerText.Length, false, false);
                                    Bitmap qrCodeBitmap = GenerateQrCode(pernr);
                                    MemoryStream memoStream = new MemoryStream();
                                    qrCodeBitmap.Save(memoStream, System.Drawing.Imaging.ImageFormat.Png);
                                    Xceed.Document.NET.Image image = document.AddImage(memoStream);
                                    var picture = image.CreatePicture(100f, 100f);
                                    paragraph.Alignment = Alignment.center;
                                    paragraph.AppendPicture(picture);
                                }
                            }
                        }
                    }
                    page++;
                }

                Stream myStream;
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();

                saveFileDialog1.Filter = "docx files (*.docx)|*.docx";
                saveFileDialog1.RestoreDirectory = true;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    if (Path.GetExtension(saveFileDialog1.FileName) != ".docx")
                    {
                        System.Windows.MessageBox.Show("Plik musi mieć rozszerzenie .docx!");
                    }
                    else
                    {
                        if ((myStream = saveFileDialog1.OpenFile()) != null)
                        {
                            myStream.Close();
                            document.SaveAs(saveFileDialog1.FileName);
                        }
                    }

                }
            }
        }

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
            Page[] pdfPages = new Page[renderer.GetPageCount()];
            for (int i = 0; i < renderer.GetPageCount(); i++)
            {
                // Render first page of the document to BMP image file.
                
                System.Drawing.Image img = renderer.GetImage(i, 118);
                Bitmap btm = img as Bitmap;
                var results = reader.Decode(btm);
                if (results != null)
                {
                    pdfPages[i] = new Page(i.ToString() , results.ToString(), btm, img);
                }
            }

            string outputFile = Path.GetDirectoryName(path);
            outputFile = outputFile + "\\" + pdfPages[0].decodedQR + ".pdf";
            ExtractPages(path, outputFile, "1, 4");
        }
        protected void ExtractPages(string sourcePDFpath, string outputFile, string pageRange)
        {

            //PdfDocument pdfDoc = new PdfDocument(new PdfReader(sourcePDFpath));
            //var split = new MySplitter(pdfDoc);
            //var result = split.ExtractPageRange(new PageRange(pageRange));
            //result.Close();

            string range = "1, 3";
            var pdfDocumentInvoiceNumber = new PdfDocument(new PdfReader(sourcePDFpath));
            var split = new ImprovedSplitter(pdfDocumentInvoiceNumber, pageRange2 => new PdfWriter(outputFile));
            var result = split.ExtractPageRange(new PageRange(range));
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

    class Page
    {
        public string pageNumber { get; set; }
        public string decodedQR { get; set; }
        public Bitmap bmp { get; set; }
        public System.Drawing.Image img { get; set; }

        public Page(string pageNumber, string decodedQR, Bitmap bmp, System.Drawing.Image img)
        {
            this.pageNumber = pageNumber;
            this.decodedQR = decodedQR;
            this.bmp = bmp;
            this.img = img;
        }
    }
}
