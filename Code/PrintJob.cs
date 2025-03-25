using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Web;
using Ghostscript.NET;
using Ghostscript.NET.Processor;

namespace ShaperPrint
{
    public enum DeviceType
    {
        None,
        File,
        Printer
    }

    public delegate void CreatingHandler();

    public class PrintJob
    {
        static object _lockLog = new object();

        PrintDocument _printDocument = null;
        DateTime _start = DateTime.Now;
        string _ip;
        int _contentLength;
        string _path;
        List<string> _messages = new List<string>();

        public event CreatingHandler Creating;

        public PrintDocument PrintDocument
        {
            get
            {
                if (_printDocument == null) _printDocument = new PrintDocument();
                _printDocument.PrinterSettings.PrinterName = Device;
                return _printDocument;
            }
        }

        public string Device { get; set; } = "";
        public DeviceType DeviceType { get; set; } = DeviceType.None;
        public bool Enqueued { get; set; } = false;
        public bool IsRaw { get; set; } = false;
        public byte[] Content { get; set; } = null;
        public byte[] Result { get; private set; } = null;
        public bool Grayscale { get; set; } = false;
        public short Copies { get; set; } = 1;
        public List<Image> Pages { get; } = new List<Image>();
        public byte[] Report { get; set; } = null;
        public Dictionary<string, DataTable> Dataset { get; } = new Dictionary<string, DataTable>();

        public PrintJob(HttpRequest request)
        {
            _ip = request.UserHostAddress;
            _contentLength = request.ContentLength;
            _path = request.PhysicalApplicationPath;
        }

        public void Create()
        {
            Creating?.Invoke();
        }

        public void AddMessage(string message)
        {
            _messages.Add(message);
        }

        public void AddException(Exception ex)
        {
            _messages.Add("Error: " + ex.Message + " " + ex.StackTrace.Replace("\r", " ").Replace("\n", " "));
        }

        public void Complete()
        {
            try
            {
                foreach (Image img in Pages)
                    img.Dispose();
            }
            catch
            {
            }

            WriteLog();
        }

        private void WriteLog()
        {
            try
            {
                string message = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " " +
                    _ip + " " +
                    _contentLength.ToString() + " " +
                    Convert.ToInt32(DateTime.Now.Subtract(_start).TotalMilliseconds).ToString() + "ms" + " " +
                    string.Join(", ", _messages);

                lock (_lockLog)
                {
                    DirectoryInfo di = new DirectoryInfo(_path + "\\App_Data\\Log");
                    if (!di.Exists) di.Create();
                    FileStream fs = new FileStream(di.FullName + "\\log_" + DateTime.Now.ToString("yyyyMMdd") + ".txt", FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read);
                    fs.Position = fs.Length;
                    StreamWriter sw = new StreamWriter(fs);
                    sw.WriteLine(message);
                    sw.Close();
                    fs.Close();
                }
            }
            catch
            {
            }
        }

        public void WriteImages()
        {
            _messages.Add(Pages.Count.ToString() + " pages printed to " + Device);

            if (Pages.Count == 0)
                return;

            int n = 0;

            PrintDocument.PrinterSettings.Copies = Copies;

            // set page format from first page, always portrait
            PrintDocument.DefaultPageSettings.Landscape = false;
            bool paperPortrait = (PrintDocument.PrinterSettings.DefaultPageSettings.PaperSize.Width <= PrintDocument.PrinterSettings.DefaultPageSettings.PaperSize.Height);
            bool documentPortrait = (Pages[0].Width <= Pages[0].Height);

            PrintDocument.PrintPage += (sender, e) =>
            {
                e.Graphics.TranslateTransform(-e.PageSettings.HardMarginX, -e.PageSettings.HardMarginY);

                if (paperPortrait ^ documentPortrait)
                    Pages[n].RotateFlip(RotateFlipType.Rotate270FlipNone);

                e.Graphics.DrawImage(Pages[n], 0, 0);

                n++;
                e.HasMorePages = (n < Pages.Count);
            };

            PrintDocument.Print();
        }

        public void WriteRaw()
        {
            _messages.Add("RAW printed to " + Device);

            int nLength = Content.Length;
            IntPtr pUnmanagedBytes = Marshal.AllocCoTaskMem(nLength);
            Marshal.Copy(Content, 0, pUnmanagedBytes, nLength);
            RawPrinterHelper.SendBytesToPrinter(Device, pUnmanagedBytes, nLength);
            Marshal.FreeCoTaskMem(pUnmanagedBytes);
        }

        public void WriteToPath()
        {
            _messages.Add("Saved to " + Device);

            FileStream fs = new FileStream(Device, FileMode.CreateNew, FileAccess.Write);
            fs.Write(Content, 0, Content.Length);
            fs.Close();
        }

        public void PrintRAW()
        {
            PrinterQueueManager.Enqueue(this);
        }

        public void PrintPDF()
        {
            Creating += () =>
            {
                _messages.Add("Processing PDF to PNG");

                Ghostscript.NET.Processor.GhostscriptProcessor processor = new GhostscriptProcessor(GhostscriptVersionInfo.GetLastInstalledVersion(), true);
                DirectoryInfo di = new DirectoryInfo(_path + "\\App_Data\\Temp");
                string pfix = Guid.NewGuid().ToString("n");

                try
                {
                    if (!di.Exists) di.Create();

                    FileStream fs = new FileStream(di.FullName + "\\" + pfix + "_in.pdf", FileMode.CreateNew, FileAccess.Write);
                    fs.Write(Content, 0, Content.Length);
                    fs.Close();

                    List<string> switches = new List<string>();
                    switches.Add("-dBATCH");
                    switches.Add("-dNOPAUSE");
                    switches.Add("-dNOSAFER");
                    switches.Add("-dNOPROMPT");
                    switches.Add("-dQUIET");

                    switches.Add("-r" + PrintDocument.PrinterSettings.DefaultPageSettings.PrinterResolution.X.ToString() +
                        "x" + PrintDocument.PrinterSettings.DefaultPageSettings.PrinterResolution.Y.ToString());  // resolution

                    if (Grayscale)
                        switches.Add("-sDEVICE=pnggray");
                    else
                        switches.Add("-sDEVICE=png16m");

                    switches.Add("-sOutputFile=" + di.FullName + "\\" + pfix + "_%04d.png");
                    switches.Add(di.FullName + "\\" + pfix + "_in.pdf");

                    processor.Process(switches.ToArray(), null);

                    foreach (FileInfo fi in di.GetFiles(pfix + "*.png"))
                    {
                        byte[] buf = new byte[0];
                        bool readed = false;

                        for (int i = 0; i < 3; i++)
                        {
                            try
                            {
                                fs = new FileStream(fi.FullName, FileMode.Open, FileAccess.Read);
                                buf = new byte[fs.Length];
                                fs.Read(buf, 0, buf.Length);
                                fs.Close();
                                readed = true;
                                break;
                            }
                            catch
                            {
                                Thread.Sleep(1000);
                            }
                        }

                        if (!readed)
                            throw new Exception("Unable to access file " + fi.Name);

                        Pages.Add(Image.FromStream(new MemoryStream(buf)));
                    }
                }
                finally
                {
                    try
                    {
                        foreach (FileInfo fi in di.GetFiles(pfix + "*"))
                            fi.Delete();
                    }
                    catch
                    {
                    }
                }
            };

            PrinterQueueManager.Enqueue(this);
        }

        public void PrintRDL()
        {
            Creating += () =>
            {
                Microsoft.Reporting.WebForms.ReportViewer repView = CreateReportViewer();
                if (DeviceType == DeviceType.File)
                {
                    _messages.Add("Rendered RDL to PDF");
                    Content = repView.LocalReport.Render("PDF");
                }
                else
                {
                    _messages.Add("Rendered RDL to PNG");

                    int dpiX = PrintDocument.PrinterSettings.DefaultPageSettings.PrinterResolution.X;
                    int dpiY = PrintDocument.PrinterSettings.DefaultPageSettings.PrinterResolution.Y;

                    var devinfo = "<DeviceInfo>" +
                        "<OutputFormat>PNG</OutputFormat>" +
                        "<DpiX>" + dpiX.ToString() + "</DpiX>" +
                        "<DpiY>" + dpiY.ToString() + "</DpiY>" +
                        "</DeviceInfo>";

                    List<MemoryStream> streams = new List<MemoryStream>();

                    Microsoft.Reporting.WebForms.Warning[] wrns;
                    repView.LocalReport.Render("Image", devinfo, (name, ext, enc, mime, seek) =>
                    {
                        MemoryStream ms = new MemoryStream();
                        streams.Add(ms);
                        return ms;
                    }, out wrns);

                    foreach (MemoryStream ms in streams)
                    {
                        var bmp = (Bitmap)Image.FromStream(ms);
                        bmp.SetResolution(dpiX, dpiY);
                        Pages.Add(bmp);
                    }
                }
            };

            PrinterQueueManager.Enqueue(this);
        }

        public void RenderRDLtoPDF()
        {
            _messages.Add("Rendered RDL and sent PDF");
            Result = CreateReportViewer().LocalReport.Render("PDF");
        }

        private Microsoft.Reporting.WebForms.ReportViewer CreateReportViewer()
        {
            Microsoft.Reporting.WebForms.ReportViewer repView = new Microsoft.Reporting.WebForms.ReportViewer();

            MemoryStream defStream = new MemoryStream(Report);
            repView.ProcessingMode = Microsoft.Reporting.WebForms.ProcessingMode.Local;
            repView.LocalReport.LoadReportDefinition(defStream);
            defStream.Close();

            foreach (string name in Dataset.Keys)
                repView.LocalReport.DataSources.Add(new Microsoft.Reporting.WebForms.ReportDataSource(name, Dataset[name]));

            return repView;
        }
    }
}