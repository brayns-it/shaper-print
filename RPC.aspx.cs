using System;
using System.Threading;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Runtime.InteropServices;
using System.Xml;
using System.Text;
using System.Drawing;
using System.Drawing.Printing;
using System.Data;

namespace ShaperPrint
{
    public partial class RPC : System.Web.UI.Page
    {
        byte[] _result = null;
        Log _log;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (Request.HttpMethod == "POST")
            {
                _log = new Log(Request);

                try
                {
                    StreamReader sr = new StreamReader(Request.InputStream);
                    JObject jRequest = JObject.Parse(sr.ReadToEnd());
                    sr.Close();

                    XmlDocument doc = new XmlDocument();
                    doc.Load(Request.PhysicalApplicationPath + "\\App_Data\\Settings.xml");
                    if (doc.SelectSingleNode("settings/authToken").InnerText != jRequest["token"].ToString())
                        throw new Exception("Invalid token");

                    switch (jRequest["request"].ToString())
                    {
                        case "print":
                            Print(jRequest);
                            break;

                        case "render":
                            Render(jRequest);
                            break;

                        default:
                            throw new Exception("Invalid request");
                    }

                    JObject jError = new JObject();
                    jError["status"] = (_log.Enqueued) ? "enqueued" : "success";
                    if (_result != null)
                        jError["result"] = Convert.ToBase64String(_result);
                    Response.Write(jError.ToString());
                }
                catch (Exception ex)
                {
                    JObject jError = new JObject();
                    jError["status"] = "error";
                    jError["message"] = ex.Message;
                    Response.Write(jError.ToString());

                    _log.Add("Error: " + ex.Message);
                }

                if (!_log.Enqueued) _log.Write();
                Response.End();
            }
        }

        private void Render(JObject jRequest)
        {
            switch (jRequest["format"].ToString())
            {
                case "RDL":
                    RenderRDL(jRequest);
                    break;
                default:
                    throw new Exception("Invalid format");
            }
        }

        private Microsoft.Reporting.WebForms.ReportViewer CreateReportViewer(JObject jRequest)
        {
            Microsoft.Reporting.WebForms.ReportViewer repView = new Microsoft.Reporting.WebForms.ReportViewer();

            MemoryStream defStream = new MemoryStream(Convert.FromBase64String(jRequest["rdl"].ToString()));
            repView.LocalReport.LoadReportDefinition(defStream);
            defStream.Close();

            foreach (var prop in ((JObject)jRequest["datasets"]).Properties())
            {
                MemoryStream dsStream = new MemoryStream(Convert.FromBase64String(prop.Value.ToString()));
                DataSet dt = new DataSet();
                dt.ReadXml(dsStream);
                dsStream.Close();

                repView.LocalReport.DataSources.Add(new Microsoft.Reporting.WebForms.ReportDataSource(prop.Name, dt.Tables[0]));
            }

            return repView;
        }

        private void RenderRDL(JObject jRequest)
        {
            _log.Add("Rendered RDL");

            if (jRequest.ContainsKey("device"))
            {
                _log.Enqueued = true;

                Thread thd = new Thread(new ThreadStart(() =>
                {
                    try
                    {
                        Microsoft.Reporting.WebForms.ReportViewer repView = CreateReportViewer(jRequest);
                        string device = jRequest["device"].ToString();

                        if (device.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
                        {
                            byte[] buf = repView.LocalReport.Render("PDF");
                            PrintToPath(device.Substring(7), buf);
                        }
                        else if (device.StartsWith("printer://", StringComparison.OrdinalIgnoreCase))
                        {
                            PrintDocument printDoc = new PrintDocument();
                            printDoc.PrinterSettings.PrinterName = device.Substring(10);

                            var devinfo = "<DeviceInfo>" +
                                "<OutputFormat>PNG</OutputFormat>" +
                                "<DpiX>" + printDoc.PrinterSettings.DefaultPageSettings.PrinterResolution.X + "</DpiX>" +
                                "<DpiY>" + printDoc.PrinterSettings.DefaultPageSettings.PrinterResolution.X + "</DpiY>" +
                                "</DeviceInfo>";

                            List<MemoryStream> streams = new List<MemoryStream>();

                            Microsoft.Reporting.WebForms.Warning[] wrns;
                            repView.LocalReport.Render("Image", devinfo, (name, ext, enc, mime, seek) =>
                            {
                                MemoryStream ms = new MemoryStream();
                                streams.Add(ms);
                                return ms;
                            }, out wrns);

                            List<Image> pages = new List<Image>();
                            foreach (MemoryStream ms in streams)
                                pages.Add(Image.FromStream(ms));

                            PrintImages(printDoc, pages, jRequest);
                        }
                        else
                            throw new Exception("Invalid device");
                    }
                    catch (Exception ex)
                    {
                        _log.Add("Error: " + ex.Message);
                    }

                    _log.Write();
                }));
                thd.Start();
            }
            else
            {
                _log.Add("Sent PDF to client");
                _result = CreateReportViewer(jRequest).LocalReport.Render("PDF");
            }
        }

        private void Print(JObject jRequest)
        {
            string device = jRequest["device"].ToString();
            byte[] content = Convert.FromBase64String(jRequest["content"].ToString());

            if (device.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
                PrintToPath(device.Substring(7), content);
            else if (device.StartsWith("printer://", StringComparison.OrdinalIgnoreCase))
                PrintToPrinter(device.Substring(10), content, jRequest);
            else
                throw new Exception("Invalid device");
        }

        private void PrintToPrinter(string printerName, byte[] content, JObject jRequest)
        {
            switch (jRequest["format"].ToString())
            {
                case "PDF":
                    PrintPDF(printerName, content, jRequest);
                    break;
                case "RAW":
                    PrintRAW(printerName, content);
                    break;
                default:
                    throw new Exception("Invalid format");
            }
        }

        private void PrintRAW(string printerName, byte[] content)
        {
            _log.Add("RAW printed to " + printerName);
            _log.Enqueued = true;

            Thread thd = new Thread(new ThreadStart(() =>
            {
                try
                {
                    int nLength = content.Length;
                    IntPtr pUnmanagedBytes = Marshal.AllocCoTaskMem(nLength);
                    Marshal.Copy(content, 0, pUnmanagedBytes, nLength);
                    RawPrinterHelper.SendBytesToPrinter("File Printer", pUnmanagedBytes, nLength);
                    Marshal.FreeCoTaskMem(pUnmanagedBytes);
                }
                catch (Exception ex)
                {
                    _log.Add("Error: " + ex.Message);
                }

                _log.Write();
            }));
            thd.Start();
        }

        private void PrintImages(PrintDocument printDoc, List<Image> pages, JObject jRequest)
        {
            _log.Add(pages.Count.ToString() + " pages printed to " + printDoc.PrinterSettings.PrinterName);

            if (pages.Count == 0)
                return;

            int n = 0;

            if (jRequest.ContainsKey("copies"))
                printDoc.PrinterSettings.Copies = Convert.ToInt16(jRequest["copies"]);

            // set page format from first page, always portrait
            printDoc.DefaultPageSettings.Landscape = false;
            bool paperPortrait = (printDoc.PrinterSettings.DefaultPageSettings.PaperSize.Width <= printDoc.PrinterSettings.DefaultPageSettings.PaperSize.Height);
            bool documentPortrait = (pages[0].Width <= pages[0].Height);

            printDoc.PrintPage += (sender, e) =>
            {
                e.Graphics.TranslateTransform(-e.PageSettings.HardMarginX, -e.PageSettings.HardMarginY);

                if (paperPortrait ^ documentPortrait)
                    pages[n].RotateFlip(RotateFlipType.Rotate270FlipNone);
                
                e.Graphics.DrawImage(pages[n], 0, 0);
                
                n++;
                e.HasMorePages = (n < pages.Count);
            };

            printDoc.Print();
        }

        private void PrintPDF(string printerName, byte[] content, JObject jRequest)
        {
            _log.Add("PDF processed");
            _log.Enqueued = true;
            string path = Request.PhysicalApplicationPath;

            Thread thd = new Thread(new ThreadStart(() =>
            {
                try
                {
                    PrintDocument printDoc = new PrintDocument();
                    printDoc.PrinterSettings.PrinterName = printerName;

                    Ghostscript.NET.Processor.GhostscriptProcessor p = new Ghostscript.NET.Processor.GhostscriptProcessor();
                    DirectoryInfo di = new DirectoryInfo(path + "\\App_Data\\Temp");
                    string pfix = Guid.NewGuid().ToString("n");

                    try
                    {
                        if (!di.Exists) di.Create();

                        FileStream fs = new FileStream(di.FullName + "\\" + pfix + "_in.pdf", FileMode.CreateNew, FileAccess.Write);
                        fs.Write(content, 0, content.Length);
                        fs.Close();

                        List<string> switches = new List<string>();
                        switches.Add("-dBATCH");
                        switches.Add("-dNOPAUSE");
                        switches.Add("-dNOSAFER");
                        switches.Add("-dNOPROMPT");
                        switches.Add("-dQUIET");

                        switches.Add("-r" + printDoc.PrinterSettings.DefaultPageSettings.PrinterResolution.X.ToString() + 
                            "x" + printDoc.PrinterSettings.DefaultPageSettings.PrinterResolution.Y.ToString());  // resolution

                        if (jRequest.ContainsKey("grayscale") && Convert.ToBoolean(jRequest["grayscale"]))
                            switches.Add("-sDEVICE=pnggray");
                        else
                            switches.Add("-sDEVICE=png16m");

                        switches.Add("-sOutputFile=" + di.FullName + "\\" + pfix + "_%04d.png");
                        switches.Add(di.FullName + "\\" + pfix + "_in.pdf");

                        p.Process(switches.ToArray(), null);

                        List<Image> pages = new List<Image>();
                        foreach (FileInfo fi in di.GetFiles(pfix + "*.png"))
                            pages.Add(Image.FromFile(fi.FullName));

                        PrintImages(printDoc, pages, jRequest);

                        foreach (Image img in pages)
                            img.Dispose();
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
                }
                catch (Exception ex)
                {
                    _log.Add("Error: " + ex.Message);
                }

                _log.Write();
            }));
            thd.Start();
        }

        private void PrintToPath(string fileName, byte[] content)
        {
            _log.Add("Saved to " + fileName);

            FileStream fs = new FileStream(fileName, FileMode.CreateNew, FileAccess.Write);
            fs.Write(content, 0, content.Length);
            fs.Close();
        }
    }
}