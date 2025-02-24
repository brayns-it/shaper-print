using System;
using System.Threading;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Xml;
using System.Text;
using System.Drawing;
using System.Drawing.Printing;
using System.Data;

namespace ShaperPrint
{
    public partial class RPC : System.Web.UI.Page
    {
        PrintJob _job = null;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (Request.HttpMethod == "POST")
            {
                _job = new PrintJob(Request);

                try
                {
                    StreamReader sr = new StreamReader(Request.InputStream);
                    JObject jRequest = JObject.Parse(sr.ReadToEnd());
                    sr.Close();

                    XmlDocument doc = new XmlDocument();
                    doc.Load(Request.PhysicalApplicationPath + "\\App_Data\\Settings.xml");
                    if (doc.SelectSingleNode("settings/authToken").InnerText != jRequest["token"].ToString())
                        throw new Exception("Invalid token");

                    if (jRequest.ContainsKey("device"))
                    {
                        string device = jRequest["device"].ToString();

                        if (device.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
                        {
                            _job.DeviceType = DeviceType.File;
                            _job.Device = device.Substring(7);
                        }
                        else if (device.StartsWith("printer://", StringComparison.OrdinalIgnoreCase))
                        {
                            _job.DeviceType = DeviceType.Printer;
                            _job.Device = device.Substring(10);
                        }
                        else
                            throw new Exception(string.Format("Invalid device {0}", device));
                    }

                    switch (jRequest["request"].ToString())
                    {
                        case "test":
                            break;

                        case "print":
                            Print(jRequest);
                            break;

                        case "render":
                            Render(jRequest);
                            break;

                        default:
                            throw new Exception("Invalid request");
                    }

                    JObject jResult = new JObject();
                    jResult["status"] = _job.Enqueued ? "enqueued" : "success";
                    if (_job.Result != null)
                        jResult["result"] = Convert.ToBase64String(_job.Result);
                    Response.Write(jResult.ToString());
                }
                catch (Exception ex)
                {
                    JObject jError = new JObject();
                    jError["status"] = "error";
                    jError["message"] = ex.Message;
                    Response.Write(jError.ToString());

                    _job.AddException(ex);
                }

                if (!_job.Enqueued) _job.Complete();
                Response.End();
            }
        }

        private void Render(JObject jRequest)
        {
            _job.Report = Convert.FromBase64String(jRequest["report"].ToString());

            foreach (var prop in ((JObject)jRequest["datasets"]).Properties())
            {
                MemoryStream dsStream = new MemoryStream(Convert.FromBase64String(prop.Value.ToString()));
                DataSet dt = new DataSet();
                dt.ReadXml(dsStream);
                dsStream.Close();

                _job.Dataset.Add(prop.Name, dt.Tables[0]);
            }

            switch (jRequest["format"].ToString())
            {
                case "RDL":
                    if (_job.DeviceType == DeviceType.None)
                        _job.RenderRDLtoPDF();
                    else
                        _job.PrintRDL();
                    break;

                default:
                    throw new Exception("Invalid format");
            }
        }

        private void Print(JObject jRequest)
        {
            _job.Content = Convert.FromBase64String(jRequest["content"].ToString());

            switch (_job.DeviceType)
            {
                case DeviceType.File:
                    _job.WriteToPath();
                    break;

                case DeviceType.Printer:
                    switch (jRequest["format"].ToString())
                    {
                        case "PDF":
                            if (jRequest.ContainsKey("grayscale"))
                                _job.Grayscale = Convert.ToBoolean(jRequest["grayscale"]);

                            if (jRequest.ContainsKey("copies"))
                                _job.Copies = Convert.ToInt16(jRequest["copies"]);

                            _job.PrintPDF();
                            break;

                        case "RAW":
                            _job.IsRaw = true;
                            _job.PrintRAW();
                            break;

                        default:
                            throw new Exception("Invalid format");
                    }
                    break;
            }
        }
    }
}