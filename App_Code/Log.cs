using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace ShaperPrint
{
    public class Log
    {
        static object _lockLog = new object();
        DateTime _start = DateTime.Now;
        string _ip;
        int _contentLength;
        string _path;
        List<string> _messages = new List<string>();

        public bool Enqueued { get; set; } = false;

        public Log(HttpRequest request)
        {
            _ip = request.UserHostAddress;
            _contentLength = request.ContentLength;
            _path = request.PhysicalApplicationPath;
        }

        public void Add(string message)
        {
            _messages.Add(message);
        }

        public void Write()
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
    }
}