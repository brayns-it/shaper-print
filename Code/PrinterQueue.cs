using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Web;
using Ghostscript.NET;
using Ghostscript.NET.Processor;

namespace ShaperPrint
{
    public class PrinterQueue
    {
        object _lockJobs = new object();
        SemaphoreSlim _semaphore = new SemaphoreSlim(0);
        List<PrintJob> _jobs = new List<PrintJob>();
        
        public PrinterQueue()
        {
            Thread thd = new Thread(new ThreadStart(Worker));
            thd.Start();
        }

        public void Enqueue(PrintJob job)
        {
            job.AddMessage("Enqueued");
            job.Enqueued = true;

            lock(_lockJobs)
            {
                _jobs.Add(job);
            }

            _semaphore.Release();
        }

        private void Worker()
        {
            try
            {
                while (!Global.StopRequest.IsCancellationRequested)
                {
                    _semaphore.Wait(Global.StopRequest.Token);

                    while (_jobs.Count > 0)
                    {
                        PrintJob job;

                        lock (_lockJobs)
                        {
                            job = _jobs[0];
                            _jobs.RemoveAt(0);
                        }

                        SpoolOne(job);
                    }
                }
            }
            catch
            {
            }
        }

        private void SpoolOne(PrintJob job)
        {
            try
            {
                job.Create();

                if (job.DeviceType == DeviceType.File)
                    job.WriteToPath();
                else if (job.IsRaw)
                    job.WriteRaw();
                else
                    job.WriteImages();
            }
            catch (Exception ex)
            {
                job.AddException(ex);
            }

            job.Complete();
        }
    }

    public static class PrinterQueueManager
    {
        static object _lockQueues = new object();
        static Dictionary<string, PrinterQueue> Queues { get; } = new Dictionary<string, PrinterQueue>();

        public static void Enqueue(PrintJob job)
        {
            bool found = false;
            foreach (string ip in PrinterSettings.InstalledPrinters)
            {
                if (ip.Equals(job.Device, StringComparison.OrdinalIgnoreCase))
                {
                    found = true;
                    break;
                }
            }
            if (!found)
                throw new Exception(string.Format("Printer {0} not found", job.Device));

            lock (_lockQueues)
            {
                if (!Queues.ContainsKey(job.Device.ToLower()))
                    Queues[job.Device.ToLower()] = new PrinterQueue();
            }

            Queues[job.Device.ToLower()].Enqueue(job);
        }
    }
}