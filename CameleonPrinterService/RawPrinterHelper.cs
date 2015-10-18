using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Management;
using System.Printing;

namespace CameleonPrinterService
{
    public class RawPrinterHelper
    {
        // Structure and API declarions:
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        public class DOCINFOA
        {
            [MarshalAs(UnmanagedType.LPStr)]
            public string pDocName;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pOutputFile;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pDataType;
        }

        [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter, out IntPtr hPrinter, IntPtr pd);

        [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool ClosePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "StartDocPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool StartDocPrinter(IntPtr hPrinter, Int32 level, [In, MarshalAs(UnmanagedType.LPStruct)] DOCINFOA di);

        [DllImport("winspool.Drv", EntryPoint = "EndDocPrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool EndDocPrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "StartPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool StartPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "EndPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool EndPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "WritePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, Int32 dwCount, out Int32 dwWritten);

        // SendBytesToPrinter()
        // When the function is given a printer name and an unmanaged array
        // of bytes, the function sends those bytes to the print queue.
        // Returns true on success, false on failure.
        public static bool SendBytesToPrinter(string szPrinterName, string szDocumentName, IntPtr pBytes, Int32 dwCount, string printDataType = "RAW")
        {
            Int32 dwError = 0, dwWritten = 0;
            IntPtr hPrinter = new IntPtr(0);
            DOCINFOA di = new DOCINFOA();
            bool bSuccess = false; // Assume failure unless you specifically succeed.

            di.pDocName = szDocumentName;
            // List of Data Types
            // RAW
            // TEXT
            // RAW [FF appended]
            // RAW [FF auto]
            // NT EMF 1.00x
            di.pDataType = printDataType;

            // Open the printer.
            if (OpenPrinter(szPrinterName.Normalize(), out hPrinter, IntPtr.Zero))
            {
                // Start a document.
                if (StartDocPrinter(hPrinter, 1, di))
                {
                    // Start a page.
                    if (StartPagePrinter(hPrinter))
                    {
                        // Write your bytes.
                        bSuccess = WritePrinter(hPrinter, pBytes, dwCount, out dwWritten);
                        EndPagePrinter(hPrinter);
                    }
                    EndDocPrinter(hPrinter);
                }
                ClosePrinter(hPrinter);
            }
            // If you did not succeed, GetLastError may give more information
            // about why not.
            if (bSuccess == false)
            {
                dwError = Marshal.GetLastWin32Error();
            }
            return bSuccess;
        }

        public static bool SendFileToPrinter(string szPrinterName, string szFileName)
        {
            // Open the file.
            FileStream fs = new FileStream(szFileName, FileMode.Open);
            // Create a BinaryReader on the file.
            BinaryReader br = new BinaryReader(fs);
            // Dim an array of bytes big enough to hold the file's contents.
            Byte[] bytes = new Byte[fs.Length];
            bool bSuccess = false;
            // Your unmanaged pointer.
            IntPtr pUnmanagedBytes = new IntPtr(0);
            int nLength;

            nLength = Convert.ToInt32(fs.Length);
            // Read the contents of the file into the array.
            bytes = br.ReadBytes(nLength);
            // Allocate some unmanaged memory for those bytes.
            pUnmanagedBytes = Marshal.AllocCoTaskMem(nLength);
            // Copy the managed byte array into the unmanaged array.
            Marshal.Copy(bytes, 0, pUnmanagedBytes, nLength);
            // Send the unmanaged bytes to the printer.
            bSuccess = SendBytesToPrinter(szPrinterName, szFileName, pUnmanagedBytes, nLength);
            // Free the unmanaged memory that you allocated earlier.
            Marshal.FreeCoTaskMem(pUnmanagedBytes);
            return bSuccess;
        }

        public static bool SendStringToPrinter(string szPrinterName, string szDocumentName, string szString, string printDataType = "RAW")
        {
            IntPtr pBytes;
            Int32 dwCount;
            // How many characters are in the string?
            dwCount = szString.Length;
            // Assume that the printer is expecting ANSI text, and then convert
            // the string to ANSI text.
            pBytes = Marshal.StringToCoTaskMemAnsi(szString);
            // Send the converted ANSI string to the printer.
            bool printRequestStatus = SendBytesToPrinter(szPrinterName, szDocumentName, pBytes, dwCount, printDataType);
            Marshal.FreeCoTaskMem(pBytes);
            return printRequestStatus;
        }

        public static List<Dictionary<string, object>> GetPrinters(string options = "")
        {
            Dictionary<string, object> properties = new Dictionary<string, object>();
            List<Dictionary<string, object>> printers = new List<Dictionary<string, object>>();

            //string printerName = "YourPrinterName";
            //string query = string.Format("SELECT * from Win32_Printer WHERE Name LIKE '%{0}'", printerName);

            string query = string.Format("SELECT * from Win32_Printer");
            
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(query);
            ManagementObjectCollection coll = searcher.Get();

            foreach (ManagementObject printer in coll)
            {
                properties = new Dictionary<string, object>();

                if (options.ToLower() == "extended")
                {
                    // Return all properties
                    foreach (PropertyData property in printer.Properties)
                    {
                        properties.Add(property.Name, property.Value);
                    }    
                }
                else
                {
                    // Return only a selected set of properties
                    properties.Add("Name", printer.Properties["Name"].Value);
                    
                    properties.Add("PortName", printer.Properties["PortName"].Value);
                    properties.Add("DriverName", printer.Properties["DriverName"].Value);
                    properties.Add("DeviceID", printer.Properties["DeviceID"].Value);
                    properties.Add("Shared", printer.Properties["Shared"].Value);
                    properties.Add("PrintJobDataType", printer.Properties["PrintJobDataType"].Value);
                    properties.Add("Local", printer.Properties["Local"].Value);
                    properties.Add("SpoolEnabled", printer.Properties["SpoolEnabled"].Value);
                    properties.Add("Location", printer.Properties["Location"].Value);
                    properties.Add("Description", printer.Properties["Description"].Value);

                }

                printers.Add(properties);
            }

            // Old method get printer name
            //List<string> printers = new List<string>();
            //foreach (string printer in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            //{
            //    printers.Add(printer);
            //}


            return printers;
        }

        public static List<Dictionary<string, object>> GetPrintQueue()
        {
            Dictionary<string, object> properties = new Dictionary<string, object>();
            Dictionary<string, object> jobs = new Dictionary<string, object>();
            List<Dictionary<string, object>> printers = new List<Dictionary<string, object>>();
            List<object> jobsQueue = new List<object>();
            
            EnumeratedPrintQueueTypes[] enumerationFlags = {EnumeratedPrintQueueTypes.Local,
                                                EnumeratedPrintQueueTypes.Shared};

            LocalPrintServer printServer = new LocalPrintServer();

            //Use the enumerationFlags to filter out unwanted print queues
            PrintQueueCollection printQueuesOnLocalServer = printServer.GetPrintQueues(enumerationFlags);

            foreach (PrintQueue printeQueue in printQueuesOnLocalServer)
            {
                properties = new Dictionary<string, object>();

                printeQueue.Refresh();
                PrintJobInfoCollection jobsQueued = printeQueue.GetPrintJobInfoCollection();
                jobsQueue.Clear();
                foreach (PrintSystemJobInfo job in jobsQueued)
                {
                    jobs = new Dictionary<string, object>();
                    // Since the user may not be able to articulate which job is problematic, 
                    // present information about each job the user has submitted. 
                    jobs.Add("Job", job.JobName);
                    jobs.Add("Name", job.Name);
                    jobs.Add("ID", job.JobIdentifier);
                    jobs.Add("Priority", job.PositionInPrintQueue);
                    jobs.Add("SummitedOn", DateTime.Parse(job.TimeJobSubmitted.ToString()).ToString("yyyy-MM-dd hh:mm:ss"));
                    jobs.Add("IsPaused", job.IsPaused);
                    jobs.Add("IsPrinting", job.IsPrinting);
                    jobs.Add("IsDeleting", job.IsDeleting);
                    jobs.Add("IsInError", job.IsInError);
                    
                    jobsQueue.Add(jobs);

                }// end for each print job    
                
                properties.Add("Printer", printeQueue.Name);
                properties.Add("Location", printeQueue.Location);
                properties.Add("NumberOfJobs", printeQueue.NumberOfJobs);
                properties.Add("Jobs", jobsQueue);

                properties.Add("QueueAttributes", printeQueue.QueueAttributes);
                properties.Add("QueueDriver", printeQueue.QueueDriver.Name);
                properties.Add("QueuePort", printeQueue.QueuePort.Name);
                properties.Add("QueuePrintProcessor", printeQueue.QueuePrintProcessor.Name);
                properties.Add("QueueStatus", printeQueue.QueueStatus);

                properties.Add("IsDirect", printeQueue.IsDirect);
                properties.Add("PrintingIsCancelled", printeQueue.PrintingIsCancelled);
                properties.Add("IsRawOnlyEnabled", printeQueue.IsRawOnlyEnabled);
                properties.Add("IsWaiting", printeQueue.IsWaiting);
                properties.Add("IsQueued", printeQueue.IsQueued);
                properties.Add("IsProcessing", printeQueue.IsProcessing);
                properties.Add("IsPrinting", printeQueue.IsPrinting);
                properties.Add("IsBusy", printeQueue.IsBusy);
                properties.Add("IsNotAvailable", printeQueue.IsNotAvailable);
                properties.Add("IsOffline", printeQueue.IsOffline);
                properties.Add("IsPaused", printeQueue.IsPaused);
                properties.Add("IsIOActive", printeQueue.IsIOActive);
                properties.Add("IsOutOfPaper", printeQueue.IsOutOfPaper);
                properties.Add("IsOutOfMemory", printeQueue.IsOutOfMemory);
                properties.Add("IsPaperJammed", printeQueue.IsPaperJammed);
                properties.Add("IsManualFeedRequired", printeQueue.IsManualFeedRequired);
                properties.Add("IsDoorOpened", printeQueue.IsDoorOpened);
                properties.Add("IsTonerLow", printeQueue.IsTonerLow);
                properties.Add("IsOutputBinFull", printeQueue.IsOutputBinFull);
                properties.Add("NeedUserIntervention", printeQueue.NeedUserIntervention);
                properties.Add("IsInError", printeQueue.IsInError);

                printers.Add(properties);
            }
            
            return printers;
        }

        public static int SendCommandToPrinter(string command, string printer, string jobId = "", string documentName = "")
        {
            Dictionary<string, object> properties = new Dictionary<string, object>();
            List<Dictionary<string, object>> printers = new List<Dictionary<string, object>>();

            EnumeratedPrintQueueTypes[] enumerationFlags = {EnumeratedPrintQueueTypes.Local,
                                                EnumeratedPrintQueueTypes.Shared};

            LocalPrintServer printServer = new LocalPrintServer();

            //Use the enumerationFlags to filter out unwanted print queues
            PrintQueueCollection printQueuesOnLocalServer = printServer.GetPrintQueues(enumerationFlags);
            int counter = 0;
            foreach (PrintQueue printeQueue in printQueuesOnLocalServer)
            {
                properties = new Dictionary<string, object>();

                printeQueue.Refresh();
                PrintJobInfoCollection jobsQueued = printeQueue.GetPrintJobInfoCollection();
                if (printeQueue.Name.ToLower().Trim() == printer.ToLower().Trim())
                {
                    foreach (PrintSystemJobInfo job in jobsQueued)
                    {
                        if (((job.Name.ToLower().Trim() == documentName.ToLower().Trim() && !string.IsNullOrEmpty(documentName.ToLower().Trim()))
                            ||
                            job.JobIdentifier.ToString().ToLower().Trim() == jobId.ToLower().Trim())
                            ||
                            jobId.ToLower().Trim() == "*"
                            )
                        {
                            if (command.ToLower().Trim() == "purge")
                            {
                                job.Cancel();
                                counter++;
                            }
                            else if (command.ToLower().Trim() == "pause")
                            {
                                job.Pause();
                                counter++;
                            }
                            else if (command.ToLower().Trim() == "resume")
                            {
                                job.Resume();
                                counter++;
                            }
                            else
                            {
                                return 0;
                            }    
                        }
                    }    
                }
            }

            return counter;
        }
    }
}
