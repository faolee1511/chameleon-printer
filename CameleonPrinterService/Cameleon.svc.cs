using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.ServiceModel.Web;
using System.Text;
using System.Web.Caching;
using System.Web.Script.Serialization;

namespace CameleonPrinterService
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "Cameleon" in code, svc and config file together.
    public class Cameleon : ICameleon
    {
        public Stream Printers()
        {
            var printers = RawPrinterHelper.GetPrinters();
            var jss = new JavaScriptSerializer();
            string jsonClient = jss.Serialize(printers);
            WebOperationContext.Current.OutgoingResponse.ContentType = "application/json; charset=utf-8";

            return new MemoryStream(Encoding.UTF8.GetBytes(jsonClient));
        }

        public Stream PrintQueue()
        {
            var printers = RawPrinterHelper.GetPrintQueue();
            var jss = new JavaScriptSerializer();
            string jsonClient = jss.Serialize(printers);
            WebOperationContext.Current.OutgoingResponse.ContentType = "application/json; charset=utf-8";

            return new MemoryStream(Encoding.UTF8.GetBytes(jsonClient));
        }

        public Stream Command(System.IO.Stream pStream)
        {
            var jss = new JavaScriptSerializer();

            string value = "";
            string response = "";
            string command = "";
            string printer = "";
            string jobId = "*";
            string documentName = "*";

            using (var reader = new StreamReader(pStream, Encoding.UTF8))
            {
                value = reader.ReadToEnd();
            }

            var commandData = jss.Deserialize<Dictionary<string, string>>(value);

            // Make sure required command paramters are complete
            if (!commandData.ContainsKey("Command") && !commandData.ContainsKey("Printer"))
            {
                response = "Command requests requires a 'Command' and 'Printer' parameters.";
            }
            else
            {
                command = commandData["Command"];
                printer = commandData["Printer"];

                // Check for command validity
                if (!IsValidCommand(command))
                {
                    response = "'" + command.ToUpper().Trim() + "' command is not supported.";
                } 
                else if (!IsExistingPrinter(printer))
                {
                    response = "'" + printer.ToUpper().Trim() + "' is not found in the Cameleon Printer Service server.";
                }
                else
                {
                    // Get specified JobId or DocumentName
                    if (commandData.ContainsKey("JobId"))
                    {
                        jobId = commandData["JobId"];
                    }
                    else if (commandData.ContainsKey("DocumentName"))
                    {
                        documentName = commandData["DocumentName"];
                        if (jobId.Trim() == "*" && !string.IsNullOrEmpty(documentName.Trim()))
                        {
                            jobId = "";
                        }
                    }

                    // Check if printer parameter is valid
                    if (printer.Trim() == "")
                    {
                        response = "Command requires a 'Printer' parameter to have a value.";
                    }
                    else
                    {
                        int totalAffected = RawPrinterHelper.SendCommandToPrinter(command, printer, jobId, documentName);
                        WebOperationContext.Current.OutgoingResponse.ContentType = "application/json; charset=utf-8";

                        response = totalAffected + " print jobs are affected by '" + command.ToUpper() + "' on '" +
                                   printer.ToUpper() + "'.";
                    }
                }                
            }

            return new MemoryStream(Encoding.UTF8.GetBytes(@"{""Response"" : " + @"""" + response + @"""}"));
        }

        private bool IsValidCommand(string command)
        {
            bool response = false;

            switch (command.ToLower().Trim())
            {
                case "purge":
                    response =  true;
                    break;
                case "pause":
                    response = true;
                    break;
                case "resume":
                    response = true;
                    break;
            }

            return response;
        }

        public Stream Print(System.IO.Stream pStream)
        {
            var jss = new JavaScriptSerializer();
            
            string value = "";
            string response = "";
            string dataType = "RAW";
            string defaultDataType = "RAW";
            string printer = "";
            string name = "";
            string data = "";

            using (var reader = new StreamReader(pStream, Encoding.UTF8))
            {
                value = reader.ReadToEnd();
            }

            var printData = jss.Deserialize<Dictionary<string, string>>(value);

            if (printData.ContainsKey("Printer") && printData.ContainsKey("Name") && printData.ContainsKey("Data"))
            {
                printer = printData["Printer"];
                name = printData["Name"];
                data = printData["Data"];

                //Check for Printer Name validity and DataType default

                if (printData.ContainsKey("DataType"))
                {
                    dataType = printData["DataType"].Trim();
                }

                var printers = RawPrinterHelper.GetPrinters();
                bool isValidPrinter = false;
                foreach (Dictionary<string, object> p in printers)
                {
                    if (p["Name"].ToString() == printer)
                    {
                        defaultDataType = p["PrintJobDataType"].ToString();
                        isValidPrinter = true;
                        break;
                    }
                }

                if (isValidPrinter)
                {
                    if (!IsValidDataType(dataType))
                    {
                        dataType = defaultDataType;
                    }
                    bool successPrint = RawPrinterHelper.SendStringToPrinter(printer, name, data, dataType);
                    if (successPrint)
                    {
                        response = "Print request sent to " + printer + " as " + dataType + " data.";
                    }
                    else
                    {
                        response = "An error occured during print request: " + Marshal.GetLastWin32Error();
                    }
                    
                    
                }
                else
                {
                    response = printer + " is not found in the Cameleon Printer Service server.";
                }
            }
            else
            {
                response = "Print requests requires a 'Printer', 'Name' and 'Data' parameters.";
            }

            return new MemoryStream(Encoding.UTF8.GetBytes(@"{""Response"" : " + @"""" + response + @"""}"));
        }

        private bool IsValidDataType(string datatType)
        {
            bool response = false;

            switch (datatType.ToUpper().Trim())
            {
                case "RAW":
                    response = true;
                    break;
                case "TEXT":
                    response = true;
                    break;
                case "RAW [FF APPENDED]":
                    response = true;
                    break;
                case "RAW [FF AUTO]":
                    response = true;
                    break;
                case "NT EMF 1.003":
                    response = true;
                    break;
                case "NT EMF 1.006":
                    response = true;
                    break;
                case "NT EMF 1.007":
                    response = true;
                    break;
                case "NT EMF 1.008":
                    response = true;
                    break;
            }

            return response;
        }

        private bool IsExistingPrinter(string printer)
        {
            var printers = RawPrinterHelper.GetPrinters();
            bool response = false;
            foreach (Dictionary<string, object> p in printers)
            {
                if (p["Name"].ToString() == printer)
                {
                    response = true;
                    break;
                }
            }

            return response;
        }
    }
}
