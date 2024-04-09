using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using DAL;

namespace Print_General
{
    public class PrintOperation
    {


        IDataLayer dal;
        private string workstationId;
        private string Type;
        private int Port = 9100;
        private string IP = "";
        private ReportStation reportStation;

        public PrintOperation(IDataLayer dal, dynamic workstationIdD, string Type)
        {
            // TODO: Complete member initialization
            this.dal = dal;
            this.workstationId = workstationIdD.ToString();
            this.Type = Type;
            Workstation ws = dal.getWorkStaitionById(long.Parse(workstationId));
            reportStation = dal.getReportStationByWorksAndType(ws.NAME, Type);
            string printerName = "";
            //string ip = GetIp(printerName);

            if (reportStation != null)
            {
                if (reportStation.Destination != null)
                {
                    IP = reportStation.Destination.ManualIP;
                }
                if (reportStation.Destination != null && reportStation.Destination.RawTcpipPort != null)
                {
                    Port = (int)reportStation.Destination.RawTcpipPort;
                }
            }
        }


        string removeBadChar(string ip)
        {
            string ret = "";
            foreach (var c in ip)
            {
                int ascii = (int)c;
                if ((ascii >= 48 && ascii <= 57) || ascii == 44 || ascii == 46)
                    ret += c;
            }
            return ret;
        }
        public string GetIp(string printerName)
        {
            string query = string.Format("SELECT * from Win32_Printer WHERE Name LIKE '%{0}'", printerName);
            string ret = "";
            var searcher = new ManagementObjectSearcher(query);
            var coll = searcher.Get();
            foreach (ManagementObject printer in coll)
            {
                foreach (PropertyData property in printer.Properties)
                {
                    if (property.Name == "PortName")
                    {
                        ret = property.Value.ToString();
                    }
                }
            }
            return ret;
        }
        public string ReverseString(string s)
        {
            var str = s;
            string[] strsubs = s.Split(Convert.ToChar(" "));
            var newstr = "";
            string substr = "";
            int i;
            int c = strsubs.Count();
            for (i = 0; i < c; ++i)
            {
                substr = strsubs[i];
                if (HasHebrewChar(strsubs[i]))
                {
                    substr = Reverse(substr);
                }

                newstr += substr + " ";
            }
            return newstr;
        }
        public string ManipulateHebrew(string value)
        {
            string retval = "";
            if (HasHebrewChar(value))
            {
                var split = ReverseString(value).Split(' ');
                split.Reverse();
                foreach (string s in split)
                {
                    retval = s + " " + retval;
                }
            }
            else
            {
                retval = value;
            }
            return retval;
        }
        private static string Reverse(string s)
        {
            char[] arr = s.ToCharArray();
            Array.Reverse(arr);
            return new string(arr);
        }

        public bool HasHebrewChar(string value)
        {
            return value.ToCharArray().Any(x => (x <= 'ת' && x >= 'א'));
        }

        public void Print(string ZPLString)
        {
            if (reportStation == null)
            {
                MessageBox.Show("לא הוגדרה תחנה");
            }
            else
            {


                try
                {

                    // Open connection
                    System.Net.Sockets.TcpClient client = new System.Net.Sockets.TcpClient();
                    client.Connect(IP, Port);
                    // Write ZPL String to connection
                    StreamWriter writer = new StreamWriter(client.GetStream());
                    writer.Write(ZPLString);
                    writer.Flush();
                    // Close Connection
                    writer.Close();
                    client.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.InnerException.Message);
                }
            }
        }
    }
}
