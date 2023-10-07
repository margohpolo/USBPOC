using System.IO.Ports;
using System.Management;
using System.Text;

namespace USBPOC
{
    /// <summary>
    /// Service for getting all connected devices' info. Maybe not only USB anymore (TBC)...
    /// </summary>
    public class UsbService
    {
        /// <summary>
        /// Main constructor.
        /// </summary>
        public UsbService() { }

        /// <summary>
        /// PowerShell has a useful method for this.
        /// May not be the best approach for scale though.
        /// </summary>
        public void ExplorePnPEntities()
        {
            #region for testing
            //string psOutput = PowerShellHandler.RunCommand("systeminfo | more").Result;
            //string psOutput = PowerShellHandler.RunCommand("Get-WmiObject -class Win32_PnPEntity").Result;
            #endregion

            string psOutput = string.Empty;

            psOutput = PowerShellHandler.RunCommand("Get-PnpDevice").Result;

            Console.WriteLine(psOutput);
        }

        /// <summary>
        /// Only checks existing COM port connections, if any. Does not return all USB connections.
        /// Also does not provide much information on the connected devices.
        /// </summary>
        public void CheckPorts()
        {
            StringBuilder sb = new();

            sb.AppendLine("SerialPort.GetPortNames(): ");

            string[] portNames = SerialPort.GetPortNames();

            if (portNames.Length.Equals(0))
            {
                sb.AppendLine("No port connections found.");
            }
            else
            {
                sb.AppendLine("Ports found:");

                foreach (string name in portNames)
                {
                    sb.AppendLine($"{name}");
                }
            }

            Console.WriteLine(sb.ToString());

            PrintSeparatorInConsole();
        }


        /// <summary>
        /// Prints to console all 3 methods. Used this to see what comes up.
        /// </summary>
        public void RunAllManagementObjectSearches()
        {
            GetAllUsbHubObjects();
            PrintSeparatorInConsole();
            GetAllPnpEntities();
            PrintSeparatorInConsole();
            GetAllSerialPortInfo();
        }

        private void GetAllUsbHubObjects() => SearchAndReturnInfo("USBHub");

        private void GetAllPnpEntities() => SearchAndReturnInfo("PnPEntity");

        private void GetAllSerialPortInfo() => SearchAndReturnInfo("SerialPort");

        private void SearchAndReturnInfo(string objectName)
        {
            StringBuilder sb = new();

            sb.AppendLine($"{objectName}:\n");

            using ManagementObjectSearcher searcher = new ManagementObjectSearcher(QueryBuilder(objectName));

            ManagementObjectCollection searchResults = searcher.Get();

            if (searchResults.Count.Equals(0))
            {
                sb.AppendLine("No items found.");
            }
            else
            {
                foreach (ManagementBaseObject? item in searchResults)
                {
                    sb.AppendLine($"{item.Properties["Name"].Value}  | {item.Properties["Description"].Value}  | {item.Properties["Status"].Value}  | {item.Properties["StatusInfo"].Value}");
                }
            }


            Console.WriteLine(sb.ToString());
        }

        private static string QueryBuilder(string objName) => $"SELECT * FROM Win32_{objName}";

        private void PrintSeparatorInConsole()
        {
            StringBuilder sb = new();
            int i = 0;
            sb.Append("\n\n\n");
            while (i < 100)
            {
                sb.Append('=');
                i++;
            }
            sb.Append("\n\n\n");
            Console.WriteLine(sb.ToString());
        }

    }
}
