using System;
using ConsoleShell;

namespace MTS.Shell
{
    class Program
    {
        /// <summary>
        /// https://blog.csdn.net/pengze0902/article/details/53366287
        /// ManagementClass
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            WMIShell Shell = new WMIShell()
            {
                Title = "Windows Management Information",
                InputHeader = "wmi",

                ServicePort = 6103,
                
                ServiceModeAllowInput = true,
            };
            ConsoleApplication.Run(Shell, args);
           
        }
        
    }
}