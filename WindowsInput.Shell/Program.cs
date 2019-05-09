using ConsoleShell;
using System;

namespace MTS.Shell
{
    class Program
    {
        static void Main(string[] args)
        {
            //Simulator
            WindowsInput Shell = new WindowsInput()
            {
                Title = "WindowInput Controller",
                InputHeader = "WINC",

                ServicePort = 6102,
                ServiceModeAllowInput = true,
            };
            ConsoleApplication.Run(Shell, args);
        }
    }
}
