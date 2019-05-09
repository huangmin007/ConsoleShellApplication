using ConsoleShell;
using System;

namespace MTS.Shell
{
    class Program
    {
        static void Main(String[] args)
        {
            PowerPointShell Shell = new PowerPointShell()
            {
                Title = "PowerPoint Controller",
                InputHeader = "PPTC",

                ServicePort = 6101,
                ServiceModeAllowInput = true,
            };
            ConsoleApplication.Run(Shell, args);
        }
    }
}
