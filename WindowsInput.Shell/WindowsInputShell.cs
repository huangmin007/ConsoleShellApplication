using ConsoleShell;
using Newtonsoft.Json.Linq;
using System;
using System.Reflection;
using System.Runtime.InteropServices;
using WindowsInput;
using WindowsInput.Native;

namespace MTS.Shell
{
    [StructLayout(LayoutKind.Sequential)]
    public struct POINT
    {
        public int X;
        public int Y;

        public POINT(int x, int y)
        {
            this.X = x;
            this.Y = y;
        }
    }


    /// <summary>
    /// Windows Input
    /// </summary>
    public partial class WindowsInput:ConsoleApplication
    {
        [DllImport("user32.dll")]
        static extern bool GetCursorPos(out POINT lpPoint);

        /// <summary>
        /// 输入模拟对象。如果重复执行，则不需要重新创建
        /// </summary>
        private InputSimulator InputSimulator = null;
        
        /// <summary>
        /// 
        /// </summary>
        public WindowsInput()
        {
            ServicePort = 6102;
            InputSimulator = new InputSimulator();
            _Author = "hm-secret@163.com";
        }
        

        [CommandLine("kd", "(enum)", "模拟输入键盘按下键，参数为十进制键值或是键的枚举字符串，参考：-kvs")]
        public void KeyDown(String args)
        {
            if (String.IsNullOrWhiteSpace(args))
            {
                String ErrorMessage = "参数输入错误，不能为空";
                WriteError(ErrorMessage, _OutputFormat, MethodBase.GetCurrentMethod());
                return;
            }

            try
            {
                VirtualKeyCode KeyCode = Uint_Regex.IsMatch(args) ? (VirtualKeyCode)Convert.ToInt32(args) : (VirtualKeyCode)Enum.Parse(typeof(VirtualKeyCode), args, true);

                if (InputSimulator == null) InputSimulator = new InputSimulator();
                InputSimulator.Keyboard.KeyDown(KeyCode);
            }
            catch (Exception e)
            {
                WriteError(e.Message, _OutputFormat, MethodBase.GetCurrentMethod());
            }
        }


        [CommandLine("ku", "(enum)", "模拟输入键盘释放键")]
        public void KeyUp(String args)
        {
            //if (String.IsNullOrWhiteSpace(args)) return;
            if (String.IsNullOrWhiteSpace(args))
            {
                String ErrorMessage = "参数输入错误，不能为空";
                WriteError(ErrorMessage, _OutputFormat, MethodBase.GetCurrentMethod());
                return;
            }

            try
            {
                VirtualKeyCode KeyCode = Uint_Regex.IsMatch(args) ? (VirtualKeyCode)Convert.ToInt32(args) : (VirtualKeyCode)Enum.Parse(typeof(VirtualKeyCode), args, true);

                if (InputSimulator == null) InputSimulator = new InputSimulator();
                InputSimulator.Keyboard.KeyUp(KeyCode);
            }
            catch (Exception e)
            {
                WriteError(e.Message, _OutputFormat, MethodBase.GetCurrentMethod());
            }
        }


        [CommandLine("kp", "(enum)", "模拟输入键盘按压并释放键")]
        public void KeyPress(String args)
        {
            //if (String.IsNullOrWhiteSpace(args)) return;
            if (String.IsNullOrWhiteSpace(args))
            {
                String ErrorMessage = "参数输入错误，不能为空";
                WriteError(ErrorMessage, _OutputFormat, MethodBase.GetCurrentMethod());
                return;
            }

            try
            {
                VirtualKeyCode KeyCode = Uint_Regex.IsMatch(args) ? (VirtualKeyCode)Convert.ToInt32(args) : (VirtualKeyCode)Enum.Parse(typeof(VirtualKeyCode), args, true);

                if (InputSimulator == null) InputSimulator = new InputSimulator();
                InputSimulator.Keyboard.KeyPress(KeyCode);
            }
            catch (Exception e)
            {
                WriteError(e.Message, _OutputFormat, MethodBase.GetCurrentMethod());
            }
        }


        [CommandLine("ks", "(keys)", "模拟输入组合键。\r\n例如：-ks CONTROL+VK_C 或 -ks CONTROL|LSHIFT+VK_A")]
        public void ModifiedKeyStroke(String args)
        {
            //if (String.IsNullOrWhiteSpace(args)) return;
            if (String.IsNullOrWhiteSpace(args))
            {
                String ErrorMessage = "参数输入错误，不能为空";
                WriteError(ErrorMessage, _OutputFormat, MethodBase.GetCurrentMethod());
                return;
            }

            try
            {
                String[] values = args.Split('+')[0].Split('|');
                VirtualKeyCode[] modifierKeyCode = new VirtualKeyCode[values.Length];
                for (int i = 0; i < values.Length; i++)
                    modifierKeyCode[i] = Uint_Regex.IsMatch(values[i]) ? (VirtualKeyCode)Convert.ToInt32(values[i]) : (VirtualKeyCode)Enum.Parse(typeof(VirtualKeyCode), values[i], true);

                values = args.Split('+')[1].Split('|');
                VirtualKeyCode[] keyCode = new VirtualKeyCode[values.Length];
                for (int i = 0; i < values.Length; i++)
                    keyCode[i] = Uint_Regex.IsMatch(values[i]) ? (VirtualKeyCode)Convert.ToInt32(values[i]) : (VirtualKeyCode)Enum.Parse(typeof(VirtualKeyCode), values[i], true);

                if (InputSimulator == null) InputSimulator = new InputSimulator();
                InputSimulator.Keyboard.ModifiedKeyStroke(modifierKeyCode, keyCode);
            }
            catch (Exception e)
            {
                WriteError(e.Message, _OutputFormat, MethodBase.GetCurrentMethod());
            }
        }


        [CommandLine("te", "(str)", "模拟输入文本字符")]
        public void TextEntry(String args)
        {
            if (String.IsNullOrWhiteSpace(args))
            {
                String ErrorMessage = "参数输入错误，不能为空";
                WriteError(ErrorMessage, _OutputFormat, MethodBase.GetCurrentMethod());
                return;
            }

            try
            {
                if (InputSimulator == null) InputSimulator = new InputSimulator();
                InputSimulator.Keyboard.TextEntry(args);
            }
            catch (Exception e)
            {
                WriteError(e.Message, _OutputFormat, MethodBase.GetCurrentMethod());
            }
        }


        [CommandLine("kvs", "键盘键值参考")]
        public void KeyValues()
        {
            Type type = typeof(VirtualKeyCode);
            String[] Names = type.GetEnumNames();

            JArray Keys = new JArray();
            foreach (String Name in Names)
            {
                Keys.Add(new JObject()
                {
                    ["Value"] = (int)Enum.Parse(type, Name),
                    ["Name"] = Name,
                });
            }

            WriteLine(new JObject()
            {
                ["Keys"] = Keys,
            }, _OutputFormat);
        }


        [CommandLine("mld", "模拟输入鼠标左键按下")]
        public void MouseLeftDown()
        {
            try
            {
                if (InputSimulator == null) InputSimulator = new InputSimulator();
                InputSimulator.Mouse.LeftButtonDown();
            }
            catch (Exception e)
            {
                WriteError(e.Message, _OutputFormat, MethodBase.GetCurrentMethod());
            }
        }


        [CommandLine("mlu", "模拟输入鼠标左键弹起")]
        public void MouseLeftUp()
        {
            try
            {
                if (InputSimulator == null) InputSimulator = new InputSimulator();
                InputSimulator.Mouse.LeftButtonUp();
            }
            catch (Exception e)
            {
                WriteError(e.Message, _OutputFormat, MethodBase.GetCurrentMethod());
            }
        }


        [CommandLine("mlc", "模拟输入鼠标左键点击")]
        public void MouseLeftClick()
        {
            try
            {
                if (InputSimulator == null) InputSimulator = new InputSimulator();
                InputSimulator.Mouse.LeftButtonClick();
            }
            catch (Exception e)
            {
                WriteError(e.Message, _OutputFormat, MethodBase.GetCurrentMethod());
            }
        }


        [CommandLine("mldc", "模拟输入鼠标左键双击")]
        public void MouseLeftDoubleClick()
        {
            try
            {
                if (InputSimulator == null) InputSimulator = new InputSimulator();
                InputSimulator.Mouse.LeftButtonDoubleClick();
            }
            catch (Exception e)
            {
                WriteError(e.Message, _OutputFormat, MethodBase.GetCurrentMethod());
            }
        }


        [CommandLine("mrd", "模拟输入鼠标右键按下")]
        public void MouseRightDown()
        {
            try
            {
                if (InputSimulator == null) InputSimulator = new InputSimulator();
                InputSimulator.Mouse.RightButtonDown();
            }
            catch (Exception e)
            {
                WriteError(e.Message, _OutputFormat, MethodBase.GetCurrentMethod());
            }
        }


        [CommandLine("mru", "模拟输入鼠标右键弹起")]
        public void MouseRightUp()
        {
            try
            {
                if (InputSimulator == null) InputSimulator = new InputSimulator();
                InputSimulator.Mouse.RightButtonUp();
            }
            catch (Exception e)
            {
                WriteError(e.Message, _OutputFormat, MethodBase.GetCurrentMethod());
            }
        }


        [CommandLine("mrc", "模拟输入鼠标右键点击")]
        public void MouseRightClick()
        {
            try
            {
                if (InputSimulator == null) InputSimulator = new InputSimulator();
                InputSimulator.Mouse.RightButtonClick();
            }
            catch (Exception e)
            {
                WriteError(e.Message, _OutputFormat, MethodBase.GetCurrentMethod());
            }
        }


        [CommandLine("mrdc", "模拟输入鼠标右键双击")]
        public void MouseRightDoubleClick()
        {
            try
            {
                if (InputSimulator == null) InputSimulator = new InputSimulator();
                InputSimulator.Mouse.RightButtonDoubleClick();
            }
            catch (Exception e)
            {
                WriteError(e.Message, _OutputFormat, MethodBase.GetCurrentMethod());
            }
        }


        [CommandLine("mmb", "(point)", "模拟输入鼠标坐标相对偏移。\r\n示例：-mmb 100,100")]
        public void MoveMouseBy(String args)
        {
            if (String.IsNullOrWhiteSpace(args))
            {
                String ErrorMessage = "参数输入错误，不能为空";
                WriteError(ErrorMessage, _OutputFormat, MethodBase.GetCurrentMethod());
                return;
            }

            try
            {
                String[] Args = args.Split(',');
                if (Args.Length < 2) return;

                int pixelDeltaX = Convert.ToInt32(Args[0]);
                int pixelDeltaY = Convert.ToInt32(Args[1]);

                if (InputSimulator == null) InputSimulator = new InputSimulator();
                InputSimulator.Mouse.MoveMouseBy(pixelDeltaX, pixelDeltaY);
            }
            catch (Exception e)
            {
                WriteError(e.Message, _OutputFormat, MethodBase.GetCurrentMethod());
            }
        }


        [CommandLine("mmt", "(point)", "模拟输入鼠标坐标移动到绝对位置。\r\n示例：-mmt 100,100")]
        public void MoveMouseTo(String args)
        {
            if (String.IsNullOrWhiteSpace(args))
            {
                String ErrorMessage = "参数输入错误，不能为空";
                WriteError(ErrorMessage, _OutputFormat, MethodBase.GetCurrentMethod());
                return;
            }

            try
            {
                String[] Args = args.Split(',');
                if (Args.Length < 2) return;
                Console.WriteLine(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width);
                double absoluteX = Convert.ToDouble(Args[0]) / System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width * 0xFFFF;
                double absoluteY = Convert.ToDouble(Args[1]) / System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height * 0xFFFF;

                if (InputSimulator == null) InputSimulator = new InputSimulator();
                InputSimulator.Mouse.MoveMouseTo(absoluteX, absoluteY);
            }
            catch (Exception e)
            {
                WriteError(e.Message, _OutputFormat, MethodBase.GetCurrentMethod());
            }
        }

        public void GetDPI()
        {
            //获取DPI的两种方式
            /*
            using (ManagementClass mc = new ManagementClass("Win32_DesktopMonitor"))
            {
                using (ManagementObjectCollection moc = mc.GetInstances())
                {

                    int PixelsPerXLogicalInch = 0; // dpi for x
                    int PixelsPerYLogicalInch = 0; // dpi for y

                    foreach (ManagementObject each in moc)
                    {
                        PixelsPerXLogicalInch = int.Parse((each.Properties["PixelsPerXLogicalInch"].Value.ToString()));
                        PixelsPerYLogicalInch = int.Parse((each.Properties["PixelsPerYLogicalInch"].Value.ToString()));
                    }

                    Console.WriteLine("PixelsPerXLogicalInch:" + PixelsPerXLogicalInch.ToString());
                    Console.WriteLine("PixelsPerYLogicalInch:" + PixelsPerYLogicalInch.ToString());
                    Console.Read();
                }
            }
            */
            /*
            using (Graphics graphics = Graphics.FromHwnd(IntPtr.Zero))
            {
                float dpiX = graphics.DpiX;
                float dpiY = graphics.DpiY;
            }
            */
        }


        [CommandLine("hs", "(int)", "模拟输入水平滚动")]
        public void HorizontalScroll(String args)
        {
            if (String.IsNullOrWhiteSpace(args))
            {
                String ErrorMessage = "参数输入错误，不能为空";
                WriteError(ErrorMessage, _OutputFormat, MethodBase.GetCurrentMethod());
                return;
            }

            if (!Uint_Regex.IsMatch(args))
            {
                String ErrorMessage = "参数 " + args + " 格式输入错误";
                WriteError(ErrorMessage, _OutputFormat, MethodBase.GetCurrentMethod());
                return;
            }
            try
            {
                int scrollAmountInClicks = Convert.ToInt32(args);

                if (InputSimulator == null) InputSimulator = new InputSimulator();
                InputSimulator.Mouse.HorizontalScroll(scrollAmountInClicks);
            }
            catch (Exception e)
            {
                WriteError(e.Message, _OutputFormat, MethodBase.GetCurrentMethod());
            }
        }


        [CommandLine("vs", "(int)", "模拟输入垂直滚动")]
        public void VerticalScroll(String args)
        {
            if (String.IsNullOrWhiteSpace(args))
            {
                String ErrorMessage = "参数输入错误，不能为空";
                WriteError(ErrorMessage, _OutputFormat, MethodBase.GetCurrentMethod());
                return;
            }
            if (!Uint_Regex.IsMatch(args))
            {
                String ErrorMessage = "参数 " + args + " 输入错误";
                WriteError(ErrorMessage, _OutputFormat, MethodBase.GetCurrentMethod());
                return;
            }

            try
            {
                int scrollAmountInClicks = Convert.ToInt32(args);

                if (InputSimulator == null) InputSimulator = new InputSimulator();
                InputSimulator.Mouse.VerticalScroll(scrollAmountInClicks);
            }
            catch (Exception e)
            {
                WriteError(e.Message, _OutputFormat, MethodBase.GetCurrentMethod());
            }
        }


        [CommandLine("mmtvd", "(point)", "模拟输入移动鼠标位置到虚拟桌面")]
        public void MoveMouseToPositionOnVirtualDesktop(String args)
        {
            if (String.IsNullOrWhiteSpace(args))
            {
                String ErrorMessage = "参数输入错误，不能为空";
                WriteError(ErrorMessage, _OutputFormat, MethodBase.GetCurrentMethod());
                return;
            }

            try
            {
                String[] Args = args.Split(',');
                if (Args.Length < 2) return;

                double absoluteX = Convert.ToDouble(Args[0]);
                double absoluteY = Convert.ToDouble(Args[1]);

                if (InputSimulator == null) InputSimulator = new InputSimulator();
                InputSimulator.Mouse.MoveMouseToPositionOnVirtualDesktop(absoluteX, absoluteY);
            }
            catch (Exception e)
            {
                WriteError(e.Message, _OutputFormat, MethodBase.GetCurrentMethod());
            }
        }


        [CommandLine("gmp", "获取鼠标相对屏幕的坐标")]
        public void GetMousePosition()
        {
            try
            {
                POINT Point;
                GetCursorPos(out Point);

                WriteLine(new JObject()
                {
                    ["X"] = Point.X,
                    ["Y"] = Point.Y,
                }, _OutputFormat);
            }
            catch (Exception e)
            {
                WriteError(e.Message, _OutputFormat, MethodBase.GetCurrentMethod());
            }
        }
        
        
    }

}