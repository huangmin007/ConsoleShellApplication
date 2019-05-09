using ConsoleShell.Service;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Xml;

namespace ConsoleShell
{
    /// <summary>
    /// 控制台启动运行模式
    /// </summary>
    public enum ConsoleStartMode
    {
        /// <summary>
        /// 默认运行模式
        /// </summary>
        Default,

        /// <summary>
        /// 连续输入模式
        /// </summary>
        Contine,

        /// <summary>
        /// 网络输入模式
        /// </summary>
        Service,

        /// <summary>
        /// 连续输入与网络输入模式
        /// </summary>
        ContineService,
    }

    /// <summary>
    /// 控制台输出字符格式
    /// </summary>
    public enum ConsoleOutputFormat
    {
        /// <summary>
        /// 默认格式
        /// </summary>
        DEFAULT,

        /// <summary>
        /// JSON字符格式
        /// </summary>
        JSON,

        /// <summary>
        /// XML字符格式
        /// </summary>
        XML,
    }

    /// <summary>
    /// 控制台应用程序接口
    /// </summary>
    public interface IConsoleApplication
    {
        void Help();

        Boolean OutputFormat(String args);

        void Sleep(String args);

        void RunProgram();

        void QuitProgram();

        void StartService(String args);

        void StopService();

        void Version();
    }


    /// <summary>
    /// 控制台程序 Ctrl+C 的代理
    /// </summary>
    /// <param name="CtrlType"></param>
    /// <returns></returns>
    public delegate bool ControlCtrlDelegate(int CtrlType);

    /// <summary>
    /// 控制台程序
    /// </summary>
    public partial class ConsoleApplication : IConsoleApplication
    {
        //protected static readonly IConsoleApplication Instance;
        
        private ConsoleStartMode _StartMode = ConsoleStartMode.Default;    
        
        public ConsoleStartMode StartMode
        {
            get { return _StartMode; }
        }    

        protected String _Author = "huangmin@spacecg.cn";
        protected String _CopyRight = "SpaceCG.CN";
        protected Version _Version = new Version(0, 0, 18, 0926);
        protected String _HelpDescription;
        protected ConsoleOutputFormat _OutputFormat = ConsoleOutputFormat.DEFAULT;

        /// <summary>
        /// 控制台输入头部字符显示
        /// </summary>
        public String InputHeader = "Input";

        /// <summary>
        /// 控制台标题
        /// </summary>
        public String Title = "Console Application";

        /// <summary>
        /// 默认的服务端口
        /// </summary>
        public int ServicePort = 6101;

        /// <summary>
        /// 服务模式下允许继续输入
        /// </summary>
        public Boolean ServiceModeAllowInput = true;

        /// <summary>
        /// ConsoleApplication
        /// </summary>
        public ConsoleApplication()
        {
            //Console.TreatControlCAsInput = true;
            CompilingChecked(this);            
        }

        
        [CommandLine("?", "显示帮助，与键入 -? -help 是一样的")]//，只能输出默认格式信息")]
        public void Help()
        {
            //是否需要打印 JSON XML 格式？？？？
            try
            {
                String ArgsName = "";
                CommandLineAttribute[] Attributes;
                StringBuilder CommandLines = new StringBuilder();

                Attributes = (CommandLineAttribute[])MethodBase.GetCurrentMethod().GetCustomAttributes(typeof(CommandLineAttribute), false);

                ArgsName = Attributes[0].Name;
                CommandLines.AppendLine(Attributes[0].ToString());

                MethodInfo[] Methods = this.GetType().GetMethods();                
                foreach (MethodInfo Method in Methods)
                {
                    if (Method.Name == MethodBase.GetCurrentMethod().Name) continue;

                    Attributes = (CommandLineAttribute[])Method.GetCustomAttributes(typeof(CommandLineAttribute), false);
                    if (Attributes != null && Attributes.Length > 0)
                    {
                        ArgsName += " | " + Attributes[0].Name;
                        CommandLines.AppendLine(Attributes[0].ToString());                        
                    }
                }

                FileInfo Info = new FileInfo(Process.GetCurrentProcess().MainModule.FileName);
                String fileName = Info.Name.ToLower().Replace(".exe", "");
                Console.WriteLine(String.Format("用法：{0}({1}) [{2}]", fileName, this.GetType().Name, ArgsName));
                Console.WriteLine(CommandLines.ToString());
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        #region Run Mode
        [CommandLine("-run", "进入持续输入模式")]
        public void RunProgram()
        {
            if (_StartMode != ConsoleStartMode.Default)
            {
                Console.WriteLine("Input Arguments Exception.");
                return;
            }

            Console.Title = Title;
            _StartMode = ConsoleStartMode.Contine;            

            while (_StartMode == ConsoleStartMode.Contine)
            {
                Console.Write("{0}>", InputHeader);

                String input = Console.ReadLine();
                if (String.IsNullOrWhiteSpace(input)) continue;

                String[] arguments = ConsoleApplication.ParserInput(input);
                if (arguments != null)  ConsoleApplication.Run(this, arguments);
            }
        }

        [CommandLine("-quit", "退出持续输入模式")]
        public void QuitProgram()
        {
            if (_StartMode != ConsoleStartMode.Contine)
            {
                Console.WriteLine("Input Arguments Exception.");
                return;
            }

            _StartMode = ConsoleStartMode.Default;
            Environment.Exit(0);
        }
        #endregion


        #region Service Mode
        private TCPServer TCPServer;
        private UDPServer UDPServer;
        private ControlCtrlDelegate ControlCtrl;

        [CommandLine("-start", "[int port]", "进入网络模式下持续输入模式，端口号空时使用默认端口")]
        public void StartService(String args)
        {
            if (_StartMode != ConsoleStartMode.Default || TCPServer != null || UDPServer != null)
            {
                Console.WriteLine("Input Arguments Exception.");
                return;
            }

            Console.Title = Title;
            _StartMode = ConsoleStartMode.Service;

            //安全退出处理
            ControlCtrl = new ControlCtrlDelegate(ConsoleExitHandler);
            ConsoleApplication.SetConsoleCtrlHandler(ControlCtrl, true);

            //服务模式启动后不可重复启动
            bool createdNew;            
            Mutex mutex = new Mutex(true, InputHeader + ".Service", out createdNew);

            if (!createdNew)
            {
                Console.WriteLine("Services Error 服务正在运行中 ...");
                MessageBox.Show("服务正在运行中...", "Error", MessageBoxButton.OK, MessageBoxImage.Warning, MessageBoxResult.OK);
                Console.WriteLine("Exiting ...");

                Thread.Sleep(100);
                Environment.Exit(0);
                return;
            }

            if (!String.IsNullOrWhiteSpace(args) && ConsoleApplication.Uint_Regex.IsMatch(args)) ServicePort = Convert.ToInt32(args);

            //独立线程运行服务
            Thread thread = new Thread(() =>
            {
                Console.Title = String.Format("{0}   Port:{1}", Title, ServicePort);
                Console.WriteLine("{0} Starting ... ", Title);

                TCPServer = new TCPServer();
                if (TCPServer.Start(ServicePort))
                {
                    TCPServer.ServerDataReceived += Server_DataReceived;
                    TCPServer.ServerStateChanged += Server_StateChanged;
                    TCPServer.ClientStateChanged += Server_ClientStateChanged;                    
                    Console.WriteLine("TCP Server 启功完成 Port:{0}，等待客户端连接 ...", ServicePort);
                }
                else
                {
                    TCPServer.Destroy();
                    TCPServer = null;
                    Console.WriteLine("TCP Server 启动失败，检查端口 {0} 是否被占用", ServicePort);
                }

                UDPServer = new UDPServer();
                if (UDPServer.Start(ServicePort))
                {
                    UDPServer.ServerDataReceived += Server_DataReceived;
                    UDPServer.ServerStateChanged += Server_StateChanged;
                    UDPServer.ClientStateChanged += Server_ClientStateChanged;
                    Console.WriteLine("UDP Server 启功完成 Port:{0}，等待客户端连接 ...", ServicePort);
                }
                else
                {
                    UDPServer.Destroy();
                    UDPServer = null;
                    Console.WriteLine("UDP Server 启动失败，检查端口 {0} 是否被占用", ServicePort);
                }
                Console.Write("{0}>", InputHeader);
            });
            thread.Name = Title;
            thread.Start();

            //可输入模式
            while (_StartMode == ConsoleStartMode.Service)
            {
                Console.Write("{0}>", InputHeader);
                String input = Console.ReadLine();
                if (String.IsNullOrWhiteSpace(input)) continue;

                String[] arguments = ConsoleApplication.ParserInput(input);                
                if (arguments != null) ConsoleApplication.Run(this, arguments);
            }
        }


        [CommandLine("-stop", "停止网络模式下持续输入模式")]
        public void StopService()
        {
            if (_StartMode != ConsoleStartMode.Service)
            {
                Console.WriteLine("Arguments Exception");
                return;
            }

            //Console.WriteLine();
            Console.WriteLine("正在退出服务 ...");
            _StartMode = ConsoleStartMode.Default;

            if (TCPServer != null)
            {
                TCPServer.ServerDataReceived -= Server_DataReceived;
                TCPServer.ServerStateChanged -= Server_StateChanged;
                TCPServer.ClientStateChanged -= Server_ClientStateChanged;
                Console.WriteLine(TCPServer.Destroy() ? "TCP Server 已停止服务 ..." : "TCP Server 服务停止失败 ...");

                TCPServer = null;
            }

            if (UDPServer != null)
            {
                UDPServer.ServerDataReceived -= Server_DataReceived;
                UDPServer.ServerStateChanged -= Server_StateChanged;
                UDPServer.ClientStateChanged -= Server_ClientStateChanged;
                Console.WriteLine(UDPServer.Destroy() ? "UDP Server 已停止服务 ..." : "UDP Server 服务停止失败 ...");

                UDPServer = null;
            }

            Thread.Sleep(500);
            Environment.Exit(0);
        }

        /// <summary>
        /// 服务广播数据
        /// </summary>
        /// <param name="data"></param>
        protected virtual void ServiceBroadcast(String data)
        {
            if (_StartMode != ConsoleStartMode.Service) return;
            
        }

        /// <summary>
        /// 控制台退出，安全处理
        /// </summary>
        /// <param name="CtrlType"></param>
        /// <returns></returns>
        private bool ConsoleExitHandler(int CtrlType)
        {
            /*
            switch (CtrlType)
            {
                case 0:
                    //Ctrl+C关闭
                    return true;

                case 2:
                    //按控制台关闭按钮 或 Alt+F4
                    return false;
            }
            */
            StopService();

            return false;
        }


        private void Server_ClientStateChanged(object sender, IntPtr connId, ServerState newState)
        {
            IServer Server = (IServer)sender;

            ushort port = 0;
            String address = "";
            if (Server.GetRemoteAddress(connId, ref address, ref port))
            {
                Console.WriteLine(String.Format("<{0}> Reomte Client {1}:{2} {3}", Server.GetType().Name, address, port, newState));
                //Console.Write("{0}>", InputHeader);
            }
        }

        private void Server_StateChanged(object sender, ServerState newState)
        {
            IServer Server = (IServer)sender;
            Console.WriteLine(String.Format("<{0}> Server State Changed:{1}", Server.GetType().Name, newState));
            //Console.Write("{0}>", InputHeader);
        }

        /// <summary>
        /// 数据接收处理
        /// <para>命令行应用设计是多命令排队处理，按设计就可能会出现多个返回值</para>
        /// <para>这里就完全不做处理了，由子类自已去处理 Socket 数据输出的问题，所以数据接收处理设计为虚函数，也可由子类完全自已解析参数</para>
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="connId"></param>
        /// <param name="data"></param>
        protected virtual void Server_DataReceived(object sender, IntPtr connId, byte[] data)
        {
            IServer Server = (IServer)sender;

            ushort port = 0;
            String address = "";
            if(Server.GetRemoteAddress(connId, ref address, ref port))
                Console.WriteLine(String.Format("<{0}> Receive Reomte Client {1}:{2} Data.", Server.GetType().Name, address, port));

            String input = Encoding.Default.GetString(data);
            Console.Write("{0}>", InputHeader);
            Console.WriteLine(input);

            String[] arguments = ConsoleApplication.ParserInput(input);
            if(arguments != null)ConsoleApplication.Run(this, arguments);

            //Console.Write("{0}>", InputHeader);
        }
        #endregion

        [CommandLine("of", "(enum)", "输出格式：0:default默认格式，1:json格式，2:xml格式。示例：-of 1 或 -of json")]
        public Boolean OutputFormat(String args)
        {
            if (String.IsNullOrWhiteSpace(args))
            {
                String ErrorMessage = "参数输入错误，不能为空";
                WriteError(ErrorMessage, _OutputFormat, MethodBase.GetCurrentMethod());
                return false;
            }

            try
            {
                _OutputFormat = Uint_Regex.IsMatch(args) ? (ConsoleOutputFormat)Convert.ToInt32(args) : (ConsoleOutputFormat)Enum.Parse(typeof(ConsoleOutputFormat), args, true);
            }
            catch (Exception e)
            {
                WriteError(e.Message, _OutputFormat, MethodBase.GetCurrentMethod());
                return false;
            }

            return true;
        }

        [CommandLine("sp", "(int ms)", "将当前线程挂起指定的时间(ms)")]
        public void Sleep(String args)
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
                int millsecondsTimeout = Convert.ToInt32(args);
                Thread.Sleep(millsecondsTimeout);
            }
            catch (Exception e)
            {
                WriteError(e.Message, _OutputFormat, MethodBase.GetCurrentMethod());
            }
        }

        [CommandLine("v", "控制台程序版本信息")]
        public void Version()
        {
            JObject Jobject = new JObject();

            if (!String.IsNullOrWhiteSpace(Title)) Jobject["Title"] = Title;
            if (!String.IsNullOrWhiteSpace(_Author)) Jobject["Author"] = _Author;
            if (!String.IsNullOrWhiteSpace(_CopyRight)) Jobject["CopyRight"] = _CopyRight;
            if (_Version != null) Jobject["Version"] = _Version.ToString();

            WriteLine(Jobject, _OutputFormat);
        }
    }

    /// <summary>
    /// 控制台应用程序
    /// </summary>
    public partial class ConsoleApplication
    {
        /// <summary>
        /// 正整数正则表达式，匹配正整数
        /// </summary>
        public static readonly Regex Uint_Regex = new Regex(@"^[1-9]\d*$");

        #region Private Static Function
        /// <summary>
        /// 编译检查重复命令行
        /// </summary>
        /// <param name="Shell"></param>
        private static void CompilingChecked(IConsoleApplication Shell)
        {
            CommandLineAttribute[] Attributes;
            MethodInfo[] Methods = Shell.GetType().GetMethods();
            Dictionary<String, MethodInfo> MethodInfos = new Dictionary<string, MethodInfo>(Methods.Length);

            foreach (MethodInfo Method in Methods)
            {
                Attributes = (CommandLineAttribute[])Method.GetCustomAttributes(typeof(CommandLineAttribute), false);
                if (Attributes.Length <= 0) continue;

                if (MethodInfos.ContainsKey(Attributes[0].Name) && MethodInfos[Attributes[0].Name].Name != Method.Name)
                    throw new NotSupportedException(String.Format("命令行设计重复，Type:{0} CommandLine:{1} Function:{2} {3}", Shell.GetType().ReflectedType, Attributes[0].Name, Method.Name, MethodInfos[Attributes[0].Name].Name));
                else
                    MethodInfos.Add(Attributes[0].Name, Method);
            }

            MethodInfos.Clear();
            MethodInfos = null;
        }


        /// <summary>
        /// 通过反射的方式来获取 IConsoleApplication 的方法
        /// </summary>
        /// <param name="type"></param>
        /// <param name="Name"></param>
        /// <param name="CommandLine"></param>
        /// <returns></returns>
        private static MethodInfo GetPublicMethod(Type type, String Name, ref CommandLineAttribute CommandLine)
        {
            Name = Name.ToLower();
            CommandLineAttribute[] Attributes;
            MethodInfo[] Methods = type.GetMethods();//多个返回的处理，需要比对参数

            foreach (MethodInfo Method in Methods)
            {
                //Method.GetGenericArguments();
                Attributes = (CommandLineAttribute[])Method.GetCustomAttributes(typeof(CommandLineAttribute), false);
                if (Attributes.Length <= 0) continue;

                CommandLine = Attributes[0];
                if (Attributes[0].Name.ToLower() == Name)
                {
                    return Method;
                }
            }

            return null;
        }
        #endregion


        #region Static Function
        /// <summary>
        /// Console Ctrl+C 的处理
        /// </summary>
        /// <param name="HandlerRoutine"></param>
        /// <param name="Add"></param>
        /// <returns></returns>
        [DllImport("kernel32.dll")]
        public static extern bool SetConsoleCtrlHandler(ControlCtrlDelegate HandlerRoutine, bool Add);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Shell"></param>
        /// <param name="args"></param>
        public static void Run(ConsoleApplication Shell, String[] args)
        {
            MethodInfo Method = null;
            Type type = Shell.GetType();
            
            //Print Help Info
            if (args.Length == 0)
            {
                Shell.Help();
                return;
            }

            int i = 0;
            String ArgsName = "";
            CommandLineAttribute Help = null;

            while (i < args.Length)
            {
                ArgsName = args[i++];
                Method = GetPublicMethod(type, ArgsName, ref Help);

                //函数没找到，输入错误
                if (Method == null)
                {
                    Shell.WriteError(String.Format("参数 {0} 输入错误", ArgsName), ConsoleOutputFormat.JSON);
                    return;
                }

                //如果下一个参数输入为 "?" 或 "help" 则打印该函数的帮助信息
                //并结束进程的执行
                if (i < args.Length && (args[i] == "?" || args[i].ToLower() == "help"))
                {
                    Console.WriteLine(Help.ToString());
                    return;
                }

                //函数本身没有参数，则执行
                if (Method.GetParameters().Length == 0)
                {
                    Method.Invoke(Shell, null);
                    continue;
                }

                //按函数的参数数量提取参数，以及下一个参数为函数则跳出参数提取
                int Length = Method.GetParameters().Length;
                Object[] Parameters = new Object[Length];
                for (int j = 0; j < Length; j++)
                {
                    if (i >= args.Length) break;
                    //下一个参数为函数，则不在提取
                    if (args[i].IndexOf(CommandLineAttribute.DefaultHeaderMark) == 0) break;

                    Parameters[j] = args[i++];
                }

                Method.Invoke(Shell, Parameters);
            }
        }

        /// <summary>
        /// 控制台输入解析
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        public static String[] ParserInput(String args)
        {
            if (String.IsNullOrWhiteSpace(args)) return null;
            if (args.IndexOf('\"') == -1)   return args.Split(' ');

            int i = 0;
            int index = 0;
            List<String> list = new List<String>(32);

            list.Add("");
            bool qmark = false;

            while(i < args.Length)
            {
                if (args[i] == ' ' && !qmark)
                {
                    i++;
                    index++;
                    list.Add("");
                    continue;
                }

                if (args[i] == '\"')
                {
                    i++;
                    qmark = !qmark;
                    continue;
                }

                list[index] += args[i++];
            }

            //Console.WriteLine(list.Count);
            //for(i = 0; i < list.Count; i ++)    Console.WriteLine(list[i]);

            return list.ToArray();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Jobject"></param>
        /// <param name="Format"></param>
        /// <returns></returns>
        public static String ConvertToString(JObject Jobject, ConsoleOutputFormat Format)
        {
            switch (Format)
            {
                case ConsoleOutputFormat.DEFAULT:
                    StringBuilder builder = new StringBuilder();
                    foreach (KeyValuePair<String, JToken> kvp in Jobject)
                        builder.AppendLine(String.Format("{0}:{1}", kvp.Key, kvp.Value));
                    return builder.ToString().TrimEnd();

                case ConsoleOutputFormat.JSON:
                    return Jobject.ToString();

                case ConsoleOutputFormat.XML:
                    XmlDocument Doc = (XmlDocument)JsonConvert.DeserializeXmlNode(Jobject.ToString(), "Root", true);
                    StringBuilder Output = new StringBuilder();

                    Doc.Save(XmlWriter.Create(Output, new XmlWriterSettings()
                    {
                        Indent = true,
                        IndentChars = "  ",
                        OmitXmlDeclaration = true,
                    }));

                    return Output.ToString();
            }

            return "";
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Jobject"></param>
        /// <param name="Format"></param>
        public void WriteLine(JObject Jobject, ConsoleOutputFormat Format)
        {
            Console.WriteLine(ConvertToString(Jobject, Format));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Message"></param>
        /// <param name="Format"></param>
        public void WriteLine(String Message, ConsoleOutputFormat Format)
        {
            JObject Jobject = new JObject()
            {
                ["Message"] = Message,
            };
            Console.WriteLine(ConvertToString(Jobject, Format));
        }

        /// <summary>
        /// write error
        /// </summary>
        /// <param name="Message"></param>
        /// <param name="Format"></param>
        public void WriteError(String Message, ConsoleOutputFormat Format)
        {
            WriteLine(new JObject()
            {
                ["Error"] = Message,
            }, Format);

            if(StartMode == ConsoleStartMode.Default)   Environment.Exit(0);
        }

        /// <summary>
        /// write error
        /// </summary>
        /// <param name="Message"></param>
        /// <param name="Format"></param>
        /// <param name="Method"></param>
        public void WriteError(String Message, ConsoleOutputFormat Format, MethodBase Method)
        {
            String hn = "";

            if (Method.IsPublic && Method.IsStatic)
            {
                CommandLineAttribute[] Attributes = (CommandLineAttribute[])Method.GetCustomAttributes(typeof(CommandLineAttribute), false);
                if (Attributes != null && Attributes.Length > 0) hn = Attributes[0].Name + " ";
            }

            WriteLine(new JObject()
            {
                ["Error"] = hn + Message,
            }, Format);

            if (StartMode == ConsoleStartMode.Default) Environment.Exit(0);
        }
        #endregion
        
    }
}
