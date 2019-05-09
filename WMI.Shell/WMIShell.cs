using ConsoleShell;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Management;
using System.Reflection;

namespace MTS.Shell
{

    /// <summary>
    /// Windows Input
    /// </summary>
    public partial class WMIShell : ConsoleApplication
    {
        /// <summary>
        /// 默认的 Query 语句
        /// </summary>
        protected Dictionary<String, String> DefaultQueryString = new Dictionary<string, string>(16);

        
        /// <summary>
        /// Windows Management Information
        /// </summary>
        public WMIShell()
        {
            ServicePort = 6103;
            Title = "Windows Management Information";
            _Author = "huangmin@spacecg.cn";
            _Version = new System.Version(0, 0, 19, 509);

            DefaultQueryString.Add("SELECT * FROM Win32_PnPEntity WHERE Caption LIKE '%(COM_)%' OR Caption LIKE '%(COM__)%'",
                "模糊查询包含 (COM_) 或 (COM__) 输出所有属性信息，其实是就是查询串口信息");

            DefaultQueryString.Add("SELECT Caption,Description,Service,Status FROM Win32_PnPEntity WHERE Caption LIKE '%(COM_)%' OR Caption LIKE '%(COM__)%'",
                "模糊查询包含 (COM_) 或 (COM__) 输出Caption,Description,Service,Status属性信息");

            DefaultQueryString.Add("SELECT ProcessorId,Name,Manufacturer,Version FROM Win32_Processor", "查询CPU编号、名称、制造厂商、版本");

            DefaultQueryString.Add("SELECT Manufacturer,SerialNumber,Product FROM Win32_BaseBoard", "查询主板制造厂商、编号、型号");

            DefaultQueryString.Add("SELECT IpAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled='true'", "查询本机可用的IP地址");

            DefaultQueryString.Add("SELECT MacAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled='true'", "查询本机可用的MAC地址");

            DefaultQueryString.Add("SELECT PNPDeviceID FROM Win32_VideoController", "查询显卡PNPDeviceID");

            DefaultQueryString.Add("SELECT Model FROM Win32_DiskDrive", "查询获取硬盘序列号");
        }

        [CommandLine("ex", "列举内置默认的信息查询语句")]
        public void ExampleQuery()
        {
            int index = 0;
            JArray Keys = new JArray();
            foreach(KeyValuePair<String, String> kvp in DefaultQueryString)
            {
                Keys.Add(new JObject()
                {
                    ["ID"] = index,
                    ["Query"] = kvp.Key,
                    ["Name"] = kvp.Value,
                });

                index++;
            }

            WriteLine(new JObject()
            {
                ["Queries"] = Keys,
            }, _OutputFormat);
        }

        [CommandLine("li", "列举 Win32 可查询的信息管理对象(表)")]
        public void MOList()
        {
            Type type = typeof(Win32APIEnum);
            String[] Names = type.GetEnumNames();

            JArray Keys = new JArray();
            foreach (String Name in Names)
            {
                Keys.Add(new JObject()
                {
                    ["ID"] = (int)Enum.Parse(type, Name),
                    ["Name"] = Name,
                    ["Description"] = WMIShell.GetDescription((Win32APIEnum)Enum.Parse(type, Name)),
                });
            }

            WriteLine(new JObject()
            {
                ["Table"] = Keys,
            }, _OutputFormat);
        }

        
        [CommandLine("mos", "(str)", "Management Object Searcher 管理信息的指定查询；参数：SQL 语句，参考表：-li -ex")]
        public void MOSearcher(String queryString)
        {
            if (String.IsNullOrWhiteSpace(queryString))
            {
                String ErrorMessage = "参数输入错误，不能为空";
                WriteError(ErrorMessage, _OutputFormat, MethodBase.GetCurrentMethod());
                return;
            }

            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(queryString);
                ManagementObjectCollection collection = searcher.Get();
                
                if (collection.Count <= 0)
                {
                    WriteLine(new JObject()
                    {
                        ["Results"] = { },
                    }, _OutputFormat);
                    return;
                }

                JArray Keys = new JArray();
                foreach (ManagementObject mo in collection)
                {
                    JObject obj = new JObject();
                    foreach (PropertyData pd in mo.Properties)
                    {
                        if (pd.Name == null || pd.Value == null) continue;

                        object value = mo.Properties[pd.Name].Value;
                        if (value == null) continue;
                        
                        if (pd.IsArray)
                             obj[pd.Name] = new JArray(value);
                        else
                            obj[pd.Name] = value.ToString();
                    }
                    Keys.Add(obj);
                }

                //JObject query = new JObject();
                //query["Query"] = queryString;
                //Keys.Add(query);

                WriteLine(new JObject()
                {
                    ["Results"] = Keys,
                    ["Query"] = queryString,
                }, _OutputFormat);
            }
            catch(Exception e)
            {
                WriteError("查询语句错误：" + queryString, _OutputFormat);
                //WriteError(e.Message, _OutputFormat);
            }
        }

        [CommandLine("mosp", "(enum)", "Management Object Searcher 管理信息的字段属性查询；参数参考：-li，使用索引或名称")]
        public void MOSearcherProperties(String win32api)
        {
            if (String.IsNullOrWhiteSpace(win32api))
            {
                String ErrorMessage = "参数输入错误，不能为空";
                WriteError(ErrorMessage, _OutputFormat, MethodBase.GetCurrentMethod());
                return;
            }

            try
            {
                Win32APIEnum s = (Win32APIEnum)Enum.Parse(typeof(Win32APIEnum), win32api, true);

                JArray Keys = new JArray();
                ManagementClass mc = new ManagementClass(Enum.GetName(typeof(Win32APIEnum), s));
                foreach (PropertyData pd in mc.Properties)
                {
                     Keys.Add(pd.Name);
                }

                WriteLine(new JObject()
                {
                    ["Results"] = Keys,
                }, _OutputFormat);
            }
            catch (Exception e)
            {
                WriteError(e.Message, _OutputFormat);
            }
        }

        [CommandLine("mosi", "(int)", "Management Object Searcher 管理信息的指定查询；参数：-ex 的索引，引用内置示例查询")]
        public void MOSearcherIndex(String queryIndex)
        {
            if (queryIndex == null)
            {
                WriteError("参数错误", _OutputFormat);
                return;
            }

            int Index = Convert.ToInt32(queryIndex);
            if (Index < 0 || Index >= DefaultQueryString.Count)
            {
                String ErrorMessage = "索引 " + Index + " 超出 QueryString 范围：" + (DefaultQueryString.Count - 1);
                WriteError(ErrorMessage, _OutputFormat, MethodBase.GetCurrentMethod());
                return;
            }

            int index = 0;
            String query = "";
            foreach (KeyValuePair<String, String> kvp in DefaultQueryString)
            {
                if(index == Index)
                {
                    query = kvp.Key;
                    break;
                }

                index++;
            }

            MOSearcher(query);
        }
        
        
        /// <summary>
        /// 枚举win32 api
        /// </summary>
        public enum Win32APIEnum
        {
            // 硬件
            /// <summary> CPU处理器信息 </summary>
            [Description("CPU处理器信息")]
            Win32_Processor,
            /// <summary> 物理内存信息 </summary>
            [Description("物理内存信息")]
            Win32_PhysicalMemory,
            /// <summary> 键盘信息 </summary>
            [Description("键盘信息")]
            Win32_Keyboard,
            /// <summary> 点输入设备 </summary>
            [Description("点输入设备")]
            Win32_PointingDevice,
            /// <summary> 软盘驱动器信息 </summary>
            [Description("软盘驱动器信息")]
            Win32_FloppyDrive,
            /// <summary> 硬盘驱动器信息 </summary>
            [Description("硬盘驱动器信息")]
            Win32_DiskDrive,
            /// <summary> 光盘驱动器 </summary>
            [Description("光盘驱动器")]
            Win32_CDROMDrive,
            /// <summary> 主板 </summary>
            [Description("主板")]
            Win32_BaseBoard,
            /// <summary> BIOS 芯片 </summary>
            [Description("BIOS 芯片")]
            Win32_BIOS,
            /// <summary> 并口 </summary>
            [Description("并口")]
            Win32_ParallelPort,
            /// <summary> 串口信息 </summary>
            [Description("串口信息")]
            Win32_SerialPort,
            /// <summary> 串口配置 </summary>
            [Description("串口配置")]
            Win32_SerialPortConfiguration,
            /// <summary> 多媒体设备信息 </summary>
            [Description("多媒体设备信息")]
            Win32_SoundDevice,
            /// <summary> 主板插槽 (ISA & PCI & AGP) </summary>
            [Description("主板插槽 (ISA & PCI & AGP)")]
            Win32_SystemSlot,
            /// <summary> USB 控制器 </summary>
            [Description("USB 控制器")]
            Win32_USBController,
            /// <summary> 网络适配器 </summary>
            [Description("网络适配器")]
            Win32_NetworkAdapter,
            /// <summary> 网络适配器配置信息 </summary>
            [Description("网络适配器配置信息")]
            Win32_NetworkAdapterConfiguration,
            /// <summary> 打印机 </summary>
            [Description("打印机")]
            Win32_Printer,
            /// <summary> 打印机配置信息 </summary>
            [Description("打印机配置信息")]
            Win32_PrinterConfiguration,
            /// <summary> 打印机任务信息 </summary>
            [Description("打印机任务信息")]
            Win32_PrintJob,
            /// <summary> 打印机端口 </summary>
            [Description("打印机端口")]
            Win32_TCPIPPrinterPort,
            /// <summary> MODEM </summary>
            [Description("MODEM")]
            Win32_POTSModem,
            /// <summary> MODEM 端口 </summary>
            [Description("MODEM 端口")]
            Win32_POTSModemToSerialPort,
            /// <summary> 显示器 </summary>
            [Description("显示器")]
            Win32_DesktopMonitor,
            /// <summary> 显示器配置信息 </summary>
            [Description("显示器配置信息")]
            Win32_DisplayConfiguration,
            /// <summary> 显卡配置 </summary>
            [Description("显卡配置")]
            Win32_DisplayControllerConfiguration,
            /// <summary> 显卡细节 </summary>
            [Description("显卡细节")]
            Win32_VideoController,
            /// <summary> 显卡支持的显示模式 </summary>
            [Description("显卡支持的显示模式")]
            Win32_VideoSettings,

            // 操作系统
            /// <summary> 时区 </summary>
            [Description("时区")]
            Win32_TimeZone,
            /// <summary> 驱动程序 </summary>
            [Description("驱动程序")]
            Win32_SystemDriver,
            /// <summary> 磁盘分区 </summary>
            [Description("磁盘分区")]
            Win32_DiskPartition,
            /// <summary> 逻辑磁盘 </summary>
            [Description("逻辑磁盘")]
            Win32_LogicalDisk,
            /// <summary> 逻辑磁盘所在分区及始末位置 </summary>
            [Description("逻辑磁盘所在分区及始末位置")]
            Win32_LogicalDiskToPartition,
            /// <summary> 逻辑内存配置 </summary>
            [Description("逻辑内存配置")]
            Win32_LogicalMemoryConfiguration,
            /// <summary> 系统页文件信息 </summary>
            [Description("系统页文件信息")]
            Win32_PageFile,
            /// <summary> 页文件设置 </summary>
            [Description("页文件设置")]
            Win32_PageFileSetting,
            /// <summary> 系统启动配置 </summary>
            [Description("系统启动配置")]
            Win32_BootConfiguration,
            /// <summary> 计算机信息简要 </summary>
            [Description("计算机信息简要")]
            Win32_ComputerSystem,
            /// <summary> 操作系统信息 </summary>
            [Description("操作系统信息")]
            Win32_OperatingSystem,
            /// <summary> 系统自动启动程序 </summary>
            [Description("系统自动启动程序")]
            Win32_StartupCommand,
            /// <summary> 系统安装的服务 </summary>
            [Description("系统安装的服务")]
            Win32_Service,
            /// <summary> 系统管理组 </summary>
            [Description("系统管理组")]
            Win32_Group,
            /// <summary> 系统组帐号 </summary>
            [Description("系统组帐号")]
            Win32_GroupUser,
            /// <summary> 用户帐号 </summary>
            [Description("用户帐号")]
            Win32_UserAccount,
            /// <summary> 系统进程 </summary>
            [Description("系统进程")]
            Win32_Process,
            /// <summary> 系统线程 </summary>
            [Description("系统线程")]
            Win32_Thread,
            /// <summary> 共享 </summary>
            [Description("共享")]
            Win32_Share,
            /// <summary> 已安装的网络客户端 </summary>
            [Description("已安装的网络客户端")]
            Win32_NetworkClient,
            /// <summary> 已安装的网络协议 </summary>
            [Description("已安装的网络协议")]
            Win32_NetworkProtocol,
            /// <summary> all device </summary>
            [Description("all device")]
            Win32_PnPEntity,
        }

        /// <summary>
        /// 扩展方法，获得枚举的Description
        /// </summary>
        /// <param name="value">枚举值</param>
        /// <param name="nameInstead">当枚举值没有定义DescriptionAttribute，是否使用枚举名代替，默认是使用</param>
        /// <returns>枚举的Description</returns>
        public static string GetDescription(Enum value, Boolean nameInstead = true)
        {
            Type type = value.GetType();
            string name = Enum.GetName(type, value);
            if (name == null)
            {
                return null;
            }

            FieldInfo field = type.GetField(name);
            DescriptionAttribute attribute = Attribute.GetCustomAttribute(field, typeof(DescriptionAttribute)) as DescriptionAttribute;

            if (attribute == null && nameInstead == true)
            {
                return name;
            }
            return attribute == null ? null : attribute.Description;
        }
    }

}