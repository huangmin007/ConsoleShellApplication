using System;
using System.ComponentModel;
using System.Reflection;

namespace ConsoleShell
{

    /// <summary>
    /// 控制台程序Help Attrubute
    /// </summary>
    [Description("控制台程序参数帮助信息")]
    [AttributeUsage(AttributeTargets.Method, AllowMultiple = false)]
    public class CommandLineAttribute : Attribute
    {
        /// <summary>
        /// 参数标记头
        /// </summary>
        public static String DefaultHeaderMark = "-";

        protected String _Name;

        /// <summary>
        /// 参数名
        /// </summary>
        public String Name
        {
            get { return String.Format("{0}{1}", DefaultHeaderMark, _Name); }
        }

        protected String _Arguments;
        /// <summary>
        /// 参数
        /// </summary>
        public String Arguments
        {
            get { return String.Format("{0}", _Arguments); }
        }

        protected String _Description;
        /// <summary>
        /// 参数描述
        /// </summary>
        public String Description
        {
            get { return _Description; }
        }

        /// <summary>
        /// Console Help Attribute
        /// </summary>
        /// <param name="name"></param>
        /// <param name="description"></param>
        public CommandLineAttribute(String name, String description)
        {
            this._Name = name;
            this._Description = description;
        }

        /// <summary>
        /// Console Help Attribute
        /// </summary>
        /// <param name="name"></param>
        /// <param name="arguments"></param>
        /// <param name="description"></param>
        public CommandLineAttribute(String name, String arguments, String description)
        {
            this._Name = name;
            this._Arguments = arguments;
            this._Description = description;
        }

        /// <summary>
        /// override ToString
        /// </summary>
        /// <returns></returns>
        public override String ToString()
        {
            String newLine = ("\r\n").PadRight(32, ' ');
            String fn = String.Format("{0}", String.IsNullOrWhiteSpace(_Arguments) ? " " : _Arguments);
            String MethodHelp = String.Format("    {0}{1}{2}", Name.PadRight(10, ' '), fn.PadRight(16, ' '), Description.Replace("\r\n", newLine));

            return MethodHelp;
        }

        /// <summary>
        /// override ToString
        /// </summary>
        /// <returns></returns>
        public String ToString(ConsoleOutputFormat Format)
        {
            String MethodHelp = "";
            if (Format == ConsoleOutputFormat.JSON)
            {
                MethodHelp = String.Format(@"{'Name':'{0}','Arguments':'{1}','Description':'{2}'}", Name, _Arguments, Description);
            }
            else if (Format == ConsoleOutputFormat.XML)
            {
                MethodHelp = String.Format(@"<Help Name='{0}' Arguments='{1}' Description='{2}' />", Name, _Arguments, Description);
            }
            else
            {
                String fn = String.Format("[{0}]", String.IsNullOrWhiteSpace(_Arguments) ? " " : _Arguments);
                MethodHelp = String.Format("    {0}{1}{2}", Name.PadRight(8, ' '), fn.PadRight(24, ' '), Description);
            }

            return MethodHelp;
        }

        /// <summary>
        /// 获取静态方法
        /// </summary>
        /// <param name="type"></param>
        /// <param name="Name">参数名称</param>
        /// <param name="help"></param>
        /// <returns></returns>
        public static MethodInfo GetStaticMethod(Type type, String Name, ref CommandLineAttribute help)
        {
            Name = Name.ToLower();
            CommandLineAttribute[] Attributes;
            MethodInfo[] Methods = type.GetMethods(BindingFlags.Public | BindingFlags.Static | BindingFlags.IgnoreCase);

            foreach (MethodInfo Method in Methods)
            {
                Attributes = (CommandLineAttribute[])Method.GetCustomAttributes(typeof(CommandLineAttribute), false);
                if (Attributes.Length <= 0) continue;

                help = Attributes[0];
                if (Attributes[0].Name.ToLower() == Name
                    // || Attributes[0].Name.Remove(0, 1).ToLower() == Name.ToLower()
                    // || Attributes[0].Arguments.ToLower() == Name.ToLower()
                    )
                {
                    return Method;
                }
            }

            return null;
        }


        public static MethodInfo GetPublicMethod(Type type, String Name, ref CommandLineAttribute CommandLine)
        {
            Name = Name.ToLower();
            CommandLineAttribute[] Attributes;
            MethodInfo[] Methods = type.GetMethods();//多个返回的处理，需要比对参数
            //MethodInfo[] Methods = type.GetMethods(BindingFlags.Public | BindingFlags.IgnoreCase);

            foreach (MethodInfo Method in Methods)
            {
                //Method.GetGenericArguments();
                Attributes = (CommandLineAttribute[])Method.GetCustomAttributes(typeof(CommandLineAttribute), false);
                if (Attributes.Length <= 0) continue;

                CommandLine = Attributes[0];
                if (Attributes[0].Name.ToLower() == Name
                    // || Attributes[0].Name.Remove(0, 1).ToLower() == Name.ToLower()
                    // || Attributes[0].Arguments.ToLower() == Name.ToLower()
                    )
                {
                    return Method;
                }
            }

            return null;
        }

    }
}