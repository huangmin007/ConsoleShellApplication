using System;
using System.Text;
using PPT = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;
using System.Diagnostics;
using Newtonsoft.Json.Linq;
using System.Windows;
using ConsoleShell;
using ConsoleShell.Service;

namespace MTS.Shell
{
    /// <summary>
    /// PowerPoint 文档控制，可以用多个文档切换。。。
    /// </summary>
    public partial class PowerPointShell:ConsoleApplication
    {
        private const String _FileName = "POWERPNT.EXE";

        /// <summary>
        /// PowerPoint Application
        /// </summary>
        private PPT.Application Application = null;
        /// <summary>
        /// 活动文档
        /// </summary>
        private PPT.Presentation ActivePresentation;


        public PowerPointShell()
        {
            ServicePort = 6101;
            _Author = "hm-secret@163.com";

            try
            {
                Application = Marshal.GetActiveObject("PowerPoint.Application") as PPT.Application;
                if (Application != null) ActivePresentation = Application.ActivePresentation;
            }
            catch(Exception)
            {
                // ...
            }
        }


        /// <summary>
        /// 是否有活动的 PowerPoint 文档内容
        /// </summary>
        /// <returns></returns>
        private Boolean HasActivePresentation()
        {
            if (ActivePresentation != null) return true;

            try
            {
                Application = Marshal.GetActiveObject("PowerPoint.Application") as PPT.Application;
                ActivePresentation = Application.ActivePresentation;
                return true;
            }
            catch (Exception)
            {
                WriteError("未获取到活动的 PowerPoint 应用文档", _OutputFormat);
                return false;
            }
        }

        
        [CommandLine("a", "激活应用窗体，将活动的演示文档窗体在最前面显示")]
        public void Activate()
        {
            if (!HasActivePresentation()) return;
            Application.Activate();
        }


        [CommandLine("aw", "(int index)", "编辑模式下有效，激活多个编辑文档中的指定一个窗体，参考：-i，示例：-aw 1")]
        public void ActivateWindow(String args)
        {
            if (!HasActivePresentation()) return;
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

            int Index = Convert.ToInt32(args);
            if (Index < 1 || Index > Application.Windows.Count)
            {
                String ErrorMessage = "索引 " + Index + " 超出 Windows 范围：" + Application.Windows.Count;
                WriteError(ErrorMessage, _OutputFormat, MethodBase.GetCurrentMethod());
                return;
            }

            Application.Windows[Index].Activate();
        }


        //[CommandLine("aws", "ApplicationWindowState", "编辑模式下有效，设置 PowerPoint 应用窗体状态，参考：1:ppWindowNormal, 2:ppWindowMinimized, 3:ppWindowMaximized")]
        public void ApplicationWindowState(String args)
        {
            if (!HasActivePresentation()) return;

            if (String.IsNullOrWhiteSpace(args))
            {
                String ErrorMessage = "参数 " + args + " 输入错误，不能为空";
                WriteError(ErrorMessage, _OutputFormat, MethodBase.GetCurrentMethod());
                return;
            }

            try
            {
                Application.WindowState = Uint_Regex.IsMatch(args) ? (PPT.PpWindowState)Convert.ToInt32(args) : (PPT.PpWindowState)Enum.Parse(typeof(PPT.PpWindowState), args, true);
            }
            catch (Exception e)
            {
                WriteError(e.Message, _OutputFormat, MethodBase.GetCurrentMethod());
            }
        }


        [CommandLine("np", "显示文档下一页，示例：-np 或 -of json -np -i")]
        public void NextPage()
        {
            if (!HasActivePresentation()) return;
            try
            {
                //要判断是不是最后一页，如果是，就不要结束了
                if (ActivePresentation.SlideShowWindow.View.Slide.SlideIndex < ActivePresentation.Slides.Count)
                    ActivePresentation.SlideShowWindow.View.Next();
            }
            catch (Exception)
            {
                int Index = ActivePresentation.Windows[1].Selection.SlideRange.SlideIndex + 1;
                ActivePresentation.Slides[Index > ActivePresentation.Slides.Count ? 1 : Index].Select();
            }
        }


        [CommandLine("pp", "显示文档上一页")]
        public void PrevPage()
        {
            if (!HasActivePresentation()) return;
            try
            {
                ActivePresentation.SlideShowWindow.View.Previous();
            }
            catch (Exception)
            {
                int Index = ActivePresentation.Windows[1].Selection.SlideRange.SlideIndex - 1;
                ActivePresentation.Slides[Index < 1 ? ActivePresentation.Slides.Count : Index].Select();
            }
        }


        [CommandLine("gp", "(int index)", "显示文档指定页面(参考：-i)。\r\n示例：-gp 2 或 -of json -gp 2 -i")]
        public void GotoPage(String args)
        {
            if (!HasActivePresentation()) return;
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

            int Index = Convert.ToInt32(args);
            if (Index < 1 || Index > ActivePresentation.Slides.Count)
            {
                String ErrorMessage = "索引 " + Index + " 超出 PowerPoint 文档页面范围：" + ActivePresentation.Slides.Count;
                WriteError(ErrorMessage, _OutputFormat, MethodBase.GetCurrentMethod());
                return;
            }

            try
            {
                ActivePresentation.SlideShowWindow.View.GotoSlide(Index);
            }
            catch (Exception)
            {
                ActivePresentation.Slides[Index].Select();
            }
        }


        [CommandLine("fp", "显示文档第一页")]
        public void FirstPage()
        {
            if (!HasActivePresentation()) return;
            try
            {
                ActivePresentation.SlideShowWindow.View.First();
            }
            catch (Exception)
            {
                ActivePresentation.Slides[1].Select();
            }
        }


        [CommandLine("lp", "显示文档最后一页")]
        public void LastPage()
        {
            if (!HasActivePresentation()) return;
            try
            {
                ActivePresentation.SlideShowWindow.View.Last();
            }
            catch (Exception)
            {
                ActivePresentation.Slides[ActivePresentation.Slides.Count].Select();
            }
        }


        [CommandLine("gc", "(int index)", "演示模式下有效，控制动画播放的索引(参考：-i)。\r\n示例：-gc 2")]
        public void GotoClick(String args)
        {
            if (!HasActivePresentation()) return;
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
                int Index = Convert.ToInt32(args);
                int ClickCount = ActivePresentation.SlideShowWindow.View.GetClickCount();
                if (Index < 0 || Index > ClickCount)
                {
                    String ErrorMessage = "索引 " + Index + " 超出 PowerPoint 文档当前页面范围：" + ClickCount;
                    WriteError(ErrorMessage, _OutputFormat, MethodBase.GetCurrentMethod());
                    return;
                }

                ActivePresentation.SlideShowWindow.View.GotoClick(Index);
            }
            catch (Exception)
            {
                String ErrorMessage = "非演示模式下的 PowerPoint 应用文档，操作失败";
                WriteError(ErrorMessage, _OutputFormat, MethodBase.GetCurrentMethod());
            }
        }


        [CommandLine("exp", "[path,width,height]", "将 PowerPoint 文档以图片导出(注意文件完整路径使用'\\')。\r\n示例：-exp \"D:\\ttp\\\" 1920 1080")]
        public void Export(String path, String width, String height)
        {
            if (!HasActivePresentation()) return;
            
            String Path = String.IsNullOrWhiteSpace(path) ? ActivePresentation.Path + "/" + ActivePresentation.Name.Split('.')[0] : path;
            //目录不存在则创建
            if (!Directory.Exists(Path)) Directory.CreateDirectory(Path);

            int ScaleWidth = String.IsNullOrWhiteSpace(width) || !Uint_Regex.IsMatch(width) ? 0 : Convert.ToInt32(width);
            int ScaleHeight = String.IsNullOrWhiteSpace(height) || !Uint_Regex.IsMatch(height) ? 0 : Convert.ToInt32(height);

            try
            {
                ActivePresentation.Export(Path.Replace("/", "\\"), "jpg", ScaleWidth, ScaleHeight);
            }
            catch (Exception e)
            {
                WriteError("导出失败：" + e.Message, _OutputFormat, MethodBase.GetCurrentMethod());
            }
        }


        [CommandLine("exps", "(index,[path,width,height])", "将 PowerPoint 文档指定的页面以图片导出(注意文件完整路径使用'\\')。\r\n示例：-exps 1 \"D:\\ttp\\\" 1920 1080")]
        public void ExportSlide(String index, String path, String width, String height)
        {
            if (!HasActivePresentation()) return;
            if (String.IsNullOrWhiteSpace(index))
            {
                String ErrorMessage = "参数 index 输入错误，不能为空";
                WriteError(ErrorMessage, _OutputFormat, MethodBase.GetCurrentMethod());
                return;
            }
            if (!Uint_Regex.IsMatch(index))
            {
                String ErrorMessage = "参数 " + index + " 格式输入错误";
                WriteError(ErrorMessage, _OutputFormat, MethodBase.GetCurrentMethod());
                return;
            }

            int Index = Convert.ToInt32(index);
            if (Index < 1 || Index > ActivePresentation.Slides.Count)
            {
                String ErrorMessage = "索引 " + Index + " 超出 PowerPoint 文档页面范围：" + ActivePresentation.Slides.Count;
                WriteError(ErrorMessage, _OutputFormat, MethodBase.GetCurrentMethod());
                return;
            }

            String Path = String.IsNullOrWhiteSpace(path) ? ActivePresentation.Path + "/" + ActivePresentation.Name.Split('.')[0] : path;
            //目录不存在则创建
            if (!Directory.Exists(Path)) Directory.CreateDirectory(Path);

            int ScaleWidth = String.IsNullOrWhiteSpace(width) || !Uint_Regex.IsMatch(width) ? 0 : Convert.ToInt32(width);
            int ScaleHeight = String.IsNullOrWhiteSpace(height) || !Uint_Regex.IsMatch(height) ? 0 : Convert.ToInt32(height);

            try
            {
                ActivePresentation.Slides[Index].Export(Path.Replace("/", "\\"), "jpg", ScaleWidth, ScaleHeight);
            }
            catch (Exception e)
            {
                WriteError("导出失败：" + e.Message, _OutputFormat, MethodBase.GetCurrentMethod());
            }
        }


        [CommandLine("od", "(path)", "打开 PowerPoint 文档，示例：-od \"D:\\ppt.pptx\"")]
        public void OpenDocument(String args)
        {
            if (String.IsNullOrWhiteSpace(args))
            {
                String ErrorMessage = "参数输入错误，不能为空";
                WriteError(ErrorMessage, _OutputFormat, MethodBase.GetCurrentMethod());
                return;
            }

            FileInfo file = new FileInfo(args);
            if (!file.Exists)
            {
                WriteError("指定的文件不存在：" + args, _OutputFormat, MethodBase.GetCurrentMethod());
                return;
            }

            if (file.Extension.ToLower().IndexOf(".pp") != 0)
            {
                WriteError("指定的文件类型错误：" + file.Extension, _OutputFormat, MethodBase.GetCurrentMethod());
                return;
            }

            try
            {
                PPT.Application App = new PPT.Application();
                String fn = App.Path + @"\" + _FileName;
                Process.Start(fn, "/S " + args);
            }
            catch (Exception e)
            {
                WriteError(e.Message, _OutputFormat, MethodBase.GetCurrentMethod());
                return;
            }
        }


        [CommandLine("cd", "关闭活动的 PowerPoint 文档")]
        public void CloseDocument()
        {
            if (!HasActivePresentation()) return;
            ActivePresentation.Close();
        }


        [CommandLine("esv", "退出活动的 PowerPoint 文档，只在演示模式下有效")]
        public void ExitShowView()
        {
            if (!HasActivePresentation()) return;
            try
            {
                ActivePresentation.SlideShowWindow.View.Exit();
            }
            catch (Exception e)
            {
                WriteError("当前文档不是演示模式：" + e.Message, _OutputFormat, MethodBase.GetCurrentMethod());
            }
        }


        [CommandLine("qa", "退出 PowerPoint 应用程序，即关闭所有活动文档")]
        public void QuitApplication()
        {
            if (!HasActivePresentation()) return;
            Application.Quit();
        }

        [CommandLine("i", "输出打开的文档/页面信息")]
        public void Infomations()
        {
            Console.WriteLine(GetInfomations());
        }

        /// <summary>
        /// 该函数只负责返回信息，不负责打印
        /// </summary>
        /// <returns></returns>
        protected String GetInfomations()
        {
            try
            {
                if (ActivePresentation == null)
                {
                    Application = Marshal.GetActiveObject("PowerPoint.Application") as PPT.Application;
                    ActivePresentation = Application.ActivePresentation;
                }
            }
            catch (Exception)
            {
                // ...
            }

            JObject RootObject = new JObject();
            if (ActivePresentation == null)
            {
                PPT.Application App = new PPT.Application();

                RootObject["Version"] = App.Version;
                RootObject["File"] = _FileName;
                RootObject["Path"] = App.Path;

                return ConvertToString(RootObject, _OutputFormat);
            }

            RootObject["Version"] = Application.Version;
            RootObject["File"] = _FileName;
            RootObject["Path"] = Application.Path;

            try
            {
                JArray Windows = new JArray();
                for (int i = 1; i <= Application.Windows.Count; i++)
                {
                    JObject Window = new JObject();

                    Window["ID"] = i;
                    Window["Active"] = Application.Windows[i].Active == MsoTriState.msoTrue || Application.Windows[i].Active == MsoTriState.msoCTrue;
                    Window["FileName"] = Application.Windows[i].Presentation.Name;
                    Window["FullName"] = Application.Windows[i].Presentation.FullName;
                    Window["SlideIndex"] = Application.Windows[i].Selection.SlideRange.SlideIndex;
                    Window["SlideCount"] = Application.Windows[i].Presentation.Slides.Count;
                    try
                    {
                        Window["ClickIndex"] = Application.Windows[i].Presentation.SlideShowWindow.View.GetClickIndex();
                        Window["ClickCount"] = Application.Windows[i].Presentation.SlideShowWindow.View.GetClickCount();
                    }
                    catch (Exception)
                    {
                        Window["ClickIndex"] = -1;
                        Window["ClickCount"] = -1;
                    }

                    Windows.Add(Window);
                }
                RootObject["Windows"] = Windows;
            }
            catch (Exception)
            {
                // ...
            }

            try
            {
                JArray Presentations = new JArray();
                for (int i = 1; i <= Application.Presentations.Count; i++)
                {
                    JObject Presentation = new JObject();
                    Presentation["ID"] = i;
                    Presentation["FileName"] = Application.Presentations[i].Name;
                    Presentation["FullName"] = Application.Presentations[i].FullName;
                    Presentation["SlideIndex"] = ActivePresentation.SlideShowWindow.View.State == PPT.PpSlideShowState.ppSlideShowDone ?
                                 Application.Presentations[i].Slides.Count : ActivePresentation.SlideShowWindow.View.Slide.SlideIndex;
                    Presentation["SlideCount"] = Application.Presentations[i].Slides.Count;

                    Presentation["ClickIndex"] = Application.Presentations[i].SlideShowWindow.View.GetClickIndex();
                    Presentation["ClickCount"] = Application.Presentations[i].SlideShowWindow.View.GetClickCount();

                    Presentations.Add(Presentation);
                }
                RootObject["Presentations"] = Presentations;
            }
            catch (Exception)
            {
                // ...
            }

            return ConvertToString(RootObject, _OutputFormat);
        }

        protected override void Server_DataReceived(object sender, IntPtr connId, byte[] data)
        {
            base.Server_DataReceived(sender, connId, data);

            IServer Server = (IServer)sender;
            Server.WriteBytes(connId, Encoding.Default.GetBytes(GetInfomations()));
        }

        
    }
    
}

