using System;
using System.ComponentModel;

namespace ConsoleShell.Service
{

    /// <summary>
    /// 通信/通道状态
    /// </summary>
    [Description("通信/通道状态")]
    public enum ServerState
    {
        [Description("连接状态")]
        Connected,

        [Description("关闭状态")]
        Closed,

        [Description("服务端接启动监听")]
        Listened,

        [Description("服务端接受客户端")]
        Accepted,

        [Description("发送数据状态")]
        Sended,

        [Description("服务端接Shutdown")]
        Shutdown,

        [Description("错误状态")]
        Error,
    }

    public delegate void DelegateServerStateChanged(object sender, ServerState newState);

    public delegate void DelegateServerDataReceived(object sender, IntPtr connId, byte[] data);

    public delegate void DelegateClientStateChanged(object sender, IntPtr connId, ServerState newState);
    
    /// <summary>
    /// 通信服务器接口
    /// </summary>
    public interface IServer
    {
        event DelegateServerDataReceived ServerDataReceived;

        event DelegateServerStateChanged ServerStateChanged;

        event DelegateClientStateChanged ClientStateChanged;
        
        int Port { get; }

        uint ClientCount { get; }

        IntPtr[] Clients { get; }

        Boolean IsStarted { get; }
        
        Boolean Start(int port);

        Boolean Stop();

        Boolean Destroy();

        Boolean WriteBytes(byte[] data);

        Boolean WriteBytes(IntPtr connId, byte[] data);

        Boolean GetRemoteAddress(IntPtr connId, ref String address, ref ushort port);
    }

}
