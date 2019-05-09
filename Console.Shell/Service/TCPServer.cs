using HPSocket;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace ConsoleShell.Service
{
    /// <summary>
    /// TCP Server
    /// </summary>
    public class TCPServer:IServer
    {
        protected IntPtr pServer = IntPtr.Zero;
        protected IntPtr pListener = IntPtr.Zero;

        protected SDK.OnSend _OnSend = null;
        protected SDK.OnClose _OnClose = null;
        protected SDK.OnAccept _OnAccept = null;
        protected SDK.OnReceive _OnReceive = null;
        protected SDK.OnShutdown _OnShutdown = null;
        protected SDK.OnPrepareListen _OnPrepareListen = null;

        private int _Port = 6100;
        private bool _IsCreate = false;

        #region Server Interface Properties
        /// <summary>
        /// 服务是否启动
        /// </summary>
        public bool IsStarted
        {
            get { return pServer == IntPtr.Zero ? false : SDK.HP_Server_HasStarted(pServer); }
        }
        
        /// <summary>
        /// 服务端口号
        /// </summary>
        public int Port
        {
            get { return _Port; }
        }

        /// <summary>
        /// 连接数
        /// </summary>
        public uint ClientCount
        {
            get { return SDK.HP_Server_GetConnectionCount(pServer); }
        }

        /// <summary>
        /// 获取所有连接
        /// </summary>
        /// <returns></returns>
        public IntPtr[] Clients
        {
            get
            {
                uint count = ClientCount;
                if (count == 0) return new IntPtr[0];
                
                IntPtr[] clients = new IntPtr[count];
                if (SDK.HP_Server_GetAllConnectionIDs(pServer, clients, ref count))
                {
                    if (clients.Length > count)
                    {
                        IntPtr[] newArr = new IntPtr[count];
                        Array.Copy(clients, newArr, count);
                        clients = newArr;
                    }
                }

                return clients;
            }
        }

        #endregion


        #region Channel Interface Event
        /// <summary>
        /// 服务对象接收数据事件
        /// </summary>
        public event DelegateServerDataReceived ServerDataReceived;
        /// <summary>
        /// 服务对象状态变化事件
        /// </summary>
        public event DelegateServerStateChanged ServerStateChanged;
        /// <summary>
        /// 客户端状态变化事件
        /// </summary>
        public event DelegateClientStateChanged ClientStateChanged;
        #endregion


        #region Channel Constructor & Initialize Function
        /// <summary>
        /// 构造函数
        /// </summary>
        public TCPServer()
        {
            InitializeChannel();
        }

        /// <summary>
        /// 初使化通道
        /// </summary>
        protected void InitializeChannel()
        {
            if (_IsCreate == true || pListener != IntPtr.Zero || pServer != IntPtr.Zero) return;

            pListener = SDK.Create_HP_TcpServerListener();
            if (pListener == IntPtr.Zero) return;

            pServer = SDK.Create_HP_TcpServer(pListener);
            if (pServer == IntPtr.Zero) return;

            _IsCreate = true;
            InitializeCallback();            
        }
        #endregion


        #region Channel Interface Function
        /// <summary>
        /// 启动服务
        /// </summary>
        /// <param name="port"></param>
        /// <returns></returns>
        public bool Start(int port)
        {
            if (!_IsCreate) return false;
            if (IsStarted) return false;

            if (port <= 0) throw new ArgumentException("端口号不能小于0.");         

            return SDK.HP_Server_Start(pServer, "0.0.0.0", (ushort)port);
        }

        /// <summary>
        /// 停止
        /// </summary>
        /// <returns></returns>
        public bool Stop()
        {
            if (!IsStarted) return false;

            IntPtr[] clients = Clients;
            for (int i = 0; i < clients.Length; i++)
                SDK.HP_Server_Disconnect(pServer, clients[i], true);

            return SDK.HP_Server_Stop(pServer);
        }

        /// <summary>
        /// 销毁服务对象
        /// </summary>
        /// <returns></returns>
        public bool Destroy()
        {
            try
            {
                Stop();

                if (pServer != IntPtr.Zero)
                {
                    SDK.Destroy_HP_TcpServer(pServer);
                    pServer = IntPtr.Zero;
                }
                if (pListener != IntPtr.Zero)
                {
                    SDK.Destroy_HP_TcpServerListener(pListener);
                    pListener = IntPtr.Zero;
                }

                _IsCreate = false;

                _OnSend = null;
                _OnClose = null;
                _OnAccept = null;
                _OnReceive = null;
                _OnShutdown = null;
                _OnPrepareListen = null;
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// 发送广播数据
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public bool WriteBytes(byte[] data)
        {
            if (!IsStarted) return false;

            IntPtr[] clients = Clients;
            for (int i = 0; i < clients.Length; i++)
                SDK.HP_Server_Send(pServer, clients[i], data, data.Length);

            return true;
        }

        /// <summary>
        /// 发送数据
        /// </summary>
        /// <param name="connId"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public bool WriteBytes(IntPtr connId, byte[] data)
        {
            if (!IsStarted) return false;
            
            return SDK.HP_Server_Send(pServer, connId, data, data.Length);
        }
        

        /// <summary>
        /// 获取远程连接的地址信息
        /// </summary>
        /// <param name="connId"></param>
        /// <param name="address"></param>
        /// <param name="port"></param>
        /// <returns></returns>
        public bool GetRemoteAddress(IntPtr connId, ref string address, ref ushort port)
        {
            int ipLength = 40;
            StringBuilder sb = new StringBuilder(ipLength);

            bool ret = SDK.HP_Server_GetRemoteAddress(pServer, connId, sb, ref ipLength, ref port) && ipLength > 0;
            if (ret == true) address = sb.ToString();

            return ret;
        }

        #endregion


        #region Server Callback Event Handler
        /// <summary>
        /// Initialize Callback
        /// </summary>
        protected void InitializeCallback()
        {
            _OnSend = new SDK.OnSend(SDK_OnSend);
            _OnClose = new SDK.OnClose(SDK_OnClose);
            _OnAccept = new SDK.OnAccept(SDK_OnAccept);
            _OnReceive = new SDK.OnReceive(SDK_OnReceive);
            _OnShutdown = new SDK.OnShutdown(SDK_OnShutdown);
            _OnPrepareListen = new SDK.OnPrepareListen(SDK_OnPrepareListen);

            SDK.HP_Set_FN_Server_OnSend(pListener, _OnSend);
            SDK.HP_Set_FN_Server_OnClose(pListener, _OnClose);
            SDK.HP_Set_FN_Server_OnAccept(pListener, _OnAccept);
            SDK.HP_Set_FN_Server_OnReceive(pListener, _OnReceive);
            SDK.HP_Set_FN_Server_OnShutdown(pListener, _OnShutdown);
            SDK.HP_Set_FN_Server_OnPrepareListen(pListener, _OnPrepareListen);
        }
        protected HandleResult SDK_OnPrepareListen(IntPtr pSender, IntPtr soListen)
        {
            ServerStateChanged?.Invoke(this, ServerState.Listened);

            return HandleResult.Ignore;
        }

        protected HandleResult SDK_OnAccept(IntPtr pSender, IntPtr connId, IntPtr pClient)
        {
            ServerStateChanged?.Invoke(this, ServerState.Accepted);
            ClientStateChanged?.Invoke(this, connId, ServerState.Connected);

            return HandleResult.Ignore;
        }

        protected HandleResult SDK_OnSend(IntPtr pSender, IntPtr connId, IntPtr pData, int length)
        {
            ServerStateChanged?.Invoke(this, ServerState.Sended);

            return HandleResult.Ignore;
        }

        protected HandleResult SDK_OnReceive(IntPtr pSender, IntPtr connId, IntPtr pData, int length)
        {
            if (ServerDataReceived != null)
            {
                byte[] bytes = new byte[length];
                Marshal.Copy(pData, bytes, 0, length);
                ServerDataReceived.Invoke(this, connId, bytes);
            }
            return HandleResult.Ignore;
        }

        protected HandleResult SDK_OnClose(IntPtr pSender, IntPtr connId, SocketOperation enOperation, int errorCode)
        {
            ClientStateChanged?.Invoke(this, connId, ServerState.Closed);

            return HandleResult.Ignore;
        }

        protected HandleResult SDK_OnShutdown(IntPtr pSender)
        {
            ServerStateChanged?.Invoke(this, ServerState.Shutdown);
            return HandleResult.Ignore;
        }

        #endregion
        
    }
}
