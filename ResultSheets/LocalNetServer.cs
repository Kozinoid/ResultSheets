using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using NET_TCP_Device;


namespace ResultSheets
{
    class LocalNetServer
    {
        public const int INC_SCORE = 3;
        public const int DEC_SCORE = 4;
        public const int NEW_TEAM = 5;
        public const int DEL_TEAM = 6;
        public const int RESET_TEAM = 7;
        public const int RESET_ALL = 8;
        public const int SORT = 9;

        private int appPort = 8060;
        private delegate void TextRefredhDelegate(string str);
        TCP_NET_Server_Device tcp_divice;

        public IPAddress ServerIPAddress
        {
            get
            {
                // Получение имени компьютера.
                String host = Dns.GetHostName();
                // Получение ip-адреса.
                IPAddress ip = Dns.GetHostByName(host).AddressList[0];
                return ip;
            }
        }

        public string StringIPAddress
        {
            get
            {
                // Получение имени компьютера.
                String host = Dns.GetHostName();
                // Получение ip-адреса.
                IPAddress ip = Dns.GetHostByName(host).AddressList[0];
                return ip.ToString();
            }
        }

        public LocalNetServer(int port)
        {
            appPort = port;
            tcp_divice = new TCP_NET_Server_Device();
            Connect();
        }

        public void Connect()
        {
            tcp_divice.Connect(GetConnectionIP()); // Autodetect server IP
            while (!tcp_divice.IsConnected) { };
            tcp_divice.onRecieveMessage += tcp_divice_onRecieveMessage;
        }

        public void Disconnect()
        {
            tcp_divice.onRecieveMessage -= tcp_divice_onRecieveMessage;
            tcp_divice.Disconnect();
        }

        void tcp_divice_onRecieveMessage(object sender, OnReciveMessageEventArgs e)
        {
            switch (e.command)
            {
                case OnReciveMessageEventArgs.CONNECTED:
                    
                    break;
                case OnReciveMessageEventArgs.DISCONNECTED:
                    
                    break;
                case OnReciveMessageEventArgs.TEXT_MESSAGE:
                    
                    break;
            }
        }

        public String GetConnectionIP()
        {
            return "tcp://" + StringIPAddress + ":" + appPort.ToString() + "/";
        }
    }
}
