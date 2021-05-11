﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Eneter.Messaging.EndPoints.TypedMessages;
using Eneter.Messaging.MessagingSystems.MessagingSystemBase;
using Eneter.Messaging.MessagingSystems.TcpMessagingSystem;
using System.Windows.Forms;


namespace NET_TCP_Device
{
    // Request message type
    public class MyRequest
    {
        public string Text { get; set; }
    }

    // Response message type
    public class MyResponse
    {
        public string Text { get; set; }
    }

    class OnReciveMessageEventArgs:EventArgs
    {
        public const int TEXT_MESSAGE = 0;
        public const int CONNECTED = 1;
        public const int DISCONNECTED = 2;

        public int command = TEXT_MESSAGE;
        public string id = "";
        public string text = "";

        public OnReciveMessageEventArgs(int com, string i, string t)
        {
            command = com;
            id = i;
            text = t;
        }
    }
    class TCP_NET_Server_Device
    {
        public delegate void OnRecieveMessageHandler(object sender, OnReciveMessageEventArgs e);
        public event OnRecieveMessageHandler onRecieveMessage;

        private IDuplexTypedMessageReceiver<MyResponse, MyRequest> myReceiver;
        private bool isConnected = false;
        private ClientList clientList = new ClientList();

        public bool IsConnected
        {
            get { return isConnected; }
        }

        // Constructor
        public TCP_NET_Server_Device()
        {
            
        }

        // Connect
        public void Connect(string channelID)
        {
            // Create message receiver receiving 'MyRequest' and receiving 'MyResponse'.
            IDuplexTypedMessagesFactory aReceiverFactory = new DuplexTypedMessagesFactory();
            myReceiver = aReceiverFactory.CreateDuplexTypedMessageReceiver<MyResponse, MyRequest>();

            // Subscribe to handle messages.
            myReceiver.MessageReceived += OnMessageReceived;

            // Create TCP messaging.
            IMessagingSystemFactory aMessaging = new TcpMessagingSystemFactory();
            IDuplexInputChannel anInputChannel =
               aMessaging.CreateDuplexInputChannel(channelID);    

            // Attach the input channel and start to listen to messages.
            myReceiver.AttachDuplexInputChannel(anInputChannel);

            isConnected = true;
        }

        // Disconnect
        public void Disconnect()
        {
            isConnected = false;

            // Detach the input channel and stop listening.
            // It releases the thread listening to messages.
            myReceiver.MessageReceived -= OnMessageReceived;
            myReceiver.DetachDuplexInputChannel();
        }

        // OnMessageReceived
        private void OnMessageReceived(object sender, TypedRequestReceivedEventArgs<MyRequest> e)
        {
            string message = e.RequestMessage.Text;
            string id = e.ResponseReceiverId;

            ServerClient foundClient = clientList.FindClientByID(id);

            if (foundClient != null)
             {
                 // Проверяем сообщение
                 if (message == "#e#n#d")   // если #e#n#d
                 {
                     // посылаем внешнему приложению сообщение о выходе клиента
                     if (onRecieveMessage != null)
                     {
                         OnReciveMessageEventArgs ea = new OnReciveMessageEventArgs(OnReciveMessageEventArgs.DISCONNECTED,
                             id, "");
                         onRecieveMessage(this, ea);
                     }
                     clientList.Remove(foundClient);    // удаляем клиентаs
                 }
                 else// иначе
                 {
                     // посылаем внешнему приложению сообщение от клиента
                     if (onRecieveMessage != null)
                     {
                         OnReciveMessageEventArgs ea = new OnReciveMessageEventArgs(OnReciveMessageEventArgs.TEXT_MESSAGE,
                                 id, message);
                         onRecieveMessage(this, ea);
                     }
                 }
             }
            else
             {
                 // Новый клиент. Прверяем запрос 
                 if (message == "#n#e#w")   // если #n#e#w
                 {
                     SendMessage(id, "#y#e#s");             // отвечаем клиенту, 
                     clientList.Add(new ServerClient(id));  // заносим слиента в список
                     // посылаем внешнему приложению сообщение о новом клиенте
                     if (onRecieveMessage != null)
                     {
                         OnReciveMessageEventArgs ea = new OnReciveMessageEventArgs(OnReciveMessageEventArgs.CONNECTED,
                                 id, "");
                         onRecieveMessage(this, ea);
                     }
                 }
             }
        }

        // Отправка сообщения message клиенту recID
        public void SendMessage(string recID, string str)
        {
            // Create the response message.
            MyResponse aResponse = new MyResponse();
            aResponse.Text = str;

            // Send the response message back to the client.
            myReceiver.SendResponseMessage(recID, aResponse);
        }

        // Отправка сообщения message клиенту client
        public void SendMessage(ServerClient client, string str)
        {
            SendMessage(client.ClientID, str);
        }

        // Отправить всем
        public void BroadcastSendMessage(string str)
        {
            foreach (ServerClient sc in clientList)
            {
                SendMessage(sc.ClientID, str);
            }
        }
    }

    class ServerClient
    {
        private string clientID = "";
        public string ClientID { get { return clientID; } }
        public ServerClient(string id)
        {
            clientID = id;
        }
        
    }

    class ClientList : List<ServerClient>
    {
        public ServerClient FindClientByID(string id)
        {
            ServerClient res = null;
            foreach (ServerClient sc in this)
            {
                if (sc.ClientID == id)
                {
                    res = sc;
                    break;
                }
            }
            return res;
        }
    }
}
