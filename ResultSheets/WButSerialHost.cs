using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SPorts = System.IO.Ports;

namespace ResultSheets
{
    class WButSerialHost
    {
        private SPorts.SerialPort arduinoPort = null;   // Главное устройство
        private byte[] sBuf = new byte[255];            // Буфер для отправки данных
        private byte[] rBuf = new byte[255];            // Буфер для приема даных
        private System.Timers.Timer timer1;
        private bool isConnected = false;

        public bool IsConnected { get {return isConnected;} }

        // Constructor
        public WButSerialHost(string portName)
        {
            isConnected = false;
            string[] ports = SPorts.SerialPort.GetPortNames();
            if (ports.Contains(portName))
            {
                timer1 = new System.Timers.Timer(100);
                timer1.Elapsed += timer1_Elapsed;
                timer1.AutoReset = true;

                try
                {
                    OpenPort(portName);
                    isConnected = true;
                }
                catch
                {

                }
            }
        }

        // Timer Callback
        void timer1_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            
        }

        // Destructor
        ~WButSerialHost()
        {
            ClosePort();
            timer1.Elapsed -= timer1_Elapsed;
        }

        // Закрыть порт
        private void ClosePort()
        {
            timer1.Stop();
            if (arduinoPort == null) return;
            if (arduinoPort.IsOpen)
            {
                arduinoPort.Close();
            }
            arduinoPort = null;
        }

        // Открыть порт
        private void OpenPort(string portName)
        {
            arduinoPort = new SPorts.SerialPort(portName, 9600);
            arduinoPort.RtsEnable = false;
            arduinoPort.DtrEnable = true;
            arduinoPort.Open();
            timer1.Start();
        }

        // Посылаем байт
        private void Send(int num)
        {
            if (arduinoPort == null) return;
            if (arduinoPort.IsOpen)
            {
                arduinoPort.Write(sBuf, 0, num);
            }
        }

        // Получаем байт
        private int RecieveBytes()
        {
            int num = 0;
            if (arduinoPort == null) return num;
            if (arduinoPort.IsOpen)
            {
                num = arduinoPort.BytesToRead;
                if (num > 0)
                {
                    arduinoPort.Read(rBuf, 0, num);
                }
            }

            return num;
        }
    }
}
