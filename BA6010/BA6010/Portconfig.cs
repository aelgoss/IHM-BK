using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Ports;
using System.Windows.Forms;

namespace BA6010
{
    public class ConfigSerialPort
    {
        /// <summary>   Name of the port. </summary>
        public string sPortName = "COM9";
        /// <summary>   Zero-based index of the baud rate. </summary>
        public int iBaudRate = 9600;
        /// <summary>   Zero-based index of the data bits. </summary>
        public int iDataBits = 8;
        /// <summary>   The parity. </summary>
        public Parity eParity = Parity.Odd;
        /// <summary>   The stop bits. </summary>
        public StopBits eStopBits = StopBits.One;
        /// <summary>   Size of the read buffer. </summary>
        public int iReadBufferSize = 1024;
        /// <summary>   Size of the write buffer. </summary>
        public int iWriteBufferSize = 1024;
        /// <summary>   true if dtr enable. </summary>
        public bool bDtrEnable = false;
        /// <summary>   true if RTS enable. </summary>
        public bool bRtsEnable = false;
        /// <summary>   Zero-based index of the read time out. </summary>
        public int iReadTimeOut = 1000;
        /// <summary>   Zero-based index of the write time out. </summary>
        public int iWriteTimeOut = 1000;
    }
}
