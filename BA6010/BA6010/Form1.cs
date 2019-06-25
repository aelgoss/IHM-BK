using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Configuration;
using System.IO.Ports;
using Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace BA6010
{
    public partial class Form1 : Form
    {
        // excel file 
        private static Microsoft.Office.Interop.Excel.Workbook mWorkBook;
        private static Microsoft.Office.Interop.Excel.Sheets mWorkSheets;
        private static Microsoft.Office.Interop.Excel.Worksheet mWSheet1;
        private static Microsoft.Office.Interop.Excel.Application oXL;


        // communication serie 
        /// <summary>   The serial. </summary>
        private SerialPort m_Serial = new SerialPort();
        /// <summary>   The configuration serial. </summary>
        private ConfigSerialPort m_ConfigSerial = new ConfigSerialPort();

        public Form1()
        {
            ConfigSerialPort _port = new ConfigSerialPort();
            bool _openserial = true;
            InitializeComponent();

            /*Init(_port);
            Open(out _openserial);*/
        }

        public void Init(ConfigSerialPort oConfig)
        {
            // Recopie de l'objet de configuration de la liaison série
            m_ConfigSerial = oConfig;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Label3_Click(object sender, EventArgs e)
        {

        }

        public void Open(out Boolean bIsSerialPortAlreadyOpen)
        {
            // Initialisation du booléen indiquant l'état du port série
            bIsSerialPortAlreadyOpen = false;

            // Récupération de l'état du port série
            if (m_Serial.IsOpen)
            {
                bIsSerialPortAlreadyOpen = true;
                return;
            }

            // Recopie de la désignation du port
            m_Serial.PortName = m_ConfigSerial.sPortName;
            // Recopie de la vitesse de transfert
            m_Serial.BaudRate = m_ConfigSerial.iBaudRate;
            // Recopie du nombre de bits de data
            m_Serial.DataBits = m_ConfigSerial.iDataBits;
            // Recopie du nombre de bits de parité
            m_Serial.Parity = m_ConfigSerial.eParity;
            // Recopie du nombre de bits de stop
            m_Serial.StopBits = m_ConfigSerial.eStopBits;
            // Recopie de la taille du buffer de lecture
            m_Serial.ReadBufferSize = m_ConfigSerial.iReadBufferSize;
            // Recopie de la taille du buffer d'écriture
            m_Serial.WriteBufferSize = m_ConfigSerial.iWriteBufferSize;
            m_Serial.DtrEnable = m_ConfigSerial.bDtrEnable;
            m_Serial.RtsEnable = m_ConfigSerial.bRtsEnable;
            m_Serial.ReadTimeout = m_ConfigSerial.iReadTimeOut;
            m_Serial.WriteTimeout = m_ConfigSerial.iWriteTimeOut;

            // Ouverture du port série
            m_Serial.Open();
        }
        public void close()
        {
            if (m_Serial.IsOpen)
            {
                m_Serial.Close();
            }
        }

        public void ClearSerial()
        {
            m_Serial.DiscardInBuffer();
            m_Serial.DiscardOutBuffer();
        }

        private void WriteSerial(String sCommande)
        {
            m_Serial.DiscardInBuffer();
            m_Serial.DiscardOutBuffer();
            m_Serial.Write(sCommande);
        }

        public void ReadSerial(Double dTimeout, out String sBufferRead)
        {
            int iSizeBuffer = m_ConfigSerial.iReadBufferSize;
            char[] arr = new char[iSizeBuffer];
            sBufferRead = "";
            bool bStop = false;
            TimeSpan ts = TimeSpan.MinValue;
            DateTime dtStart = DateTime.Now;
            int iNbData = 0;
            do
            {
                try
                {
                    iNbData = m_Serial.Read(arr, 0, iSizeBuffer - 1);
                    for (int i = 0; i < iNbData; i++)
                    {
                        sBufferRead += arr[i];
                    }
                }
                catch (TimeoutException)
                {
                    iNbData = 0;
                }

                //Gestion timeout
                ts = DateTime.Now - dtStart;
                if (ts.TotalMilliseconds > dTimeout)
                {
                    throw new TimeoutException("Le délai d'attente de la lecture sur le port série a expirée");
                }

                //Si on a déjà lu des données et qu'il n'y en a plus maintenant
                if (sBufferRead.Length != 0 && iNbData == 0)
                {
                    bStop = true;
                }
                System.Windows.Forms.Application.DoEvents();
            }
            while (!bStop);
        }

        public void WriteCom(String sCommand)
        {
            ClearSerial();
            WriteSerial(sCommand);
        }

        public void QueryCom(String sCommand, Double dTimeout, out String sBufferRead)
        {
            sBufferRead = String.Empty;

            ClearSerial();

            WriteSerial(sCommand);

            ReadSerial(dTimeout, "", "\r\n", out sBufferRead, true);
        }
        private void ReadSerial(Double dTimeOut, String sPrefix, String sSuffix, out String sBufferRead, Boolean bStopReadWhenSuffixFind)
        {
            sBufferRead = String.Empty;

            Boolean bSearchPrefix = !String.IsNullOrEmpty(sPrefix);
            Boolean bSearchSuffix = !String.IsNullOrEmpty(sSuffix);
            Boolean bPrefixFound = false;
            Boolean bSuffixFound = false;

            int iSizeBuffer = m_ConfigSerial.iReadBufferSize;
            char[] arr = new char[iSizeBuffer];
            sBufferRead = "";

            bool bStop = false;
            int iNbData = 0;

            TimeSpan ts = TimeSpan.MinValue;
            DateTime dtStart = DateTime.Now;

            do
            {
                try
                {
                    iNbData = m_Serial.Read(arr, 0, iSizeBuffer - 1);
                    for (int i = 0; i < iNbData; i++)
                    {
                        sBufferRead += arr[i];
                    }

                    //si on attend un préfixe que l'on a pas encore trouvé
                    if (bSearchPrefix && !bPrefixFound)
                    {
                        //regarde la présence du préfixe dans la chaine reçue
                        bPrefixFound = sBufferRead.Contains(sPrefix);
                    }

                    //si on attend un suffixe que l'on a pas encore trouvé
                    if (bSearchSuffix && !bSuffixFound)
                    {
                        //si pas de prefixe à chercher
                        if (!bSearchPrefix)
                        {
                            //regarde la présence du suffixe dans la chaine reçue
                            bSuffixFound = sBufferRead.Contains(sSuffix);
                        }
                        else
                        {
                            //si le préfixe a été trouvé
                            if (bPrefixFound)
                            {
                                //regarde la présence du suffixe dans la chaine reçue à partir de l'emplacement du préfixe
                                int iIndexPrefixe = sBufferRead.IndexOf(sPrefix);
                                if (sBufferRead.Substring(iIndexPrefixe).Contains(sSuffix))
                                {
                                    bSuffixFound = true;
                                }
                            }
                        }
                    }
                }
                catch (TimeoutException)
                {
                    iNbData = 0;
                }

                //Gestion timeout
                ts = DateTime.Now - dtStart;
                if (ts.TotalMilliseconds > dTimeOut)
                {
                    throw new TimeoutException("Le délai d'attente de la lecture sur le port série a expirée");
                }


                //Si déjà lu des données
                if (sBufferRead.Length != 0)
                {
                    //Gestion des cas de sortie de la boucle en fonction de la recherche des prefixe et/ou suffixe

                    //PAS de recherche de prefixe ET PAS de recherche de suffixe
                    if (!bSearchPrefix && !bSearchSuffix)
                    {
                        if (iNbData == 0)
                            bStop = true;
                    }
                    //Recherche de prefixe ET Pas de recherche de suffixe
                    else if (bSearchPrefix && !bSearchSuffix)
                    {
                        if (iNbData == 0 && bPrefixFound)
                            bStop = true;
                    }
                    //PAS de recherche de prefixe ET Recherche de suffixe
                    else if (!bSearchPrefix && bSearchSuffix)
                    {
                        if (bSuffixFound && (bStopReadWhenSuffixFind || iNbData == 0))
                            bStop = true;
                    }
                    //Recherche de prefixe ET Recherche de suffixe
                    else
                    {
                        if (bPrefixFound && bSuffixFound && (bStopReadWhenSuffixFind || iNbData == 0))
                            bStop = true;
                    }
                }
                System.Windows.Forms.Application.DoEvents();
            }
            while (!bStop);
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            //m_Serial.DataReceived += serialPort1_DataReceived;
            /*string value1 = "";
            string value2 = "";
            double voltage, impedance;

            QueryCom("*IDN?", 200, out value1);
            textBox1.Text = Convert.ToString(voltagevalue);
            ClearSerial();
            QueryCom("FUNCtion:IMPedance R", 1000, out value2);
            textBox2.Text = Convert.ToString(impedancevalue);
            ClearSerial();
            voltage = Convert.ToDouble(value1);
            impedance = Convert.ToDouble(value2);
            comparing(voltage, impedance);*/
            // example
            voltagevalue = 3.7;
            impedancevalue = 1.2;
            textBox1.Text = Convert.ToString(voltagevalue);
            textBox2.Text = Convert.ToString(impedancevalue);
            System.Threading.Thread.Sleep(500);
            comparing(voltagevalue, impedancevalue);
            ReadExistingExcel();
        }


        public double voltagevalue;
        public double impedancevalue;
        public string result = "";
        private void comparing(in double voltage, in double impedance)
        {
            // on compare entre 3.6 et 4 V pour la tension
            // on compare entre 1 et 2 ohm
            if (voltagevalue >= 3.60 && voltagevalue <= 4.10)
            {
                if (impedancevalue >= 1 && impedancevalue <= 2)
                {
                    pictureBox1.Image = new Bitmap(@"C:\Users\ayoubexo.CEBONGROUP\source\repos\BA6010\vert.PNG");
                    result = "PASS";
                }
                else
                {
                    pictureBox1.Image = new Bitmap(@"C:\Users\ayoubexo.CEBONGROUP\source\repos\BA6010\rouge.PNG");
                    result = "FAILED";
                }  
            }
            else
            {
                pictureBox1.Image = new Bitmap(@"C:\Users\ayoubexo.CEBONGROUP\source\repos\BA6010\rouge.PNG");
                result = "FAILED";
            }
        }

        public bool _Isactivated = true;


        public void ReadExistingExcel()
        {
            string path = @"C:\Users\ayoubexo.CEBONGROUP\source\repos\BA6010\BA6010\excel file result\Test result.xlsx";
            
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;
            oXL.DisplayAlerts = false;

            mWorkBook = oXL.Workbooks.Open(path, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //Get all the sheets in the workbook
            mWorkSheets = mWorkBook.Worksheets;
            
            //Get the allready exists sheet
            mWSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item("Feuil1");
            Microsoft.Office.Interop.Excel.Range range = mWSheet1.UsedRange;

            int colCount = range.Columns.Count;
            int rowCount = range.Rows.Count;

            // Appending to test result file
            mWSheet1.Cells[rowCount+1, 1] = DateTime.Now.ToString();
            mWSheet1.Cells[rowCount+1, 2] = voltagevalue;
            mWSheet1.Cells[rowCount+1, 3] = impedancevalue;
            mWSheet1.Cells[rowCount+1, 4] = result;

            mWorkBook.SaveAs(path, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
            false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            mWorkBook.Close();
            mWSheet1 = null;
            mWorkBook = null;
            oXL.Quit();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

    }
}
