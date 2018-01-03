using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Ports;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using Testsoft;

namespace stm32load {
    public partial class Stm32 {
        private const byte ACK = 0x79;
        private const byte NACK = 0x1F;
        private const byte EVERY_WRITE_LEN = 128;
        private const int EVERY_PACK_SIZE = EVERY_WRITE_LEN + 2;//2 is for length and checksum ;
        private const int EVERY_PAGE_BYTE_NUMBERS = 2048;

        private BinaryReader binFile;
        private FileStream fs;
        private byte tmp_ack = 0x00;
        private bool GetCMD_Flag = false;
        private int Rxcount = 0;
        private bool Extended_Erase = false;

        public Stm32() {
        }

        public void STM32DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e) {     //接收数据的原始函数
            int i;
            byte RXbyte = 0;
            //读取数据
            try {
                i = Form1.form1.myCOM.BytesToRead;
            } catch {
                return;
            }

            while (0 != i) {
                try {
                    RXbyte = (byte)Form1.form1.myCOM.ReadByte();
                    if (RXbyte == ACK) {
                        tmp_ack = RXbyte;
                    }
                    if (GetCMD_Flag) {
                        Rxcount++;
                        if (Rxcount == 10) {
                            if (RXbyte == 0x43) {
                                Extended_Erase = false;
                            } else if (RXbyte == 0x44) {
                                Extended_Erase = true;
                            }
                            GetCMD_Flag = false;
                        }
                    }
                    i--;
                } catch {
                    return;
                }
            }
        }

        //Send information to STM32.If STM32 return ACK,SendInfo's return
        //value will be true,else false.
        private bool SendInfo(byte[] info, int offset, int length) {
            if (Form1.form1.myCOM.IsOpen == false) return false;
            tmp_ack = NACK;
            Form1.form1.myCOM.Write(info, offset, length);
            Thread.Sleep(30);
            if (tmp_ack == ACK) return true;
            else return false;
        }

        private bool Connect() {
            byte[] cmd = new byte[1] { 0x7F };
            if (SendInfo(cmd, 0, 1) == false) return false;
            //cmd = new byte[2] { 0x02, 0xFD };
            //SendInfo(cmd, 0, 2);
            return true;
        }

        private bool GetCMD() {
            GetCMD_Flag = true;
            Rxcount = 0;
            byte[] cmd = new byte[2] { 0x00, 0xFF };
            if (SendInfo(cmd, 0, 2) == false) return false;
            Thread.Sleep(100);
            return true;
        }

        private double GetBootloaderVersion() {
            if (Form1.form1.myCOM.IsOpen == false) return 0;

            Form1.form1.myCOM.DiscardInBuffer();
            byte[] cmd = new byte[2] { 0x00, 0xFF };
            if (SendInfo(cmd, 0, 2) == false) return 0;

            int n = Form1.form1.myCOM.ReadByte();//the number of bytes to follow,except acks,
                                                 //in there ,n=11
            int version = Form1.form1.myCOM.ReadByte();//the version,the first of n

            int tmp;
            for (int i = 0; i < n; i++) { //read 11 times
                try {
                    tmp = Form1.form1.myCOM.ReadByte();
                } catch (Exception) {
                    return 0;
                }
            }
            if (Form1.form1.myCOM.ReadByte() == ACK) {
                return ((version >> 4) + (version & 0x0f) * 0.1);
            } else {
                return 0;
            }
        }

        private bool EraseAll() {
            //erase page number must more than 1!
            if (Form1.form1.myCOM.IsOpen == false) return false;

            //prepare data
            byte[] eraseCmd;
            byte[] eraseInfo;
            if (Extended_Erase) {
                eraseCmd = new byte[2] { 0x44, 0xBB };
                eraseInfo = new byte[3];
                eraseInfo[0] = 0xFF;
                eraseInfo[1] = 0xFF;
                eraseInfo[2] = 0x00;
            } else {
                eraseCmd = new byte[2] { 0x43, 0xBC };
                eraseInfo = new byte[2];
                eraseInfo[0] = 0xFF;
                eraseInfo[1] = 0x00;
            }

            //write erase cmd
            if (SendInfo(eraseCmd, 0, eraseCmd.Length) == false) return false;
            Thread.Sleep(100);

            //write erase cmd succ,do next!
            if (SendInfo(eraseInfo, 0, eraseInfo.Length) == false) return false;

            return true;
        }

        private bool EnableReadProtect() {
            byte[] ERcmd = new byte[2] { 0x82, 0x7D };
            if (SendInfo(ERcmd, 0, 2) == false) return false;
            Thread.Sleep(100);
            if (Connect() == false) {
                Form1.form1.Digital_Download_Message("The STM32 not avaliable,please check it!");
                return false;
            }
            return true;
        }

        private bool CancelReadProtect() {
            byte[] CRcmd = new byte[2] { 0x92, 0x6D };
            if (SendInfo(CRcmd, 0, 2) == false) return false;
            Thread.Sleep(100);
            if (Connect() == false) {
                Form1.form1.Digital_Download_Message("The STM32 not avaliable,please check it!");
                return false;
            }
            return true;
        }

        private bool CancelWriteProtect() {
            byte[] CWcmd = new byte[2] { 0x73, 0x8C };
            SendInfo(CWcmd, 0, 2);
            Thread.Sleep(100);
            if (Connect() == false) {
                Form1.form1.Digital_Download_Message("The STM32 not avaliable,please check it!");
                return false;
            }
            return true;
        }

        private bool Load(string fileName, UInt32 firstAddr) {
            string oldstring = "";
            int tick = 0;
            UInt32 addr = firstAddr;
            byte[] writeMemoryCmd = new byte[2] { 0x31, 0xCE };
            if (fileName.EndsWith(".bin") == false) {
                Form1.form1.Digital_Download_Message("请选择正确的.bin文件");
                return false;
            }
            try {
                fs = new FileStream(fileName, FileMode.Open);
                binFile = new BinaryReader(fs);
            } catch {
                Form1.form1.Digital_Download_Message("打开.bin文件失败，请检查文件的正确性");
                return false;
            }
            int binFileLen = (int)fs.Length;
            if (binFileLen == 0) {
                Form1.form1.Digital_Download_Message("文件内容为空!");
                return false;
            }

            int erasePageNumber = binFileLen / EVERY_PAGE_BYTE_NUMBERS;
            erasePageNumber = (binFileLen % EVERY_PAGE_BYTE_NUMBERS == 0 ? erasePageNumber : erasePageNumber + 1);
            if (CancelReadProtect() == false) {
                return false;
            }

            if (CancelWriteProtect() == false) {
                return false;
            }

            if (GetCMD() == false) {
                return false;
            }

            if (EraseAll() == false) { //erase page
                Form1.form1.Digital_Download_Message("全片擦除失败");
                return false;
            } else {
                Form1.form1.Digital_Download_Message("全片擦除成功");
            }
            Form1.form1.Digital_Download_Message("固件下载中，请稍候");
            byte[] buffer_addr = new byte[4 + 1];//1 is for checksum
            if (binFileLen <= EVERY_WRITE_LEN) {
                buffer_addr[0] = Bits.MsbOfUInt32(addr);
                buffer_addr[1] = Bits.ThreeByteOfUInt32(addr);
                buffer_addr[2] = Bits.TwoByteOfUInt32(addr);
                buffer_addr[3] = Bits.LsbOfUInt32(addr);
                buffer_addr[4] = Bits.CheckSum(buffer_addr, 0, 3);

                byte[] buffer = new byte[binFileLen + 2];//2 is for length and checkSum
                buffer[0] = Convert.ToByte(binFileLen - 1);
                for (int i = 0; i < binFileLen; i++) {
                    buffer[i + 1] = binFile.ReadByte();
                }
                buffer[binFileLen + 1] = Bits.CheckSum(buffer, 0, binFileLen);

                //add your code in there to send write memory cmd!!!
                if (SendInfo(writeMemoryCmd, 0, 2) == false) return false;
                //add your code in there to send address!!!
                if (SendInfo(buffer_addr, 0, 5) == false) return false;
                //in there ,add code to send the buffer!!!
                if (SendInfo(buffer, 0, EVERY_PACK_SIZE) == false) return false;
            } else {
                byte[] buffer = new byte[EVERY_PACK_SIZE];//2 is for length and checkSum
                int leftByteNumber = binFileLen;
                int sendIndex = 0;//sends index
                oldstring = Form1.form1.Digital_Get_Message();
                for (int i = 0; i < binFileLen / EVERY_WRITE_LEN; i++) { //every package
                    buffer_addr[0] = Bits.MsbOfUInt32(addr);
                    buffer_addr[1] = Bits.ThreeByteOfUInt32(addr);
                    buffer_addr[2] = Bits.TwoByteOfUInt32(addr);
                    buffer_addr[3] = Bits.LsbOfUInt32(addr);
                    buffer_addr[4] = Bits.CheckSum(buffer_addr, 0, 3);

                    buffer[0] = EVERY_WRITE_LEN - 1;//the first byte is buffer is length
                    for (int j = 1; j < EVERY_PACK_SIZE - 1; j++) {
                        //buffer[EVERY_PACK_SIZE-1] is for checkSum
                        buffer[j] = binFile.ReadByte();
                    }
                    //the last byte in buffer is checkSum
                    buffer[EVERY_PACK_SIZE - 1] = Bits.CheckSum(buffer, EVERY_WRITE_LEN + 1);

                    //add your code in there to send write memory cmd!!!
                    if (SendInfo(writeMemoryCmd, 0, 2) == false) return false;
                    //add your code in there to send address!!!
                    if (SendInfo(buffer_addr, 0, 5) == false) return false;
                    //in there ,add code to send the buffer!!!
                    if (SendInfo(buffer, 0, EVERY_PACK_SIZE) == false) return false;


                    //this code must in the last!!!
                    addr += EVERY_WRITE_LEN;
                    sendIndex += EVERY_WRITE_LEN;
                    leftByteNumber -= EVERY_WRITE_LEN;
                    Form1.form1.Digital_Download_Message(oldstring, tick++);
                }
                if (binFileLen % EVERY_WRITE_LEN != 0) {
                    buffer_addr[0] = Bits.MsbOfUInt32(addr);
                    buffer_addr[1] = Bits.ThreeByteOfUInt32(addr);
                    buffer_addr[2] = Bits.TwoByteOfUInt32(addr);
                    buffer_addr[3] = Bits.LsbOfUInt32(addr);
                    buffer_addr[4] = Bits.CheckSum(buffer_addr, 0, 3);

                    //add lefts
                    int leftPackSize = (int)leftByteNumber + 2;//2 is for length and checkSum
                    byte[] leftBuffer = new byte[leftPackSize];

                    leftBuffer[0] = Convert.ToByte(leftByteNumber - 1);
                    for (int i = 1; i < leftPackSize - 1; i++) {
                        leftBuffer[i] = binFile.ReadByte();
                    }
                    leftBuffer[leftPackSize - 1] = Bits.CheckSum(leftBuffer, leftByteNumber + 1);

                    //add your code in there to send write memory cmd!!!
                    if (SendInfo(writeMemoryCmd, 0, 2) == false) return false;
                    //add your code in there to send address!!!
                    if (SendInfo(buffer_addr, 0, 5) == false) return false;
                    //in there ,add code to send the buffer!!!
                    if (SendInfo(leftBuffer, 0, leftPackSize) == false) return false;
                }
            }
            if (EnableReadProtect() == false) {
                return false;
            }
            return true;//write memory succ
        }

        public void STM32Load() {
            if (Connect() == false) {
                Form1.form1.Digital_Download_Message("CPU识别失败");
                return;
            } else {
                Form1.form1.Digital_Download_Message("CPU识别成功");
            }

            Form1.form1.BIN_Path = Directory.GetCurrentDirectory() + "\\被测物固件\\firmware.bin";
            if (File.Exists(Form1.form1.BIN_Path) == false) {
                MessageBox.Show("清先选择要下载的.bin文件");
                Form1.form1.SelectBIN();
            }
            string binFileName = Form1.form1.BIN_Path;
            if (Load(binFileName, 0x08000000) == false) {
                Form1.form1.Digital_Download_Message("固件下载失败!");
            } else {
                Form1.form1.Digital_Download_Message("固件下载成功!");
            }
            try {
                fs.Close();
                binFile.Close();
            } catch { }
        }

    }//end class
}//end namespace
