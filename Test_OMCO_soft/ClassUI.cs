using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Windows.Forms;
using System.IO;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using Aspose.Cells;
using stm32load;

namespace Testsoft {
    public unsafe partial class Form1 : Form {
        public string BIN_Path = "";
        public Stm32 stmLoad = new Stm32();
        public ClassCompor myCOM = new ClassCompor();
        public Stream myCOMPortStream;

        public static CheckBox[] checkBox_CH_Result_Sel = new CheckBox[CHANNELS_NUMBER];
        public static TextBox[] textBox_CH_Result = new TextBox[CHANNELS_NUMBER];
        public static CheckBox[] checkBox_CH_Switch_Sel = new CheckBox[CHANNELS_NUMBER];
        public static TextBox[] textBox_CH_Switch = new TextBox[CHANNELS_NUMBER];
        public static CheckBox[] checkBox_CH_Type_Sel = new CheckBox[CHANNELS_NUMBER];
        public static TextBox[] textBox_CH_Type = new TextBox[CHANNELS_NUMBER];
        public static CheckBox[] checkBox_CH_TC_CJC_Mode_Sel = new CheckBox[CHANNELS_NUMBER];
        public static ComboBox[] comboBox_CH_TC_CJC_Mode = new ComboBox[CHANNELS_NUMBER];
        public static CheckBox[] checkBox_CH_TC_Sensor_Sel = new CheckBox[CHANNELS_NUMBER];
        public static ComboBox[] comboBox_CH_TC_Sensor = new ComboBox[CHANNELS_NUMBER];
        public static CheckBox[] checkBox_CH_TC_CJC_T_Sel = new CheckBox[CHANNELS_NUMBER];
        public static TextBox[] textBox_CH_TC_CJC_T = new TextBox[CHANNELS_NUMBER];
        public static CheckBox[] checkBox_CH_Original_Value_Sel = new CheckBox[CHANNELS_NUMBER];
        public static TextBox[] textBox_CH_Original_Value = new TextBox[CHANNELS_NUMBER];
        public static CheckBox[] checkBox_Cal_Phy_CH_Sel = new CheckBox[CHANNELS_NUMBER];
        public static TextBox[] textBox_Cal_Phy = new TextBox[CHANNELS_NUMBER];
        public static CheckBox[] checkBox_Get_Phy_CH_Sel = new CheckBox[CHANNELS_NUMBER];
        public static TextBox[] textBox_Get_Phy = new TextBox[CHANNELS_NUMBER];
        public static CheckBox[] checkBox_CH_Cal_Par_Sel = new CheckBox[CHANNELS_NUMBER];
        public static TextBox[,] textBox_CH_Cal_Par = new TextBox[2, CHANNELS_NUMBER];
        public static TextBox[] textBox_Infrared = new TextBox[4];
        public static TextBox[] textBox_Digital_CHV = new TextBox[8];
        public static TextBox[] textBox_Digital_CHI = new TextBox[8];

        public void softinitial() {   //初始化软件刚打开时界面要加载的值
            //comboBoxPort.Text = "COM1";
            //comboBoxBaud.Text = "115200";
            //comboBoxVerify.Text = "NONE";
            //comboBoxData.Text = "8";
            //comboBoxStop.Text = "1";
            Get_Configuration();

            buttonOpenCOM.BackColor = Color.LightBlue;
            button_OpenDigital.BackColor = Color.LightBlue;
            button_OpenAnalog.BackColor = Color.LightBlue;
            radioButton_ADI.Checked = true;
            radioButton_Res1.Checked = true;
            radioButton_G1.Checked = true;
            textBox_BIN_Path.Enabled = false;
            tabControl1.TabPages.Remove(tabPage7);
            tabControl1.TabPages.Remove(tabPage6);
            tabControl1.TabPages.Remove(tabPage1);
            tabControl1.TabPages.Remove(tabPage3);
            tabControl1.TabPages.Remove(tabPage9);
            tabControl1.TabPages.Remove(tabPage2);
            tabControl1.TabPages.Remove(tabPage4);
            tabControl1.TabPages.Remove(tabPage5);
            tabControl1.Enabled = false;
            checkBox_CH_Result_Sel[0] = checkBox_CH_Result_Sel1;
            checkBox_CH_Result_Sel[1] = checkBox_CH_Result_Sel2;
            checkBox_CH_Result_Sel[2] = checkBox_CH_Result_Sel3;
            checkBox_CH_Result_Sel[3] = checkBox_CH_Result_Sel4;
            checkBox_CH_Result_Sel[4] = checkBox_CH_Result_Sel5;
            checkBox_CH_Result_Sel[5] = checkBox_CH_Result_Sel6;
            checkBox_CH_Result_Sel[6] = checkBox_CH_Result_Sel7;
            checkBox_CH_Result_Sel[7] = checkBox_CH_Result_Sel8;
            textBox_CH_Result[0] = textBox_CH_Result_1;
            textBox_CH_Result[1] = textBox_CH_Result_2;
            textBox_CH_Result[2] = textBox_CH_Result_3;
            textBox_CH_Result[3] = textBox_CH_Result_4;
            textBox_CH_Result[4] = textBox_CH_Result_5;
            textBox_CH_Result[5] = textBox_CH_Result_6;
            textBox_CH_Result[6] = textBox_CH_Result_7;
            textBox_CH_Result[7] = textBox_CH_Result_8;
            checkBox_CH_Switch_Sel[0] = checkBox_CH_Switch_Sel1;
            checkBox_CH_Switch_Sel[1] = checkBox_CH_Switch_Sel2;
            checkBox_CH_Switch_Sel[2] = checkBox_CH_Switch_Sel3;
            checkBox_CH_Switch_Sel[3] = checkBox_CH_Switch_Sel4;
            checkBox_CH_Switch_Sel[4] = checkBox_CH_Switch_Sel5;
            checkBox_CH_Switch_Sel[5] = checkBox_CH_Switch_Sel6;
            checkBox_CH_Switch_Sel[6] = checkBox_CH_Switch_Sel7;
            checkBox_CH_Switch_Sel[7] = checkBox_CH_Switch_Sel8;
            textBox_CH_Switch[0] = textBox_CH_Switch_1;
            textBox_CH_Switch[1] = textBox_CH_Switch_2;
            textBox_CH_Switch[2] = textBox_CH_Switch_3;
            textBox_CH_Switch[3] = textBox_CH_Switch_4;
            textBox_CH_Switch[4] = textBox_CH_Switch_5;
            textBox_CH_Switch[5] = textBox_CH_Switch_6;
            textBox_CH_Switch[6] = textBox_CH_Switch_7;
            textBox_CH_Switch[7] = textBox_CH_Switch_8;
            checkBox_CH_Type_Sel[0] = checkBox_CH_Type_Sel1;
            checkBox_CH_Type_Sel[1] = checkBox_CH_Type_Sel2;
            checkBox_CH_Type_Sel[2] = checkBox_CH_Type_Sel3;
            checkBox_CH_Type_Sel[3] = checkBox_CH_Type_Sel4;
            checkBox_CH_Type_Sel[4] = checkBox_CH_Type_Sel5;
            checkBox_CH_Type_Sel[5] = checkBox_CH_Type_Sel6;
            checkBox_CH_Type_Sel[6] = checkBox_CH_Type_Sel7;
            checkBox_CH_Type_Sel[7] = checkBox_CH_Type_Sel8;
            textBox_CH_Type[0] = textBox_CH_Type_1;
            textBox_CH_Type[1] = textBox_CH_Type_2;
            textBox_CH_Type[2] = textBox_CH_Type_3;
            textBox_CH_Type[3] = textBox_CH_Type_4;
            textBox_CH_Type[4] = textBox_CH_Type_5;
            textBox_CH_Type[5] = textBox_CH_Type_6;
            textBox_CH_Type[6] = textBox_CH_Type_7;
            textBox_CH_Type[7] = textBox_CH_Type_8;
            checkBox_CH_TC_CJC_Mode_Sel[0] = checkBox_CH_TC_CJC_Mode_Sel1;
            checkBox_CH_TC_CJC_Mode_Sel[1] = checkBox_CH_TC_CJC_Mode_Sel2;
            checkBox_CH_TC_CJC_Mode_Sel[2] = checkBox_CH_TC_CJC_Mode_Sel3;
            checkBox_CH_TC_CJC_Mode_Sel[3] = checkBox_CH_TC_CJC_Mode_Sel4;
            checkBox_CH_TC_CJC_Mode_Sel[4] = checkBox_CH_TC_CJC_Mode_Sel5;
            checkBox_CH_TC_CJC_Mode_Sel[5] = checkBox_CH_TC_CJC_Mode_Sel6;
            checkBox_CH_TC_CJC_Mode_Sel[6] = checkBox_CH_TC_CJC_Mode_Sel7;
            checkBox_CH_TC_CJC_Mode_Sel[7] = checkBox_CH_TC_CJC_Mode_Sel8;
            comboBox_CH_TC_CJC_Mode[0] = comboBox_CH_TC_CJC_Mode_1;
            comboBox_CH_TC_CJC_Mode[1] = comboBox_CH_TC_CJC_Mode_2;
            comboBox_CH_TC_CJC_Mode[2] = comboBox_CH_TC_CJC_Mode_3;
            comboBox_CH_TC_CJC_Mode[3] = comboBox_CH_TC_CJC_Mode_4;
            comboBox_CH_TC_CJC_Mode[4] = comboBox_CH_TC_CJC_Mode_5;
            comboBox_CH_TC_CJC_Mode[5] = comboBox_CH_TC_CJC_Mode_6;
            comboBox_CH_TC_CJC_Mode[6] = comboBox_CH_TC_CJC_Mode_7;
            comboBox_CH_TC_CJC_Mode[7] = comboBox_CH_TC_CJC_Mode_8;
            checkBox_CH_TC_Sensor_Sel[0] = checkBox_CH_TC_Sensor_Sel1;
            checkBox_CH_TC_Sensor_Sel[1] = checkBox_CH_TC_Sensor_Sel2;
            checkBox_CH_TC_Sensor_Sel[2] = checkBox_CH_TC_Sensor_Sel3;
            checkBox_CH_TC_Sensor_Sel[3] = checkBox_CH_TC_Sensor_Sel4;
            checkBox_CH_TC_Sensor_Sel[4] = checkBox_CH_TC_Sensor_Sel5;
            checkBox_CH_TC_Sensor_Sel[5] = checkBox_CH_TC_Sensor_Sel6;
            checkBox_CH_TC_Sensor_Sel[6] = checkBox_CH_TC_Sensor_Sel7;
            checkBox_CH_TC_Sensor_Sel[7] = checkBox_CH_TC_Sensor_Sel8;
            comboBox_CH_TC_Sensor[0] = comboBox_CH_TC_Sensor_1;
            comboBox_CH_TC_Sensor[1] = comboBox_CH_TC_Sensor_2;
            comboBox_CH_TC_Sensor[2] = comboBox_CH_TC_Sensor_3;
            comboBox_CH_TC_Sensor[3] = comboBox_CH_TC_Sensor_4;
            comboBox_CH_TC_Sensor[4] = comboBox_CH_TC_Sensor_5;
            comboBox_CH_TC_Sensor[5] = comboBox_CH_TC_Sensor_6;
            comboBox_CH_TC_Sensor[6] = comboBox_CH_TC_Sensor_7;
            comboBox_CH_TC_Sensor[7] = comboBox_CH_TC_Sensor_8;
            checkBox_CH_TC_CJC_T_Sel[0] = checkBox_CH_TC_CJC_T_Sel1;
            checkBox_CH_TC_CJC_T_Sel[1] = checkBox_CH_TC_CJC_T_Sel2;
            checkBox_CH_TC_CJC_T_Sel[2] = checkBox_CH_TC_CJC_T_Sel3;
            checkBox_CH_TC_CJC_T_Sel[3] = checkBox_CH_TC_CJC_T_Sel4;
            checkBox_CH_TC_CJC_T_Sel[4] = checkBox_CH_TC_CJC_T_Sel5;
            checkBox_CH_TC_CJC_T_Sel[5] = checkBox_CH_TC_CJC_T_Sel6;
            checkBox_CH_TC_CJC_T_Sel[6] = checkBox_CH_TC_CJC_T_Sel7;
            checkBox_CH_TC_CJC_T_Sel[7] = checkBox_CH_TC_CJC_T_Sel8;
            textBox_CH_TC_CJC_T[0] = textBox_CH_TC_CJC_T_1;
            textBox_CH_TC_CJC_T[1] = textBox_CH_TC_CJC_T_2;
            textBox_CH_TC_CJC_T[2] = textBox_CH_TC_CJC_T_3;
            textBox_CH_TC_CJC_T[3] = textBox_CH_TC_CJC_T_4;
            textBox_CH_TC_CJC_T[4] = textBox_CH_TC_CJC_T_5;
            textBox_CH_TC_CJC_T[5] = textBox_CH_TC_CJC_T_6;
            textBox_CH_TC_CJC_T[6] = textBox_CH_TC_CJC_T_7;
            textBox_CH_TC_CJC_T[7] = textBox_CH_TC_CJC_T_8;
            checkBox_CH_Original_Value_Sel[0] = checkBox_CH_Original_Value_Sel1;
            checkBox_CH_Original_Value_Sel[1] = checkBox_CH_Original_Value_Sel2;
            checkBox_CH_Original_Value_Sel[2] = checkBox_CH_Original_Value_Sel3;
            checkBox_CH_Original_Value_Sel[3] = checkBox_CH_Original_Value_Sel4;
            checkBox_CH_Original_Value_Sel[4] = checkBox_CH_Original_Value_Sel5;
            checkBox_CH_Original_Value_Sel[5] = checkBox_CH_Original_Value_Sel6;
            checkBox_CH_Original_Value_Sel[6] = checkBox_CH_Original_Value_Sel7;
            checkBox_CH_Original_Value_Sel[7] = checkBox_CH_Original_Value_Sel8;
            textBox_CH_Original_Value[0] = textBox_CH_Original_Value_1;
            textBox_CH_Original_Value[1] = textBox_CH_Original_Value_2;
            textBox_CH_Original_Value[2] = textBox_CH_Original_Value_3;
            textBox_CH_Original_Value[3] = textBox_CH_Original_Value_4;
            textBox_CH_Original_Value[4] = textBox_CH_Original_Value_5;
            textBox_CH_Original_Value[5] = textBox_CH_Original_Value_6;
            textBox_CH_Original_Value[6] = textBox_CH_Original_Value_7;
            textBox_CH_Original_Value[7] = textBox_CH_Original_Value_8;
            checkBox_Cal_Phy_CH_Sel[0] = checkBox_Cal_Phy_CH_Sel1;
            checkBox_Cal_Phy_CH_Sel[1] = checkBox_Cal_Phy_CH_Sel2;
            checkBox_Cal_Phy_CH_Sel[2] = checkBox_Cal_Phy_CH_Sel3;
            checkBox_Cal_Phy_CH_Sel[3] = checkBox_Cal_Phy_CH_Sel4;
            checkBox_Cal_Phy_CH_Sel[4] = checkBox_Cal_Phy_CH_Sel5;
            checkBox_Cal_Phy_CH_Sel[5] = checkBox_Cal_Phy_CH_Sel6;
            checkBox_Cal_Phy_CH_Sel[6] = checkBox_Cal_Phy_CH_Sel7;
            checkBox_Cal_Phy_CH_Sel[7] = checkBox_Cal_Phy_CH_Sel8;
            textBox_Cal_Phy[0] = textBox_Cal_Phy_1;
            textBox_Cal_Phy[1] = textBox_Cal_Phy_2;
            textBox_Cal_Phy[2] = textBox_Cal_Phy_3;
            textBox_Cal_Phy[3] = textBox_Cal_Phy_4;
            textBox_Cal_Phy[4] = textBox_Cal_Phy_5;
            textBox_Cal_Phy[5] = textBox_Cal_Phy_6;
            textBox_Cal_Phy[6] = textBox_Cal_Phy_7;
            textBox_Cal_Phy[7] = textBox_Cal_Phy_8;
            checkBox_Get_Phy_CH_Sel[0] = checkBox_Get_Phy_CH_Sel1;
            checkBox_Get_Phy_CH_Sel[1] = checkBox_Get_Phy_CH_Sel2;
            checkBox_Get_Phy_CH_Sel[2] = checkBox_Get_Phy_CH_Sel3;
            checkBox_Get_Phy_CH_Sel[3] = checkBox_Get_Phy_CH_Sel4;
            checkBox_Get_Phy_CH_Sel[4] = checkBox_Get_Phy_CH_Sel5;
            checkBox_Get_Phy_CH_Sel[5] = checkBox_Get_Phy_CH_Sel6;
            checkBox_Get_Phy_CH_Sel[6] = checkBox_Get_Phy_CH_Sel7;
            checkBox_Get_Phy_CH_Sel[7] = checkBox_Get_Phy_CH_Sel8;
            textBox_Get_Phy[0] = textBox_Get_Phy_1;
            textBox_Get_Phy[1] = textBox_Get_Phy_2;
            textBox_Get_Phy[2] = textBox_Get_Phy_3;
            textBox_Get_Phy[3] = textBox_Get_Phy_4;
            textBox_Get_Phy[4] = textBox_Get_Phy_5;
            textBox_Get_Phy[5] = textBox_Get_Phy_6;
            textBox_Get_Phy[6] = textBox_Get_Phy_7;
            textBox_Get_Phy[7] = textBox_Get_Phy_8;
            checkBox_CH_Cal_Par_Sel[0] = checkBox_CH_Cal_Par_Sel1;
            checkBox_CH_Cal_Par_Sel[1] = checkBox_CH_Cal_Par_Sel2;
            checkBox_CH_Cal_Par_Sel[2] = checkBox_CH_Cal_Par_Sel3;
            checkBox_CH_Cal_Par_Sel[3] = checkBox_CH_Cal_Par_Sel4;
            checkBox_CH_Cal_Par_Sel[4] = checkBox_CH_Cal_Par_Sel5;
            checkBox_CH_Cal_Par_Sel[5] = checkBox_CH_Cal_Par_Sel6;
            checkBox_CH_Cal_Par_Sel[6] = checkBox_CH_Cal_Par_Sel7;
            checkBox_CH_Cal_Par_Sel[7] = checkBox_CH_Cal_Par_Sel8;
            textBox_CH_Cal_Par[0, 0] = textBox_CH_Cal_ParA_1;
            textBox_CH_Cal_Par[0, 1] = textBox_CH_Cal_ParA_2;
            textBox_CH_Cal_Par[0, 2] = textBox_CH_Cal_ParA_3;
            textBox_CH_Cal_Par[0, 3] = textBox_CH_Cal_ParA_4;
            textBox_CH_Cal_Par[0, 4] = textBox_CH_Cal_ParA_5;
            textBox_CH_Cal_Par[0, 5] = textBox_CH_Cal_ParA_6;
            textBox_CH_Cal_Par[0, 6] = textBox_CH_Cal_ParA_7;
            textBox_CH_Cal_Par[0, 7] = textBox_CH_Cal_ParA_8;
            textBox_CH_Cal_Par[1, 0] = textBox_CH_Cal_ParB_1;
            textBox_CH_Cal_Par[1, 1] = textBox_CH_Cal_ParB_2;
            textBox_CH_Cal_Par[1, 2] = textBox_CH_Cal_ParB_3;
            textBox_CH_Cal_Par[1, 3] = textBox_CH_Cal_ParB_4;
            textBox_CH_Cal_Par[1, 4] = textBox_CH_Cal_ParB_5;
            textBox_CH_Cal_Par[1, 5] = textBox_CH_Cal_ParB_6;
            textBox_CH_Cal_Par[1, 6] = textBox_CH_Cal_ParB_7;
            textBox_CH_Cal_Par[1, 7] = textBox_CH_Cal_ParB_8;
            textBox_Infrared[0] = textBox_Infrared1;
            textBox_Infrared[1] = textBox_Infrared2;
            textBox_Infrared[2] = textBox_Infrared3;
            textBox_Infrared[3] = textBox_Infrared4;
            textBox_Digital_CHV[0] = textBox_Digital_CHV_1;
            textBox_Digital_CHV[1] = textBox_Digital_CHV_2;
            textBox_Digital_CHV[2] = textBox_Digital_CHV_3;
            textBox_Digital_CHV[3] = textBox_Digital_CHV_4;
            textBox_Digital_CHV[4] = textBox_Digital_CHV_5;
            textBox_Digital_CHV[5] = textBox_Digital_CHV_6;
            textBox_Digital_CHV[6] = textBox_Digital_CHV_7;
            textBox_Digital_CHV[7] = textBox_Digital_CHV_8;
            textBox_Digital_CHI[0] = textBox_Digital_CHI_1;
            textBox_Digital_CHI[1] = textBox_Digital_CHI_2;
            textBox_Digital_CHI[2] = textBox_Digital_CHI_3;
            textBox_Digital_CHI[3] = textBox_Digital_CHI_4;
            textBox_Digital_CHI[4] = textBox_Digital_CHI_5;
            textBox_Digital_CHI[5] = textBox_Digital_CHI_6;
            textBox_Digital_CHI[6] = textBox_Digital_CHI_7;
            textBox_Digital_CHI[7] = textBox_Digital_CHI_8;
        }

        //初使化并打开端口,add in v1.83
        public bool iniCOM() {
            if (myCOM.IsOpen == true) {
                try {
                    myCOMPortStream.Close();
                } catch {
                    return (false);
                }
                try {
                    myCOM.Close();
                } catch {
                    return (false);
                }
                return (false);
            } else {
                if (myCOM.PortName != comboBoxPort.Text) {
                    myCOM.PortName = comboBoxPort.Text;
                }
                //波特率设置
                //myCOM.BaudRate = Convert.ToInt32(comboBoxBaud.Text);
                myCOM.BaudRate = 115200;
                //数据位设置
                //myCOM.DataBits = Convert.ToInt32(comboBoxData.Text);
                myCOM.DataBits = 8;
                //停止位设置
                //if (comboBoxStop.Text == "1")
                //{
                //    myCOM.StopBits = System.IO.Ports.StopBits.One;
                //}
                //else if (comboBoxStop.Text == "1.5")
                //{
                //    myCOM.StopBits = System.IO.Ports.StopBits.OnePointFive;
                //}
                //else if (comboBoxStop.Text == "2")
                //{
                //    myCOM.StopBits = System.IO.Ports.StopBits.Two;
                //}
                myCOM.StopBits = System.IO.Ports.StopBits.One;
                //设置奇偶校验协议
                //if (comboBoxVerify.Text == "NONE")
                //{
                //    myCOM.Parity = System.IO.Ports.Parity.None; ;
                //}
                //else if (comboBoxVerify.Text == "ODD")
                //{
                //    myCOM.Parity = System.IO.Ports.Parity.Odd;
                //}
                //else if (comboBoxVerify.Text == "EVEN")
                //{
                //    myCOM.Parity = System.IO.Ports.Parity.Even;
                //}
                //else if (comboBoxVerify.Text == "MARK")
                //{
                //    myCOM.Parity = System.IO.Ports.Parity.Mark;
                //}
                //else if (comboBoxVerify.Text == "SPACE")
                //{
                //    myCOM.Parity = System.IO.Ports.Parity.Space;
                //}
                myCOM.Parity = System.IO.Ports.Parity.None;

                myCOM.DtrEnable = true;     //mod in v3.22
                myCOM.RtsEnable = true;    //mod in v3.22

                myCOM.DataReceived += DataReceived;

                try {
                    myCOM.Open();
                    myCOMPortStream = myCOM.BaseStream;
                } catch {
                    MessageBoxERRwarning(4);
                    return (false);
                }
                return (true);
            }
        }

        private void DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e) {      //接收数据的原始函数
            if (DIGITAL_DATA.DOWNLOAD == 0) {
                stmLoad.STM32DataReceived(sender, e);
                return;
            }
            int i;
            byte rxbyte;
            //读取数据
            try {
                i = myCOM.BytesToRead;
            } catch {
                return;
            }

            while (0 != i) {
                try {
                    rxbyte = (byte)myCOM.ReadByte();
                    i--;
                } catch {
                    return;
                }
                fixed (PROTOCOL_FRAME_TypeDef* temp_p = &g_Protocol_Buff) {
                    g_Protocol_Buff_p = temp_p;
                    ReceiveFrameAnalysis(rxbyte);      //解析接收到的数据
                }
            }
        }

        //显示接收数据
        public void DisplayRXData(int bytenumber, byte[] RXBUFbytearray) {
            byte i;
            string temps = "";
            for (i = 0; i < bytenumber; i++) {
                temps += RXBUFbytearray[i].ToString("X2") + " ";
            }
            richTextBoxDataRecieve.Text = temps + "\r\n" + "\r\n" + richTextBoxDataRecieve.Text;    //本次接收的数据
        }

        public void MessageBoxERRwarning(int ERRnumber) {
            if (ERRnumber == 1) {
                MessageBox.Show("串口未打开，请检查串口连接");
            } else if (ERRnumber == 2) {
                MessageBox.Show("数据发送失败，请检查串口连接");
            } else if (ERRnumber == 3) {
                MessageBox.Show("请输入正确的值");
            } else if (ERRnumber == 4) {
                MessageBox.Show("串口打开失败，请检查端口是否存在或占用");
            } else if (ERRnumber == 5) {
                MessageBox.Show("TC冷端模式错误");
            } else if (ERRnumber == 6) {
                MessageBox.Show("输入值超出范围-327.68~327.67");
            } else if (ERRnumber == 7) {
                MessageBox.Show("请输入正确的校验和");
            } else if (ERRnumber == 8) {
                MessageBox.Show("输入值超出范围-3276.8~3276.7");
            } else if (ERRnumber == 9) {
                MessageBox.Show("请检查USB连接");
            } else if (ERRnumber == 10) {
                MessageBox.Show("输入值超出范围-10~10");
            } else if (ERRnumber == 11) {
                MessageBox.Show("输入值超出范围0~20");
            }
        }

        public bool SendData_New(byte[] TXBUFbytearray, byte offset, int totallen) {     //发送数据的原始函数
            int i = 0;
            string temps = "";

            if (myCOM.IsOpen == true) {
                try {
                    myCOM.Write(TXBUFbytearray, offset, totallen);
                    for (i = 0; i < totallen; i++) {//把要发送的数据转换成字符串
                        temps += TXBUFbytearray[i].ToString("X2") + " ";
                    }
                    richTextBoxDataSend.Text = temps + "\r\n" + "\r\n" + richTextBoxDataSend.Text;    //本次发送的数据
                    return true;
                } catch {
                    try {
                        myCOMPortStream.Close();
                    } catch {
                        ;
                    }
                    try {
                        myCOM.Close();
                    } catch {
                        ;
                    }
                    MessageBoxERRwarning(2);
                    return false;
                }
            } else {
                MessageBoxERRwarning(1);
                return false;
            }
        }

        public ushort Make_CH_Mask(CheckBox[] checkBox_CH) {
            int i = 0;
            ushort mask = 0;
            for (i = 0; i < CHANNELS_NUMBER; i++) {
                if (checkBox_CH[i].Checked == true) {
                    mask |= (ushort)(1 << i);
                }
            }
            return mask;
        }

        public void Master_Send() {
            toolStripStatusLabelCOM.Text = "";
            g_Protocol_Buff_p->Head[0] = 0xAA;
            g_Protocol_Buff_p->Head[1] = 0x55;
            g_Protocol_Buff_p->Station_NUM = 0x00;
            ushort CRC_VERIFY_DATA_LEN = (ushort)(FOREPART_LEN - 2 + *(ushort*)g_Protocol_Buff_p->Data_LEN);
            ushort crc = GetCRC(&g_Protocol_Buff_p->Station_NUM, CRC_VERIFY_DATA_LEN);
            *(ushort*)g_Protocol_Buff_p->Crc = crc;
            byte[] bySendBuffer = new byte[512];
            int totallen = FrameSendBuffer(ref bySendBuffer);
            if (SendData_New(bySendBuffer, 0, totallen)) {
                Thread.Sleep(300);
                if (toolStripStatusLabelCOM.Text == "") {
                    toolStripStatusLabelCOM.Text = "应答失败";
                }
            }
        }

        public void Get_Software_Ver() {
            fixed (PROTOCOL_FRAME_TypeDef* p = &g_Protocol_Buff) {
                g_Protocol_Buff_p = p;
                g_Protocol_Buff_p->Function_CMD = (byte)CMD_TypeDef.CMD_SOFTWARE_VERSION;
                g_Protocol_Buff_p->RW_CMD = READ_CMD;
                *(ushort*)g_Protocol_Buff_p->Data_LEN = 0;
                Master_Send();
            }
        }

        public void Get_Station_Num() {
            fixed (PROTOCOL_FRAME_TypeDef* p = &g_Protocol_Buff) {
                g_Protocol_Buff_p = p;
                g_Protocol_Buff_p->Function_CMD = (byte)CMD_TypeDef.CMD_STATION_NUM;
                g_Protocol_Buff_p->RW_CMD = READ_CMD;
                *(ushort*)g_Protocol_Buff_p->Data_LEN = 0;
                Master_Send();
            }
        }

        public void Get_Calibration_State() {
            fixed (PROTOCOL_FRAME_TypeDef* p = &g_Protocol_Buff) {
                g_Protocol_Buff_p = p;
                g_Protocol_Buff_p->Function_CMD = (byte)CMD_TypeDef.CMD_CALIBRATION_STATE;
                g_Protocol_Buff_p->RW_CMD = READ_CMD;
                *(ushort*)g_Protocol_Buff_p->Data_LEN = 0;
                Master_Send();
            }
        }

        public void Set_Station_Num() {
            fixed (PROTOCOL_FRAME_TypeDef* p = &g_Protocol_Buff) {
                g_Protocol_Buff_p = p;
                g_Protocol_Buff_p->Function_CMD = (byte)CMD_TypeDef.CMD_STATION_NUM;
                g_Protocol_Buff_p->RW_CMD = WRITE_CMD;
                *(ushort*)g_Protocol_Buff_p->Data_LEN = 1;
                try {
                    g_Protocol_Buff_p->Data[0] = Convert.ToByte(textBox_Station_Num.Text);
                } catch {
                    MessageBoxERRwarning(3);
                }
                Master_Send();
            }
        }

        public void Get_Device_Type() {
            fixed (PROTOCOL_FRAME_TypeDef* temp_p = &g_Protocol_Buff) {
                g_Protocol_Buff_p = temp_p;
                g_Protocol_Buff_p->Function_CMD = (byte)CMD_TypeDef.CMD_DEVICE_TYPE;
                g_Protocol_Buff_p->RW_CMD = READ_CMD;
                *(ushort*)g_Protocol_Buff_p->Data_LEN = 0;
                Master_Send();
            }
        }

        public void Set_CH_Result() {
            fixed (PROTOCOL_FRAME_TypeDef* temp_p = &g_Protocol_Buff) {
                g_Protocol_Buff_p = temp_p;
                g_Protocol_Buff_p->Function_CMD = (byte)CMD_TypeDef.CMD_CH_RESULT;
                g_Protocol_Buff_p->RW_CMD = WRITE_CMD;
                byte* p = g_Protocol_Buff_p->Data;
                ushort ch_mask = 0;
                ch_mask = Make_CH_Mask(checkBox_CH_Result_Sel);
                *(ushort*)p = ch_mask;
                p += 2;
                try {
                    byte i = 0;
                    for (i = 0; i < CHANNELS_NUMBER; i++) {
                        if (0x01 == ((ch_mask >> i) & 0x01)) {
                            *(ushort*)p = (ushort)(Convert.ToDouble(textBox_CH_Result[i].Text) * 100);
                            p += 2;
                        }
                    }
                    *(ushort*)g_Protocol_Buff_p->Data_LEN = (ushort)(p - g_Protocol_Buff_p->Data);
                    Master_Send();
                } catch {
                    MessageBoxERRwarning(3);
                }
            }
        }

        public void Get_CH_Result() {
            fixed (PROTOCOL_FRAME_TypeDef* temp_p = &g_Protocol_Buff) {
                g_Protocol_Buff_p = temp_p;
                g_Protocol_Buff_p->Function_CMD = (byte)CMD_TypeDef.CMD_CH_RESULT;
                g_Protocol_Buff_p->RW_CMD = READ_CMD;
                byte* p = g_Protocol_Buff_p->Data;
                ushort ch_mask = 0;
                ch_mask = Make_CH_Mask(checkBox_CH_Result_Sel);
                *(ushort*)p = ch_mask;
                p += 2;
                *(ushort*)g_Protocol_Buff_p->Data_LEN = (ushort)(p - g_Protocol_Buff_p->Data);
                Master_Send();
            }
        }

        public void Set_CH_Switch() {
            fixed (PROTOCOL_FRAME_TypeDef* temp_p = &g_Protocol_Buff) {
                g_Protocol_Buff_p = temp_p;
                g_Protocol_Buff_p->Function_CMD = (byte)CMD_TypeDef.CMD_CH_SWITCH;
                g_Protocol_Buff_p->RW_CMD = WRITE_CMD;
                byte* p = g_Protocol_Buff_p->Data;
                ushort ch_mask = 0;
                ch_mask = Make_CH_Mask(checkBox_CH_Switch_Sel);
                *(ushort*)p = ch_mask;
                p += 2;
                try {
                    byte i = 0;
                    for (i = 0; i < CHANNELS_NUMBER; i++) {
                        if (0x01 == ((ch_mask >> i) & 0x01)) {
                            *(byte*)p = (byte)(Convert.ToByte(textBox_CH_Switch[i].Text));
                            p += 1;
                        }
                    }
                    *(ushort*)g_Protocol_Buff_p->Data_LEN = (ushort)(p - g_Protocol_Buff_p->Data);
                    Master_Send();
                } catch {
                    MessageBoxERRwarning(3);
                }
            }
        }

        public void Get_CH_Switch() {
            fixed (PROTOCOL_FRAME_TypeDef* temp_p = &g_Protocol_Buff) {
                g_Protocol_Buff_p = temp_p;
                g_Protocol_Buff_p->Function_CMD = (byte)CMD_TypeDef.CMD_CH_SWITCH;
                g_Protocol_Buff_p->RW_CMD = READ_CMD;
                byte* p = g_Protocol_Buff_p->Data;
                ushort ch_mask = 0;
                ch_mask = Make_CH_Mask(checkBox_CH_Switch_Sel);
                *(ushort*)p = ch_mask;
                p += 2;
                *(ushort*)g_Protocol_Buff_p->Data_LEN = (ushort)(p - g_Protocol_Buff_p->Data);
                Master_Send();
            }
        }

        public void Set_CH_Type() {
            fixed (PROTOCOL_FRAME_TypeDef* temp_p = &g_Protocol_Buff) {
                g_Protocol_Buff_p = temp_p;
                g_Protocol_Buff_p->Function_CMD = (byte)CMD_TypeDef.CMD_CH_TYPE;
                g_Protocol_Buff_p->RW_CMD = WRITE_CMD;
                byte* p = g_Protocol_Buff_p->Data;
                ushort ch_mask = 0;
                ch_mask = Make_CH_Mask(checkBox_CH_Type_Sel);
                *(ushort*)p = ch_mask;
                p += 2;
                try {
                    byte i = 0;
                    for (i = 0; i < CHANNELS_NUMBER; i++) {
                        if (0x01 == ((ch_mask >> i) & 0x01)) {
                            *(byte*)p = (byte)(Convert.ToByte(textBox_CH_Type[i].Text));
                            p += 1;
                        }
                    }
                    *(ushort*)g_Protocol_Buff_p->Data_LEN = (ushort)(p - g_Protocol_Buff_p->Data);
                    Master_Send();
                } catch {
                    MessageBoxERRwarning(3);
                }
            }
        }

        public void Get_CH_Type() {
            fixed (PROTOCOL_FRAME_TypeDef* temp_p = &g_Protocol_Buff) {
                g_Protocol_Buff_p = temp_p;
                g_Protocol_Buff_p->Function_CMD = (byte)CMD_TypeDef.CMD_CH_TYPE;
                g_Protocol_Buff_p->RW_CMD = READ_CMD;
                byte* p = g_Protocol_Buff_p->Data;
                ushort ch_mask = 0;
                ch_mask = Make_CH_Mask(checkBox_CH_Type_Sel);
                *(ushort*)p = ch_mask;
                p += 2;
                *(ushort*)g_Protocol_Buff_p->Data_LEN = (ushort)(p - g_Protocol_Buff_p->Data);
                Master_Send();
            }
        }
        
        public void Set_TC_CJC_Mode() {
            fixed (PROTOCOL_FRAME_TypeDef* temp_p = &g_Protocol_Buff) {
                g_Protocol_Buff_p = temp_p;
                g_Protocol_Buff_p->Function_CMD = (byte)CMD_TypeDef.CMD_CH_TC_CJC_MODEL;
                g_Protocol_Buff_p->RW_CMD = WRITE_CMD;
                byte* p = g_Protocol_Buff_p->Data;
                ushort ch_mask = 0;
                ch_mask = Make_CH_Mask(checkBox_CH_TC_CJC_Mode_Sel);
                *(ushort*)p = ch_mask;
                p += 2;
                try {
                    byte i = 0;
                    for (i = 0; i < CHANNELS_NUMBER; i++) {
                        if (0x01 == ((ch_mask >> i) & 0x01)) {
                            switch (comboBox_CH_TC_CJC_Mode[i].Text) {
                                case "内置冷端": {
                                        *(byte*)p = (0x00);
                                        break;
                                    }
                                case "外置冷端": {
                                        *(byte*)p = (0x01);
                                        break;
                                    }
                                case "冰点冷端": {
                                        *(byte*)p = (0x02);
                                        break;
                                    }
                                default: {
                                        MessageBoxERRwarning(5);
                                        return;
                                    }
                            }
                            p += 1;
                        }
                    }
                    *(ushort*)g_Protocol_Buff_p->Data_LEN = (ushort)(p - g_Protocol_Buff_p->Data);
                    Master_Send();
                } catch {
                    MessageBoxERRwarning(3);
                }
            }
        }

        public void Get_TC_CJC_Mode() {
            fixed (PROTOCOL_FRAME_TypeDef* temp_p = &g_Protocol_Buff) {
                g_Protocol_Buff_p = temp_p;
                g_Protocol_Buff_p->Function_CMD = (byte)CMD_TypeDef.CMD_CH_TC_CJC_MODEL;
                g_Protocol_Buff_p->RW_CMD = READ_CMD;
                byte* p = g_Protocol_Buff_p->Data;
                ushort ch_mask = 0;
                ch_mask = Make_CH_Mask(checkBox_CH_TC_CJC_Mode_Sel);
                *(ushort*)p = ch_mask;
                p += 2;
                *(ushort*)g_Protocol_Buff_p->Data_LEN = (ushort)(p - g_Protocol_Buff_p->Data);
                Master_Send();
            }
        }

        public void Set_TC_Sensor() {
            fixed (PROTOCOL_FRAME_TypeDef* temp_p = &g_Protocol_Buff) {
                g_Protocol_Buff_p = temp_p;
                g_Protocol_Buff_p->Function_CMD = (byte)CMD_TypeDef.CMD_CH_TC_SENSOR;
                g_Protocol_Buff_p->RW_CMD = WRITE_CMD;
                byte* p = g_Protocol_Buff_p->Data;
                ushort ch_mask = 0;
                ch_mask = Make_CH_Mask(checkBox_CH_TC_Sensor_Sel);
                *(ushort*)p = ch_mask;
                p += 2;
                try {
                    byte i = 0;
                    for (i = 0; i < CHANNELS_NUMBER; i++) {
                        if (0x01 == ((ch_mask >> i) & 0x01)) {
                            switch (comboBox_CH_TC_Sensor[i].Text) {
                                case "K": {
                                        *(byte*)p = (0x00);
                                        break;
                                    }
                                case "J": {
                                        *(byte*)p = (0x01);
                                        break;
                                    }
                                case "T": {
                                        *(byte*)p = (0x02);
                                        break;
                                    }
                                case "E": {
                                        *(byte*)p = (0x03);
                                        break;
                                    }
                                case "N": {
                                        *(byte*)p = (0x04);
                                        break;
                                    }
                                case "B": {
                                        *(byte*)p = (0x05);
                                        break;
                                    }
                                case "R": {
                                        *(byte*)p = (0x06);
                                        break;
                                    }
                                case "S": {
                                        *(byte*)p = (0x07);
                                        break;
                                    }
                                default: {
                                        MessageBoxERRwarning(5);
                                        return;
                                    }
                            }
                            p += 1;
                        }
                    }
                    *(ushort*)g_Protocol_Buff_p->Data_LEN = (ushort)(p - g_Protocol_Buff_p->Data);
                    Master_Send();
                } catch {
                    MessageBoxERRwarning(3);
                }
            }
        }

        public void Get_TC_Sensor() {
            fixed (PROTOCOL_FRAME_TypeDef* temp_p = &g_Protocol_Buff) {
                g_Protocol_Buff_p = temp_p;
                g_Protocol_Buff_p->Function_CMD = (byte)CMD_TypeDef.CMD_CH_TC_SENSOR;
                g_Protocol_Buff_p->RW_CMD = READ_CMD;
                byte* p = g_Protocol_Buff_p->Data;
                ushort ch_mask = 0;
                ch_mask = Make_CH_Mask(checkBox_CH_TC_Sensor_Sel);
                *(ushort*)p = ch_mask;
                p += 2;
                *(ushort*)g_Protocol_Buff_p->Data_LEN = (ushort)(p - g_Protocol_Buff_p->Data);
                Master_Send();
            }
        }

        public void Get_TC_CJC_T() {
            fixed (PROTOCOL_FRAME_TypeDef* temp_p = &g_Protocol_Buff) {
                g_Protocol_Buff_p = temp_p;
                g_Protocol_Buff_p->Function_CMD = (byte)CMD_TypeDef.CMD_CH_TC_CJC_T;
                g_Protocol_Buff_p->RW_CMD = READ_CMD;
                byte* p = g_Protocol_Buff_p->Data;
                ushort ch_mask = 0;
                ch_mask = Make_CH_Mask(checkBox_CH_TC_CJC_T_Sel);
                *(ushort*)p = ch_mask;
                p += 2;
                *(ushort*)g_Protocol_Buff_p->Data_LEN = (ushort)(p - g_Protocol_Buff_p->Data);
                Master_Send();
            }
        }

        public void Get_Original_Value() {
            fixed (PROTOCOL_FRAME_TypeDef* temp_p = &g_Protocol_Buff) {
                g_Protocol_Buff_p = temp_p;
                g_Protocol_Buff_p->Function_CMD = (byte)CMD_TypeDef.CMD_CH_ORIGINAL_VALUE;
                g_Protocol_Buff_p->RW_CMD = READ_CMD;
                byte* p = g_Protocol_Buff_p->Data;
                ushort ch_mask = 0;
                try {
                    ch_mask = Make_CH_Mask(checkBox_CH_Original_Value_Sel);
                    *(ushort*)p = ch_mask;
                    p += 2;
                    byte ori_num = (byte)(Convert.ToByte(textBox_Ori_Num.Text));
                    *(byte*)p = ori_num;
                    p += 1;
                    *(ushort*)g_Protocol_Buff_p->Data_LEN = (ushort)(p - g_Protocol_Buff_p->Data);
                    Master_Send();
                } catch {
                    MessageBoxERRwarning(3);
                }
            }
        }

        public void Ch_Calibration() {
            fixed (PROTOCOL_FRAME_TypeDef* temp_p = &g_Protocol_Buff) {
                g_Protocol_Buff_p = temp_p;
                g_Protocol_Buff_p->Function_CMD = (byte)CMD_TypeDef.CMD_CH_CALIBRATION;
                g_Protocol_Buff_p->RW_CMD = WRITE_CMD;
                byte* p = g_Protocol_Buff_p->Data;
                ushort ch_mask = 0;
                ch_mask = Make_CH_Mask(checkBox_Cal_Phy_CH_Sel);
                *(ushort*)p = ch_mask;
                p += 2;
                byte phy_num = (byte)(Convert.ToByte(textBox_Phy_Num_Cal.Text));
                *(byte*)p = phy_num;
                p += 1;
                try {
                    byte i = 0;
                    for (i = 0; i < CHANNELS_NUMBER; i++) {
                        if (0x01 == ((ch_mask >> i) & 0x01)) {
                            *(ushort*)p = (ushort)(Convert.ToDouble(textBox_Cal_Phy[i].Text) * 100);
                            p += 2;
                        }
                    }
                    *(ushort*)g_Protocol_Buff_p->Data_LEN = (ushort)(p - g_Protocol_Buff_p->Data);
                    Master_Send();
                } catch {
                    MessageBoxERRwarning(3);
                }
            }
        }

        public void Get_Physical_Variable() {
            fixed (PROTOCOL_FRAME_TypeDef* temp_p = &g_Protocol_Buff) {
                g_Protocol_Buff_p = temp_p;
                g_Protocol_Buff_p->Function_CMD = (byte)CMD_TypeDef.CMD_CH_PHYSICAL_VARIABLE;
                g_Protocol_Buff_p->RW_CMD = READ_CMD;
                byte* p = g_Protocol_Buff_p->Data;
                ushort ch_mask = 0;
                try {
                    ch_mask = Make_CH_Mask(checkBox_Get_Phy_CH_Sel);
                    *(ushort*)p = ch_mask;
                    p += 2;
                    byte phy_num = (byte)(Convert.ToByte(textBox_Phy_Num_Get.Text));
                    *(byte*)p = phy_num;
                    p += 1;
                    *(ushort*)g_Protocol_Buff_p->Data_LEN = (ushort)(p - g_Protocol_Buff_p->Data);
                    Master_Send();
                } catch {
                    MessageBoxERRwarning(3);
                }
            }
        }

        public void Get_Cal_Par() {
            fixed (PROTOCOL_FRAME_TypeDef* temp_p = &g_Protocol_Buff) {
                g_Protocol_Buff_p = temp_p;
                g_Protocol_Buff_p->Function_CMD = (byte)CMD_TypeDef.CMD_CH_CALIBRATION_PARAMETER;
                g_Protocol_Buff_p->RW_CMD = READ_CMD;
                byte* p = g_Protocol_Buff_p->Data;
                ushort ch_mask = 0;
                try {
                    ch_mask = Make_CH_Mask(checkBox_CH_Cal_Par_Sel);
                    *(ushort*)p = ch_mask;
                    p += 2;
                    byte phy_num = (byte)(Convert.ToByte(textBox_Phy_Num_Par.Text));
                    *(byte*)p = phy_num;
                    p += 1;
                    *(ushort*)g_Protocol_Buff_p->Data_LEN = (ushort)(p - g_Protocol_Buff_p->Data);
                    Master_Send();
                } catch {
                    MessageBoxERRwarning(3);
                }
            }
        }

    }
}


