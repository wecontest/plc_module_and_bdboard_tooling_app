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
using HID;
using System.IO.Ports;

namespace Testsoft {
    public partial class Form1 : Form {
        /*****************************************模拟模块部分***************************************************/
        //模拟模块的通讯数据定义
        Hid myHid_An = new Hid(); //模拟模块USB口
        IntPtr myHidPtr_An = new IntPtr();
        bool Rec_Start_An = false;
        bool Rec_Completed_An = false;
        byte[] RxDataBuffer_An = new byte[200]; //接收缓冲
        byte RecDataIndex_An = 0; //当前缓冲接收数据的位置
        byte[] TxDataPart_An = new byte[200]; //发送的数据段

        public enum FX2N_CMD_TypeDef {
            FX2N_READ,
            FX2N_WRITE,
        }

        public enum ANALYTE_TYPE_TypeDef {
            ANALYTE_ADI,
            ANALYTE_DAI,
            ANALYTE_ADV,
            ANALYTE_DAV,
            ANALYTE_PT,
            ANALYTE_TCV,
            ANALYTE_TCR,
            ANALYTE_NTC
        }
        public ANALYTE_TYPE_TypeDef ANALYTE_TYPE;

        public enum ANALYTE_CHANNEL_TypeDef {
            ANALYTE_CH1 = 0x01,
            ANALYTE_CH2 = 0x02,
            ANALYTE_CH3 = 0x04,
            ANALYTE_CH4 = 0x08,
            ANALYTE_CH5 = 0x10,
            ANALYTE_CH6 = 0x20,
            ANALYTE_CH7 = 0x40,
            ANALYTE_CH8 = 0x80,
            ANALYTE_CH_ALL = 0xFF
        }
        public byte ANALYTE_CHANNEL;

        public enum ANALYTE_RES_TypeDef {
            ANALYTE_RES1_PT100F,
            ANALYTE_RES2_PT100S,
            ANALYTE_RES3_CU50F,
            ANALYTE_RES4_CU50S,
            ANALYTE_RES5_PT1000F,
            ANALYTE_RES6_PT1000S,
            ANALYTE_RES7_NTCF,
            ANALYTE_RES8_NTCS
        }
        public ANALYTE_RES_TypeDef ANALYTE_RES;

        public enum ADC7190_G_TypeDef {
            ADC7190_G1 = 1,
            ADC7190_G8 = 8,
            ADC7190_G16 = 16,
            ADC7190_G32 = 32,
            ADC7190_G64 = 64,
            ADC7190_G128 = 128
        }
        public ADC7190_G_TypeDef ADC7190_G;

        public struct ANALOG_DATA_TypeDef {
            public uint ADC7190_DATA_AIN1;
            public uint ADC7190_DATA_AIN2;
            public uint ADC7190_DATA_AIN3;
            public uint ADC7190_DATA_AIN4;
            public uint ADS8699_DATA;
            public ushort DAC8760_V_OUT;
            public ushort DAC8760_I_OUT;
            public ushort Analyte_Type;
            public ushort Analyte_Channel;
            public ushort Analyte_Res;
            public ushort ADC7190_G;
            public ushort DAC8760_Output_Switch_Flag;
            public double Infrared_temperature1;
            public double Infrared_temperature2;
            public double Infrared_temperature3;
            public double Infrared_temperature4;
            public ushort Auto_Cal_Res;
            public double ADC7190_V_AIN1;
            public double ADC7190_V_AIN2;
            public double ADC7190_V_AIN3;
            public double ADC7190_V_AIN4;
            public double ADS8699_AD_DA_V;
            public double ADS8699_AD_DA_I;
        }
        ANALOG_DATA_TypeDef ANALOG_DATA;

        double[] ANALOG_RES = new double[8];

        //USB数据到达事件
        protected void myhid_DataReceived_An(object sender, report e) {
            FX2N_Read(e, ref Rec_Start_An, ref Rec_Completed_An, ref RecDataIndex_An, ref RxDataBuffer_An);
        }

        //数据已处理完成事件
        protected void myhid_DataProcessed_An(object sender, EventArgs e) {
            byte[] DealDataBuffer = new byte[100]; //接收缓冲
            if (FX2N_Read_DataDeal(ref DealDataBuffer, ref Rec_Completed_An, ref RecDataIndex_An, ref RxDataBuffer_An)) {
                myhid_RxDataDeal_An(DealDataBuffer);
                //DisplayRXData(j, DealDataBuffer);
                if (checkBox_ContinueRead2.Checked == true) {
                    Thread.Sleep(500);
                    myhid_UIDeal_An();
                    Analog_Read_All();
                }
            }
        }

        //设备移除事件
        protected void myhid_DeviceRemoved(object sender, EventArgs e) {

        }

        private void Analog_Read_All() {
            TxDataPart_An = new byte[2]; //重置发送数据的缓冲
            FX2N_Send(myHid_An, FX2N_CMD_TypeDef.FX2N_READ, 0x0000, 52, TxDataPart_An);
        }

        protected void myhid_RxDataDeal_An(byte[] DealDataBuffer) {
            byte index = 0;
            ANALOG_DATA.ADC7190_DATA_AIN1 = DealDataBuffer[index++];
            ANALOG_DATA.ADC7190_DATA_AIN1 |= (uint)DealDataBuffer[index++] << 8;
            ANALOG_DATA.ADC7190_DATA_AIN1 |= (uint)DealDataBuffer[index++] << 16;
            ANALOG_DATA.ADC7190_DATA_AIN1 |= (uint)DealDataBuffer[index++] << 24;
            ANALOG_DATA.ADC7190_DATA_AIN2 = DealDataBuffer[index++];
            ANALOG_DATA.ADC7190_DATA_AIN2 |= (uint)DealDataBuffer[index++] << 8;
            ANALOG_DATA.ADC7190_DATA_AIN2 |= (uint)DealDataBuffer[index++] << 16;
            ANALOG_DATA.ADC7190_DATA_AIN2 |= (uint)DealDataBuffer[index++] << 24;
            ANALOG_DATA.ADC7190_DATA_AIN3 = DealDataBuffer[index++];
            ANALOG_DATA.ADC7190_DATA_AIN3 |= (uint)DealDataBuffer[index++] << 8;
            ANALOG_DATA.ADC7190_DATA_AIN3 |= (uint)DealDataBuffer[index++] << 16;
            ANALOG_DATA.ADC7190_DATA_AIN3 |= (uint)DealDataBuffer[index++] << 24;
            ANALOG_DATA.ADC7190_DATA_AIN4 = DealDataBuffer[index++];
            ANALOG_DATA.ADC7190_DATA_AIN4 |= (uint)DealDataBuffer[index++] << 8;
            ANALOG_DATA.ADC7190_DATA_AIN4 |= (uint)DealDataBuffer[index++] << 16;
            ANALOG_DATA.ADC7190_DATA_AIN4 |= (uint)DealDataBuffer[index++] << 24;
            ANALOG_DATA.ADS8699_DATA = DealDataBuffer[index++];
            ANALOG_DATA.ADS8699_DATA |= (uint)DealDataBuffer[index++] << 8;
            ANALOG_DATA.ADS8699_DATA |= (uint)DealDataBuffer[index++] << 16;
            ANALOG_DATA.ADS8699_DATA |= (uint)DealDataBuffer[index++] << 24;
            ANALOG_DATA.DAC8760_V_OUT = DealDataBuffer[index++];
            ANALOG_DATA.DAC8760_V_OUT |= (ushort)(DealDataBuffer[index++] << 8);
            ANALOG_DATA.DAC8760_I_OUT = DealDataBuffer[index++];
            ANALOG_DATA.DAC8760_I_OUT |= (ushort)(DealDataBuffer[index++] << 8);
            ANALOG_DATA.Analyte_Type = DealDataBuffer[index++];
            ANALOG_DATA.Analyte_Type |= (ushort)(DealDataBuffer[index++] << 8);
            ANALOG_DATA.Analyte_Channel = DealDataBuffer[index++];
            ANALOG_DATA.Analyte_Channel |= (ushort)(DealDataBuffer[index++] << 8);
            ANALOG_DATA.Analyte_Res = DealDataBuffer[index++];
            ANALOG_DATA.Analyte_Res |= (ushort)(DealDataBuffer[index++] << 8);
            ANALOG_DATA.ADC7190_G = DealDataBuffer[index++];
            ANALOG_DATA.ADC7190_G |= (ushort)(DealDataBuffer[index++] << 8);
            ANALOG_DATA.DAC8760_Output_Switch_Flag = DealDataBuffer[index++];
            ANALOG_DATA.DAC8760_Output_Switch_Flag |= (ushort)(DealDataBuffer[index++] << 8);
            ANALOG_DATA.Infrared_temperature1 = BitConverter.ToSingle(DealDataBuffer, index);
            ANALOG_DATA.Infrared_temperature2 = BitConverter.ToSingle(DealDataBuffer, index + 4);
            ANALOG_DATA.Infrared_temperature3 = BitConverter.ToSingle(DealDataBuffer, index + 8);
            ANALOG_DATA.Infrared_temperature4 = BitConverter.ToSingle(DealDataBuffer, index + 12);
            index += 16;
            ANALOG_DATA.Auto_Cal_Res = DealDataBuffer[index++];
            ANALOG_DATA.Auto_Cal_Res |= (ushort)(DealDataBuffer[index++] << 8);
        }

        private void myhid_UIDeal_An() {
            try {
                textBox_V.Text = ((float)ANALOG_DATA.DAC8760_V_OUT / 65535 * 20 - 10).ToString("F4");
                textBox_I.Text = ((float)ANALOG_DATA.DAC8760_I_OUT / 65535 * 20).ToString("F4");
                textBox_Infrared1.Text = (ANALOG_DATA.Infrared_temperature1).ToString("F4");
                textBox_Infrared2.Text = (ANALOG_DATA.Infrared_temperature2).ToString("F4");
                textBox_Infrared3.Text = (ANALOG_DATA.Infrared_temperature3).ToString("F4");
                textBox_Infrared4.Text = (ANALOG_DATA.Infrared_temperature4).ToString("F4");
                textBox_ADC7190_CH1.Text = "0x" + (ANALOG_DATA.ADC7190_DATA_AIN1).ToString("X8");
                textBox_ADC7190_CH2.Text = "0x" + (ANALOG_DATA.ADC7190_DATA_AIN2).ToString("X8");
                textBox_ADC7190_CH3.Text = "0x" + (ANALOG_DATA.ADC7190_DATA_AIN3).ToString("X8");
                textBox_ADC7190_CH4.Text = "0x" + (ANALOG_DATA.ADC7190_DATA_AIN4).ToString("X8");
                textBox_ADS8699.Text = "0x" + (ANALOG_DATA.ADS8699_DATA).ToString("X8");
                Analyte_Type_UISet((ANALYTE_TYPE_TypeDef)ANALOG_DATA.Analyte_Type);
                Analyte_Channel_UISet((byte)ANALOG_DATA.Analyte_Channel);
                Analyte_Res_UISet((ANALYTE_RES_TypeDef)ANALOG_DATA.Analyte_Res);
                Analyte_ADC7190_G_UISet((ADC7190_G_TypeDef)ANALOG_DATA.ADC7190_G);
                if (ANALOG_DATA.DAC8760_Output_Switch_Flag == 0x0001)
                    label_OutputEnable.Text = ":开";
                else
                    label_OutputEnable.Text = ":关";

                if (ANALOG_DATA.Auto_Cal_Res == 0x0001)
                    label_Auto_Cal_RES_Enable.Text = ":开";
                else
                    label_Auto_Cal_RES_Enable.Text = ":关";

                ANALOG_DATA.ADC7190_V_AIN1 = Analog_Count_ADC7190_V(ANALOG_DATA.ADC7190_DATA_AIN1);
                ANALOG_DATA.ADC7190_V_AIN2 = Analog_Count_ADC7190_V(ANALOG_DATA.ADC7190_DATA_AIN2);
                ANALOG_DATA.ADC7190_V_AIN3 = Analog_Count_ADC7190_V(ANALOG_DATA.ADC7190_DATA_AIN3);
                ANALOG_DATA.ADC7190_V_AIN4 = Analog_Count_ADC7190_V(ANALOG_DATA.ADC7190_DATA_AIN4);
                ANALOG_DATA.ADS8699_AD_DA_V = Analog_Count_ADS8699_V(ANALOG_DATA.ADS8699_DATA);
                ANALOG_DATA.ADS8699_AD_DA_I = Analog_Count_ADS8699_I(ANALOG_DATA.ADS8699_DATA);
                //richTextBox3.Text = "ADC7190第1通道电压：" + ANALOG_DATA.ADC7190_V_AIN1.ToString("F") + "\r\n";
                //richTextBox3.Text += "ADC7190第2通道电压：" + ANALOG_DATA.ADC7190_V_AIN2.ToString("F") + "\r\n";
                //richTextBox3.Text += "ADC7190第3通道电压：" + ANALOG_DATA.ADC7190_V_AIN3.ToString("F") + "\r\n";
                //richTextBox3.Text += "ADC7190第4通道电压：" + ANALOG_DATA.ADC7190_V_AIN4.ToString("F") + "\r\n";
            } catch {

            }
        }

        private void Analyte_Type_UISet(ANALYTE_TYPE_TypeDef type) {
            switch (type) {
                case ANALYTE_TYPE_TypeDef.ANALYTE_ADI:
                    radioButton_ADI.Checked = true;
                    break;
                case ANALYTE_TYPE_TypeDef.ANALYTE_DAI:
                    radioButton_DAI.Checked = true;
                    break;
                case ANALYTE_TYPE_TypeDef.ANALYTE_ADV:
                    radioButton_ADV.Checked = true;
                    break;
                case ANALYTE_TYPE_TypeDef.ANALYTE_DAV:
                    radioButton_DAV.Checked = true;
                    break;
                case ANALYTE_TYPE_TypeDef.ANALYTE_PT:
                    radioButton_PT.Checked = true;
                    break;
                case ANALYTE_TYPE_TypeDef.ANALYTE_TCV:
                    radioButton_TCV.Checked = true;
                    break;
                case ANALYTE_TYPE_TypeDef.ANALYTE_TCR:
                    radioButton_TCR.Checked = true;
                    break;
                case ANALYTE_TYPE_TypeDef.ANALYTE_NTC:
                    radioButton_NTC.Checked = true;
                    break;
            }
        }

        private void Analyte_Channel_UISet(byte channel) {
            byte locate = 0;
            locate = (byte)((channel >> 0) & 0x01);
            if (locate == 0x01) {
                checkBox_CH1.Checked = true;
            } else {
                checkBox_CH1.Checked = false;
            }
            locate = (byte)((channel >> 1) & 0x01);
            if (locate == 0x01) {
                checkBox_CH2.Checked = true;
            } else {
                checkBox_CH2.Checked = false;
            }
            locate = (byte)((channel >> 2) & 0x01);
            if (locate == 0x01) {
                checkBox_CH3.Checked = true;
            } else {
                checkBox_CH3.Checked = false;
            }
            locate = (byte)((channel >> 3) & 0x01);
            if (locate == 0x01) {
                checkBox_CH4.Checked = true;
            } else {
                checkBox_CH4.Checked = false;
            }
            locate = (byte)((channel >> 4) & 0x01);
            if (locate == 0x01) {
                checkBox_CH5.Checked = true;
            } else {
                checkBox_CH5.Checked = false;
            }
            locate = (byte)((channel >> 5) & 0x01);
            if (locate == 0x01) {
                checkBox_CH6.Checked = true;
            } else {
                checkBox_CH6.Checked = false;
            }
            locate = (byte)((channel >> 6) & 0x01);
            if (locate == 0x01) {
                checkBox_CH7.Checked = true;
            } else {
                checkBox_CH7.Checked = false;
            }
            locate = (byte)((channel >> 7) & 0x01);
            if (locate == 0x01) {
                checkBox_CH8.Checked = true;
            } else {
                checkBox_CH8.Checked = false;
            }
        }

        double RF = 249;
        private void Analyte_Res_UISet(ANALYTE_RES_TypeDef res) {
            switch (res) {
                case ANALYTE_RES_TypeDef.ANALYTE_RES1_PT100F:
                    radioButton_Res1.Checked = true;
                    break;
                case ANALYTE_RES_TypeDef.ANALYTE_RES2_PT100S:
                    radioButton_Res2.Checked = true;
                    break;
                case ANALYTE_RES_TypeDef.ANALYTE_RES3_CU50F:
                    radioButton_Res3.Checked = true;
                    break;
                case ANALYTE_RES_TypeDef.ANALYTE_RES4_CU50S:
                    radioButton_Res4.Checked = true;
                    break;
                case ANALYTE_RES_TypeDef.ANALYTE_RES5_PT1000F:
                    radioButton_Res5.Checked = true;
                    break;
                case ANALYTE_RES_TypeDef.ANALYTE_RES6_PT1000S:
                    radioButton_Res6.Checked = true;
                    break;
                case ANALYTE_RES_TypeDef.ANALYTE_RES7_NTCF:
                    radioButton_Res7.Checked = true;
                    break;
                case ANALYTE_RES_TypeDef.ANALYTE_RES8_NTCS:
                    radioButton_Res8.Checked = true;
                    break;
            }
            RF = ANALOG_RES[(byte)res];
        }

        private void Analyte_ADC7190_G_UISet(ADC7190_G_TypeDef g) {
            switch (g) {
                case ADC7190_G_TypeDef.ADC7190_G1:
                    radioButton_G1.Checked = true;
                    break;
                case ADC7190_G_TypeDef.ADC7190_G8:
                    radioButton_G8.Checked = true;
                    break;
                case ADC7190_G_TypeDef.ADC7190_G16:
                    radioButton_G16.Checked = true;
                    break;
                case ADC7190_G_TypeDef.ADC7190_G32:
                    radioButton_G32.Checked = true;
                    break;
                case ADC7190_G_TypeDef.ADC7190_G64:
                    radioButton_G64.Checked = true;
                    break;
                case ADC7190_G_TypeDef.ADC7190_G128:
                    radioButton_G128.Checked = true;
                    break;
            }
        }

        //V：要输出的电压，单位V
        private void Analog_DAC8760_V_Set(double V) {
            try {
                ushort data = 0;
                data = (ushort)(((V + 10) / 20) * 65535);
                TxDataPart_An = new byte[2]; //重置发送数据的缓冲
                TxDataPart_An[0] = (byte)(data & 0xFF);
                TxDataPart_An[1] = (byte)(data >> 8 & 0xFF);
            } catch {
                MessageBoxERRwarning(3);
            }
            FX2N_Send(myHid_An, FX2N_CMD_TypeDef.FX2N_WRITE, 0x0014, 2, TxDataPart_An);
        }

        //I：要输出的电流，单位mA
        private void Analog_DAC8760_I_Set(double I) {
            textBox_I.Text = I.ToString("F4");
            try {
                ushort data = 0;
                data = (ushort)((I / 20) * 65535);
                TxDataPart_An = new byte[2]; //重置发送数据的缓冲
                TxDataPart_An[0] = (byte)(data & 0xFF);
                TxDataPart_An[1] = (byte)(data >> 8 & 0xFF);
            } catch {
                MessageBoxERRwarning(3);
            }
            FX2N_Send(myHid_An, FX2N_CMD_TypeDef.FX2N_WRITE, 0x0016, 2, TxDataPart_An);
        }

        private void Analog_Analyte_Type_Select(ANALYTE_TYPE_TypeDef type) {
            Analyte_Type_UISet(type);
            ushort data = (ushort)type;
            TxDataPart_An = new byte[2]; //重置发送数据的缓冲
            TxDataPart_An[0] = (byte)(data & 0xFF);
            TxDataPart_An[1] = (byte)(data >> 8 & 0xFF);
            FX2N_Send(myHid_An, FX2N_CMD_TypeDef.FX2N_WRITE, 0x0018, 2, TxDataPart_An);
        }

        private void Analog_Analyte_Channel_Select(byte channel) {
            Analyte_Channel_UISet(channel);
            ushort data = (ushort)channel;
            TxDataPart_An = new byte[2]; //重置发送数据的缓冲
            TxDataPart_An[0] = (byte)(data & 0xFF);
            TxDataPart_An[1] = (byte)(data >> 8 & 0xFF);
            FX2N_Send(myHid_An, FX2N_CMD_TypeDef.FX2N_WRITE, 0x001A, 2, TxDataPart_An);
        }

        private void Analog_Analyte_Res_Select(ANALYTE_RES_TypeDef res) {
            Analyte_Res_UISet(res);
            ushort data = (ushort)res;
            TxDataPart_An = new byte[2]; //重置发送数据的缓冲
            TxDataPart_An[0] = (byte)(data & 0xFF);
            TxDataPart_An[1] = (byte)(data >> 8 & 0xFF);
            FX2N_Send(myHid_An, FX2N_CMD_TypeDef.FX2N_WRITE, 0x001C, 2, TxDataPart_An);
        }

        private void Analog_ADC7190_G_Set(ADC7190_G_TypeDef g) {
            Analyte_ADC7190_G_UISet(g);
            ushort data = (ushort)g;
            TxDataPart_An = new byte[2]; //重置发送数据的缓冲
            TxDataPart_An[0] = (byte)(data & 0xFF);
            TxDataPart_An[1] = (byte)(data >> 8 & 0xFF);
            FX2N_Send(myHid_An, FX2N_CMD_TypeDef.FX2N_WRITE, 0x001E, 2, TxDataPart_An);
        }

        private void Analog_DAC8760_Output_Switch_Flag_Set(string outputswitch) {
            TxDataPart_An = new byte[2]; //重置发送数据的缓冲
            if (outputswitch == ":关") {
                label_OutputEnable.Text = ":开";
                TxDataPart_An[0] = 0x01;
            } else if (outputswitch == ":开") {
                label_OutputEnable.Text = ":关";
                TxDataPart_An[0] = 0x00;
            }
            TxDataPart_An[1] = 0x00;
            FX2N_Send(myHid_An, FX2N_CMD_TypeDef.FX2N_WRITE, 0x0020, 2, TxDataPart_An);
        }

        private void Analog_Auto_Cal_RES_Set(string autocal) {
            TxDataPart_An = new byte[2]; //重置发送数据的缓冲
            if (autocal == ":关") {
                label_Auto_Cal_RES_Enable.Text = ":开";
                TxDataPart_An[0] = 0x01;
            } else if (autocal == ":开") {
                label_Auto_Cal_RES_Enable.Text = ":关";
                TxDataPart_An[0] = 0x00;
            }
            TxDataPart_An[1] = 0x00;
            FX2N_Send(myHid_An, FX2N_CMD_TypeDef.FX2N_WRITE, 0x0032, 2, TxDataPart_An);
        }

        private void Analog_OneKey_Calibration() {
            checkBox_ContinueRead2.Checked = false;
            richTextBox_AnalogMessage.Text = "";
            richTextBox_AnalogMessage.BackColor = SystemColors.Window;
            Digital_Read_All();
            myhid_UIDeal_Di();
            if (DIGITAL_DATA.DOWNLOAD == 0) {
                richTextBox_AnalogMessage.Text += "固件尚未下载完成";
                return;
            }
            Analog_Read_All();
            myhid_UIDeal_An();
            Analog_Calibrate_Verify(CAL_VER_FLAG_TypeDef.CAL_FLAG);

            Onekey_Calibration_Flag = false;
        }

        private void Analog_OneKey_Verify() {
            checkBox_ContinueRead2.Checked = false;
            richTextBox_AnalogMessage.Text = "";
            richTextBox_AnalogMessage.BackColor = SystemColors.Window;
            Digital_Read_All();
            myhid_UIDeal_Di();
            if (DIGITAL_DATA.DOWNLOAD == 0) {
                richTextBox_AnalogMessage.Text += "固件尚未下载完成";
                return;
            }
            Analog_Read_All();
            myhid_UIDeal_An();
            Analog_Calibrate_Verify(CAL_VER_FLAG_TypeDef.VER_FLAG);

            Onekey_Verify_Flag = false;
        }

        private float Analog_Count_ADC7190_V(uint ADC7190_DATA) {
            float ADC7190_V = (float)(((float)ADC7190_DATA - (1 << 23)) * 4.096 / (ANALOG_DATA.ADC7190_G * (1 << 23)));
            return ADC7190_V;
        }

        private float Analog_Count_ADS8699_V(uint ADS8699_DATA) {
            float ADS8699_V = (float)(78.125 * (float)ADS8699_DATA / 1000000 - 10240);
            return ADS8699_V;
        }

        private float Analog_Count_ADS8699_I(uint ADS8699_DATA) {
            float ADS8699_mA = (float)((78.125 * (float)ADS8699_DATA / 1000000 - 10240) / RF) * 1000;
            return ADS8699_mA;
        }

        private void Analog_ADV_Standard_Set(double standard_mV) {
            Analog_DAC8760_V_Set(standard_mV / 1000);
        }

        private double Analog_ADV_DAV_Standard_Read() {
            Thread.Sleep(2000);
            Analog_Read_All();
            myhid_UIDeal_An();
            return ANALOG_DATA.ADS8699_AD_DA_V * 1000;
        }

        private void Analog_ADI_Standard_Set(double standard_mA) {
            Analog_DAC8760_I_Set(standard_mA);
        }

        private double Analog_ADI_DAI_Standard_Read() {
            Thread.Sleep(2000);
            Analog_Read_All();
            myhid_UIDeal_An();
            return ANALOG_DATA.ADS8699_AD_DA_I;
        }

        private void Analog_TCV_Standard_Set(double standard_mV) {
            Analog_DAC8760_V_Set(1.385 * standard_mV / 1000);
        }

        private double Analog_TCV_Standard_Read() {
            Analog_ADC7190_G_Set(ADC7190_G_TypeDef.ADC7190_G8);
            Thread.Sleep(2000);
            Analog_Read_All();
            myhid_UIDeal_An();
            return ANALOG_DATA.ADC7190_V_AIN4 * 1000;
        }

        private void Analog_Res_Standard_Set(ANALYTE_RES_TypeDef standard_R) {
            Analog_Analyte_Res_Select(standard_R);
        }

        private double Analog_PT100_PT1000_TCR_Standard_Read() {
            Analog_ADC7190_G_Set(ADC7190_G_TypeDef.ADC7190_G8);
            Thread.Sleep(1500);
            Analog_Read_All();
            myhid_UIDeal_An();
            return ANALOG_DATA.ADC7190_V_AIN3 * RF / (ANALOG_DATA.ADC7190_V_AIN2 - ANALOG_DATA.ADC7190_V_AIN1);
        }

        private double Analog_NTC_Standard_Read() {
            Analog_ADC7190_G_Set(ADC7190_G_TypeDef.ADC7190_G1);
            Thread.Sleep(1500);
            Analog_Read_All();
            myhid_UIDeal_An();
            return ANALOG_DATA.ADC7190_V_AIN3 * RF / (ANALOG_DATA.ADC7190_V_AIN2 - ANALOG_DATA.ADC7190_V_AIN1);
        }

        /*****************************************数字模块部分***************************************************/
        //数字模块的通讯数据定义
        Hid myHid_Di = new Hid(); //数字模块USB口
        IntPtr myHidPtr_Di = new IntPtr();
        bool Rec_Start_Di = false;
        bool Rec_Completed_Di = false;
        byte[] RxDataBuffer_Di = new byte[200]; //接收缓冲
        byte RecDataIndex_Di = 0; //当前缓冲接收数据的位置
        byte[] TxDataPart_Di = new byte[200]; //发送的数据段

        public struct DIGITAL_DATA_TypeDef {
            public ushort ADC_CH1_D1;
            public ushort ADC_CH1_D2;
            public ushort ADC_CH2_D1;
            public ushort ADC_CH2_D2;
            public ushort ADC_CH3_D1;
            public ushort ADC_CH3_D2;
            public ushort ADC_CH4_D1;
            public ushort ADC_CH4_D2;
            public ushort ADC_CH5_D1;
            public ushort ADC_CH5_D2;
            public ushort ADC_CH6_D1;
            public ushort ADC_CH6_D2;
            public ushort ADC_CH7_D1;
            public ushort ADC_CH7_D2;
            public ushort ADC_CH8_D1;
            public ushort ADC_CH8_D2;
            public ushort CH1_V_P24;
            public ushort CH1_I_P24;
            public ushort CH2_V_P15;
            public ushort CH2_I_P15;
            public ushort CH3_V_N15;
            public ushort CH3_I_N15;
            public ushort CH4_V_L;
            public ushort CH4_I_L;
            public ushort CH1_I_P24_CAL;
            public ushort CH2_I_P15_CAL;
            public ushort CH3_I_N15_CAL;
            public ushort CH4_I_L_CAL;
            public ushort MODULE_TYPE;
            public ushort BD_TYPE;
            public ushort POWERUP;
            public ushort DOWNLOAD;
        }
        DIGITAL_DATA_TypeDef DIGITAL_DATA;

        public enum MODULE_TYPE_TypeDef {
            MODULE_TYPE_8IN = (0x00),
            MODULE_TYPE_16IN = (0x09),
            MODULE_TYPE_8OUT = (0x01),
            MODULE_TYPE_16OUT = (0x12),
            MODULE_TYPE_8IN8OUT = (0x02),
            MODULE_TYPE_2AD = (0x2D),
            MODULE_TYPE_4AD = (0x25),
            MODULE_TYPE_2DA = (0x3F),
            MODULE_TYPE_4DA = (0x37),
            MODULE_TYPE_2PT = (0x26),
            MODULE_TYPE_4PT = (0x2F),
            MODULE_TYPE_2TC = (0x27),
            MODULE_TYPE_4TC = (0x04),
            MODULE_TYPE_LC = (0x0D),
            MODULE_TYPE_2WT = (0x0F),
            MODULE_TYPE_4TC_Y = (0X03),
            MODULE_TYPE_4PG_ADV = (0x06),
            MODULE_TYPE_4PG_BAS = (0x17),
            MODULE_TYPE_8PT = (0x16),
            MODULE_TYPE_8TC = (0x36)
        }

        public enum BD_TYPE_TypeDef {
            BD_TYPE_2AD = (0xFF01),
            BD_TYPE_2DA = (0xFF02),
            BD_TYPE_2ADV = (0xFF03),
            BD_TYPE_2DAV = (0xFF04),
            BD_TYPE_2PT = (0xFF05),
            BD_TYPE_2TC = (0xFF06),
            BD_TYPE_1AD1DA = (0xFF07),
            BD_TYPE_1ADV1DAV = (0xFF08),
            BD_TYPE_2AD2DA = (0x0212),
            BD_TYPE_2PT2DA = (0x0222),
            BD_TYPE_2TC2DA = (0x0232),
            BD_TYPE_2PT2DAV = (0x6222),
            BD_TYPE_2PT2ADV = (0x7222),
            BD_TYPE_2PT1K2NTC2DAV = (0x6282)
        }

        //USB数据到达事件
        protected void myhid_DataReceived_Di(object sender, report e) {
            FX2N_Read(e, ref Rec_Start_Di, ref Rec_Completed_Di, ref RecDataIndex_Di, ref RxDataBuffer_Di);
        }

        //数据已处理完成事件
        protected void myhid_DataProcessed_Di(object sender, EventArgs e) {
            byte[] DealDataBuffer = new byte[100]; //接收缓冲
            if (FX2N_Read_DataDeal(ref DealDataBuffer, ref Rec_Completed_Di, ref RecDataIndex_Di, ref RxDataBuffer_Di)) {
                myhid_RxDataDeal_Di(DealDataBuffer);
                //DisplayRXData(j, DealDataBuffer);
                if (checkBox_ContinueRead1.Checked == true) {
                    Thread.Sleep(500);
                    myhid_UIDeal_Di();
                    Digital_Read_All();
                }
            }
        }

        private void Digital_Read_All() {
            TxDataPart_Di = new byte[2]; //重置发送数据的缓冲
            FX2N_Send(myHid_Di, FX2N_CMD_TypeDef.FX2N_READ, 0x6EE0, 64, TxDataPart_Di);
        }

        private void myhid_RxDataDeal_Di(byte[] DealDataBuffer) {
            byte index = 0;
            DIGITAL_DATA.ADC_CH1_D1 = DealDataBuffer[index++];
            DIGITAL_DATA.ADC_CH1_D1 |= (ushort)(DealDataBuffer[index++] << 8);
            DIGITAL_DATA.ADC_CH1_D2 = DealDataBuffer[index++];
            DIGITAL_DATA.ADC_CH1_D2 |= (ushort)(DealDataBuffer[index++] << 8);

            DIGITAL_DATA.ADC_CH2_D1 = DealDataBuffer[index++];
            DIGITAL_DATA.ADC_CH2_D1 |= (ushort)(DealDataBuffer[index++] << 8);
            DIGITAL_DATA.ADC_CH2_D2 = DealDataBuffer[index++];
            DIGITAL_DATA.ADC_CH2_D2 |= (ushort)(DealDataBuffer[index++] << 8);

            DIGITAL_DATA.ADC_CH3_D1 = DealDataBuffer[index++];
            DIGITAL_DATA.ADC_CH3_D1 |= (ushort)(DealDataBuffer[index++] << 8);
            DIGITAL_DATA.ADC_CH3_D2 = DealDataBuffer[index++];
            DIGITAL_DATA.ADC_CH3_D2 |= (ushort)(DealDataBuffer[index++] << 8);

            DIGITAL_DATA.ADC_CH4_D1 = DealDataBuffer[index++];
            DIGITAL_DATA.ADC_CH4_D1 |= (ushort)(DealDataBuffer[index++] << 8);
            DIGITAL_DATA.ADC_CH4_D2 = DealDataBuffer[index++];
            DIGITAL_DATA.ADC_CH4_D2 |= (ushort)(DealDataBuffer[index++] << 8);

            DIGITAL_DATA.ADC_CH5_D1 = DealDataBuffer[index++];
            DIGITAL_DATA.ADC_CH5_D1 |= (ushort)(DealDataBuffer[index++] << 8);
            DIGITAL_DATA.ADC_CH5_D2 = DealDataBuffer[index++];
            DIGITAL_DATA.ADC_CH5_D2 |= (ushort)(DealDataBuffer[index++] << 8);

            DIGITAL_DATA.ADC_CH6_D1 = DealDataBuffer[index++];
            DIGITAL_DATA.ADC_CH6_D1 |= (ushort)(DealDataBuffer[index++] << 8);
            DIGITAL_DATA.ADC_CH6_D2 = DealDataBuffer[index++];
            DIGITAL_DATA.ADC_CH6_D2 |= (ushort)(DealDataBuffer[index++] << 8);

            DIGITAL_DATA.ADC_CH7_D1 = DealDataBuffer[index++];
            DIGITAL_DATA.ADC_CH7_D1 |= (ushort)(DealDataBuffer[index++] << 8);
            DIGITAL_DATA.ADC_CH7_D2 = DealDataBuffer[index++];
            DIGITAL_DATA.ADC_CH7_D2 |= (ushort)(DealDataBuffer[index++] << 8);

            DIGITAL_DATA.ADC_CH8_D1 = DealDataBuffer[index++];
            DIGITAL_DATA.ADC_CH8_D1 |= (ushort)(DealDataBuffer[index++] << 8);
            DIGITAL_DATA.ADC_CH8_D2 = DealDataBuffer[index++];
            DIGITAL_DATA.ADC_CH8_D2 |= (ushort)(DealDataBuffer[index++] << 8);

            DIGITAL_DATA.CH1_V_P24 = DealDataBuffer[index++];
            DIGITAL_DATA.CH1_V_P24 |= (ushort)(DealDataBuffer[index++] << 8);

            DIGITAL_DATA.CH1_I_P24 = DealDataBuffer[index++];
            DIGITAL_DATA.CH1_I_P24 |= (ushort)(DealDataBuffer[index++] << 8);

            DIGITAL_DATA.CH2_V_P15 = DealDataBuffer[index++];
            DIGITAL_DATA.CH2_V_P15 |= (ushort)(DealDataBuffer[index++] << 8);

            DIGITAL_DATA.CH2_I_P15 = DealDataBuffer[index++];
            DIGITAL_DATA.CH2_I_P15 |= (ushort)(DealDataBuffer[index++] << 8);

            DIGITAL_DATA.CH3_V_N15 = DealDataBuffer[index++];
            DIGITAL_DATA.CH3_V_N15 |= (ushort)(DealDataBuffer[index++] << 8);

            DIGITAL_DATA.CH3_I_N15 = DealDataBuffer[index++];
            DIGITAL_DATA.CH3_I_N15 |= (ushort)(DealDataBuffer[index++] << 8);

            DIGITAL_DATA.CH4_V_L = DealDataBuffer[index++];
            DIGITAL_DATA.CH4_V_L |= (ushort)(DealDataBuffer[index++] << 8);

            DIGITAL_DATA.CH4_I_L = DealDataBuffer[index++];
            DIGITAL_DATA.CH4_I_L |= (ushort)(DealDataBuffer[index++] << 8);

            DIGITAL_DATA.CH1_I_P24_CAL = DealDataBuffer[index++];
            DIGITAL_DATA.CH1_I_P24_CAL |= (ushort)(DealDataBuffer[index++] << 8);

            DIGITAL_DATA.CH2_I_P15_CAL = DealDataBuffer[index++];
            DIGITAL_DATA.CH2_I_P15_CAL |= (ushort)(DealDataBuffer[index++] << 8);

            DIGITAL_DATA.CH3_I_N15_CAL = DealDataBuffer[index++];
            DIGITAL_DATA.CH3_I_N15_CAL |= (ushort)(DealDataBuffer[index++] << 8);

            DIGITAL_DATA.CH4_I_L_CAL = DealDataBuffer[index++];
            DIGITAL_DATA.CH4_I_L_CAL |= (ushort)(DealDataBuffer[index++] << 8);

            DIGITAL_DATA.MODULE_TYPE = DealDataBuffer[index++];
            DIGITAL_DATA.MODULE_TYPE |= (ushort)(DealDataBuffer[index++] << 8);

            DIGITAL_DATA.BD_TYPE = DealDataBuffer[index++];
            DIGITAL_DATA.BD_TYPE |= (ushort)(DealDataBuffer[index++] << 8);

            DIGITAL_DATA.POWERUP = DealDataBuffer[index++];
            DIGITAL_DATA.POWERUP |= (ushort)(DealDataBuffer[index++] << 8);

            DIGITAL_DATA.DOWNLOAD = DealDataBuffer[index++];
            DIGITAL_DATA.DOWNLOAD |= (ushort)(DealDataBuffer[index++] << 8);
        }

        private void myhid_UIDeal_Di() {
            try {
                short VCH = 0, ICH = 0;
                textBox_Digital_CHV_1.Text = ((short)DIGITAL_DATA.CH1_V_P24).ToString("D");
                textBox_Digital_CHV_2.Text = ((short)DIGITAL_DATA.CH2_V_P15).ToString("D");
                textBox_Digital_CHV_3.Text = ((short)DIGITAL_DATA.CH3_V_N15).ToString("D");
                textBox_Digital_CHV_4.Text = ((short)DIGITAL_DATA.CH4_V_L).ToString("D");
                VCH = (short)((3300 * (float)DIGITAL_DATA.ADC_CH5_D1 / 4096) * 16 - (3300 * (float)DIGITAL_DATA.ADC_CH5_D2 / 4096) * 16);
                textBox_Digital_CHV_5.Text = (VCH).ToString("D");
                VCH = (short)((3300 * (float)DIGITAL_DATA.ADC_CH6_D1 / 4096) * 16 - (3300 * (float)DIGITAL_DATA.ADC_CH6_D2 / 4096) * 16);
                textBox_Digital_CHV_6.Text = (VCH).ToString("D");
                VCH = (short)((3300 * (float)DIGITAL_DATA.ADC_CH7_D1 / 4096) * 16 - (3300 * (float)DIGITAL_DATA.ADC_CH7_D2 / 4096) * 16);
                textBox_Digital_CHV_7.Text = (VCH).ToString("D");
                VCH = (short)((3300 * (float)DIGITAL_DATA.ADC_CH8_D1 / 4096) * 16 - (3300 * (float)DIGITAL_DATA.ADC_CH8_D2 / 4096) * 16);
                textBox_Digital_CHV_8.Text = (VCH).ToString("D");

                textBox_Digital_CHI_1.Text = ((short)DIGITAL_DATA.CH1_I_P24).ToString("D");
                textBox_Digital_CHI_2.Text = ((short)DIGITAL_DATA.CH2_I_P15).ToString("D");
                textBox_Digital_CHI_3.Text = ((short)DIGITAL_DATA.CH3_I_N15).ToString("D");
                textBox_Digital_CHI_4.Text = ((short)DIGITAL_DATA.CH4_I_L).ToString("D");
                ICH = (short)(((3300 * DIGITAL_DATA.ADC_CH5_D1 / 4096) * 16 - (3300 * DIGITAL_DATA.ADC_CH5_D2 / 4096) * 16) / 1000);
                textBox_Digital_CHI_5.Text = (ICH).ToString("D");
                ICH = (short)(((3300 * DIGITAL_DATA.ADC_CH6_D1 / 4096) * 16 - (3300 * DIGITAL_DATA.ADC_CH6_D2 / 4096) * 16) / 1000);
                textBox_Digital_CHI_6.Text = (ICH).ToString("D");
                ICH = (short)(((3300 * DIGITAL_DATA.ADC_CH7_D1 / 4096) * 16 - (3300 * DIGITAL_DATA.ADC_CH7_D2 / 4096) * 16) / 1000);
                textBox_Digital_CHI_7.Text = (ICH).ToString("D");
                ICH = (short)(((3300 * DIGITAL_DATA.ADC_CH8_D1 / 4096) * 16 - (3300 * DIGITAL_DATA.ADC_CH8_D2 / 4096) * 16) / 1000);
                textBox_Digital_CHI_8.Text = (ICH).ToString("D");

                textBox_ModulType.Text = DIGITAL_DATA.MODULE_TYPE.ToString("X4");
                textBox_BDType.Text = DIGITAL_DATA.BD_TYPE.ToString("X4");

                if (DIGITAL_DATA.POWERUP == 0x0001)
                    label_PowerUp.Text = ":开";
                else
                    label_PowerUp.Text = ":关";

                if (DIGITAL_DATA.DOWNLOAD == 0x0001)
                    label_DownLoadFlag.Text = "已完成";
                else
                    label_DownLoadFlag.Text = "未完成";
            } catch {

            }
        }

        private void Digital_PowerUp_Set(string powerup) {
            TxDataPart_Di = new byte[2]; //重置发送数据的缓冲
            if (powerup == ":关") {
                label_PowerUp.Text = ":开";
                TxDataPart_Di[0] = 0x01;
            } else if (powerup == ":开") {
                label_PowerUp.Text = ":关";
                TxDataPart_Di[0] = 0x00;
            }
            TxDataPart_Di[1] = 0x00;
            FX2N_Send(myHid_Di, FX2N_CMD_TypeDef.FX2N_WRITE, 0x6F1C, 2, TxDataPart_Di);
        }

        private void Digital_I_Verify() {
            Digital_PowerUp_Set(":关"); //给被测物供电
            Thread.Sleep(3000);
            Digital_Read_All();
            myhid_UIDeal_Di();
            short data = 0;
            TxDataPart_Di = new byte[2]; //重置发送数据的缓冲
            data = Convert.ToInt16(textBox_Digital_CHI_1.Text);
            TxDataPart_Di[0] = (byte)(data & 0xFF);
            TxDataPart_Di[1] = (byte)(data >> 8 & 0xFF);
            FX2N_Send(myHid_Di, FX2N_CMD_TypeDef.FX2N_WRITE, 0x6F10, 2, TxDataPart_Di);
            data = Convert.ToInt16(textBox_Digital_CHI_2.Text);
            TxDataPart_Di[0] = (byte)(data & 0xFF);
            TxDataPart_Di[1] = (byte)(data >> 8 & 0xFF);
            FX2N_Send(myHid_Di, FX2N_CMD_TypeDef.FX2N_WRITE, 0x6F12, 2, TxDataPart_Di);
            data = Convert.ToInt16(textBox_Digital_CHI_3.Text);
            TxDataPart_Di[0] = (byte)(data & 0xFF);
            TxDataPart_Di[1] = (byte)(data >> 8 & 0xFF);
            FX2N_Send(myHid_Di, FX2N_CMD_TypeDef.FX2N_WRITE, 0x6F14, 2, TxDataPart_Di);
            data = Convert.ToInt16(textBox_Digital_CHI_4.Text);
            TxDataPart_Di[0] = (byte)(data & 0xFF);
            TxDataPart_Di[1] = (byte)(data >> 8 & 0xFF);
            FX2N_Send(myHid_Di, FX2N_CMD_TypeDef.FX2N_WRITE, 0x6F16, 2, TxDataPart_Di);
        }

        public string Digital_Get_Message() {
            return richTextBox_DigitalMessage.Text;
        }

        public void Digital_Download_Message(string oldstring, int tick) {
            richTextBox_DigitalMessage.Text = oldstring + tick.ToString("D") + "\r\n";
        }

        public void Digital_Download_Message(string message) {
            richTextBox_DigitalMessage.Text += message + "\r\n";
        }

        public void SelectBIN() {
            OpenFileDialog openfile = new OpenFileDialog();
            openfile.Filter = "bin files(*.bin)|*.bin";
            if (openfile.ShowDialog(Form1.form1) == DialogResult.OK) {
                BIN_Path = openfile.FileName;
                textBox_BIN_Path.Text = BIN_Path;
            }
        }

        private bool Rang_Verify(double be_test_value, double min_value, double max_value, string message) {
            if ((be_test_value < min_value) || (be_test_value > max_value)) {
                richTextBox_DigitalMessage.Text += message + "异常" + "\r\n";
                return false;
            } else {
                richTextBox_DigitalMessage.Text += message + "正常" + "\r\n";
                return true;
            }
        }

        private void Digital_OneKey_Detection() {
            richTextBox_DigitalMessage.Text = "开始硬件检测" + "\r\n";
            richTextBox_DigitalMessage.BackColor = SystemColors.Window;
            if (comboBox_ModulType.Text == "") {
                MessageBox.Show("请先选择被测扩展模块型号");
                return;
            }
            //Digital_I_Verify(); //校准一下硬件检测电流
            checkBox_ContinueRead1.Checked = false;
            Thread.Sleep(2000);

            Digital_Read_All();
            myhid_UIDeal_Di();
            bool error = false;
            int arg_location = 0;
            int ch = 0;
            string message = "";
            try {
                if (textBox_ModulType.Text == comboBox_ModulType.Text) {
                    richTextBox_DigitalMessage.Text += "被测物型号正确" + "\r\n";
                    //检测通道1
                    message = "通道" + (ch + 1).ToString() + "电压";
                    error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHV[ch].Text), 22800, 25000, message);
                    message = "通道" + (ch + 1).ToString() + "电流";
                    arg_location = 10;
                    if (1 == cell[arg_location, 3].IntValue) {
                        if (1 == DIGITAL_DATA.DOWNLOAD) {
                            error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHI[ch].Text), cell[arg_location, 1].DoubleValue, cell[arg_location, 2].DoubleValue, message);
                        }
                    } else {
                        error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHI[ch].Text), cell[arg_location, 1].DoubleValue, cell[arg_location, 2].DoubleValue, message);
                    }
                    ch++;
                    //检测通道2
                    message = "通道" + (ch + 1).ToString() + "电压";
                    error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHV[ch].Text), 14300, 15800, message);
                    message = "通道" + (ch + 1).ToString() + "电流";
                    arg_location = 11;
                    if (1 == cell[arg_location, 3].IntValue) {
                        if (1 == DIGITAL_DATA.DOWNLOAD) {
                            error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHI[ch].Text), cell[arg_location, 1].DoubleValue, cell[arg_location, 2].DoubleValue, message);
                        }
                    } else {
                        error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHI[ch].Text), cell[arg_location, 1].DoubleValue, cell[arg_location, 2].DoubleValue, message);
                    }
                    ch++;
                    //检测通道3
                    message = "通道" + (ch + 1).ToString() + "电压";
                    error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHV[ch].Text), -15800, -14300, message);
                    message = "通道" + (ch + 1).ToString() + "电流";
                    arg_location = 12;
                    if (1 == cell[arg_location, 3].IntValue) {
                        if (1 == DIGITAL_DATA.DOWNLOAD) {
                            error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHI[ch].Text), cell[arg_location, 1].DoubleValue, cell[arg_location, 2].DoubleValue, message);
                        }
                    } else {
                        error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHI[ch].Text), cell[arg_location, 1].DoubleValue, cell[arg_location, 2].DoubleValue, message);
                    }
                    ch++;
                    //检测通道4
                    message = "通道" + (ch + 1).ToString() + "电压";
                    if (Convert.ToUInt16(textBox_ModulType.Text) == (ushort)MODULE_TYPE_TypeDef.MODULE_TYPE_2WT) {
                        error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHV[ch].Text), 6500, 7200, message);
                    } else {
                        error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHV[ch].Text), 4500, 5200, message);
                    }
                    message = "通道" + (ch + 1).ToString() + "电流";
                    arg_location = 13;
                    if (1 == cell[arg_location, 3].IntValue) {
                        if (1 == DIGITAL_DATA.DOWNLOAD) {
                            error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHI[ch].Text), cell[arg_location, 1].DoubleValue, cell[arg_location, 2].DoubleValue, message);
                        }
                    } else {
                        error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHI[ch].Text), cell[arg_location, 1].DoubleValue, cell[arg_location, 2].DoubleValue, message);
                    }
                    ch++;
                    //检测通道5
                    message = "通道" + (ch + 1).ToString() + "电压";
                    arg_location = 14;
                    if (1 == cell[arg_location, 3].IntValue) {
                        if (1 == DIGITAL_DATA.DOWNLOAD) {
                            error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHV[ch].Text), cell[arg_location, 1].DoubleValue, cell[arg_location, 2].DoubleValue, message);
                        }
                    } else {
                        error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHV[ch].Text), cell[arg_location, 1].DoubleValue, cell[arg_location, 2].DoubleValue, message);
                    }
                    message = "通道" + (ch + 1).ToString() + "电流";
                    arg_location = 15;
                    if (1 == cell[arg_location, 3].IntValue) {
                        if (1 == DIGITAL_DATA.DOWNLOAD) {
                            error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHI[ch].Text), cell[arg_location, 1].DoubleValue, cell[arg_location, 2].DoubleValue, message);
                        }
                    } else {
                        error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHI[ch].Text), cell[arg_location, 1].DoubleValue, cell[arg_location, 2].DoubleValue, message);
                    }
                    ch++;
                    //检测通道6
                    message = "通道" + (ch + 1).ToString() + "电压";
                    arg_location = 16;
                    if (1 == cell[arg_location, 3].IntValue) {
                        if (1 == DIGITAL_DATA.DOWNLOAD) {
                            error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHV[ch].Text), cell[arg_location, 1].DoubleValue, cell[arg_location, 2].DoubleValue, message);
                        }
                    } else {
                        error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHV[ch].Text), cell[arg_location, 1].DoubleValue, cell[arg_location, 2].DoubleValue, message);
                    }
                    message = "通道" + (ch + 1).ToString() + "电流";
                    arg_location = 17;
                    if (1 == cell[arg_location, 3].IntValue) {
                        if (1 == DIGITAL_DATA.DOWNLOAD) {
                            error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHI[ch].Text), cell[arg_location, 1].DoubleValue, cell[arg_location, 2].DoubleValue, message);
                        }
                    } else {
                        error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHI[ch].Text), cell[arg_location, 1].DoubleValue, cell[arg_location, 2].DoubleValue, message);
                    }
                    ch++;
                    //检测通道7
                    message = "通道" + (ch + 1).ToString() + "电压";
                    arg_location = 18;
                    if (1 == cell[arg_location, 3].IntValue) {
                        if (1 == DIGITAL_DATA.DOWNLOAD) {
                            error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHV[ch].Text), cell[arg_location, 1].DoubleValue, cell[arg_location, 2].DoubleValue, message);
                        }
                    } else {
                        error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHV[ch].Text), cell[arg_location, 1].DoubleValue, cell[arg_location, 2].DoubleValue, message);
                    }
                    message = "通道" + (ch + 1).ToString() + "电流";
                    arg_location = 19;
                    if (1 == cell[arg_location, 3].IntValue) {
                        if (1 == DIGITAL_DATA.DOWNLOAD) {
                            error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHI[ch].Text), cell[arg_location, 1].DoubleValue, cell[arg_location, 2].DoubleValue, message);
                        }
                    } else {
                        error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHI[ch].Text), cell[arg_location, 1].DoubleValue, cell[arg_location, 2].DoubleValue, message);
                    }
                    ch++;
                    //检测通道8
                    message = "通道" + (ch + 1).ToString() + "电压";
                    arg_location = 20;
                    if (1 == cell[arg_location, 3].IntValue) {
                        if (1 == DIGITAL_DATA.DOWNLOAD) {
                            error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHV[ch].Text), cell[arg_location, 1].DoubleValue, cell[arg_location, 2].DoubleValue, message);
                        }
                    } else {
                        error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHV[ch].Text), cell[arg_location, 1].DoubleValue, cell[arg_location, 2].DoubleValue, message);
                    }
                    message = "通道" + (ch + 1).ToString() + "电流";
                    arg_location = 21;
                    if (1 == cell[arg_location, 3].IntValue) {
                        if (1 == DIGITAL_DATA.DOWNLOAD) {
                            error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHI[ch].Text), cell[arg_location, 1].DoubleValue, cell[arg_location, 2].DoubleValue, message);
                        }
                    } else {
                        error = Rang_Verify(Convert.ToDouble(textBox_Digital_CHI[ch].Text), cell[arg_location, 1].DoubleValue, cell[arg_location, 2].DoubleValue, message);
                    }
                    ch++;
                } else {
                    richTextBox_DigitalMessage.Text += "被测物型号错误" + "\r\n";
                    error = false;
                }
                if (!error) {
                    Digital_PowerUp_Set(":开"); //被测物断电
                    richTextBox_DigitalMessage.BackColor = Color.Red;
                } else {
                    MessageBox.Show("硬件检测成功");
                }
            } catch {
                richTextBox_DigitalMessage.Text += "配置文件错误" + "\r\n";
            }

            Onekey_Detection_Flag = false;
        }

        private void Digital_OneKey_DownLoad() {
            //开始固件下载
            richTextBox_DigitalMessage.Text = "开始下载固件" + "\r\n";
            richTextBox_DigitalMessage.BackColor = SystemColors.Window;
            Parity myParity = myCOM.Parity;
            myCOM.Parity = Parity.Even;
            stmLoad.STM32Load();
            myCOM.Parity = myParity;
            while (DIGITAL_DATA.DOWNLOAD == 0) {
                MessageBox.Show("固件下载成功，请按下BOOT按钮进行确认,确认后会再次进行硬件检测");
                Digital_Read_All();
                myhid_UIDeal_Di();
                Thread.Sleep(500);
            }
            Digital_OneKey_Detection();
            Onekey_DownLoad_Flag = false;
        }

        /*****************************************FX2N协议部分***************************************************/
        private bool FX2N_Read_DataDeal(ref byte[] DealDataBuffer, ref bool Rec_Completed, ref byte RecDataIndex, ref byte[] RxDataBuffer) {
            if (Rec_Completed) {
                byte i = 0;
                byte j = 0;
                string receivechar = "";
                Rec_Completed = false;
                while (RecDataIndex > i) {
                    receivechar = Convert.ToChar(RxDataBuffer[i++]).ToString();
                    receivechar += Convert.ToChar(RxDataBuffer[i++]).ToString();
                    DealDataBuffer[j++] = Convert.ToByte(receivechar, 16);
                    receivechar = "";
                }
                RecDataIndex = 0;
                return true;
            }
            return false;
        }

        private void FX2N_Read(report e, ref bool Rec_Start, ref bool Rec_Completed, ref byte RecDataIndex, ref byte[] RxDataBuffer) {
            //DisplayRXData(e.reportBuff.Length, e.reportBuff);
            byte Deal_Len = 2;
            if (e.reportBuff[2] == 0x02) {
                Rec_Start = true;
                RecDataIndex = 0;
                Deal_Len = 3;
                while (e.reportBuff[0] + 2 > Deal_Len) {
                    if (e.reportBuff[Deal_Len] == 0x03 && Deal_Len == e.reportBuff[0] - 1) {
                        Rec_Completed = true;
                        Rec_Start = false;
                        break;
                    }
                    RxDataBuffer[RecDataIndex++] = e.reportBuff[Deal_Len++];
                }
            } else if (Rec_Start) {
                Deal_Len = 2;
                while (e.reportBuff[0] + 2 > Deal_Len) {
                    if (e.reportBuff[Deal_Len] == 0x03 && Deal_Len == e.reportBuff[0] - 1) {
                        Rec_Completed = true;
                        Rec_Start = false;
                        break;
                    }
                    RxDataBuffer[RecDataIndex++] = e.reportBuff[Deal_Len++];
                }
            } else if (e.reportBuff[0] == 0x01 && e.reportBuff[2] == 0x06) {
                toolStripStatusLabelCOM.Text = "设置成功";
                //MessageBox.Show("设置成功");
            }
        }

        private void FX2N_Send(Hid hid_tar, FX2N_CMD_TypeDef cmd, ushort addr, byte RWlen, byte[] TxDataPart) {
            if (hid_tar.Opened == false) {
                return;
            }
            byte Len = 1;
            byte[] array;
            Byte[] data = new byte[24];
            data[Len++] = 0x00; //没有意义，只是填充用

            data[Len++] = 0x02; //STX

            FX2N_Send_Set_CMD(cmd, ref data, ref Len);
            array = Encoding.ASCII.GetBytes(addr.ToString("X4"));  //数组array为对应的ASCII数组
            data[Len++] = array[0];
            data[Len++] = array[1];
            data[Len++] = array[2];
            data[Len++] = array[3];

            array = Encoding.ASCII.GetBytes(RWlen.ToString("X2"));  //数组array为对应的ASCII数组
            data[Len++] = array[0];
            data[Len++] = array[1];

            if (cmd == FX2N_CMD_TypeDef.FX2N_WRITE) {
                byte i = 0;
                for (i = 0; i < RWlen; i++) {
                    array = Encoding.ASCII.GetBytes(TxDataPart[i].ToString("X2"));  //数组array为对应的ASCII数组
                    data[Len++] = array[0];
                    data[Len++] = array[1];
                }
            }

            data[Len++] = 0x03; //ETX

            byte lrc = LRC(data, 3, (byte)(Len - 3));
            array = Encoding.ASCII.GetBytes(lrc.ToString("X2"));  //数组array为对应的ASCII数组
            data[Len++] = array[0];
            data[Len++] = array[1];

            data[0] = (byte)(Len - 2); //发送的FX2N的数据的长度

            report r = new report(0, data);
            hid_tar.Write(r);
            Thread.Sleep(50);
        }

        private byte LRC(byte[] data, byte startaddr, byte Len) {
            byte lrc = 0;
            byte i = 0;
            for (i = 0; i < Len; i++) {
                lrc += data[startaddr + i];
            }
            return lrc;
        }

        private void FX2N_Send_Set_CMD(FX2N_CMD_TypeDef cmd, ref byte[] data, ref byte Len) {
            switch (cmd) {
                case FX2N_CMD_TypeDef.FX2N_READ:
                    data[Len++] = 0x45;
                    data[Len++] = 0x30;
                    data[Len++] = 0x30;
                    break;
                case FX2N_CMD_TypeDef.FX2N_WRITE:
                    data[Len++] = 0x45;
                    data[Len++] = 0x31;
                    data[Len++] = 0x30;
                    break;
            }
        }

        /*****************************************其他部分***************************************************/
        private void Tooling_Enable() {
            if (myHid_An.Opened == true && myHid_Di.Opened == true && myCOM.IsOpen == true) {
                tabControl1.Enabled = true;
            } else {
                tabControl1.Enabled = false;
            }
        }

        //private void Save_Res_Value() {
        //    //获取用户选择的excel文件名称
        //    string path;
        //    //获取保存路径
        //    path = Directory.GetCurrentDirectory() + "\\配置文件\\Configuration.xls";
        //    Workbook wb = new Workbook();
        //    Worksheet ws = wb.Worksheets[0];
        //    Cells cell = ws.Cells;
        //    try {
        //        cell[0, 0].PutValue(label152.Text); //标题
        //        cell[0, 1].PutValue(textBoxRES1.Text);
        //        cell[1, 0].PutValue(label151.Text); //标题
        //        cell[1, 1].PutValue(textBoxRES2.Text);
        //        cell[2, 0].PutValue(label150.Text); //标题
        //        cell[2, 1].PutValue(textBoxRES3.Text);
        //        cell[3, 0].PutValue(label149.Text); //标题
        //        cell[3, 1].PutValue(textBoxRES4.Text);
        //        cell[4, 0].PutValue(label156.Text); //标题
        //        cell[4, 1].PutValue(textBoxRES5.Text);
        //        cell[5, 0].PutValue(label155.Text); //标题
        //        cell[5, 1].PutValue(textBoxRES6.Text);
        //        cell[6, 0].PutValue(label154.Text); //标题
        //        cell[6, 1].PutValue(textBoxRES7.Text);
        //        cell[7, 0].PutValue(label153.Text); //标题
        //        cell[7, 1].PutValue(textBoxRES8.Text);
        //        //保存excel表格
        //        wb.Save(path);
        //    } catch {

        //    }
        //}

        public Cells cell;
        private void Get_Configuration() {
            //获取用户选择的excel文件名称
            string path;
            path = Directory.GetCurrentDirectory() + "\\配置文件\\Configuration.xls";
            //打开excel表格
            try {
                Workbook wb = new Workbook();
                wb.Open(path);
                Worksheet ws = wb.Worksheets[0];
                cell = ws.Cells;
                textBoxRES1.Text = cell[0, 1].StringValue;
                textBoxRES2.Text = cell[1, 1].StringValue;
                textBoxRES3.Text = cell[2, 1].StringValue;
                textBoxRES4.Text = cell[3, 1].StringValue;
                textBoxRES5.Text = cell[4, 1].StringValue;
                textBoxRES6.Text = cell[5, 1].StringValue;
                textBoxRES7.Text = cell[6, 1].StringValue;
                textBoxRES8.Text = cell[7, 1].StringValue;

                ANALOG_RES[0] = Convert.ToDouble(textBoxRES1.Text);
                ANALOG_RES[1] = Convert.ToDouble(textBoxRES2.Text);
                ANALOG_RES[2] = Convert.ToDouble(textBoxRES3.Text);
                ANALOG_RES[3] = Convert.ToDouble(textBoxRES4.Text);
                ANALOG_RES[4] = Convert.ToDouble(textBoxRES5.Text);
                ANALOG_RES[5] = Convert.ToDouble(textBoxRES6.Text);
                ANALOG_RES[6] = Convert.ToDouble(textBoxRES7.Text);
                ANALOG_RES[7] = Convert.ToDouble(textBoxRES8.Text);

                comboBoxPort.Text = cell[23, 1].StringValue;
                comboBox_ModulType.Text = Convert.ToUInt16(cell[24, 1].StringValue, 16).ToString("X4");
            } catch {
                //Save_Res_Value();
                MessageBox.Show("Configuration.xls文件丢失或错误，即将推出程序");
                Environment.Exit(0);
            }
        }
    }
}