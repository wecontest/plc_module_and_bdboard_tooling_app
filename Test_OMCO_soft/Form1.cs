using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

using System.IO.Ports;
using HID;

namespace Testsoft {
    public partial class Form1 : Form {
        public static Form1 form1;
        public Form1() {
            Control.CheckForIllegalCrossThreadCalls = false;
            InitializeComponent();
            softinitial();
            form1 = this;
        }

        private void buttonOpenCOM_Click(object sender, EventArgs e) {
            if (iniCOM() == true) {
                buttonOpenCOM.BackColor = Color.Red;
                buttonOpenCOM.Text = "关闭串口";
                groupBox1.Enabled = false;
                toolStripStatusLabelCOM.Text = "端口打开";
            } else {
                buttonOpenCOM.BackColor = Color.LightBlue;
                buttonOpenCOM.Text = "打开串口";
                groupBox1.Enabled = true;
                toolStripStatusLabelCOM.Text = "串口未打开";
            }

            Tooling_Enable();
        }

        private void toolStripContainer1_ContentPanel_Load(object sender, System.EventArgs e) {
            throw new System.NotImplementedException();
        }

        private void label5_Click(object sender, System.EventArgs e) {
            throw new System.NotImplementedException();
        }

        private void buttonCleanReceive_Click(object sender, System.EventArgs e) {
            richTextBoxDataRecieve.Text = "";
        }

        private void buttonCleanSend_Click(object sender, System.EventArgs e) {
            richTextBoxDataSend.Text = "";
        }

        private void tBox_KeyPress(object sender, KeyPressEventArgs e) {
            if (e.KeyChar == 0x20) e.KeyChar = (char)0;  //禁止空格键
            if (e.KeyChar == 0x2D) return;   //处理负数
            if (e.KeyChar > 0x20) {
                try {
                    double.Parse(((TextBox)sender).Text + e.KeyChar.ToString());
                } catch {
                    e.KeyChar = (char)0;   //处理非法字符
                }
            }

        }

        private void tBox_KeyUp(object sender, KeyEventArgs e) {
            if (((TextBox)sender).Text.Length == 1) return;
            try {
                if ((Convert.ToDouble(((TextBox)sender).Text) * 100) > 32767 || (Convert.ToDouble(((TextBox)sender).Text) * 100) < -32768) {
                    MessageBoxERRwarning(6);
                    ((TextBox)sender).Text = "";
                }
            } catch {
                ((TextBox)sender).Text = "";
            }
        }

        private void tBox_KeyUpT(object sender, KeyEventArgs e) {
            if (((TextBox)sender).Text.Length == 1) return;
            try {
                if ((Convert.ToDouble(((TextBox)sender).Text) * 100) > 32767 || (Convert.ToDouble(((TextBox)sender).Text) * 100) < -32768) {
                    MessageBoxERRwarning(8);
                    ((TextBox)sender).Text = "";
                }
            } catch {
                ((TextBox)sender).Text = "";
            }
        }

        private void tBox_KeyUpV(object sender, KeyEventArgs e) {
            if (((TextBox)sender).Text.Length == 1) return;
            try {
                if (Convert.ToDouble(((TextBox)sender).Text) > 10 || Convert.ToDouble(((TextBox)sender).Text) < -10) {
                    MessageBoxERRwarning(10);
                    ((TextBox)sender).Text = "";
                }
            } catch {
                ((TextBox)sender).Text = "";
            }
        }

        private void tBox_KeyUpI(object sender, KeyEventArgs e) {
            if (((TextBox)sender).Text.Length == 1) return;
            try {
                if (Convert.ToDouble(((TextBox)sender).Text) > 20 || Convert.ToDouble(((TextBox)sender).Text) < 0) {
                    MessageBoxERRwarning(11);
                    ((TextBox)sender).Text = "";
                }
            } catch {
                ((TextBox)sender).Text = "";
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
            //myhid_RxDataDeal_Di((byte[])e.Result);
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e) {
            //myhid_RxDataDeal_Di((byte[])e.Argument);
            e.Result = e.Argument;
        }

        private void button_OpenDigital_Click(object sender, EventArgs e) {
            if (myHid_Di.Opened == false) {
                myHid_Di = new Hid(); //初始化该USB端口，在重新打开USB时一定要这步
                myHidPtr_Di = new IntPtr();
                myHid_Di.DataReceived += myhid_DataReceived_Di; //订阅DataRec事件
                myHid_Di.DataProcessed += myhid_DataProcessed_Di;
                myHid_Di.DeviceRemoved += myhid_DeviceRemoved;
                UInt16 myVendorID = Convert.ToUInt16("0483", 16);
                UInt16 myProductID = Convert.ToUInt16("5750", 16);
                string productname = "DIGITAL MODULE VER1";
                if ((int)(myHidPtr_Di = myHid_Di.OpenDevice(myVendorID, myProductID, productname)) != -1) {
                    button_OpenDigital.BackColor = Color.Red;
                    Digital_Read_All();
                    myhid_UIDeal_Di();
                } else {
                    button_OpenDigital.BackColor = Color.LightBlue;
                    MessageBoxERRwarning(9);
                }
            } else {
                myHid_Di.CloseDevice(myHidPtr_Di);
                button_OpenDigital.BackColor = Color.LightBlue;
            }

            Tooling_Enable();
        }

        private void button_OpenAnalog_Click(object sender, EventArgs e) {
            if (myHid_An.Opened == false) {
                myHid_An = new Hid(); //初始化该USB端口，在重新打开USB时一定要这步
                myHidPtr_An = new IntPtr();
                myHid_An.DataReceived += myhid_DataReceived_An; //订阅DataRec事件
                myHid_An.DataProcessed += myhid_DataProcessed_An;
                myHid_An.DeviceRemoved += myhid_DeviceRemoved;
                UInt16 myVendorID = Convert.ToUInt16("0483", 16);
                UInt16 myProductID = Convert.ToUInt16("5750", 16);
                string productname = "ANALOG MODULE VER1";
                if ((int)(myHidPtr_An = myHid_An.OpenDevice(myVendorID, myProductID, productname)) != -1) {
                    button_OpenAnalog.BackColor = Color.Red;
                    Analog_Read_All();
                    myhid_UIDeal_An();
                } else {
                    button_OpenAnalog.BackColor = Color.LightBlue;
                    MessageBoxERRwarning(9);
                }
            } else {
                myHid_An.CloseDevice(myHidPtr_An);
                button_OpenAnalog.BackColor = Color.LightBlue;
            }

            Tooling_Enable();
        }

        private void button_Set_V_Click(object sender, EventArgs e) {
            Analog_DAC8760_V_Set(Convert.ToDouble(textBox_V.Text));
        }

        private void button_Set_I_Click(object sender, EventArgs e) {
            Analog_DAC8760_I_Set(Convert.ToDouble(textBox_I.Text));
        }

        private void radioButton_ADI_CheckedChanged(object sender, EventArgs e) {
            ANALYTE_TYPE = ANALYTE_TYPE_TypeDef.ANALYTE_ADI;
        }

        private void radioButton_DAI_CheckedChanged(object sender, EventArgs e) {
            ANALYTE_TYPE = ANALYTE_TYPE_TypeDef.ANALYTE_DAI;
        }

        private void radioButton_ADV_CheckedChanged(object sender, EventArgs e) {
            ANALYTE_TYPE = ANALYTE_TYPE_TypeDef.ANALYTE_ADV;
        }

        private void radioButton_DAV_CheckedChanged(object sender, EventArgs e) {
            ANALYTE_TYPE = ANALYTE_TYPE_TypeDef.ANALYTE_DAV;
        }

        private void radioButton_PT_CheckedChanged(object sender, EventArgs e) {
            ANALYTE_TYPE = ANALYTE_TYPE_TypeDef.ANALYTE_PT;
        }

        private void radioButton_TCV_CheckedChanged(object sender, EventArgs e) {
            ANALYTE_TYPE = ANALYTE_TYPE_TypeDef.ANALYTE_TCV;
        }

        private void radioButton_TCR_CheckedChanged(object sender, EventArgs e) {
            ANALYTE_TYPE = ANALYTE_TYPE_TypeDef.ANALYTE_TCR;
        }

        private void radioButton_NTC_CheckedChanged(object sender, EventArgs e) {
            ANALYTE_TYPE = ANALYTE_TYPE_TypeDef.ANALYTE_NTC;
        }

        private void button_SetType_Click(object sender, EventArgs e) {
            Analog_Analyte_Type_Select(ANALYTE_TYPE);
        }

        private void checkBox_CH1_CheckedChanged(object sender, EventArgs e) {
            byte channel = (byte)ANALYTE_CHANNEL_TypeDef.ANALYTE_CH1;
            if (checkBox_CH1.Checked) {
                ANALYTE_CHANNEL |= channel;
            } else {
                ANALYTE_CHANNEL &= (byte)~channel;
            }
        }

        private void checkBox_CH2_CheckedChanged(object sender, EventArgs e) {
            byte channel = (byte)ANALYTE_CHANNEL_TypeDef.ANALYTE_CH2;
            if (checkBox_CH2.Checked) {
                ANALYTE_CHANNEL |= channel;
            } else {
                ANALYTE_CHANNEL &= (byte)~channel;
            }
        }

        private void checkBox_CH3_CheckedChanged(object sender, EventArgs e) {
            byte channel = (byte)ANALYTE_CHANNEL_TypeDef.ANALYTE_CH3;
            if (checkBox_CH3.Checked) {
                ANALYTE_CHANNEL |= channel;
            } else {
                ANALYTE_CHANNEL &= (byte)~channel;
            }
        }

        private void checkBox_CH4_CheckedChanged(object sender, EventArgs e) {
            byte channel = (byte)ANALYTE_CHANNEL_TypeDef.ANALYTE_CH4;
            if (checkBox_CH4.Checked) {
                ANALYTE_CHANNEL |= channel;
            } else {
                ANALYTE_CHANNEL &= (byte)~channel;
            }
        }

        private void checkBox_CH5_CheckedChanged(object sender, EventArgs e) {
            byte channel = (byte)ANALYTE_CHANNEL_TypeDef.ANALYTE_CH5;
            if (checkBox_CH5.Checked) {
                ANALYTE_CHANNEL |= channel;
            } else {
                ANALYTE_CHANNEL &= (byte)~channel;
            }
        }

        private void checkBox_CH6_CheckedChanged(object sender, EventArgs e) {
            byte channel = (byte)ANALYTE_CHANNEL_TypeDef.ANALYTE_CH6;
            if (checkBox_CH6.Checked) {
                ANALYTE_CHANNEL |= channel;
            } else {
                ANALYTE_CHANNEL &= (byte)~channel;
            }
        }

        private void checkBox_CH7_CheckedChanged(object sender, EventArgs e) {
            byte channel = (byte)ANALYTE_CHANNEL_TypeDef.ANALYTE_CH7;
            if (checkBox_CH7.Checked) {
                ANALYTE_CHANNEL |= channel;
            } else {
                ANALYTE_CHANNEL &= (byte)~channel;
            }
        }

        private void checkBox_CH8_CheckedChanged(object sender, EventArgs e) {
            byte channel = (byte)ANALYTE_CHANNEL_TypeDef.ANALYTE_CH8;
            if (checkBox_CH8.Checked) {
                ANALYTE_CHANNEL |= channel;
            } else {
                ANALYTE_CHANNEL &= (byte)~channel;
            }
        }

        private void button_SetChannel_Click(object sender, EventArgs e) {
            Analog_Analyte_Channel_Select(ANALYTE_CHANNEL);
        }

        private void radioButton_Res1_CheckedChanged(object sender, EventArgs e) {
            ANALYTE_RES = ANALYTE_RES_TypeDef.ANALYTE_RES1_PT100F;
        }

        private void radioButton_Res2_CheckedChanged(object sender, EventArgs e) {
            ANALYTE_RES = ANALYTE_RES_TypeDef.ANALYTE_RES2_PT100S;
        }

        private void radioButton_Res3_CheckedChanged(object sender, EventArgs e) {
            ANALYTE_RES = ANALYTE_RES_TypeDef.ANALYTE_RES3_CU50F;
        }

        private void radioButton_Res4_CheckedChanged(object sender, EventArgs e) {
            ANALYTE_RES = ANALYTE_RES_TypeDef.ANALYTE_RES4_CU50S;
        }

        private void radioButton_Res5_CheckedChanged(object sender, EventArgs e) {
            ANALYTE_RES = ANALYTE_RES_TypeDef.ANALYTE_RES5_PT1000F;
        }

        private void radioButton_Res6_CheckedChanged(object sender, EventArgs e) {
            ANALYTE_RES = ANALYTE_RES_TypeDef.ANALYTE_RES6_PT1000S;
        }

        private void radioButton_Res7_CheckedChanged(object sender, EventArgs e) {
            ANALYTE_RES = ANALYTE_RES_TypeDef.ANALYTE_RES7_NTCF;
        }

        private void radioButton_Res8_CheckedChanged(object sender, EventArgs e) {
            ANALYTE_RES = ANALYTE_RES_TypeDef.ANALYTE_RES8_NTCS;
        }

        private void button_SetRes_Click(object sender, EventArgs e) {
            Analog_Analyte_Res_Select(ANALYTE_RES);
        }

        private void radioButton_G1_CheckedChanged(object sender, EventArgs e) {
            ADC7190_G = ADC7190_G_TypeDef.ADC7190_G1;
        }

        private void radioButton_G8_CheckedChanged(object sender, EventArgs e) {
            ADC7190_G = ADC7190_G_TypeDef.ADC7190_G8;
        }

        private void radioButton_G16_CheckedChanged(object sender, EventArgs e) {
            ADC7190_G = ADC7190_G_TypeDef.ADC7190_G16;
        }

        private void radioButton_G32_CheckedChanged(object sender, EventArgs e) {
            ADC7190_G = ADC7190_G_TypeDef.ADC7190_G32;
        }

        private void radioButton_G64_CheckedChanged(object sender, EventArgs e) {
            ADC7190_G = ADC7190_G_TypeDef.ADC7190_G64;
        }

        private void radioButton_G128_CheckedChanged(object sender, EventArgs e) {
            ADC7190_G = ADC7190_G_TypeDef.ADC7190_G128;
        }

        private void button_SetG_Click(object sender, EventArgs e) {
            Analog_ADC7190_G_Set(ADC7190_G);
        }

        private void button_OutputEnable_Click(object sender, EventArgs e) {
            Analog_DAC8760_Output_Switch_Flag_Set(label_OutputEnable.Text);
        }

        private void button_Auto_Cal_RES_Enable_Click(object sender, EventArgs e) {
            Analog_Auto_Cal_RES_Set(label_Auto_Cal_RES_Enable.Text);
        }

        private void button_ReadRes_Click(object sender, EventArgs e) {
            Get_Configuration();
            MessageBox.Show("文件读取成功");
        }

        private void button_Analog_Read_Click(object sender, EventArgs e) {
            Analog_Read_All();
            myhid_UIDeal_An();
        }

        private void checkBox_ContinueRead2_CheckedChanged(object sender, EventArgs e) {
            if (checkBox_ContinueRead2.Checked) {
                button_Set_V.Enabled = false;
                button_Set_I.Enabled = false;
                button_SetG.Enabled = false;
                button_SetType.Enabled = false;
                button_SetChannel.Enabled = false;
                button_SetRes.Enabled = false;
                button_OutputEnable.Enabled = false;
            } else {
                button_Set_V.Enabled = true;
                button_Set_I.Enabled = true;
                button_SetG.Enabled = true;
                button_SetType.Enabled = true;
                button_SetChannel.Enabled = true;
                button_SetRes.Enabled = true;
                button_OutputEnable.Enabled = true;
            }
        }

        Thread Digital_thread;
        bool Onekey_Detection_Flag = false;
        private void button_Digital_Onekey_Detection_Click(object sender, EventArgs e) {
            Digital_PowerUp_Set(":开"); //被测物断电
            Digital_PowerUp_Set(":关"); //给被测物供电
            try {
                if (Digital_thread.IsAlive == true) {
                    if (Onekey_Detection_Flag == true) {
                        Digital_thread.Abort();
                        Onekey_Detection_Flag = false;
                    } else {
                        MessageBox.Show("请先等待固件烧写完成");
                    }
                    return;
                }

            } catch { }
            Onekey_Detection_Flag = true;
            Digital_thread = new Thread(Digital_OneKey_Detection);
            Digital_thread.IsBackground = true;
            Digital_thread.Start();
            //Digital_OneKey_Detection();
        }

        bool Onekey_DownLoad_Flag = false;
        private void button_Digital_Onekey_DownLoad_Click(object sender, EventArgs e) {
            try {
                if (Digital_thread.IsAlive == true) {
                    if (Onekey_DownLoad_Flag == true) {
                        Digital_thread.Abort();
                        Onekey_DownLoad_Flag = false;
                        myCOM.Parity = Parity.None;
                    } else {
                        MessageBox.Show("请先等待一键硬件检测完成");
                    }
                    return;
                }
            } catch { }
            Onekey_DownLoad_Flag = true;
            Digital_thread = new Thread(Digital_OneKey_DownLoad);
            Digital_thread.IsBackground = true;
            Digital_thread.Start();
        }

        Thread Analog_thread;
        bool Onekey_Calibration_Flag = false;
        private void button_Analog_Onekey_Calibration_Click(object sender, EventArgs e) {
            try {
                if (Analog_thread.IsAlive == true) {
                    if (Onekey_Calibration_Flag == true) {
                        Analog_thread.Abort();
                        Onekey_Calibration_Flag = false;
                    } else {
                        MessageBox.Show("请先等待一键出厂校验完成");
                    }
                    return;
                }
            } catch { }
            Onekey_Calibration_Flag = true;
            Analog_thread = new Thread(Analog_OneKey_Calibration);
            Analog_thread.IsBackground = true;
            Analog_thread.Start();
            //Analog_OneKey_Detection();
        }

        bool Onekey_Verify_Flag = false;
        private void button_Analog_Onekey_Verify_Click(object sender, EventArgs e) {
            try {
                if (Analog_thread.IsAlive == true) {
                    if (Onekey_Verify_Flag == true) {
                        Analog_thread.Abort();
                        Onekey_Verify_Flag = false;
                    } else {
                        MessageBox.Show("请先等待一键出厂校准完成");
                    }
                    return;
                }
            } catch { }
            Onekey_Verify_Flag = true;
            Analog_thread = new Thread(Analog_OneKey_Verify);
            Analog_thread.IsBackground = true;
            Analog_thread.Start();
        }

        private void button_Digital_Read_Click(object sender, EventArgs e) {
            Digital_Read_All();
            myhid_UIDeal_Di();
        }

        private void button_Digital_Power_Click(object sender, EventArgs e) {
            Digital_PowerUp_Set(label_PowerUp.Text);
        }

        Thread Digital_Verify_thread;
        private void button_Digital_I_Verify_Click(object sender, EventArgs e) {
            try {
                if (Digital_Verify_thread.IsAlive == true) {
                    Digital_thread.Abort();
                }
            } catch { }
            Digital_Verify_thread = new Thread(Digital_I_Verify);
            Digital_Verify_thread.IsBackground = true;
            Digital_Verify_thread.Start();
        }

        private void checkBox_ContinueRead1_CheckedChanged(object sender, EventArgs e) {
            if (checkBox_ContinueRead1.Checked) {
                button_Digital_Power.Enabled = false;
            } else {
                button_Digital_Power.Enabled = true;
            }
        }

        private void button_SelectBIN_Click(object sender, EventArgs e) {
            SelectBIN();
        }

        private bool unlock_flag = false;
        private void button_OpenHideInterface_Click(object sender, EventArgs e) {
            if (textBox_PassWord.Text == "we-con" && !unlock_flag) {
                unlock_flag = true;
                tabControl1.TabPages.Add(tabPage7);
                tabControl1.TabPages.Add(tabPage6);
                tabControl1.TabPages.Add(tabPage1);
                tabControl1.TabPages.Add(tabPage3);
                tabControl1.TabPages.Add(tabPage9);
                tabControl1.TabPages.Add(tabPage2);
                tabControl1.TabPages.Add(tabPage5);
                tabControl1.TabPages.Add(tabPage4);
            }
        }

        private void textBox_PassWord_KeyPress(object sender, KeyPressEventArgs e) {
            if (e.KeyChar == (char)Keys.Enter) {
                button_OpenHideInterface_Click(null, null);
            }
        }

        Thread Analyte_thread;
        private void button_Get_Software_Ver_Click(object sender, EventArgs e) {
            try {
                if (Analyte_thread.IsAlive == true) {
                    Analyte_thread.Abort();
                }
            } catch { }
            Analyte_thread = new Thread(Get_Software_Ver);
            Analyte_thread.IsBackground = true;
            Analyte_thread.Start();
        }

        private void button_Get_Station_Num_Click(object sender, EventArgs e) {
            Get_Station_Num();
        }

        private void button_Set_Station_Num_Click(object sender, EventArgs e) {
            Set_Station_Num();
        }

        private void button_Get_Calibration_State_Click(object sender, EventArgs e) {
            Get_Calibration_State();
        }

        private void button_Get_Device_Type_Click(object sender, EventArgs e) {
            Get_Device_Type();
        }

        private void button_Set_CH_Result_Click(object sender, EventArgs e) {
            Set_CH_Result();
        }

        private void button_Get_CH_Result_Click(object sender, EventArgs e) {
            Get_CH_Result();
        }

        private void button_Set_CH_Switch_Click(object sender, EventArgs e) {
            Set_CH_Switch();
        }

        private void button_Get_CH_Switch_Click(object sender, EventArgs e) {
            Get_CH_Switch();
        }

        private void button_Set_CH_Type_Click(object sender, EventArgs e) {
            Set_CH_Type();
        }

        private void button_Get_CH_Type_Click(object sender, EventArgs e) {
            Get_CH_Type();
        }

        private void button_Set_TC_CJC_Mode_Click(object sender, EventArgs e) {
            Set_TC_CJC_Mode();
        }

        private void button_Get_TC_CJC_Mode_Click(object sender, EventArgs e) {
            Get_TC_CJC_Mode();
        }

        private void button_Set_TC_Sensor_Click(object sender, EventArgs e) {
            Set_TC_Sensor();
        }

        private void button_Get_TC_Sensor_Click(object sender, EventArgs e) {
            Get_TC_Sensor();
        }

        private void button_Get_TC_CJC_T_Click(object sender, EventArgs e) {
            Get_TC_CJC_T();
        }

        private void button_Get_Original_Value_Click(object sender, EventArgs e) {
            Get_Original_Value();
        }

        private void button_CH_Calibration_Click(object sender, EventArgs e) {
            Ch_Calibration();
        }

        private void button_Get_PHY_Click(object sender, EventArgs e) {
            Get_Physical_Variable();
        }

        private void button_Get_Cal_Par_Click(object sender, EventArgs e) {
            Get_Cal_Par();
        }

    }
}
