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

namespace Testsoft {
    public partial class Form1 : Form {
        //校准、校验标志定义
        enum CAL_VER_FLAG_TypeDef {
            CAL_FLAG, //进行校准
            VER_FLAG, //进行校验
        };

        public struct OTHER_TYPE_TypeDef {
            public double First_Point_Value;
            public double Second_Point_Value;
            public double Max_D_Value;
        }
        public OTHER_TYPE_TypeDef g_Standard_Values = new OTHER_TYPE_TypeDef() {
            First_Point_Value = 0,
            Second_Point_Value = 0,
        };

        public struct RES_TYPE_TypeDef {
            public ANALYTE_RES_TypeDef First_Point_Res;
            public ANALYTE_RES_TypeDef Second_Point_Res;
            public double Max_D_Value;
        }

        public RES_TYPE_TypeDef g_PT100_RES = new RES_TYPE_TypeDef() {
            First_Point_Res = ANALYTE_RES_TypeDef.ANALYTE_RES1_PT100F,
            Second_Point_Res = ANALYTE_RES_TypeDef.ANALYTE_RES2_PT100S,
        };

        public RES_TYPE_TypeDef g_CU50_RES = new RES_TYPE_TypeDef() {
            First_Point_Res = ANALYTE_RES_TypeDef.ANALYTE_RES3_CU50F,
            Second_Point_Res = ANALYTE_RES_TypeDef.ANALYTE_RES4_CU50S,
        };

        public RES_TYPE_TypeDef g_PT1000_RES = new RES_TYPE_TypeDef() {
            First_Point_Res = ANALYTE_RES_TypeDef.ANALYTE_RES5_PT1000F,
            Second_Point_Res = ANALYTE_RES_TypeDef.ANALYTE_RES6_PT1000S,
        };

        public RES_TYPE_TypeDef g_NTC_RES = new RES_TYPE_TypeDef() {
            First_Point_Res = ANALYTE_RES_TypeDef.ANALYTE_RES7_NTCF,
            Second_Point_Res = ANALYTE_RES_TypeDef.ANALYTE_RES8_NTCS,
        };

        //private delegate void FUNC();
        private bool Verify(ushort ch_mask, byte device_ch_num, double max_d_value) {
            int i = 0;
            double d_value = 0;
            for (i = 0; i < device_ch_num; i++) {
                if (0x01 == ((ch_mask >> i) & 0x01)) {
                    checkBox_Get_Phy_CH_Sel[i].Checked = true;
                } else {
                    checkBox_Get_Phy_CH_Sel[i].Checked = false;
                }
            }
            Get_Physical_Variable(); //校验
            for (i = 0; i < device_ch_num; i++) {
                if (0x01 == ((ch_mask >> i) & 0x01)) {
                    d_value = Convert.ToDouble(textBox_Cal_Phy[i].Text) - Convert.ToDouble(textBox_Get_Phy[i].Text);
                    if (d_value < -max_d_value || d_value > max_d_value) {
                        richTextBox_AnalogMessage.Text += "通道" + (i + 1).ToString() + "校验不通过" + "\r\n";
                        richTextBox_AnalogMessage.BackColor = Color.Red;
                        return false;
                    } else {
                        richTextBox_AnalogMessage.Text += "通道" + (i + 1).ToString() + "校验通过" + "\r\n";
                    }
                }
            }
            return true;
        }

        private void Analog_Calibrate_Verify(CAL_VER_FLAG_TypeDef c_v_flag) {
            //校准部分
            int i = 0;
            byte device_ch_num = 0;
            byte phy_num = 0;
            ushort ch_mask = 0;
            device_ch_num = 8;
            int arg_location = 0;
            string handle_message = "";
            if (CAL_VER_FLAG_TypeDef.CAL_FLAG == c_v_flag) {
                handle_message = "校准";
            } else {
                handle_message = "校验";
            }

            try {
                for (i = 0; i < cell[25, 1].IntValue; i++) {
                    arg_location = cell[26, i + 1].IntValue - 1;
                    richTextBox_AnalogMessage.Text += handle_message + cell[arg_location, 1].StringValue + "\r\n";
                    switch (cell[arg_location, 1].StringValue) {
                        case "TC内置冷端": {
                                //校准内置冷端温度
                                phy_num = (byte)(cell[arg_location, 2].IntValue);
                                textBox_Phy_Num_Cal.Text = phy_num.ToString();
                                textBox_Phy_Num_Get.Text = phy_num.ToString();
                                ch_mask = Convert.ToUInt16(cell[arg_location, 3].StringValue, 16);
                                g_Standard_Values.Max_D_Value = cell[arg_location, 4].DoubleValue;
                                int infrared_sel = cell[arg_location, 5].IntValue - 1;
                                g_Standard_Values.First_Point_Value = Convert.ToDouble(textBox_Infrared[infrared_sel].Text);
                                if (!Cal_Ver_TC_CJC_T(ch_mask, device_ch_num, g_Standard_Values, c_v_flag)) {
                                    return;
                                }
                                break;
                            }
                        case "TCV": {
                                //校准TCV
                                phy_num = (byte)(cell[arg_location, 2].IntValue);
                                textBox_Phy_Num_Cal.Text = phy_num.ToString();
                                textBox_Phy_Num_Get.Text = phy_num.ToString();
                                ch_mask = Convert.ToUInt16(cell[arg_location, 3].StringValue, 16);
                                g_Standard_Values.Max_D_Value = cell[arg_location, 4].DoubleValue;
                                if (CAL_VER_FLAG_TypeDef.CAL_FLAG == c_v_flag) {
                                    g_Standard_Values.First_Point_Value = cell[arg_location, 5].DoubleValue;
                                    g_Standard_Values.Second_Point_Value = cell[arg_location, 6].DoubleValue;
                                } else {
                                    g_Standard_Values.First_Point_Value = cell[arg_location, 7].DoubleValue;
                                    g_Standard_Values.Second_Point_Value = cell[arg_location, 8].DoubleValue;
                                }
                                if (!Cal_Ver_TCV(ch_mask, device_ch_num, g_Standard_Values, c_v_flag)) {
                                    return;
                                }
                                break;
                            }
                        case "TCR": {
                                //校准TCR
                                phy_num = (byte)(cell[arg_location, 2].IntValue);
                                textBox_Phy_Num_Cal.Text = phy_num.ToString();
                                textBox_Phy_Num_Get.Text = phy_num.ToString();
                                ch_mask = Convert.ToUInt16(cell[arg_location, 3].StringValue, 16);
                                g_CU50_RES.Max_D_Value = cell[arg_location, 4].DoubleValue;
                                if (!Cal_Ver_TCR(ch_mask, device_ch_num, g_CU50_RES, c_v_flag)) {
                                    return;
                                }
                                break;
                            }
                        case "NTC": {
                                //校准NTC
                                phy_num = (byte)(cell[arg_location, 2].IntValue);
                                textBox_Phy_Num_Cal.Text = phy_num.ToString();
                                textBox_Phy_Num_Get.Text = phy_num.ToString();
                                ch_mask = Convert.ToUInt16(cell[arg_location, 3].StringValue, 16);
                                g_NTC_RES.Max_D_Value = cell[arg_location, 4].DoubleValue;
                                if (!Cal_Ver_NTC(ch_mask, device_ch_num, g_NTC_RES, c_v_flag)) {
                                    return;
                                }
                                break;
                            }
                        case "PT100": {
                                //校准PT100
                                phy_num = (byte)(cell[arg_location, 2].IntValue);
                                textBox_Phy_Num_Cal.Text = phy_num.ToString();
                                textBox_Phy_Num_Get.Text = phy_num.ToString();
                                ch_mask = Convert.ToUInt16(cell[arg_location, 3].StringValue, 16);
                                g_PT100_RES.Max_D_Value = cell[arg_location, 4].DoubleValue;
                                if (!Cal_Ver_PT100(ch_mask, device_ch_num, g_PT100_RES, c_v_flag)) {
                                    return;
                                }
                                break;
                            }
                        case "PT1000": {
                                //校准PT1000
                                phy_num = (byte)(cell[arg_location, 2].IntValue);
                                textBox_Phy_Num_Cal.Text = phy_num.ToString();
                                textBox_Phy_Num_Get.Text = phy_num.ToString();
                                ch_mask = Convert.ToUInt16(cell[arg_location, 3].StringValue, 16);
                                g_PT1000_RES.Max_D_Value = cell[arg_location, 4].DoubleValue;
                                if (!Cal_Ver_PT1000(ch_mask, device_ch_num, g_PT1000_RES, c_v_flag)) {
                                    return;
                                }
                                break;
                            }
                        case "ADV": {
                                //校准ADV
                                phy_num = (byte)(cell[arg_location, 2].IntValue);
                                textBox_Phy_Num_Cal.Text = phy_num.ToString();
                                textBox_Phy_Num_Get.Text = phy_num.ToString();
                                ch_mask = Convert.ToUInt16(cell[arg_location, 3].StringValue, 16);
                                g_Standard_Values.Max_D_Value = cell[arg_location, 4].DoubleValue;
                                if (CAL_VER_FLAG_TypeDef.CAL_FLAG == c_v_flag) {
                                    g_Standard_Values.First_Point_Value = cell[arg_location, 5].DoubleValue;
                                    g_Standard_Values.Second_Point_Value = cell[arg_location, 6].DoubleValue;
                                } else {
                                    g_Standard_Values.First_Point_Value = cell[arg_location, 7].DoubleValue;
                                    g_Standard_Values.Second_Point_Value = cell[arg_location, 8].DoubleValue;
                                }
                                if (!Cal_Ver_ADV(ch_mask, device_ch_num, g_Standard_Values, c_v_flag)) {
                                    return;
                                }
                                break;
                            }
                        case "ADI": {
                                //校准ADI
                                phy_num = (byte)(cell[arg_location, 2].IntValue);
                                textBox_Phy_Num_Cal.Text = phy_num.ToString();
                                textBox_Phy_Num_Get.Text = phy_num.ToString();
                                ch_mask = Convert.ToUInt16(cell[arg_location, 3].StringValue, 16);
                                g_Standard_Values.Max_D_Value = cell[arg_location, 4].DoubleValue;
                                if (CAL_VER_FLAG_TypeDef.CAL_FLAG == c_v_flag) {
                                    g_Standard_Values.First_Point_Value = cell[arg_location, 5].DoubleValue;
                                    g_Standard_Values.Second_Point_Value = cell[arg_location, 6].DoubleValue;
                                } else {
                                    g_Standard_Values.First_Point_Value = cell[arg_location, 7].DoubleValue;
                                    g_Standard_Values.Second_Point_Value = cell[arg_location, 8].DoubleValue;
                                }
                                if (!Cal_Ver_ADI(ch_mask, device_ch_num, g_Standard_Values, c_v_flag)) {
                                    return;
                                }
                                break;
                            }
                        case "DAV": {
                                //校准DAV
                                phy_num = (byte)(cell[arg_location, 2].IntValue);
                                textBox_Phy_Num_Cal.Text = phy_num.ToString();
                                textBox_Phy_Num_Get.Text = phy_num.ToString();
                                ch_mask = Convert.ToUInt16(cell[arg_location, 3].StringValue, 16);
                                g_Standard_Values.Max_D_Value = cell[arg_location, 4].DoubleValue;
                                if (CAL_VER_FLAG_TypeDef.CAL_FLAG == c_v_flag) {
                                    g_Standard_Values.First_Point_Value = cell[arg_location, 5].DoubleValue;
                                    g_Standard_Values.Second_Point_Value = cell[arg_location, 6].DoubleValue;
                                } else {
                                    g_Standard_Values.First_Point_Value = cell[arg_location, 7].DoubleValue;
                                    g_Standard_Values.Second_Point_Value = cell[arg_location, 8].DoubleValue;
                                }
                                if (!Cal_Ver_DAV(ch_mask, device_ch_num, g_Standard_Values, c_v_flag)) {
                                    return;
                                }
                                break;
                            }
                        case "DAI": {
                                //校准DAI
                                phy_num = (byte)(cell[arg_location, 2].IntValue);
                                textBox_Phy_Num_Cal.Text = phy_num.ToString();
                                textBox_Phy_Num_Get.Text = phy_num.ToString();
                                ch_mask = Convert.ToUInt16(cell[arg_location, 3].StringValue, 16);
                                g_Standard_Values.Max_D_Value = cell[arg_location, 4].DoubleValue;
                                if (CAL_VER_FLAG_TypeDef.CAL_FLAG == c_v_flag) {
                                    g_Standard_Values.First_Point_Value = cell[arg_location, 5].DoubleValue;
                                    g_Standard_Values.Second_Point_Value = cell[arg_location, 6].DoubleValue;
                                } else {
                                    g_Standard_Values.First_Point_Value = cell[arg_location, 7].DoubleValue;
                                    g_Standard_Values.Second_Point_Value = cell[arg_location, 8].DoubleValue;
                                }
                                if (!Cal_Ver_DAI(ch_mask, device_ch_num, g_Standard_Values, c_v_flag)) {
                                    return;
                                }
                                break;
                            }
                        default: {
                                richTextBox_AnalogMessage.Text += "配置文件错误" + "\r\n";
                                return;
                            }
                    }
                }
                if (CAL_VER_FLAG_TypeDef.CAL_FLAG == c_v_flag) {
                    richTextBox_AnalogMessage.Text += "校准操作完成" + "\r\n";
                    MessageBox.Show("校准操作完成");
                } else {
                    richTextBox_AnalogMessage.Text += "校验通过" + "\r\n";
                    MessageBox.Show("校验通过");
                }
            } catch {
                richTextBox_AnalogMessage.Text += "配置文件错误" + "\r\n";
            }
        }

        /*****************************************各校准、校验函数***************************************************/
        //多通道，同一值填入,当然也可对单通道适用
        private void Multichannel_One_Value_Fill(ushort ch_mask, byte device_ch_num, CheckBox[] sel, TextBox[] fill_value, double standard_value) {
            int i = 0;
            for (i = 0; i < device_ch_num; i++) {
                if (0x01 == ((ch_mask >> i) & 0x01)) {
                    checkBox_Cal_Phy_CH_Sel[i].Checked = true;
                    textBox_Cal_Phy[i].Text = standard_value.ToString("F2");
                } else {
                    checkBox_Cal_Phy_CH_Sel[i].Checked = false;
                }
            }
        }

        //进行校准\校验
        private bool Handle_Cal_Ver(ushort ch_mask, byte device_ch_num, double max_d_value, CAL_VER_FLAG_TypeDef c_v_flag) {
            if (CAL_VER_FLAG_TypeDef.CAL_FLAG == c_v_flag) {
                Ch_Calibration(); //校准
            } else if (CAL_VER_FLAG_TypeDef.VER_FLAG == c_v_flag) {
                //校验
                if (!Verify(ch_mask, device_ch_num, max_d_value)) {
                    return false;
                }
            }
            richTextBox_AnalogMessage.Text += toolStripStatusLabelCOM.Text + "\r\n";
            return true;
        }

        //电阻型校准\校验
        private bool Cal_Ver_R_Type(ushort ch_mask, byte device_ch_num, RES_TYPE_TypeDef standard_R, CAL_VER_FLAG_TypeDef c_v_flag) {
            bool ret = true;
            int i = 0;
            double truth_r = 0;
            for (i = 0; i < device_ch_num; i++) {
                checkBox_Cal_Phy_CH_Sel[i].Checked = false; //清空通道选择
            }
            for (i = 0; i < device_ch_num; i++) { //选择通道，并填入发送值
                if (0x01 == ((ch_mask >> i) & 0x01)) {
                    ushort current_mask = (ushort)(ch_mask & (1 << i));
                    checkBox_Cal_Phy_CH_Sel[i].Checked = true;
                    //选中当前通道
                    Analog_Analyte_Channel_Select((byte)current_mask);
                    //第一点电阻校准\校验
                    Analog_Res_Standard_Set(standard_R.First_Point_Res); //选中第1点的基准电阻
                    truth_r = Analog_PT100_PT1000_TCR_Standard_Read();//获取实际电阻
                    Multichannel_One_Value_Fill(current_mask, device_ch_num, checkBox_Cal_Phy_CH_Sel, textBox_Cal_Phy, truth_r); //填入基准值
                    ret = Handle_Cal_Ver(current_mask, device_ch_num, standard_R.Max_D_Value, c_v_flag);
                    if (!ret) {
                        return false;
                    }
                    richTextBox_AnalogMessage.Text += toolStripStatusLabelCOM.Text + "\r\n";

                    //第二点电阻校准\校验
                    Analog_Res_Standard_Set(standard_R.Second_Point_Res); //选中第2点的基准电阻
                    truth_r = Analog_PT100_PT1000_TCR_Standard_Read();//获取实际电阻
                    Multichannel_One_Value_Fill(current_mask, device_ch_num, checkBox_Cal_Phy_CH_Sel, textBox_Cal_Phy, truth_r); //填入基准值
                    ret = Handle_Cal_Ver(current_mask, device_ch_num, standard_R.Max_D_Value, c_v_flag);
                    if (!ret) {
                        return false;
                    }
                    richTextBox_AnalogMessage.Text += toolStripStatusLabelCOM.Text + "\r\n";
                    checkBox_Cal_Phy_CH_Sel[i].Checked = false;
                }
            }
            return true;
        }

        /*****************************************各校准、校验函数***************************************************/
        //TC的中间物理量编号
        enum PHY_NUM_TypeDef {
            PHYSICS_TC_HOT_V = 1, //TC热端电压
            PHYSICS_TC_OUT_CJ_R, //TC外置冷端电阻
            PHYSICS_TC_IN_CJ_T, //TC内置冷端温度
        };

        //TC原始值编号
        enum ORI_NUM_TypeDef {
            ORIGINAL_VALUE_D1 = 1, //TC原始值D1
            ORIGINAL_VALUE_D2, //TC原始值D2
            ORIGINAL_VALUE_IN_CJ, //TC内置冷端原始值
        };


        //TC的内置冷端温度校准\校验
        private bool Cal_Ver_TC_CJC_T(ushort ch_mask, byte device_ch_num, OTHER_TYPE_TypeDef standard_T, CAL_VER_FLAG_TypeDef c_v_flag) {
            bool ret = true;
            Multichannel_One_Value_Fill(ch_mask, device_ch_num, checkBox_Cal_Phy_CH_Sel, textBox_Cal_Phy, standard_T.First_Point_Value); //填入基准值
            ret = Handle_Cal_Ver(ch_mask, device_ch_num, standard_T.Max_D_Value, c_v_flag);
            return ret;
        }

        //TCV校准\校验
        private bool Cal_Ver_TCV(ushort ch_mask, byte device_ch_num, OTHER_TYPE_TypeDef standard_V, CAL_VER_FLAG_TypeDef c_v_flag) {
            bool ret = true;
            int i = 0;
            double truth_mv = 0;
            Analog_Analyte_Type_Select(ANALYTE_TYPE_TypeDef.ANALYTE_TCV);
            Analog_Analyte_Channel_Select((byte)ch_mask); //打开全部路通道一次性校准
            Analog_DAC8760_Output_Switch_Flag_Set(":关");//打开电压电流输出
            //第一点电压校准\校验:
            Analog_TCV_Standard_Set(standard_V.First_Point_Value);
            truth_mv = Analog_TCV_Standard_Read(); //获取TCV的实际输入电压
            Multichannel_One_Value_Fill(ch_mask, device_ch_num, checkBox_Cal_Phy_CH_Sel, textBox_Cal_Phy, truth_mv); //填入基准值
            ret = Handle_Cal_Ver(ch_mask, device_ch_num, standard_V.Max_D_Value, c_v_flag);
            if (!ret) {
                return false;
            }
            //第二点电压校准\校验:
            Analog_TCV_Standard_Set(standard_V.Second_Point_Value);
            truth_mv = Analog_TCV_Standard_Read(); //获取TCV的实际输入电压
            Multichannel_One_Value_Fill(ch_mask, device_ch_num, checkBox_Cal_Phy_CH_Sel, textBox_Cal_Phy, truth_mv); //填入基准值
            ret = Handle_Cal_Ver(ch_mask, device_ch_num, standard_V.Max_D_Value, c_v_flag);
            if (!ret) {
                return false;
            }
            Analog_DAC8760_Output_Switch_Flag_Set(":开");//关闭电压电流输出
            return ret;
        }

        //TCR校准\校验
        private bool Cal_Ver_TCR(ushort ch_mask, byte device_ch_num, RES_TYPE_TypeDef standard_R, CAL_VER_FLAG_TypeDef c_v_flag) {
            Analog_Analyte_Type_Select(ANALYTE_TYPE_TypeDef.ANALYTE_TCR);
            return Cal_Ver_R_Type(ch_mask, device_ch_num, standard_R, c_v_flag);
        }

        //NTC校准\校验
        private bool Cal_Ver_NTC(ushort ch_mask, byte device_ch_num, RES_TYPE_TypeDef standard_R, CAL_VER_FLAG_TypeDef c_v_flag) {
            Analog_Analyte_Type_Select(ANALYTE_TYPE_TypeDef.ANALYTE_NTC);
            return Cal_Ver_R_Type(ch_mask, device_ch_num, standard_R, c_v_flag);
        }

        //PT100校准\校验
        private bool Cal_Ver_PT100(ushort ch_mask, byte device_ch_num, RES_TYPE_TypeDef standard_R, CAL_VER_FLAG_TypeDef c_v_flag) {
            Analog_Analyte_Type_Select(ANALYTE_TYPE_TypeDef.ANALYTE_PT);
            return Cal_Ver_R_Type(ch_mask, device_ch_num, standard_R, c_v_flag);
        }

        //PT1000校准\校验
        private bool Cal_Ver_PT1000(ushort ch_mask, byte device_ch_num, RES_TYPE_TypeDef standard_R, CAL_VER_FLAG_TypeDef c_v_flag) {
            Analog_Analyte_Type_Select(ANALYTE_TYPE_TypeDef.ANALYTE_PT);
            return Cal_Ver_R_Type(ch_mask, device_ch_num, standard_R, c_v_flag);
        }

        //ADV校准\校验
        private bool Cal_Ver_ADV(ushort ch_mask, byte device_ch_num, OTHER_TYPE_TypeDef standard_V, CAL_VER_FLAG_TypeDef c_v_flag) {
            bool ret = true;
            int i = 0;
            double truth_mv = 0;
            Analog_Analyte_Type_Select(ANALYTE_TYPE_TypeDef.ANALYTE_ADV);
            Analog_Analyte_Channel_Select((byte)ch_mask); //打开全部路通道一次性校准
            Analog_DAC8760_Output_Switch_Flag_Set(":关");//打开电压电流输出
            //第一点电压校准\校验:
            Analog_ADV_Standard_Set(standard_V.First_Point_Value);
            truth_mv = Analog_ADV_DAV_Standard_Read(); //获取ADV的实际输入电压
            Multichannel_One_Value_Fill(ch_mask, device_ch_num, checkBox_Cal_Phy_CH_Sel, textBox_Cal_Phy, truth_mv); //填入基准值  
            ret = Handle_Cal_Ver(ch_mask, device_ch_num, standard_V.Max_D_Value, c_v_flag);
            if (!ret) {
                return false;
            }
            //第二点电压校准\校验:
            Analog_ADV_Standard_Set(standard_V.Second_Point_Value);
            truth_mv = Analog_ADV_DAV_Standard_Read(); //获取ADV的实际输入电压
            Multichannel_One_Value_Fill(ch_mask, device_ch_num, checkBox_Cal_Phy_CH_Sel, textBox_Cal_Phy, truth_mv); //填入基准值
            ret = Handle_Cal_Ver(ch_mask, device_ch_num, standard_V.Max_D_Value, c_v_flag);
            if (!ret) {
                return false;
            }
            Analog_DAC8760_Output_Switch_Flag_Set(":开");//关闭电压电流输出
            return ret;
        }

        //ADI校准\校验
        private bool Cal_Ver_ADI(ushort ch_mask, byte device_ch_num, OTHER_TYPE_TypeDef standard_mA, CAL_VER_FLAG_TypeDef c_v_flag) {
            bool ret = true;
            int i = 0;
            double truth_mA = 0;
            Analog_Analyte_Type_Select(ANALYTE_TYPE_TypeDef.ANALYTE_ADI);
            Analog_DAC8760_Output_Switch_Flag_Set(":关");//打开电压电流输出
            for (i = 0; i < device_ch_num; i++) {
                checkBox_Cal_Phy_CH_Sel[i].Checked = false; //清空通道选择
            }
            for (i = 0; i < device_ch_num; i++) { //选择通道，并填入发送值
                if (0x01 == ((ch_mask >> i) & 0x01)) {
                    ushort current_mask = (ushort)(ch_mask & (1 << i));
                    checkBox_Cal_Phy_CH_Sel[i].Checked = true;
                    //选中当前通道
                    Analog_Analyte_Channel_Select((byte)current_mask);
                    //第一点电流校准\校验
                    Analog_ADI_Standard_Set(standard_mA.First_Point_Value); //设置第1点电流基准
                    truth_mA = Analog_ADI_DAI_Standard_Read();//获取实际电流
                    Multichannel_One_Value_Fill(current_mask, device_ch_num, checkBox_Cal_Phy_CH_Sel, textBox_Cal_Phy, truth_mA); //填入基准值
                    ret = Handle_Cal_Ver(current_mask, device_ch_num, standard_mA.Max_D_Value, c_v_flag);
                    if (!ret) {
                        return false;
                    }
                    richTextBox_AnalogMessage.Text += toolStripStatusLabelCOM.Text + "\r\n";

                    //第二点电流校准\校验
                    Analog_ADI_Standard_Set(standard_mA.Second_Point_Value); //设置第2点电流基准
                    truth_mA = Analog_ADI_DAI_Standard_Read();//获取实际电流
                    Multichannel_One_Value_Fill(current_mask, device_ch_num, checkBox_Cal_Phy_CH_Sel, textBox_Cal_Phy, truth_mA); //填入基准值
                    ret = Handle_Cal_Ver(current_mask, device_ch_num, standard_mA.Max_D_Value, c_v_flag);
                    if (!ret) {
                        return false;
                    }
                    richTextBox_AnalogMessage.Text += toolStripStatusLabelCOM.Text + "\r\n";
                    checkBox_Cal_Phy_CH_Sel[i].Checked = false;
                }
            }
            Analog_DAC8760_Output_Switch_Flag_Set(":开");//关闭电压电流输出
            return true;
        }

        //DAV校准\校验
        private bool Cal_Ver_DAV(ushort ch_mask, byte device_ch_num, OTHER_TYPE_TypeDef standard_V, CAL_VER_FLAG_TypeDef c_v_flag) {
            bool ret = true;
            int i = 0;
            double truth_mv = 0;
            Analog_Analyte_Type_Select(ANALYTE_TYPE_TypeDef.ANALYTE_DAV);
            Analog_Analyte_Channel_Select((byte)ch_mask); //打开全部路通道一次性校准
            Analog_DAC8760_Output_Switch_Flag_Set(":关");//打开电压电流输出
            //第一点电压校准\校验:
            Multichannel_One_Value_Fill(ch_mask, device_ch_num, checkBox_CH_Result_Sel, textBox_CH_Result, standard_V.First_Point_Value);
            Set_CH_Result();
            Thread.Sleep(2000);
            truth_mv = Analog_ADV_DAV_Standard_Read(); //获取DAV的实际输入电压
            Multichannel_One_Value_Fill(ch_mask, device_ch_num, checkBox_Cal_Phy_CH_Sel, textBox_Cal_Phy, truth_mv); //填入基准值
            ret = Handle_Cal_Ver(ch_mask, device_ch_num, standard_V.Max_D_Value, c_v_flag);
            if (!ret) {
                return false;
            }
            //第二点电压校准\校验:
            Multichannel_One_Value_Fill(ch_mask, device_ch_num, checkBox_CH_Result_Sel, textBox_CH_Result, standard_V.Second_Point_Value);
            Set_CH_Result();
            Thread.Sleep(2000);
            truth_mv = Analog_ADV_DAV_Standard_Read(); //获取DAV的实际输入电压
            Multichannel_One_Value_Fill(ch_mask, device_ch_num, checkBox_Cal_Phy_CH_Sel, textBox_Cal_Phy, truth_mv); //填入基准值
            ret = Handle_Cal_Ver(ch_mask, device_ch_num, standard_V.Max_D_Value, c_v_flag);
            if (!ret) {
                return false;
            }
            Analog_DAC8760_Output_Switch_Flag_Set(":开");//关闭电压电流输出
            return ret;
        }

        //DAI校准\校验
        private bool Cal_Ver_DAI(ushort ch_mask, byte device_ch_num, OTHER_TYPE_TypeDef standard_mA, CAL_VER_FLAG_TypeDef c_v_flag) {
            bool ret = true;
            int i = 0;
            double truth_mA = 0;
            Analog_Analyte_Type_Select(ANALYTE_TYPE_TypeDef.ANALYTE_DAI);
            Analog_DAC8760_Output_Switch_Flag_Set(":关");//打开电压电流输出
            for (i = 0; i < device_ch_num; i++) {
                checkBox_Cal_Phy_CH_Sel[i].Checked = false; //清空通道选择
            }
            for (i = 0; i < device_ch_num; i++) { //选择通道，并填入发送值
                if (0x01 == ((ch_mask >> i) & 0x01)) {
                    ushort current_mask = (ushort)(ch_mask & (1 << i));
                    checkBox_Cal_Phy_CH_Sel[i].Checked = true;
                    //选中当前通道
                    Analog_Analyte_Channel_Select((byte)current_mask);
                    //第一点电流校准\校验
                    Multichannel_One_Value_Fill(ch_mask, device_ch_num, checkBox_CH_Result_Sel, textBox_CH_Result, standard_mA.First_Point_Value);
                    Set_CH_Result();
                    Thread.Sleep(2000);
                    truth_mA = Analog_ADI_DAI_Standard_Read();//获取实际电流
                    Multichannel_One_Value_Fill(current_mask, device_ch_num, checkBox_Cal_Phy_CH_Sel, textBox_Cal_Phy, truth_mA); //填入基准值
                    ret = Handle_Cal_Ver(current_mask, device_ch_num, standard_mA.Max_D_Value, c_v_flag);
                    if (!ret) {
                        return false;
                    }
                    richTextBox_AnalogMessage.Text += toolStripStatusLabelCOM.Text + "\r\n";

                    //第二点电流校准\校验
                    Multichannel_One_Value_Fill(ch_mask, device_ch_num, checkBox_CH_Result_Sel, textBox_CH_Result, standard_mA.Second_Point_Value);
                    Set_CH_Result();
                    Thread.Sleep(2000);
                    truth_mA = Analog_ADI_DAI_Standard_Read();//获取实际电流
                    Multichannel_One_Value_Fill(current_mask, device_ch_num, checkBox_Cal_Phy_CH_Sel, textBox_Cal_Phy, truth_mA); //填入基准值
                    ret = Handle_Cal_Ver(current_mask, device_ch_num, standard_mA.Max_D_Value, c_v_flag);
                    if (!ret) {
                        return false;
                    }
                    richTextBox_AnalogMessage.Text += toolStripStatusLabelCOM.Text + "\r\n";
                    checkBox_Cal_Phy_CH_Sel[i].Checked = false;
                }
            }
            Analog_DAC8760_Output_Switch_Flag_Set(":开");//关闭电压电流输出
            return true;
        }

        /********************************************************************************************/

    }
}