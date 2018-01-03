using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;


namespace Testsoft {
    public unsafe partial class Form1 : Form {
        /************************************************************************/
        //协议相关定义
        public const byte COMMON_STATION_NUM = (0x00);
        public const ushort PROTOCOL_DATA_MAXLEN = 500;
        public const byte START_BYTE = (0xAA);
        public const byte FOREPART_LEN = (7);
        //public unsafe ushort TOTAL_LEN = (FOREPART_LEN + *(ushort*)g_Protocol_Buff.Data_LEN + 2);
        //public ushort CRC_VERIFY_DATA_LEN = (FOREPART_LEN - 2 + *(WORD*)g_Protocol_Buff.Data_LEN);
        public const byte CHANNELS_NUMBER = (8);

        /************************************************************************/
        //定义“读写指令码”和“指令码“
        public const byte READ_CMD = (0x00);
        public const byte WRITE_CMD = (0x01);
        public const byte RESPONSE_READ_CMD = (0x80);
        public const byte RESPONSE_WRITE_CMD = (0x81);
        public const byte RESPONSE_CORRECT = (0x06);
        public const byte RESPONSE_INCORRECT = (0x15);

        enum CMD_TypeDef {
            CMD_SOFTWARE_VERSION = 0x01, //²Ù×÷ Èí¼þ°æ±¾ºÅ
            CMD_DEVICE_TYPE, //²Ù×÷ Éè±¸ÐÍºÅ
            CMD_STATION_NUM, //²Ù×÷ Éè±¸Õ¾ºÅ
            CMD_CALIBRATION_STATE, //²Ù×÷ Éè±¸Ð£×¼×´Ì¬

            CMD_CH_RESULT = 0x81, //²Ù×÷ Í¨µÀ½á¹û
            CMD_CH_PHYSICAL_VARIABLE, //²Ù×÷ ÖÐ¼äÎïÀíÁ¿£¨Ð£×¼Ê±Ê¹ÓÃ£©
            CMD_CH_CALIBRATION_PARAMETER, //²Ù×÷ Ð£×¼²ÎÊý
            CMD_CH_CALIBRATION, //²Ù×÷ Í¨µÀÐ£×¼
            CMD_CH_ORIGINAL_VALUE, //²Ù×÷ Í¨µÀÔ­Ê¼Öµ
            CMD_CH_TYPE, //²Ù×÷ Í¨µÀÀàÐÍ
            CMD_CH_SWITCH, //²Ù×÷ Í¨µÀ¿ª¹Ø
            CMD_CH_TC_CJC_MODEL, //²Ù×÷ TCÍ¨µÀÀä¶ËÄ£Ê½
            CMD_CH_TC_SENSOR, //²Ù×÷ TCÍ¨µÀ´«¸ÐÆ÷ÀàÐÍ
            CMD_CH_TC_CJC_T, //
            CMD_CH_LTC_PWM, //²Ù×÷ LTCÍ¨µÀµÄPWM 
            CMD_CH_WT_CPU_G, //²Ù×÷ WTµÄCPUÔöÒæ£¨¾ÉWTÊ¹ÓÃ£©
            CMD_CH_WT_LOOP_FILTER, //²Ù×÷ WTÑ­»·ÂË²¨
            CMD_CH_WT_POLARITY, //²Ù×÷ WTµ¥Ë«¼«ÐÔ²ÉÑù·½Ê½
            CMD_CH_WT_FILTER_MODE, //²Ù×÷ WTÂË²¨·½Ê½¡¢Ç¿¶È
            CMD_CH_WT_FILTER_RESET, //²Ù×÷ WTÂË²¨¸´Î»
            CMD_CH_WT_ADC_G, //²Ù×÷ WTµÄADÐ¾Æ¬ÔöÒæ
        };

        /************************************************************************/
        //通讯帧结构
        public enum LOCATE_TypeDef {
            LOCATE_HEAD_AA, //
            LOCATE_HEAD_55, //
            LOCATE_STATION_NUM, //
            LOCATE_FUNCTION_CMD, //
            LOCATE_RW_CMD, //
            LOCATE_DATA_LEN_LO, //
            LOCATE_DATA_LEN_HI, //
        };

        public struct PROTOCOL_FRAME_TypeDef {
            public fixed byte Head[2]; //帧头
            public byte Station_NUM; //设备站号
            public byte Function_CMD; //功能指令码
            public byte RW_CMD; //读写指令码
            public fixed byte Data_LEN[2]; //数据长度
            public fixed byte Data[500]; //数据内容
            public fixed byte Crc[2]; //CRC校验
        };

        //特别注意：c#中要使用指针若是指向托管代码，如全局变量的话，使用部分的代码要用fixed(){}语句包起来并重新赋值指针，要不然变量的地址随时可能会变
        static PROTOCOL_FRAME_TypeDef* g_Protocol_Buff_p;
        static PROTOCOL_FRAME_TypeDef g_Protocol_Buff = new PROTOCOL_FRAME_TypeDef();// {
                                                                              //    Head = new byte[2],
                                                                              //    Data_LEN = new byte[2],
                                                                              //    Data = new byte[500],
                                                                              //    Crc = new byte[2],
                                                                              //};

        static byte g_character_count = 0; //转义字接收计数
        static ushort g_Rxcount = 0;
        static byte g_Device_Station_NUM = 0x00;

        public ushort GetCRC(byte* crc_array, ushort len) {
            uint i, j;
            ushort modbus_crc;
            modbus_crc = 0xffff;
            for (i = 0; i < len; i++) {
                modbus_crc = (ushort)((modbus_crc & 0xFF00) | ((modbus_crc & 0x00FF) ^ crc_array[i]));
                for (j = 1; j <= 8; j++) {
                    if ((modbus_crc & 0x01) == 1) {
                        modbus_crc = (ushort)(modbus_crc >> 1);
                        modbus_crc ^= 0XA001;
                    } else {
                        modbus_crc = (ushort)(modbus_crc >> 1);
                    }
                }
            }
            return modbus_crc;
        }

        public bool RecvDataAnalysis(byte RecvByte) {
            switch ((LOCATE_TypeDef)g_Rxcount) {//ÅÐ¶Ï¸ÃÌõÊý¾ÝµÄ½ÓÊÕ¸öÊý
                case LOCATE_TypeDef.LOCATE_HEAD_AA: {
                        //½âÎöÖ¡Í·0xAA£¬ÔÚÊý¾ÝÄÚÈÝÖÐ0xAAÓÃ×ªÒå×Ö0xAA AA
                        if (START_BYTE != RecvByte) {
                            g_Rxcount = 0;
                        } else {
                            g_Protocol_Buff_p->Head[0] = RecvByte;
                            g_Rxcount++;
                        }
                        break;
                    }
                case LOCATE_TypeDef.LOCATE_HEAD_55: {
                        //½âÎöÖ¡Í·0x55
                        if (0x55 != RecvByte) {
                            g_Rxcount = 0;
                        } else {
                            g_Protocol_Buff_p->Head[1] = RecvByte;
                            g_Rxcount++;
                        }
                        break;
                    }
                case LOCATE_TypeDef.LOCATE_STATION_NUM: {
                        //½âÎöÉè±¸Õ¾ºÅ
                        if (g_Device_Station_NUM != RecvByte && COMMON_STATION_NUM != RecvByte) {
                            g_Rxcount = 0;
                        } else {
                            g_Protocol_Buff_p->Station_NUM = RecvByte;
                            g_Rxcount++;
                        }
                        break;
                    }
                case LOCATE_TypeDef.LOCATE_FUNCTION_CMD: {
                        //½âÎöÖ¸ÁîÂë,ÕâÀïÖ»½ÓÊÕ£¬µ½Ó¦ÓÃ²ãÔÙÅÐ¶Ï¶Ô´í
                        g_Protocol_Buff_p->Function_CMD = RecvByte;
                        g_Rxcount++;
                        break;
                    }
                case LOCATE_TypeDef.LOCATE_RW_CMD: {
                        //½âÎö¶ÁÐ´Ö¸Áî0x00¡¢0x01
                        if (RESPONSE_READ_CMD != RecvByte && RESPONSE_WRITE_CMD != RecvByte) {
                            g_Rxcount = 0;
                        } else {
                            g_Protocol_Buff_p->RW_CMD = RecvByte;
                            g_Rxcount++;
                        }
                        break;
                    }
                case LOCATE_TypeDef.LOCATE_DATA_LEN_LO: {
                        //½âÎöÊý¾Ý³¤¶ÈµÄµÍ×Ö½Ú£¬ÕâÀïÖ»½ÓÊÕ£¬µ½Ó¦ÓÃ²ã¸ù¾Ý¾ßÌåÓ¦ÓÃÔÙÅÐ¶Ï¶Ô´í
                        g_Protocol_Buff_p->Data_LEN[0] = RecvByte;//(BYTE)(RecvByte & 0xFF);
                        g_Rxcount++;
                        break;
                    }
                case LOCATE_TypeDef.LOCATE_DATA_LEN_HI: {
                        //½âÎöÊý¾Ý³¤¶ÈµÄ¸ß×Ö½Ú£¬ÕâÀïÖ»½ÓÊÕ£¬µ½Ó¦ÓÃ²ã¸ù¾Ý¾ßÌåÓ¦ÓÃÔÙÅÐ¶Ï¶Ô´í
                        g_Protocol_Buff_p->Data_LEN[1] = RecvByte;// |= (BYTE)(RecvByte << 8);
                        g_Rxcount++;
                        break;
                    }
                default: {
                        //½ÓÊÕÊý¾ÝÄÚÈÝÓëCRC£¬ÕâÀïÖ»½ÓÊÕ£¬µ½Ó¦ÓÃ²ãÔÙÅÐ¶Ï¶Ô´í£¬ÔÚÊý¾ÝÄÚÈÝÕâÀïÒª´¦Àí×ªÒå×Ö0xAA AA
                        if (1 <= g_character_count) { //Èç¹û´ËÊ±ÉÏÒ»¸ö½ÓÊÕµ½µÄÊÇ0xAA£¬Ò²¾ÍÊÇ¼ÆÊýÆ÷ÒÑ¾­Îª1ÁË£¬ÕâÀïÒ²¾ÍÊÇµÚ¶þ¸ö0xAA
                            g_character_count = 0;
                            if (START_BYTE != RecvByte) { //Èç¹û½ô½Ó×ÅµÚÒ»¸ö0xAAºóÃæµÄ²»ÊÇÒ»¸ö0xAA£¬ÕâÌõÊý¾Ý×÷·Ï
                                g_Rxcount = 0;
                            }
                            return false;
                        }
                        if (FOREPART_LEN <= g_Rxcount && (FOREPART_LEN + *(ushort*)g_Protocol_Buff_p->Data_LEN) > g_Rxcount) {
                            //ÔÚÕâÀï½ÓÊÕÊý¾ÝÄÚÈÝ
                            if (START_BYTE == RecvByte) { //µ±½ÓÊÕµ½0xAAÊ±£¬¼ÆÊýÆ÷¼Ó1£¬¼ÆÊýÆ÷»áÔÚ½ÓÊÕÁËÒ»¸ö0xAAºóµÄÏÂÒ»¸ö×Ö½Ú½ÓÊÕÊ±½øÐÐÅÐ¶ÏºÍÇåÁã
                                g_character_count++;
                            }
                            g_Protocol_Buff_p->Data[g_Rxcount - FOREPART_LEN] = RecvByte;
                        } else if ((FOREPART_LEN + *(ushort*)g_Protocol_Buff_p->Data_LEN) <= g_Rxcount) {
                            //ÔÚÕâÀï½ÓÊÕCRC£¬ÕâÀïÖ»½ÓÊÕ
                            g_Protocol_Buff_p->Crc[g_Rxcount - FOREPART_LEN - *(ushort*)g_Protocol_Buff_p->Data_LEN] = RecvByte;
                        }
                        g_Rxcount++;
                        ushort TOTAL_LEN = (ushort)(FOREPART_LEN + *(ushort*)g_Protocol_Buff_p->Data_LEN + 2);
                        if (TOTAL_LEN <= g_Rxcount) { //´ËÊ±CRCÒ²½ÓÊÕÍê³É,½øÐÐCRCÐ£Ñé
                            g_Rxcount = 0;
                            ushort CRC_VERIFY_DATA_LEN = (ushort)(FOREPART_LEN - 2 + *(ushort*)g_Protocol_Buff_p->Data_LEN);
                            ushort crc = GetCRC(&g_Protocol_Buff_p->Station_NUM, CRC_VERIFY_DATA_LEN); //ÕâÀï×¢ÒâÎÒÃÇµÄCRC¼ÆËã´úÂëÓÐÎÊÌâ£¬¸ßµÍ×Ö½ÚÊÇ·´µÄ
                            if (*(ushort*)g_Protocol_Buff_p->Crc == crc) {
                                return true; //¸ÃÌõÊý¾ÝÒÑÍêÕû½ÓÊÕ
                            }
                        }
                        break;
                    }
            }
            return false;
        }

        public int FrameSendBuffer(ref byte[] bySendBuffer) {
            ushort i;
            int len_count = 0;
            byte* p = g_Protocol_Buff_p->Head;
            for (i = 0; i < FOREPART_LEN; i++) { //Ìî³ä£ºÊý¾ÝÄÚÈÝ
                bySendBuffer[len_count++] = *(byte*)p;
                p++;
            }
            for (i = 0; i < *(ushort*)g_Protocol_Buff_p->Data_LEN; i++) { //Ìî³ä£ºÊý¾ÝÄÚÈÝ
                bySendBuffer[len_count++] = g_Protocol_Buff_p->Data[i];
                if (START_BYTE == g_Protocol_Buff_p->Data[i]) { //µ±·¢ËÍ0xAAÊ±Òª·¢ËÍ0xAA AA
                    bySendBuffer[len_count++] = START_BYTE;
                }
            }
            bySendBuffer[len_count++] = g_Protocol_Buff_p->Crc[0]; //Ìî³äCRC
            bySendBuffer[len_count++] = g_Protocol_Buff_p->Crc[1];
            return len_count;
        }

        byte[] RXBUFbytearray = new byte[512];
        int Recbytenumber = 0;
        public void ReceiveFrameAnalysis(byte RecvByte) {
            RXBUFbytearray[Recbytenumber++] = RecvByte;
            if (RecvDataAnalysis(RecvByte)) {
                SortProcessingCmd();
                DisplayRXData(Recbytenumber, RXBUFbytearray);
                Recbytenumber = 0;
            }
        }

        public void SortProcessingCmd() {
            switch ((CMD_TypeDef)g_Protocol_Buff_p->Function_CMD) {
                case CMD_TypeDef.CMD_SOFTWARE_VERSION: {//Èí¼þ°æ±¾ºÅ
                        Handle_Version();
                        break;
                    }
                case CMD_TypeDef.CMD_DEVICE_TYPE: {//Éè±¸ÐÍºÅ
                        Handle_Device_Type();
                        break;
                    }
                case CMD_TypeDef.CMD_STATION_NUM: {//Éè±¸Õ¾ºÅ
                        Handle_Station_Num();
                        break;
                    }
                case CMD_TypeDef.CMD_CALIBRATION_STATE: {//Éè±¸Õ¾ºÅ
                        Handle_Calibration_State();
                        break;
                    }
                case CMD_TypeDef.CMD_CH_RESULT: {//Í¨µÀ½á¹û
                        Handle_Ch_Result();
                        break;
                    }
                case CMD_TypeDef.CMD_CH_PHYSICAL_VARIABLE: {//ÖÐ¼äÎïÀíÁ¿
                        Handle_Ch_Physical_Variable();
                        break;
                    }
                case CMD_TypeDef.CMD_CH_CALIBRATION_PARAMETER: {//Ð£×¼²ÎÊý
                        Handle_Ch_Calibration_Parameter();
                        break;
                    }
                case CMD_TypeDef.CMD_CH_CALIBRATION: {//Í¨µÀÁ½µãÐ£×¼
                        Handle_Ch_Calibration();
                        break;
                    }
                case CMD_TypeDef.CMD_CH_ORIGINAL_VALUE: {//Í¨µÀÔ­Ê¼Öµ
                        Handle_Ch_Original_Value();
                        break;
                    }
                case CMD_TypeDef.CMD_CH_TYPE: {//Í¨µÀÀàÐÍ
                        Handle_Ch_Type();
                        break;
                    }
                case CMD_TypeDef.CMD_CH_SWITCH: {//Í¨µÀ¿ª¹Ø
                        Handle_Ch_Switch();
                        break;
                    }
                case CMD_TypeDef.CMD_CH_TC_CJC_MODEL: {//TCÍ¨µÀÀä¶ËÄ£Ê½
                        Handle_Ch_TC_CJC_Mode();
                        break;
                    }
                case CMD_TypeDef.CMD_CH_TC_SENSOR: {//TCÍ¨µÀÀä¶ËÄ£Ê½
                        Handle_Ch_TC_Sensor();
                        break;
                    }
                case CMD_TypeDef.CMD_CH_TC_CJC_T: {//TCÍ¨µÀÀä¶ËÄ£Ê½
                        Handle_Ch_TC_CJC_T();
                        break;
                    }
                default: {// ´íÎó
                        toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        break;
                    }
            }
        }

        public void Handle_Version() {
            string showword = "软件版本号";
            switch (g_Protocol_Buff_p->RW_CMD) {
                case RESPONSE_READ_CMD: {//¶ÁÖ¸Áî
                        if (2 == *(ushort*)g_Protocol_Buff_p->Data_LEN) {
                            ushort ver = *(ushort*)g_Protocol_Buff_p->Data;
                            try {
                                textBox_Software_Ver.Text = ver.ToString();//"X4"
                                toolStripStatusLabelCOM.Text = "读取" + showword + "成功";
                            } catch {
                                toolStripStatusLabelCOM.Text = "接收数值解析错误";
                            }
                        }
                        break;
                    }
                case RESPONSE_WRITE_CMD: {//Ð´Ö¸Áî
                        if (RESPONSE_CORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "成功";
                        } else if (RESPONSE_INCORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "失败";
                        } else {
                            toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        }
                        break;
                    }
                default: {// ´íÎó
                        toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        break;
                    }
            }
        }

        public void Handle_Device_Type() {
            string showword = "设备型号";
            switch (g_Protocol_Buff_p->RW_CMD) {
                case RESPONSE_READ_CMD: {//¶ÁÖ¸Áî
                        if (2 == *(ushort*)g_Protocol_Buff_p->Data_LEN) {
                            ushort device_type = *(ushort*)g_Protocol_Buff_p->Data;
                            try {
                                textBox_Device_Type.Text = device_type.ToString("X4");
                                toolStripStatusLabelCOM.Text = "读取" + showword + "成功";
                            } catch {
                                toolStripStatusLabelCOM.Text = "接收数值解析错误";
                            }
                        }
                        break;
                    }
                case RESPONSE_WRITE_CMD: {//Ð´Ö¸Áî
                        if (RESPONSE_CORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "成功";
                        } else if (RESPONSE_INCORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "失败";
                        } else {
                            toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        }
                        break;
                    }
                default: {// ´íÎó
                        toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        break;
                    }
            }
        }

        public void Handle_Station_Num() {
            string showword = "设备站号";
            switch (g_Protocol_Buff_p->RW_CMD) {
                case RESPONSE_READ_CMD: {//¶ÁÖ¸Áî
                        if (1 == *(ushort*)g_Protocol_Buff_p->Data_LEN) {
                            byte station_num = *(byte*)g_Protocol_Buff_p->Data;
                            try {
                                textBox_Station_Num.Text = station_num.ToString();
                                toolStripStatusLabelCOM.Text = "读取" + showword + "成功";
                            } catch {
                                toolStripStatusLabelCOM.Text = "接收数值解析错误";
                            }
                        }
                        break;
                    }
                case RESPONSE_WRITE_CMD: {//Ð´Ö¸Áî
                        if (RESPONSE_CORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "成功";
                        } else if (RESPONSE_INCORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "失败";
                        } else {
                            toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        }
                        break;
                    }
                default: {// ´íÎó
                        toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        break;
                    }
            }
        }

        public void Handle_Calibration_State() {
            string showword = "设备校准状态";
            switch (g_Protocol_Buff_p->RW_CMD) {
                case RESPONSE_READ_CMD: {//¶ÁÖ¸Áî
                        if (1 == *(ushort*)g_Protocol_Buff_p->Data_LEN) {
                            byte calibration_state = *(byte*)g_Protocol_Buff_p->Data;
                            try {
                                textBox_Calibration_State.Text = calibration_state.ToString();
                                toolStripStatusLabelCOM.Text = "读取" + showword + "成功";
                            } catch {
                                toolStripStatusLabelCOM.Text = "接收数值解析错误";
                            }
                        }
                        break;
                    }
                case RESPONSE_WRITE_CMD: {//Ð´Ö¸Áî
                        if (RESPONSE_CORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "成功";
                        } else if (RESPONSE_INCORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "失败";
                        } else {
                            toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        }
                        break;
                    }
                default: {// ´íÎó
                        toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        break;
                    }
            }
        }

        public void Handle_Ch_Result() {
            string showword = "通道结果";
            byte i = 0;
            byte* p = g_Protocol_Buff_p->Data; //»ñÈ¡Í¨µÀÊý¾Ý¶ÎµØÖ·
            byte temp_len = 0;
            ushort ch_mask = 0;
            ch_mask = *(ushort*)p; //»ñÈ¡Í¨µÀÑÚÂë
            p += 2;
            switch (g_Protocol_Buff_p->RW_CMD) {
                case RESPONSE_READ_CMD: {//¶ÁÖ¸Áî
                        try {
                            for (i = 0; i < CHANNELS_NUMBER; i++) {
                                if (0x01 == ((ch_mask >> i) & 0x01)) {
                                    //¸ù¾ÝÍ¨µÀÑÚÂë£¬°Ñ¸÷¸öÍ¨µÀµÄÊý¾Ý°´ÕÕÍ¨µÀºÅÓÉ´óµ½Ð¡£¬¸´ÖÆµ½Í¨Ñ¶Êý¾ÝÄÚÈÝ
                                    temp_len = 2;
                                    double getvalue = (double)(*(short*)p) / 100;
                                    textBox_CH_Result[i].Text = getvalue.ToString();
                                    p += temp_len;
                                }
                            }
                            toolStripStatusLabelCOM.Text = "读取" + showword + "成功";
                        } catch {
                            toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        }
                        break;
                    }
                case RESPONSE_WRITE_CMD: {//Ð´Ö¸Áî
                        if (RESPONSE_CORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "成功";
                        } else if (RESPONSE_INCORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "失败";
                        } else {
                            toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        }
                        break;
                    }
                default: {// ´íÎó
                        toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        break;
                    }
            }
        }

        public void Handle_Ch_Physical_Variable() {
            string showword = "通道物理量";
            byte i = 0;
            byte* p = g_Protocol_Buff_p->Data; //»ñÈ¡Í¨µÀÊý¾Ý¶ÎµØÖ·
            byte temp_len = 0;
            ushort ch_mask = 0;
            ch_mask = *(ushort*)p; //»ñÈ¡Í¨µÀÑÚÂë
            p += 2;
            byte phy_num = *(byte*)p;
            p += 1;
            switch (g_Protocol_Buff_p->RW_CMD) {
                case RESPONSE_READ_CMD: {//¶ÁÖ¸Áî
                        try {
                            for (i = 0; i < CHANNELS_NUMBER; i++) {
                                if (0x01 == ((ch_mask >> i) & 0x01)) {
                                    //¸ù¾ÝÍ¨µÀÑÚÂë£¬°Ñ¸÷¸öÍ¨µÀµÄÊý¾Ý°´ÕÕÍ¨µÀºÅÓÉ´óµ½Ð¡£¬¸´ÖÆµ½Í¨Ñ¶Êý¾ÝÄÚÈÝ
                                    temp_len = 2;
                                    double getvalue = (double)(*(short*)p) / 100;
                                    textBox_Get_Phy[i].Text = getvalue.ToString();
                                    p += temp_len;
                                }
                            }
                            toolStripStatusLabelCOM.Text = "读取" + showword + "成功";
                        } catch {
                            toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        }
                        break;
                    }
                case RESPONSE_WRITE_CMD: {//Ð´Ö¸Áî
                        if (RESPONSE_CORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "成功";
                        } else if (RESPONSE_INCORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "失败";
                        } else {
                            toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        }
                        break;
                    }
                default: {// ´íÎó
                        toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        break;
                    }
            }
        }

        public void Handle_Ch_Calibration_Parameter() {
            string showword = "通道校准参数";
            byte i = 0;
            byte* p = g_Protocol_Buff_p->Data; //»ñÈ¡Í¨µÀÊý¾Ý¶ÎµØÖ·
            byte temp_len = 0;
            ushort ch_mask = 0;
            ch_mask = *(ushort*)p; //»ñÈ¡Í¨µÀÑÚÂë
            p += 2;
            byte phy_num = *(byte*)p;
            p += 1;
            switch (g_Protocol_Buff_p->RW_CMD) {
                case RESPONSE_READ_CMD: {//¶ÁÖ¸Áî
                        try {
                            int ch_par_num = 0;
                            int ch_count = 0;
                            for (i = 0; i < CHANNELS_NUMBER; i++) {
                                if (0x01 == ((ch_mask >> i) & 0x01)) {
                                    ch_count++;
                                }
                            }
                            ch_par_num = (*(ushort*)g_Protocol_Buff_p->Data_LEN - 3) / ch_count / 4;
                            int j = 0;
                            for (i = 0; i < CHANNELS_NUMBER; i++) {
                                if (0x01 == ((ch_mask >> i) & 0x01)) {
                                    //¸ù¾ÝÍ¨µÀÑÚÂë£¬°Ñ¸÷¸öÍ¨µÀµÄÊý¾Ý°´ÕÕÍ¨µÀºÅÓÉ´óµ½Ð¡£¬¸´ÖÆµ½Í¨Ñ¶Êý¾ÝÄÚÈÝ
                                    temp_len = 4;
                                    if (1 == ch_par_num) {
                                        textBox_CH_Cal_Par[1, i].Text = (*(float*)p).ToString("F");
                                        p += temp_len;
                                    } else if (2 == ch_par_num) {
                                        textBox_CH_Cal_Par[0, i].Text = (*(float*)p).ToString("F");
                                        p += temp_len;
                                        textBox_CH_Cal_Par[1, i].Text = (*(float*)p).ToString("F");
                                        p += temp_len;
                                    }
                                }
                            }
                            toolStripStatusLabelCOM.Text = "读取" + showword + "成功";
                        } catch {
                            toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        }
                        break;
                    }
                case RESPONSE_WRITE_CMD: {//Ð´Ö¸Áî
                        if (RESPONSE_CORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "成功";
                        } else if (RESPONSE_INCORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "失败";
                        } else {
                            toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        }
                        break;
                    }
                default: {// ´íÎó
                        toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        break;
                    }
            }
            //                byte num = 0;
            //                unrespond_cmd = 0;
            //                getvalue = BitConverter.ToSingle(RXBUFbytearray, (num++ * 4 + 2));
            //                textBox81.Text = getvalue.ToString();
        }

        public void Handle_Ch_Calibration() {
            string showword = "通道校准";
            switch (g_Protocol_Buff_p->RW_CMD) {
                case RESPONSE_READ_CMD: {//¶ÁÖ¸Áî

                        break;
                    }
                case RESPONSE_WRITE_CMD: {//Ð´Ö¸Áî
                        if (RESPONSE_CORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "成功";
                        } else if (RESPONSE_INCORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "失败";
                        } else {
                            toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        }
                        break;
                    }
                default: {// ´íÎó
                        toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        break;
                    }
            }
        }

        public void Handle_Ch_Original_Value() {
            string showword = "通道原始值";
            byte i = 0;
            byte* p = g_Protocol_Buff_p->Data; //»ñÈ¡Í¨µÀÊý¾Ý¶ÎµØÖ·
            byte temp_len = 0;
            ushort ch_mask = 0;
            ch_mask = *(ushort*)p; //»ñÈ¡Í¨µÀÑÚÂë
            p += 2;
            byte ori_num = *(byte*)p;
            p += 1;
            switch (g_Protocol_Buff_p->RW_CMD) {
                case RESPONSE_READ_CMD: {//¶ÁÖ¸Áî
                        try {
                            for (i = 0; i < CHANNELS_NUMBER; i++) {
                                if (0x01 == ((ch_mask >> i) & 0x01)) {
                                    //¸ù¾ÝÍ¨µÀÑÚÂë£¬°Ñ¸÷¸öÍ¨µÀµÄÊý¾Ý°´ÕÕÍ¨µÀºÅÓÉ´óµ½Ð¡£¬¸´ÖÆµ½Í¨Ñ¶Êý¾ÝÄÚÈÝ
                                    temp_len = 2;
                                    short getvalue = (*(short*)p);
                                    textBox_CH_Original_Value[i].Text = getvalue.ToString("X4");
                                    p += temp_len;
                                }
                            }
                            toolStripStatusLabelCOM.Text = "读取" + showword + "成功";
                        } catch {
                            toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        }
                        break;
                    }
                case RESPONSE_WRITE_CMD: {//Ð´Ö¸Áî
                        if (RESPONSE_CORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "成功";
                        } else if (RESPONSE_INCORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "失败";
                        } else {
                            toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        }
                        break;
                    }
                default: {// ´íÎó
                        toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        break;
                    }
            }
        }

        public void Handle_Ch_Type() {
            string showword = "通道类型";
            byte i = 0;
            byte* p = g_Protocol_Buff_p->Data; //»ñÈ¡Í¨µÀÊý¾Ý¶ÎµØÖ·
            byte temp_len = 0;
            ushort ch_mask = 0;
            ch_mask = *(ushort*)p; //»ñÈ¡Í¨µÀÑÚÂë
            p += 2;
            switch (g_Protocol_Buff_p->RW_CMD) {
                case RESPONSE_READ_CMD: {//¶ÁÖ¸Áî
                        try {
                            for (i = 0; i < CHANNELS_NUMBER; i++) {
                                if (0x01 == ((ch_mask >> i) & 0x01)) {
                                    //¸ù¾ÝÍ¨µÀÑÚÂë£¬°Ñ¸÷¸öÍ¨µÀµÄÊý¾Ý°´ÕÕÍ¨µÀºÅÓÉ´óµ½Ð¡£¬¸´ÖÆµ½Í¨Ñ¶Êý¾ÝÄÚÈÝ
                                    temp_len = 1;
                                    double getvalue = (double)(*(byte*)p);
                                    textBox_CH_Type[i].Text = getvalue.ToString();
                                    p += temp_len;
                                }
                            }
                            toolStripStatusLabelCOM.Text = "读取" + showword + "成功";
                        } catch {
                            toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        }
                        break;
                    }
                case RESPONSE_WRITE_CMD: {//Ð´Ö¸Áî
                        if (RESPONSE_CORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "成功";
                        } else if (RESPONSE_INCORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "失败";
                        } else {
                            toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        }
                        break;
                    }
                default: {// ´íÎó
                        toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        break;
                    }
            }
        }

        public void Handle_Ch_Switch() {
            string showword = "通道开关";
            byte i = 0;
            byte* p = g_Protocol_Buff_p->Data; //»ñÈ¡Í¨µÀÊý¾Ý¶ÎµØÖ·
            byte temp_len = 0;
            ushort ch_mask = 0;
            ch_mask = *(ushort*)p; //»ñÈ¡Í¨µÀÑÚÂë
            p += 2;
            switch (g_Protocol_Buff_p->RW_CMD) {
                case RESPONSE_READ_CMD: {//¶ÁÖ¸Áî
                        try {
                            for (i = 0; i < CHANNELS_NUMBER; i++) {
                                if (0x01 == ((ch_mask >> i) & 0x01)) {
                                    //¸ù¾ÝÍ¨µÀÑÚÂë£¬°Ñ¸÷¸öÍ¨µÀµÄÊý¾Ý°´ÕÕÍ¨µÀºÅÓÉ´óµ½Ð¡£¬¸´ÖÆµ½Í¨Ñ¶Êý¾ÝÄÚÈÝ
                                    temp_len = 1;
                                    double getvalue = (double)(*(byte*)p);
                                    textBox_CH_Switch[i].Text = getvalue.ToString();
                                    p += temp_len;
                                }
                            }
                            toolStripStatusLabelCOM.Text = "读取" + showword + "成功";
                        } catch {
                            toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        }
                        break;
                    }
                case RESPONSE_WRITE_CMD: {//Ð´Ö¸Áî
                        if (RESPONSE_CORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "成功";
                        } else if (RESPONSE_INCORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "失败";
                        } else {
                            toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        }
                        break;
                    }
                default: {// ´íÎó
                        toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        break;
                    }
            }
        }

        public void Handle_Ch_TC_CJC_Mode() {
            string showword = "通道TC冷端模式";
            byte i = 0;
            byte* p = g_Protocol_Buff_p->Data; //»ñÈ¡Í¨µÀÊý¾Ý¶ÎµØÖ·
            byte temp_len = 0;
            ushort ch_mask = 0;
            ch_mask = *(ushort*)p; //»ñÈ¡Í¨µÀÑÚÂë
            p += 2;
            switch (g_Protocol_Buff_p->RW_CMD) {
                case RESPONSE_READ_CMD: {//¶ÁÖ¸Áî
                        try {
                            for (i = 0; i < CHANNELS_NUMBER; i++) {
                                if (0x01 == ((ch_mask >> i) & 0x01)) {
                                    //¸ù¾ÝÍ¨µÀÑÚÂë£¬°Ñ¸÷¸öÍ¨µÀµÄÊý¾Ý°´ÕÕÍ¨µÀºÅÓÉ´óµ½Ð¡£¬¸´ÖÆµ½Í¨Ñ¶Êý¾ÝÄÚÈÝ
                                    temp_len = 1;
                                    switch (*(byte*)p) {
                                        case 0x00: {
                                                comboBox_CH_TC_CJC_Mode[i].Text = "内置冷端";
                                                break;
                                            }
                                        case 0x01: {
                                                comboBox_CH_TC_CJC_Mode[i].Text = "外置冷端";
                                                break;
                                            }
                                        case 0x02: {
                                                comboBox_CH_TC_CJC_Mode[i].Text = "冰点冷端";
                                                break;
                                            }
                                        default: {
                                                MessageBoxERRwarning(5);
                                                return;
                                            }
                                    }
                                    p += temp_len;
                                }
                            }
                            toolStripStatusLabelCOM.Text = "读取" + showword + "成功";
                        } catch {
                            toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        }
                        break;
                    }
                case RESPONSE_WRITE_CMD: {//Ð´Ö¸Áî
                        if (RESPONSE_CORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "成功";
                        } else if (RESPONSE_INCORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "失败";
                        } else {
                            toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        }
                        break;
                    }
                default: {// ´íÎó
                        toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        break;
                    }
            }
        }

        public void Handle_Ch_TC_Sensor() {
            string showword = "通道TC传感器型号";
            byte i = 0;
            byte* p = g_Protocol_Buff_p->Data; //»ñÈ¡Í¨µÀÊý¾Ý¶ÎµØÖ·
            byte temp_len = 0;
            ushort ch_mask = 0;
            ch_mask = *(ushort*)p; //»ñÈ¡Í¨µÀÑÚÂë
            p += 2;
            switch (g_Protocol_Buff_p->RW_CMD) {
                case RESPONSE_READ_CMD: {//¶ÁÖ¸Áî
                        try {
                            for (i = 0; i < CHANNELS_NUMBER; i++) {
                                if (0x01 == ((ch_mask >> i) & 0x01)) {
                                    //¸ù¾ÝÍ¨µÀÑÚÂë£¬°Ñ¸÷¸öÍ¨µÀµÄÊý¾Ý°´ÕÕÍ¨µÀºÅÓÉ´óµ½Ð¡£¬¸´ÖÆµ½Í¨Ñ¶Êý¾ÝÄÚÈÝ
                                    temp_len = 1;
                                    switch (*(byte*)p) {
                                        case 0x00: {
                                                comboBox_CH_TC_Sensor[i].Text = "K";
                                                break;
                                            }
                                        case 0x01: {
                                                comboBox_CH_TC_Sensor[i].Text = "J";
                                                break;
                                            }
                                        case 0x02: {
                                                comboBox_CH_TC_Sensor[i].Text = "T";
                                                break;
                                            }
                                        case 0x03: {
                                                comboBox_CH_TC_Sensor[i].Text = "E";
                                                break;
                                            }
                                        case 0x04: {
                                                comboBox_CH_TC_Sensor[i].Text = "N";
                                                break;
                                            }
                                        case 0x05: {
                                                comboBox_CH_TC_Sensor[i].Text = "B";
                                                break;
                                            }
                                        case 0x06: {
                                                comboBox_CH_TC_Sensor[i].Text = "R";
                                                break;
                                            }
                                        case 0x07: {
                                                comboBox_CH_TC_Sensor[i].Text = "S";
                                                break;
                                            }
                                        default: {
                                                MessageBoxERRwarning(5);
                                                return;
                                            }
                                    }
                                    p += temp_len;
                                }
                            }
                            toolStripStatusLabelCOM.Text = "读取" + showword + "成功";
                        } catch {
                            toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        }
                        break;
                    }
                case RESPONSE_WRITE_CMD: {//Ð´Ö¸Áî
                        if (RESPONSE_CORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "成功";
                        } else if (RESPONSE_INCORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "失败";
                        } else {
                            toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        }
                        break;
                    }
                default: {// ´íÎó
                        toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        break;
                    }
            }
        }

        public void Handle_Ch_TC_CJC_T() {
            string showword = "通道TC冷端温度";
            byte i = 0;
            byte* p = g_Protocol_Buff_p->Data; //»ñÈ¡Í¨µÀÊý¾Ý¶ÎµØÖ·
            byte temp_len = 0;
            ushort ch_mask = 0;
            ch_mask = *(ushort*)p; //»ñÈ¡Í¨µÀÑÚÂë
            p += 2;
            switch (g_Protocol_Buff_p->RW_CMD) {
                case RESPONSE_READ_CMD: {//¶ÁÖ¸Áî
                        try {
                            for (i = 0; i < CHANNELS_NUMBER; i++) {
                                if (0x01 == ((ch_mask >> i) & 0x01)) {
                                    //¸ù¾ÝÍ¨µÀÑÚÂë£¬°Ñ¸÷¸öÍ¨µÀµÄÊý¾Ý°´ÕÕÍ¨µÀºÅÓÉ´óµ½Ð¡£¬¸´ÖÆµ½Í¨Ñ¶Êý¾ÝÄÚÈÝ
                                    temp_len = 2;
                                    double getvalue = (double)(*(short*)p) / 100;
                                    textBox_CH_TC_CJC_T[i].Text = getvalue.ToString();
                                    p += temp_len;
                                }
                            }
                            toolStripStatusLabelCOM.Text = "读取" + showword + "成功";
                        } catch {
                            toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        }
                        break;
                    }
                case RESPONSE_WRITE_CMD: {//Ð´Ö¸Áî
                        if (RESPONSE_CORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "成功";
                        } else if (RESPONSE_INCORRECT == g_Protocol_Buff_p->Data[0]) {
                            toolStripStatusLabelCOM.Text = "设置" + showword + "失败";
                        } else {
                            toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        }
                        break;
                    }
                default: {// ´íÎó
                        toolStripStatusLabelCOM.Text = "接收数值解析错误";
                        break;
                    }
            }
        }

        /************************************************************************/
        //end
    }
}
