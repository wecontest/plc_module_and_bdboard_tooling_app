using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace stm32load
{
    class Bits
    {
        public static byte MsbOfUInt32(UInt32 addr)
        {
            return Convert.ToByte((addr >> 24) & 0xff);
        }

        public static byte ThreeByteOfUInt32(UInt32 addr)
        {
            return Convert.ToByte((addr >> 16) & 0xff);
        }

        public static byte TwoByteOfUInt32(UInt32 addr)
        {
            return Convert.ToByte((addr >> 8) & 0xff);
        }

        public static byte LsbOfUInt32(UInt32 addr)
        {
            return Convert.ToByte((addr & 0xff));
        }

        public static byte checkSum(byte num)
        {
            return Convert.ToByte(~num);
        }

        public static byte CheckSum(byte[] nums)
        {
            byte tmp = 0x00;

            foreach (byte b in nums)
            {
                tmp ^= b;
            }

            return tmp;
        }

        public static byte CheckSum(byte[] nums, int length)
        {
            byte tmp = 0x00;
            for (int i = 0; i < length; i++)
            {
                tmp ^= nums[i];
            }
            return tmp;
        }

        //from firstIndex to lastIndex,include firstIndex and lastIndex;
        public static byte CheckSum(byte[] nums, int firstIndex, int lastIndex)
        {
            byte tmp = 0x00;
            for (int i = firstIndex; i <= lastIndex; i++)
            {
                tmp ^= nums[i];
            }
            return tmp;
        }
    }//end class
}//end namespaces
