using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PhHslComm
{
    public class ToolAPI
    {
        #region Convert Float Array To Ascii

        public StringBuilder ConvertFloatToAscii(float value)
        {
            StringBuilder asciiString = new StringBuilder(512);

             
            if (value >0 && value <= 255)  //value不会是0 if (value >= 0 && value <= 255)  
            {
                System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
                byte[] byteArray = new byte[] { (byte)value };
                asciiString.Append(asciiEncoding.GetString(byteArray));
            }
            else if (value == 0)
            {
                asciiString.Append("");

            }
            else
            {
                throw new Exception("ASCII Code is not valid.");
            }


            return asciiString;
        }



        public StringBuilder ConvertFloatArrayToAscii(float[] value, int startIndex, int endIndex)
        {
            StringBuilder asciiString = new StringBuilder(512);
            for (int i = startIndex; i < (endIndex + 1); i++)
            {
                asciiString.Append(ConvertFloatToAscii(value[i]));
            }
            asciiString.Append(",");
            return asciiString;
        }

        public StringBuilder ConvertFloatArrayToAscii(float[] value)
        {
            StringBuilder asciiString = new StringBuilder(512);
            foreach (float f in value)
            {
                if (f != 0)
                {
                    asciiString.Append(ConvertFloatToAscii(f));
                }
            }
            return asciiString;
        }

        #endregion



       
    }
}
