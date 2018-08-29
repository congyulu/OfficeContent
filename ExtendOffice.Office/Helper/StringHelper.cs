using System;
using System.Text;

namespace ExtendOffice.Office.Helper
{
    public static class StringHelper
    {
        internal static String GetString(Boolean isUnicode, Byte[] data)
        {
            if (isUnicode)
            {
                return Encoding.Unicode.GetString(data).Replace("\0", "");
            }
            else
            {
                return Encoding.GetEncoding("Windows-1252").GetString(data).Replace("\0", "");
            }
        }
        internal static String GetString(Encoding encoding, Byte[] data)
        {
            return encoding.GetString(data);
        }

        internal static String ReplaceString(String origin)
        {
            StringBuilder sb = new StringBuilder();

            for (Int32 i = 0; i < origin.Length; i++)
            {
                if (origin[i] == '\r')
                {
                    sb.Append("\r\n");
                    continue;
                }
                else
                {
                    sb.Append(origin[i]);
                }
            }
            sb = sb.Replace("", "").Replace("", "").Replace("", "");
            return System.Text.RegularExpressions.Regex.Replace(sb.ToString(), @"[\x00-\x08]|[\x0B-\x0C]|[\x0E-\x1F]", "");
        }
        internal static String ReplaceStringWord(String origin)
        {
            StringBuilder sb = new StringBuilder();
            bool isRead = true;
            for (Int32 i = 0; i < origin.Length; i++)
            {
                if (origin[i] == '\r')
                {
                    sb.Append("\r\n");
                    continue;
                }
                else if (origin[i] == 19)
                {
                    isRead = false;
                }
                else if (origin[i] == 20)
                {
                    isRead = true;
                }
                else if (origin[i] == 21)
                {
                    isRead = true;
                }
                else
                {
                    if (isRead)
                    {
                        sb.Append(origin[i]);
                    }
                }
            }
            sb = sb.Replace("", "").Replace("", "").Replace("", "");
            return System.Text.RegularExpressions.Regex.Replace(sb.ToString(), @"[\x00-\x08]|[\x0B-\x0C]|[\x0E-\x1F]", "");
        }
        public static string byteToHexStr(byte[] bytes)
        {
            string returnStr = "";
            if (bytes != null)
            {
                for (int i = 0; i < bytes.Length; i++)
                {
                    returnStr += bytes[i].ToString("X2");
                }
            }
            return returnStr;
        }
    }
}