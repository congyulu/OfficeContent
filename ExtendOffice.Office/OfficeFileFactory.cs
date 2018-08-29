using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExtendOffice.Office
{
    public class DocFormat
    {
        public const string Binary = "D0CF11E0";
        public const string XML = "504B0304";
    }

    public static class OfficeFileFactory
    {
        public static IOfficeFile CreateOfficeFile(String filePath)
        {
            String extension = Path.GetExtension(filePath).ToLower();
            if (File.Exists(filePath))
            {
                string hexstr = null;
                using (FileStream sf = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    if (sf.Length < 4)
                    {
                        return null;
                    }
                    byte[] header = new byte[4];
                    sf.Read(header, 0, header.Length);
                    try
                    {
                        hexstr = Helper.StringHelper.byteToHexStr(header);
                    }
                    catch (Exception)
                    {

                    }
                }
                if (!string.IsNullOrEmpty(hexstr))
                {
                    switch (hexstr)
                    {
                        case DocFormat.Binary:
                            if (extension.Contains(".doc"))
                            {
                                return new WordBinaryFile(filePath);
                            }
                            else if (extension.Contains(".ppt"))
                            {
                                return new PowerPointFile(filePath);
                            }
                            else
                            {
                                return null;
                            }
                        case DocFormat.XML:
                            if (extension.Contains(".doc"))
                            {
                                return new WordOOXMLFile(filePath);
                            }
                            else if (extension.Contains(".ppt"))
                            {
                                return new PowerPointOOXMLFile(filePath);
                            }
                            else
                            {
                                return null;
                            }
                        default:
                            return null;
                    }
                }
                else
                {
                    return null;
                }

            }
            else
            {
                return null;
            }
        }

    }
}