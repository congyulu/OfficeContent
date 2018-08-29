using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

using ExtendOffice.Office.Entity;
using ExtendOffice.Office.Helper;

namespace ExtendOffice.Office
{
    public class PowerPointFile : CompoundBinaryFile, IPowerPointFile
    {
        #region 字段
        private MemoryStream _contentStream;

        private List<PowerPointRecord> _records;
        private StringBuilder _allText;
        #region 测试方法
        private StringBuilder _recordTree;
        #endregion
        #endregion

        #region 属性
        /// <summary>
        /// 获取PowerPoint幻灯片中所有文本
        /// </summary>
        public String AllText
        {
            get { return this._allText.ToString(); }
        }

        #region 测试方法
        /// <summary>
        /// 获取PowerPoint中Record的树形结构
        /// </summary>
        public String RecordTree
        {
            get { return this._recordTree.ToString(); }
        }
        #endregion
        #endregion

        #region 构造函数
        /// <summary>
        /// 初始化PptFile
        /// </summary>
        /// <param name="filePath">文件路径</param>
        public PowerPointFile(String filePath) :
            base(filePath)
        { }
        #endregion

        #region 读取内容
        protected override void ReadContent()
        {
            DirectoryEntry entry = this._dirRootEntry.GetChild("PowerPoint Document");

            if (entry == null)
            {
                return;
            }

            try
            {
                this._contentStream = new MemoryStream(this._entryData[entry.EntryID]);

                #region 测试方法
                this._recordTree = new StringBuilder();
                #endregion

                this._allText = new StringBuilder();
                this._records = new List<PowerPointRecord>();
                PowerPointRecord record = null;
                while (this._contentStream.Position < this._contentStream.Length)
                {
                    bool isPass = false;
                    record = this.ReadRecord(null, out isPass);
                    if (record == null || record.RecordType == 0)
                    {
                        break;
                    }
                }
                this._allText = new StringBuilder(StringHelper.ReplaceString(this._allText.ToString()));
            }
            finally
            {

                if (this._contentStream != null)
                {
                    this._contentStream.Close();
                }
            }
        }

        private PowerPointRecord ReadRecord(PowerPointRecord parent, out bool isPass)
        {
            isPass = false;
            PowerPointRecord record = GetRecordBy8(parent);
            if (record == null)
            {
                return null;
            }

            #region 测试方法
            //this._allText.Append('-', record.Deepth * 2);
            //this._allText.AppendFormat("[{0}]-[{1}]-[Len:{2}] [{3}] [{4}] ", record.RecordType, record.Deepth, record.RecordLength, record.RecordInstance, record.RecordVersion);
            //this._allText.AppendLine();
            #endregion

            if (parent == null)
            {
                this._records.Add(record);
            }
            else
            {
                parent.AddChild(record);
            }

            switch (record.RecordType)
            {
                //移除备注、模板
                case RecordType.Notes:
                case RecordType.MainMaster:
                case RecordType.Handout:
                case RecordType.DocumentContainer:
                    this._contentStream.Seek(record.RecordLength, SeekOrigin.Current);
                    break;
                case RecordType.SlideAtom:
                    if (ProcessSlideAtom(record))
                    {
                        isPass = true;
                    }
                    return record;
                case RecordType.TextCharsAtom:
                case RecordType.CString:
                    if (ProcessText(record))
                    {
                        return record;
                    }
                    break;
                case RecordType.TextBytesAtom:
                    if (ProcessByte(record))
                    {
                        return record;
                    }
                    break;
                case RecordType.OfficeArtFOPT:
                    ProcessFOPT(record);
                    return record;
            }
            if (record.RecordVersion == 0xF)
            {
                while (this._contentStream.Position < record.Offset + record.RecordLength)
                {
                    PowerPointRecord code = this.ReadRecord(record, out isPass);
                    if (isPass == true)
                    {
                        this._contentStream.Seek(record.RecordLength - (code.RecordLength + 8), SeekOrigin.Current);
                        break;
                    }
                    if (code == null)
                    {
                        this._contentStream.Seek(record.RecordLength, SeekOrigin.Current);
                    }
                }
            }
            else
            {
                try
                {
                    this._contentStream.Seek(record.RecordLength, SeekOrigin.Current);
                }
                catch (Exception)
                {

                }
            }
            return record;
        }
        private bool ProcessSlideAtom(PowerPointRecord record)
        {
            SlideAtom atom = GetSlideAtom(record);
            //this._recordTree.Append('-', record.Deepth * 2);
            //this._recordTree.AppendFormat("Geom[{0}]MasterIdRef[{1}]NotesIdRef[{2}]RgPlaceholderTypes[{3}]SlideFlags[{4}] ",
            //atom.Geom,
            //atom.MasterIdRef,
            //atom.NotesIdRef,
            //atom.RgPlaceholderTypes,
            //atom.SlideFlags);
            //this._recordTree.AppendLine();
            if (atom.SlideFlags == 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }


        private bool ProcessText(PowerPointRecord record)
        {
            if (record.Parent != null && (record.Parent.RecordType == RecordType.OfficeArtClientTextbox))
            {
                Byte[] data = new Byte[(Int32)record.RecordLength];
                this._contentStream.Read(data, 0, data.Length);
                String value = StringHelper.GetString(true, data);
                this._allText.Append(value);
                return true;
            }
            else
            {
                return false;
            }
        }
        private bool ProcessByte(PowerPointRecord record)
        {
            if (record.Parent != null && record.Parent.RecordType == RecordType.OfficeArtClientTextbox)
            {
                Byte[] data = new Byte[(Int32)record.RecordLength];
                this._contentStream.Read(data, 0, data.Length);
                String value = StringHelper.GetString(false, data);
                this._allText.Append(value);
                return true;
            }
            else
            {
                return false;
            }
        }
        private void ProcessFOPT(PowerPointRecord record)
        {
            List<Fopte> list = new List<Fopte>();
            int len = 0;
            for (int i = 0; i < record.RecordInstance; i++)
            {
                Fopte fopt = GetFopt(record);
                list.Add(fopt);
            }
            foreach (var item in list)
            {
                //this._allText.Append("-Bid:" + item.FBid);
                //this._allText.AppendFormat("[Complex:{0}]-[PID:{1}]-[Len:{2}] ", item.FComplex, item.PID, item.Op);
                //this._allText.AppendLine();
                if (item.FComplex == 1)
                {
                    if (item.PID == 16576)//文字
                    {
                        Byte[] data = new Byte[item.Op];
                        this._contentStream.Read(data, 0, item.Op);
                        String value = StringHelper.GetString(true, data);
                        this._allText.Append(value);
                    }
                    else
                    {
                        this._contentStream.Seek(item.Op, SeekOrigin.Current);
                    }
                    len += item.Op;
                }
            }
            if (len > 0 && record.RecordLength != len + 6 * record.RecordInstance)
            {
                this._contentStream.Seek(-len - 6 * record.RecordInstance, SeekOrigin.Current);
                this._contentStream.Seek(record.RecordLength, SeekOrigin.Current);
            }
        }
        private PowerPointRecord GetRecordBy8(PowerPointRecord parent)
        {
            if (this._contentStream.Position + 8 >= this._contentStream.Length)
            {
                return null;
            }
            Byte[] buf = new Byte[2];
            this._contentStream.Read(buf, 0, buf.Length);
            UInt16 version = BitConverter.ToUInt16(buf, 0);
            buf = new Byte[2];
            this._contentStream.Read(buf, 0, buf.Length);
            UInt16 type = BitConverter.ToUInt16(buf, 0);
            buf = new Byte[4];
            this._contentStream.Read(buf, 0, buf.Length);
            UInt32 length = BitConverter.ToUInt32(buf, 0);

            return new PowerPointRecord(parent, version, type, length, this._contentStream.Position);
        }
        private Fopte GetFopt(PowerPointRecord parent)
        {
            if (this._contentStream.Position + 6 >= this._contentStream.Length)
            {
                return null;
            }
            Byte[] buf = new Byte[2];
            this._contentStream.Read(buf, 0, buf.Length);
            UInt16 pid = BitConverter.ToUInt16(buf, 0);
            buf = new Byte[4];
            this._contentStream.Read(buf, 0, buf.Length);
            Int32 op = BitConverter.ToInt32(buf, 0);
            return new Fopte(pid, op);
        }
        private SlideAtom GetSlideAtom(PowerPointRecord parent)
        {
            Byte[] buf = new Byte[4];
            this._contentStream.Read(buf, 0, buf.Length);
            UInt32 geom = BitConverter.ToUInt32(buf, 0);
            buf = new Byte[8];
            this._contentStream.Read(buf, 0, buf.Length);
            UInt64 rgPlaceholderTypes = BitConverter.ToUInt64(buf, 0);
            buf = new Byte[4];
            this._contentStream.Read(buf, 0, buf.Length);
            UInt32 masterIdRef = BitConverter.ToUInt32(buf, 0);
            buf = new Byte[4];
            this._contentStream.Read(buf, 0, buf.Length);
            UInt32 notesIdRef = BitConverter.ToUInt32(buf, 0);
            buf = new Byte[2];
            this._contentStream.Read(buf, 0, buf.Length);
            UInt16 slideFlags = BitConverter.ToUInt16(buf, 0);
            buf = new Byte[2];
            this._contentStream.Read(buf, 0, buf.Length);
            UInt16 unused = BitConverter.ToUInt16(buf, 0);
            return new SlideAtom(geom, rgPlaceholderTypes, masterIdRef, notesIdRef, slideFlags, unused);
        }
        #endregion
    }
}