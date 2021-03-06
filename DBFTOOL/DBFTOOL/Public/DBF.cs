﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/*
1、DBF数据库文件结构
   DBF文件由两部分组成，第一部分是文件头，其前32个字节是文件的整体描述，接着每32个字节定义一个字段，
   直到碰到一个0DH (字段描述结束符或称为文件头结束标志)为止；第二部分是实际存放每一个记录的数据部分。
   文件头部分的前32个字节说明如下：
   起止字节      长  度                 含  义 
   0             1个字节         03H表示无备注型字段，83H表示有
   1～3          3个字节         最后一次修改日期(yy／mm／dd)
   4～7          4个字节         DBF文件的记录数，低字节在前
   8～9          2个字节         文件头的长度，低字节在前
   10～11        2个字节         记录长度，低字节在前
   12～31        20个字节        保留字节

   从第32个字节开始到0DH为止是字段描述区，每32个字节定义一个字段，包括字段名、字段类型、字段长度、
   小数位数等。字段描述的各字节意义如表2。
   起止字节        长  度        含  义 
   0～10          11个字节          字段名 
   11             1个字节       字段类型(ASCII码) 
   12～15         4个字节       字段数据在内存中的地址
   16             1个字节       字段长度(二进制数) 
   17             1个字节       数值字段小数位数(二进制数)
   18～31         14个字节      保留字节

   文件头长度＝32+(32*定义的字段个数)+1在0DH后面，紧接着存放数据记录。记录以定长格式顺序存贮，
   每个记录的第一个字节是删除标识，有删除标记的记录，该字节是2AH(对应符号 “*”)，无删除标记
   的记录，该字节为空格(20H)。每个记录的各字段之间没有分隔符，记录无终止符，各种类型的数据均
   以ASCII码存放。数据记录之后为一个字节的文尾标识(1AH)。
   C-字符型     Y-货币型         N-数值型        F-浮点型    
   D-日期型     T-日期时间型     B-双精度型      I-整型    
   L-逻辑型     M-备注型         G-通用型        P-图片型
 */


using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace dbfcomp
{
    public class TDbfHeader
    {
        public const int HeaderSize = 32;
        public sbyte Version;
        public byte LastModifyYear;
        public byte LastModifyMonth;
        public byte LastModifyDay;
        public int RecordCount;
        public ushort HeaderLength;
        public ushort RecordLength;
        public byte[] Reserved = new byte[16];
        public sbyte TableFlag;
        public sbyte CodePageFlag;
        public byte[] Reserved2 = new byte[2];
    }

    public class TDbfField
    {
        public const int FieldSize = 32;
        public byte[] NameBytes = new byte[11];  // 字段名称  
        public byte TypeChar;
        public byte Length;
        public byte Precision;
        public byte[] Reserved = new byte[2];
        public sbyte DbaseivID;
        public byte[] Reserved2 = new byte[10];
        public sbyte ProductionIndex;

        public bool IsString
        {
            get
            {
                if (TypeChar == 'C')
                {
                    return true;
                }

                return false;
            }
        }

        public bool IsMoney
        {
            get
            {
                if (TypeChar == 'Y')
                {
                    return true;
                }

                return false;
            }
        }

        public bool IsNumber
        {
            get
            {
                if (TypeChar == 'N')
                {
                    return true;
                }

                return false;
            }
        }

        public bool IsFloat
        {
            get
            {
                if (TypeChar == 'F')
                {
                    return true;
                }

                return false;
            }
        }

        public bool IsDate
        {
            get
            {
                if (TypeChar == 'D')
                {
                    return true;
                }

                return false;
            }
        }

        public bool IsTime
        {
            get
            {
                if (TypeChar == 'T')
                {
                    return true;
                }

                return false;
            }
        }

        public bool IsDouble
        {
            get
            {
                if (TypeChar == 'B')
                {
                    return true;
                }

                return false;
            }
        }

        public bool IsInt
        {
            get
            {
                if (TypeChar == 'I')
                {
                    return true;
                }

                return false;
            }
        }

        public bool IsLogic
        {
            get
            {
                if (TypeChar == 'L')
                {
                    return true;
                }

                return false;
            }
        }

        public bool IsMemo
        {
            get
            {
                if (TypeChar == 'M')
                {
                    return true;
                }

                return false;
            }
        }

        public bool IsGeneral
        {
            get
            {
                if (TypeChar == 'G')
                {
                    return true;
                }

                return false;
            }
        }

        public Type FieldType
        {
            get
            {
                if (this.IsString == true)
                {
                    return typeof(string);
                }
                else if (this.IsMoney == true || this.IsNumber == true || this.IsFloat == true)
                {
                    return typeof(decimal);
                }
                else if (this.IsDate == true || this.IsTime == true)
                {
                    return typeof(System.DateTime);
                }
                else if (this.IsDouble == true)
                {
                    return typeof(double);
                }
                else if (this.IsInt == true)
                {
                    return typeof(System.Int32);
                }
                else if (this.IsLogic == true)
                {
                    return typeof(bool);
                }
                else if (this.IsMemo == true)
                {
                    return typeof(string);
                }
                else if (this.IsMemo == true)
                {
                    return typeof(string);
                }
                else
                {
                    return typeof(string);
                }
            }
        }

        public string GetFieldName()
        {
            return GetFieldName(System.Text.Encoding.Default);
        }

        public string GetFieldName(System.Text.Encoding encoding)
        {
            string fieldName = encoding.GetString(NameBytes);
            int i = fieldName.IndexOf('\0');
            if (i > 0)
            {
                return fieldName.Substring(0, i).Trim();
            }

            return fieldName.Trim();
        }
    }

    public class TDbfTable : IDisposable
    {
        private const byte DeletedFlag = 0x2A;
        private DateTime NullDateTime = new DateTime(1899, 12, 30);  // odbc中空日期对应的转换日期  

        private string _dbfFileName = null;

        private System.Text.Encoding _encoding = System.Text.Encoding.Default;
        private System.IO.FileStream _fileStream = null;
        private System.IO.BinaryReader _binaryReader = null;

        private bool _isFileOpened;
        private byte[] _recordBuffer;
        private int _fieldCount = 0;

        private TDbfHeader _dbfHeader = null;
        private TDbfField[] _dbfFields;
        private System.Data.DataTable _dbfTable = null;

        public TDbfTable(string fileName)
        {
            this._dbfFileName = fileName.Trim();
            try
            {
                this.OpenDbfFile();
            }
            finally
            {
                this.CloseFileStream();
            }
        }

        public TDbfTable(string fileName, string encodingName)
        {
            this._dbfFileName = fileName.Trim();
            this._encoding = GetEncoding(encodingName);
            try
            {
                this.OpenDbfFile();
            }
            finally
            {
                this.CloseFileStream();
            }
        }

        void System.IDisposable.Dispose()
        {
            this.Dispose(true);  // TODO:  添加 DBFFile.System.IDisposable.Dispose 实现  
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing == true)
            {
                this.Close();
            }
        }

        private System.Text.Encoding GetEncoding(string encodingName)
        {
            if (string.IsNullOrEmpty(encodingName) == true)
            {
                return System.Text.Encoding.Default;
            }

            if (encodingName.ToUpper() == "GB2313")
            {
                return System.Text.Encoding.GetEncoding("GB2312");
            }

            if (encodingName.ToUpper() == "UNICODE")
            {
                return System.Text.Encoding.Unicode;
            }

            if (encodingName.ToUpper() == "UTF8")
            {
                return System.Text.Encoding.UTF8;
            }

            if (encodingName.ToUpper() == "UTF7")
            {
                return System.Text.Encoding.UTF7;
            }

            if (encodingName.ToUpper() == "UTF32")
            {
                return System.Text.Encoding.UTF32;
            }

            if (encodingName.ToUpper() == "ASCII")
            {
                return System.Text.Encoding.ASCII;
            }

            return System.Text.Encoding.Default;
        }

        public void Close()
        {
            this.CloseFileStream();

            _recordBuffer = null;
            _dbfHeader = null;
            _dbfFields = null;

            _isFileOpened = false;
            _fieldCount = 0;
        }

        private void CloseFileStream()
        {
            if (_fileStream != null)
            {
                _fileStream.Close();
                _fileStream = null;
            }

            if (_binaryReader != null)
            {
                _binaryReader.Close();
                _binaryReader = null;
            }
        }

        private void OpenDbfFile()
        {
            this.Close();

            if (string.IsNullOrEmpty(_dbfFileName) == true)
            {
                throw new Exception("filename is empty or null.");
            }

            if (System.IO.File.Exists(_dbfFileName) == false)
            {
                throw new Exception(this._dbfFileName + " does not exist.");
            }

            try
            {
                this.GetFileStream();
                this.ReadHeader();
                this.ReadFields();
                this.GetRecordBufferBytes();
                this.CreateDbfTable();
                this.GetDbfRecords();
            }
            catch (Exception e)
            {
                this.Close();
                throw e;
            }
        }

        public void GetFileStream()
        {
            try
            {
                this._fileStream = File.Open(this._dbfFileName, FileMode.Open, FileAccess.Read, FileShare.Read);
                this._binaryReader = new BinaryReader(this._fileStream, _encoding);
                this._isFileOpened = true;
            }
            catch
            {
                throw new Exception("fail to read  " + this._dbfFileName + ".");
            }
        }

        private void ReadHeader()
        {
            this._dbfHeader = new TDbfHeader();

            try
            {
                this._dbfHeader.Version = this._binaryReader.ReadSByte();         //第1字节  
                this._dbfHeader.LastModifyYear = this._binaryReader.ReadByte();   //第2字节  
                this._dbfHeader.LastModifyMonth = this._binaryReader.ReadByte();  //第3字节  
                this._dbfHeader.LastModifyDay = this._binaryReader.ReadByte();    //第4字节  
                this._dbfHeader.RecordCount = this._binaryReader.ReadInt32();     //第5-8字节  
                this._dbfHeader.HeaderLength = this._binaryReader.ReadUInt16();   //第9-10字节  
                this._dbfHeader.RecordLength = this._binaryReader.ReadUInt16();   //第11-12字节  
                this._dbfHeader.Reserved = this._binaryReader.ReadBytes(16);      //第13-14字节  
                this._dbfHeader.TableFlag = this._binaryReader.ReadSByte();       //第15字节  
                this._dbfHeader.CodePageFlag = this._binaryReader.ReadSByte();    //第16字节  
                this._dbfHeader.Reserved2 = this._binaryReader.ReadBytes(2);      //第17-18字节  

                this._fieldCount = GetFieldCount();
            }
            catch
            {
                throw new Exception("fail to read file header.");
            }
        }

        private int GetFieldCount()
        {
            // 由于有些dbf文件的文件头最后有附加区段，但是有些文件没有，在此使用笨方法计算字段数目  
            // 就是测试每一个存储字段结构区域的第一个字节的值，如果不为0x0D，表示存在一个字段  
            // 否则从此处开始不再存在字段信息  

            int fCount = (this._dbfHeader.HeaderLength - TDbfHeader.HeaderSize - 1) / TDbfField.FieldSize;

            for (int k = 0; k < fCount; k++)
            {
                _fileStream.Seek(TDbfHeader.HeaderSize + k * TDbfField.FieldSize, SeekOrigin.Begin);  // 定位到每个字段结构区，获取第一个字节的值  
                byte flag = this._binaryReader.ReadByte();

                if (flag == 0x0D)  // 如果获取到的标志不为0x0D，则表示该字段存在；否则从此处开始后面再没有字段信息  
                {
                    return k;
                }
            }

            return fCount;
        }

        private void ReadFields()
        {
            _dbfFields = new TDbfField[_fieldCount];

            try
            {
                _fileStream.Seek(TDbfHeader.HeaderSize, SeekOrigin.Begin);
                for (int k = 0; k < _fieldCount; k++)
                {
                    this._dbfFields[k] = new TDbfField();
                    this._dbfFields[k].NameBytes = this._binaryReader.ReadBytes(11);
                    this._dbfFields[k].TypeChar = this._binaryReader.ReadByte();

                    this._binaryReader.ReadBytes(4);  // 保留, 源代码是读 UInt32()给 Offset  

                    this._dbfFields[k].Length = this._binaryReader.ReadByte();
                    this._dbfFields[k].Precision = this._binaryReader.ReadByte();
                    this._dbfFields[k].Reserved = this._binaryReader.ReadBytes(2);
                    this._dbfFields[k].DbaseivID = this._binaryReader.ReadSByte();
                    this._dbfFields[k].Reserved2 = this._binaryReader.ReadBytes(10);
                    this._dbfFields[k].ProductionIndex = this._binaryReader.ReadSByte();
                }
            }
            catch
            {
                throw new Exception("fail to read field information.");
            }
        }

        private void GetRecordBufferBytes()
        {
            this._recordBuffer = new byte[this._dbfHeader.RecordLength];

            if (this._recordBuffer == null)
            {
                throw new Exception("fail to allocate memory .");
            }
        }

        private void CreateDbfTable()
        {
            if (_dbfTable != null)
            {
                _dbfTable.Clear();
                _dbfTable = null;
            }

            _dbfTable = new System.Data.DataTable();
            _dbfTable.TableName = this.TableName;

            for (int k = 0; k < this._fieldCount; k++)
            {
                System.Data.DataColumn col = new System.Data.DataColumn();
                string colText = this._dbfFields[k].GetFieldName(_encoding);

                if (string.IsNullOrEmpty(colText) == true)
                {
                    throw new Exception("the " + (k + 1) + "th column name is null.");
                }

                col.ColumnName = colText;
                col.Caption = colText;
                col.DataType = this._dbfFields[k].FieldType;
                _dbfTable.Columns.Add(col);
            }
        }

        public void GetDbfRecords()
        {
            try
            {
                this._fileStream.Seek(this._dbfHeader.HeaderLength, SeekOrigin.Begin);

                for (int k = 0; k < this.RecordCount; k++)
                {
                    if (ReadRecordBuffer(k) != DeletedFlag)
                    {
                        System.Data.DataRow row = _dbfTable.NewRow();
                        for (int i = 0; i < this._fieldCount; i++)
                        {
                            row[i] = this.GetFieldValue(i);
                        }
                        _dbfTable.Rows.Add(row);
                    }
                }
            }
            catch (ArgumentOutOfRangeException e)
            {
                throw e;
            }
            catch
            {
                throw new Exception("fail to get dbf table.");
            }
        }

        private byte ReadRecordBuffer(int recordIndex)
        {
            byte deleteFlag = this._binaryReader.ReadByte();  // 删除标志  
            this._recordBuffer = this._binaryReader.ReadBytes(this._dbfHeader.RecordLength - 1);  // 标志位已经读取  
            return deleteFlag;
        }

        private string GetFieldValue(int fieldIndex)
        {
            string fieldValue = null;

            int offset = 0;
            for (int i = 0; i < fieldIndex; i++)
            {
                offset += _dbfFields[i].Length;
            }

            byte[] tmp = CopySubBytes(this._recordBuffer, offset, this._dbfFields[fieldIndex].Length);

            if (this._dbfFields[fieldIndex].IsInt == true)
            {
                int val = System.BitConverter.ToInt32(tmp, 0);
                fieldValue = val.ToString();
            }
            else if (this._dbfFields[fieldIndex].IsDouble == true)
            {
                double val = System.BitConverter.ToDouble(tmp, 0);
                fieldValue = val.ToString();
            }
            else if (this._dbfFields[fieldIndex].IsMoney == true)
            {
                long val = System.BitConverter.ToInt64(tmp, 0);  // 将字段值放大10000倍，变成long型存储，然后缩小10000倍。  
                fieldValue = ((decimal)val / 10000).ToString();
            }
            else if (this._dbfFields[fieldIndex].IsDate == true)
            {
                DateTime date = ToDate(tmp);
                fieldValue = date.ToString();

            }
            else if (this._dbfFields[fieldIndex].IsTime == true)
            {
                DateTime time = ToTime(tmp);
                fieldValue = time.ToString();

            }
            else
            {
                fieldValue = this._encoding.GetString(tmp);
            }

            fieldValue = fieldValue.Trim();

            // 如果本子段类型是数值相关型，进一步处理字段值  
            if (this._dbfFields[fieldIndex].IsNumber == true || this._dbfFields[fieldIndex].IsFloat == true)    // N - 数值型, F - 浮点型                      
            {
                if (fieldValue.Length == 0)
                {
                    fieldValue = "0";
                }
                else if (fieldValue == ".")
                {
                    fieldValue = "0";
                }
                else
                {
                    decimal val = 0;

                    if (decimal.TryParse(fieldValue, out val) == false)  // 将字段值先转化为Decimal类型然后再转化为字符串型，消除类似“.000”的内容, 如果不能转化则为0  
                    {
                        val = 0;
                    }

                    fieldValue = val.ToString();
                }
            }
            else if (this._dbfFields[fieldIndex].IsLogic == true)    // L - 逻辑型  
            {
                if (fieldValue != "T" && fieldValue != "Y")
                {
                    fieldValue = "false";
                }
                else
                {
                    fieldValue = "true";
                }
            }
            else if (this._dbfFields[fieldIndex].IsDate == true || this._dbfFields[fieldIndex].IsTime == true)   // D - 日期型  T - 日期时间型                      
            {
                // 暂时不做任何处理  
            }

            return fieldValue;
        }

        private static byte[] CopySubBytes(byte[] buf, int startIndex, long length)
        {
            if (startIndex >= buf.Length)
            {
                throw new ArgumentOutOfRangeException("startIndex");
            }

            if (length == 0)
            {
                throw new ArgumentOutOfRangeException("length", "length must be great than 0.");
            }

            if (length > buf.Length - startIndex)
            {
                length = buf.Length - startIndex;  // 子数组的长度超过从startIndex起到buf末尾的长度时，修正为剩余长度  
            }

            byte[] target = new byte[length];
            Array.Copy(buf, startIndex, target, 0, length);
            return target;
        }

        private DateTime ToDate(byte[] buf)
        {
            if (buf.Length != 8)
            {
                throw new ArgumentException("date array length must be 8.", "buf");
            }

            string dateStr = System.Text.Encoding.ASCII.GetString(buf).Trim();
            if (dateStr.Length < 8)
            {
                return NullDateTime;
            }

            int year = int.Parse(dateStr.Substring(0, 4));
            int month = int.Parse(dateStr.Substring(4, 2));
            int day = int.Parse(dateStr.Substring(6, 2));

            return new DateTime(year, month, day);
        }

        private DateTime ToTime(byte[] buf)
        {
            if (buf.Length != 8)
            {
                throw new ArgumentException("time array length must be 8.", "buf");
            }

            try
            {
                byte[] tmp = CopySubBytes(buf, 0, 4);
                tmp.Initialize();
                int days = System.BitConverter.ToInt32(tmp, 0);  // ( ToInt32(tmp); // 获取天数                  

                tmp = CopySubBytes(buf, 4, 4);  // 获取毫秒数  
                int milliSeconds = System.BitConverter.ToInt32(tmp, 0);  // ToInt32(tmp);  

                if (days == 0 && milliSeconds == 0)
                {
                    return NullDateTime;
                }

                int seconds = milliSeconds / 1000;
                int milli = milliSeconds % 1000;  // vfp实际上没有毫秒级, 是秒转换来的, 测试时发现2秒钟转换为1999毫秒的情况  
                if (milli > 0)
                {
                    seconds += 1;
                }

                DateTime date = DateTime.MinValue;  // 在最小日期时间的基础上添加刚获取的天数和秒数，得到日期字段数值  
                date = date.AddDays(days - 1721426);
                date = date.AddSeconds(seconds);

                return date;
            }
            catch
            {
                return new DateTime();
            }
        }

        public string TableName
        {
            get { return System.IO.Path.GetFileNameWithoutExtension(this._dbfFileName); }
        }

        public System.Text.Encoding Encoding
        {
            get { return this._encoding; }
        }

        public int RecordLength
        {
            get
            {
                if (this.IsFileOpened == false)
                {
                    return 0;
                }

                return this._dbfHeader.RecordLength;
            }
        }

        public int FieldCount
        {
            get
            {
                if (this.IsFileOpened == false)
                {
                    return 0;
                }

                return this._dbfFields.Length;
            }
        }

        public int RecordCount
        {
            get
            {
                if (this.IsFileOpened == false || this._dbfHeader == null)
                {
                    return 0;
                }

                return this._dbfHeader.RecordCount;
            }
        }

        public bool IsFileOpened
        {
            get
            {
                return this._isFileOpened;
            }
        }

        public System.Data.DataTable Table
        {
            get
            {
                if (_isFileOpened == false)
                {
                    return null;
                }
                return _dbfTable;
            }
        }

        public TDbfField[] DbfFields
        {
            get
            {
                if (_isFileOpened == false)
                {
                    return null;
                }

                return _dbfFields;
            }
        }
    }
}