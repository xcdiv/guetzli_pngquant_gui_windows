using System;
using System.Collections.Generic;
 
using System.Text;
 

namespace ad_optimization
{
   public class PicList
    {

        private string _filename = "";

        public string Filename
        {
            get { return _filename; }
            set { _filename = value; }
        }
        private string _ext = "";

        public string Ext
        {
            get { return _ext; }
            set { _ext = value; }
        }
        private int _size = 0;

        public int Size
        {
            get { return _size; }
            set { _size = value; }
        }
        private string _filenameOut = "";

        public string FilenameOut
        {
            get { return _filenameOut; }
            set { _filenameOut = value; }
        }
        private int _size_Out = 0;

        public int Size_Out
        {
            get { return _size_Out; }
            set { _size_Out = value; }
        }

        private string _status = "未处理";

        public string Status
        {
            get { return _status; }
            set { _status = value; }
        }
    }
}
