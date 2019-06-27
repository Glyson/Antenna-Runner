using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AntRunner
{
    public class SingleData
    {
        private string _filename;

        public string Filename
        {
            get { return _filename; }
            set { _filename = value; }
        }
        private string _traceType;

        public string TraceType
        {
            get { return _traceType; }
            set { _traceType = value; }
        }
        private string _result;

        public string Result
        {
            get { return _result; }
            set { _result = value; }
        }
        private string _code;

        public string Code
        {
            get { return _code; }
            set { _code = value; }
        }
        private string _errors;

        public string Errors
        {
            get { return _errors; }
            set { _errors = value; }
        }
        private string _memo;

        public string Memo
        {
            get { return _memo; }
            set { _memo = value; }
        }
        private SortedList<double, double> _referData;

        public SortedList<double, double> ReferData
        {
            get { return _referData; }
            set { _referData = value; }
        }

        private SortedList<double, double> _listData;

        public SortedList<double, double> ListData
        {
            get { return _listData; }
            set { _listData = value; }
        }

        private SortedList<double, double> _listData2;

        public SortedList<double, double> ListData2
        {
            get { return _listData2; }
            set { _listData2 = value; }
        }
    }
}
