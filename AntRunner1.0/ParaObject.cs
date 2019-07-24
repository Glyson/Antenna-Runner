using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Windows.Input;

namespace AntRunner
{
    public class ParaObject : INotifyPropertyChanged
    {
        private string _code = null;

        public string Code
        {
            get { return _code; }
            set
            {
                _code = value;
            }
        }

        private bool _enable = true;

        public bool Enable
        {
            get { return _enable; }
            set
            {
                _enable = value;
                OnPropertyChanged("Enable");
            }
        }
        private double _power = 0;

        public double Power
        {
            get { return _power; }
            set
            {
                _power = value;
                OnPropertyChanged("Power");
            }
        }
        private double _bandwidth = 100;

        public double Bandwidth
        {
            get { return _bandwidth; }
            set
            {
                _bandwidth = value;
                OnPropertyChanged("Bandwidth");
            }
        }
        private double _freqStart = 1800;

        public double FreqStart
        {
            get { return _freqStart; }
            set
            {
                _freqStart = value;
                OnPropertyChanged("FreqStart");
            }
        }
        private double _freqStop = 2000;

        public double FreqStop
        {
            get { return _freqStop; }
            set
            {
                _freqStop = value;
                OnPropertyChanged("FreqStop");
            }
        }
        private int _points = 401;

        public int Points
        {
            get { return _points; }
            set
            {
                _points = value;
                OnPropertyChanged("Points");
            }
        }

        private string _markerType = "Points";

        public string MarkerType
        {
            get { return _markerType; }
            set
            {
                _markerType = value;
                OnPropertyChanged("MarkerType");
            }
        }
        private int _markerPoints = 11;

        public int MarkerPoints
        {
            get { return _markerPoints; }
            set
            {
                _markerPoints = value;
                OnPropertyChanged("MarkerPoints");
            }
        }
        private string _markerText = "";

        public string MarkerText
        {
            get { return _markerText; }
            set
            {
                _markerText = value;
                OnPropertyChanged("MarkerText");
            }
        }

        private double _markerStart = 1800;

        public double MarkerStart
        {
            get { return _markerStart; }
            set
            {
                _markerStart = value;
                OnPropertyChanged("MarkerStart");
            }
        }
        private double _markerStop = 2000;

        public double MarkerStop
        {
            get { return _markerStop; }
            set
            {
                _markerStop = value;
                OnPropertyChanged("MarkerStop");
            }
        }

        private string _trace = "S11";

        public string Trace
        {
            get { return _trace; }
            set
            {
                _trace = value;
                OnPropertyChanged("Trace");
            }
        }

        private State _state = State.Stoped;

        public State State
        {
            get { return _state; }
            set { _state = value; }
        }

        private Key _fastKey = Key.F1;

        public Key FastKey
        {
            get { return _fastKey; }
            set { _fastKey = value; }
        }

        private string _referTracePath = string.Empty;

        public string ReferTracePath
        {
            get { return _referTracePath; }
            set
            {
                _referTracePath = value;
                OnPropertyChanged("ReferTracePath");
            }
        }

        private double _referDiff = 5;

        public double ReferDiff
        {
            get { return _referDiff; }
            set
            {
                _referDiff = value;
                OnPropertyChanged("ReferDiff");
            }
        }

        private string _scannerCOM = "COM1";

        public string ScannerCOM
        {
            get { return _scannerCOM; }
            set { _scannerCOM = value; }
        }

        private double _cutBW = 950;

        public double CutBW
        {
            get { return _cutBW; }
            set
            {
                _cutBW = value;
                OnPropertyChanged("CutBW");
            }
        }
        private double _cutLeftFreq = 950;

        public double CutLeftFreq
        {
            get { return _cutLeftFreq; }
            set
            {
                _cutLeftFreq = value;
                OnPropertyChanged("CutLeftFreq");
            }
        }
        private double _cutRightFreq = 950;

        public double CutRightFreq
        {
            get { return _cutRightFreq; }
            set
            {
                _cutRightFreq = value;
                OnPropertyChanged("CutRightFreq");
            }
        }


        private double _cutPow = 10;

        public double CutPow
        {
            get { return _cutPow; }
            set
            {
                _cutPow = value;
                OnPropertyChanged("CutPow");
            }
        }
        private double _diffBW = 100;

        public double DiffBW
        {
            get { return _diffBW; }
            set
            {
                _diffBW = value;
                OnPropertyChanged("DiffBW");
            }
        }
        private double _diffFreq = 100;

        public double DiffFreq
        {
            get { return _diffFreq; }
            set
            {
                _diffFreq = value;
                OnPropertyChanged("DiffFreq");
            }
        }
        private double _diffPower = 10;

        public double DiffPower
        {
            get { return _diffPower; }
            set
            {
                _diffPower = value;
                OnPropertyChanged("DiffPower");
            }
        }
        private double _diffFreq_Bad = 100;

        public double DiffFreq_Bad
        {
            get { return _diffFreq_Bad; }
            set
            {
                _diffFreq_Bad = value;
                OnPropertyChanged("DiffFreq_Bad");
            }
        }
        private double _diffPower_Bad = 10;

        public double DiffPower_Bad
        {
            get { return _diffPower_Bad; }
            set
            {
                _diffPower_Bad = value;
                OnPropertyChanged("DiffPower_Bad");
            }
        }
        private double _s21Min = 10;

        public double S21Min
        {
            get { return _s21Min; }
            set
            {
                _s21Min = value;
                OnPropertyChanged("S21Min");
            }
        }
        private double _s21Max = 10;

        public double S21Max
        {
            get { return _s21Max; }
            set
            {
                _s21Max = value;
                OnPropertyChanged("S21Max");
            }
        }
        private double _s22Max = 10;

        public double S22Max
        {
            get { return _s22Max; }
            set
            {
                _s22Max = value;
                OnPropertyChanged("S22Max");
            }
        }

        private S22TraceFormat _s22TraceFormat = S22TraceFormat.SWR;

        public S22TraceFormat S22TraceFormat
        {
            get { return _s22TraceFormat; }
            set
            {
                _s22TraceFormat = value;
                OnPropertyChanged("S22TraceFormat");
            }
        }

        private object _markers;
        public object Markers
        {
            get { return _markers; }
            set
            {
                _markers = value;
                //OnPropertyChanged("Markers");
            }
        }
        private object _markersCal;
        public object MarkersCal
        {
            get { return _markersCal; }
            set
            {
                _markersCal = value;
                //OnPropertyChanged("MarkersCal");
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged(string propName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propName));
            }
        }

    }
}
