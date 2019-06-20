using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NationalInstruments.VisaNS;
using System.Windows;
using AntRunner.Properties;

namespace AntRunner
{
    public class VNA
    {
        public bool IsOK = false;
        protected ResourceManager rm = null;
        protected IMessageBasedSession ses = null;
        protected List<int> portList = new List<int>();
        private static VNA vna = null;
        private static Instrument ins;

        #region public method
        public static VNA CreateVNA()
        {
            if (vna != null && Settings.Default.Instrument == ins.ToString())
            {
                return vna;
            }
            else
            {
                ins = Helper.String2Enum<Instrument>(Settings.Default.Instrument);
                switch (ins)
                {
                    case Instrument.Agilent_5071C:
                        return new VNA_AT5071C();
                    case Instrument.Agilent_8753ES:
                        return new VNA_AT8753ES();
                    default:
                        throw new Exception(string.Format("Do not implement VNA({0}).", ins));
                }
            }
        }
        public static string[] ScanGPIB()
        {
            ResourceManager manager = ResourceManager.GetLocalManager();
            string[] listGPIB = manager.FindResources("GPIB?*INSTR");
            return listGPIB;
        }
        public bool Init(string gpib, int timeout = 60000)
        {
            try
            {
                rm = ResourceManager.GetLocalManager();
                Session s = rm.Open(gpib);
                s.Timeout = timeout;//1 mins
                ses = s as IMessageBasedSession;
                IsOK = true;
                return true;
            }
            catch (Exception ex)
            {
                AppLog.Error("Init has error.", ex);
                MainWindow.Self.Dispatcher.Invoke(new Action(delegate
                {
                    MessageBox.Show(MainWindow.Self, ex.Message, "Error");
                }));
                IsOK = false;
                return false;
            }
        }
        public virtual void Config() { }
        public virtual void Setup(ParaObject para) { }
        public virtual SortedList<double, double> ReadTrace(ParaObject para, int traceNum = 1) { return null; }
        #endregion


        protected void GetPorts()
        {
            portList = new List<int>();
            if (Settings.Default.Para1.Enable)
                portList.Add(1);
            if (Settings.Default.Para2.Enable)
                portList.Add(2);
            if (Settings.Default.Para3.Enable)
                portList.Add(3);
            if (Settings.Default.Para4.Enable)
                portList.Add(4);
        }

        protected SortedList<double, double> FixTrace(ParaObject para, string[] arrCur, int dimCnt = 1, int dim = 1)
        {
            SortedList<double, double> list = new SortedList<double, double>();
            double freq = para.FreqStart;
            double step = (para.FreqStop - para.FreqStart) / (para.Points - 1);
            for (int i = 0; i < para.Points; i++, freq += step)
                list.Add(freq, double.Parse(arrCur[i * dimCnt + dim - 1]));
            return list;
        }
        public void Write(string cmd, params object[] paras)
        {
            if (paras != null && paras.Length > 0)
            {
                cmd = string.Format(cmd, paras);
            }
            ses.Write(cmd);
        }
        public string Read(string cmd, params object[] paras)
        {
            if (paras != null && paras.Length > 0)
            {
                cmd = string.Format(cmd, paras);
            }
            return ses.Query(cmd);
        }
        public static string ReadIDN(string gpib)
        {
            try
            {
                ResourceManager r = ResourceManager.GetLocalManager();
                Session session = r.Open(gpib);
                session.Timeout = 1000;//1 mins
                IMessageBasedSession ises = session as IMessageBasedSession;
                return ises.Query("*IDN?");
            }
            catch (Exception ex)
            {
                AppLog.Error("ReadIDN has error.", ex);
                return ex.Message;
            }
        }

        ~VNA()
        {
            if (ses != null)
            {
                (ses as Session).Dispose();
            }
        }
    }
}
