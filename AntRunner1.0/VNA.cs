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
        public bool Init(string gpib)
        {
            try
            {
                rm = ResourceManager.GetLocalManager();
                Session s = rm.Open(gpib);
                s.Timeout = 60000;//1 mins
                ses = s as IMessageBasedSession;
                IsOK = true;
                return true;
            }
            catch (Exception ex)
            {
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
        public virtual SortedList<double, double> ReadTrace(ParaObject para) { return null; }
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
        protected void Write(string cmd, params object[] paras)
        {
            if (paras != null && paras.Length > 0)
            {
                cmd = string.Format(cmd, paras);
            }
            ses.Write(cmd);
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
