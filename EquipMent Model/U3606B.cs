using System;

namespace WindowsFormsApplication1.EquipMent_Model
{
    class U3606B
    {
        private static Connect Equipment = new Connect();
        public CONFIG CONFIG;
        public MEASURE MEASURE;// = new MEASURE();
        public SOURCE SOURCE;// = new SOURCE();
        public U3606B(string U3606Name)
        {
            String aaa = Equipment.ConnectInstr(U3606Name);
            CONFIG = new CONFIG(Equipment);
            MEASURE = new MEASURE(Equipment);
            SOURCE = new SOURCE(Equipment);
        }
        public string Connect(string ResourceName)
        {
            return Equipment.ConnectInstr(ResourceName);
        }
        public bool Disconnect()
        {
            return Equipment.DisConnectInstr();
        }
        public string RequestID()
        {
            return Equipment.QueryString("*IDN?");
        }
        public string ReadErr()
        {
            return Equipment.QueryString("SYSTem:ERRor?");
        }
        public string CheckConFig()
        {
            return Equipment.QueryString("CONF?\n");
        }
        public bool EnableForntPanel()
        {
            return Equipment.SendSLICCmd(":SYST:LOC\n");
        }
        public int GetTimeOut()
        {
            return Equipment.TimeOut;
        }
        public int SetTimeOut(int time)
        {
            return Equipment.TimeOut = time;
        }
    }

    class CONFIG
    {
        public RESISTANCE RES;// = new RESISTANCE();
        public DC_VOLTAGE DC_VOLT;// = new DC_VOLTAGE();
        public AC_VOLTAGE AC_VOLT;// = new AC_VOLTAGE();
        public DC_CURR DC_CURR;// = new DC_CURR();
        public AC_CURR AC_CURR;// = new AC_CURR();
        public CALCULateFUNCTion CALCULate;// = new CALCULateFUNCTion();
        public FREQ FREQUENCY;// = new FREQ();

        public CONFIG(Connect Equipment)
        {
            RES = new RESISTANCE(Equipment);
            DC_VOLT = new DC_VOLTAGE(Equipment);
            AC_VOLT = new AC_VOLTAGE(Equipment);
            DC_CURR = new DC_CURR(Equipment);
            AC_CURR = new AC_CURR(Equipment);
            CALCULate = new CALCULateFUNCTion(Equipment);
            FREQUENCY = new FREQ(Equipment);
        }
    }
    class RESISTANCE
    {
        private Connect Equipment = new Connect();
        public RESISTANCE(Connect Agu)
        {
            Equipment = Agu;
        }

        #region RESISTANCE 
        public bool RANG_100()
        {
            return Equipment.SendSLICCmd("CONF:RES 100");
        }
        public bool RANG_1K()
        {
            return Equipment.SendSLICCmd("CONF:RES 1K");
        }
        public bool RANG_10K()
        {
            return Equipment.SendSLICCmd("CONF:RES 10K");
        }
        public bool RANG_100K()
        {
            return Equipment.SendSLICCmd("CONF:RES 100K");
        }
        public bool RANG_1M()
        {
            return Equipment.SendSLICCmd("CONF:RES 1M");
        }
        public bool RANG_10M()
        {
            return Equipment.SendSLICCmd("CONF:RES 10M");
        }
        public bool RANG_100M()
        {
            return Equipment.SendSLICCmd("CONF:RES 100M");
        }
        #endregion
    }
    class DC_VOLTAGE
    {
        private Connect Equipment = new Connect();
        public DC_VOLTAGE(Connect Agu)
        {
            Equipment = Agu;
        }
        public string Connect(string ResourceName)
        {
            return Equipment.ConnectInstr(ResourceName);
        }

        #region DC_VOLTAGE        
        public bool RANG_20mV()
        {
            return Equipment.SendSLICCmd("CONF:VOLT:DC 0.02");
        }
        public bool RANG_100mV()
        {
            return Equipment.SendSLICCmd("CONF:VOLT:DC 0.1");
        }
        public bool RANG_1V()
        {
            return Equipment.SendSLICCmd("CONF:VOLT:DC 1");
        }
        public bool RANG_10V()
        {
            return Equipment.SendSLICCmd("CONF:VOLT:DC 10");
        }
        public bool RANG_100V()
        {
            return Equipment.SendSLICCmd("CONF:VOLT:DC 100");
        }
        public bool RANG_1000V()
        {
            return Equipment.SendSLICCmd("CONF:VOLT:DC 1000");
        }
        #endregion
    }
    class DC_CURR
    {
        private Connect Equipment = new Connect();
        public DC_CURR(Connect Agu)
        {
            Equipment = Agu;
        }

        #region DC_CURRENT        
        public bool RANG_10mA()
        {
            return Equipment.SendSLICCmd("CONF:CURR:DC 0.01");
        }
        public bool RANG_100mA()
        {
            return Equipment.SendSLICCmd("CONF:CURR:DC 0.1");
        }
        public bool RANG_1A()
        {
            return Equipment.SendSLICCmd("CONF:CURR:DC 1");
        }
        public bool RANG_3A()
        {
            return Equipment.SendSLICCmd("CONF:CURR:DC 3");
        }
        #endregion
    }
    class AC_VOLTAGE
    {
        private Connect Equipment = new Connect();
        public AC_VOLTAGE(Connect Agu)
        {
            Equipment = Agu;
        }

        #region AC_VOLTAGE
        public bool RANG_100mV()
        {
            return Equipment.SendSLICCmd("CONF:VOLT:AC 0.1");
        }
        public bool RANG_1V()
        {
            return Equipment.SendSLICCmd("CONF:VOLT:AC 1");
        }
        public bool RANG_10V()
        {
            return Equipment.SendSLICCmd("CONF:VOLT:AC 10");
        }
        public bool RANG_100V()
        {
            return Equipment.SendSLICCmd("CONF:VOLT:AC 100");
        }
        public bool RANG_750V()
        {
            return Equipment.SendSLICCmd("CONF:VOLT:AC 750");
        }
        #endregion
    }
    class AC_CURR
    {
        private Connect Equipment = new Connect();
        public AC_CURR(Connect Agu)
        {
            Equipment = Agu;
        }

        #region AC_CURRENT
        public bool RANG_10mA()
        {
            return Equipment.SendSLICCmd("CONF:CURR:AC 0.01");
        }
        public bool RANG_100mA()
        {
            return Equipment.SendSLICCmd("CONF:CURR:AC 0.1");
        }
        public bool RANG_1A()
        {
            return Equipment.SendSLICCmd("CONF:CURR:AC 1");
        }
        public bool RANG_3A()
        {
            return Equipment.SendSLICCmd("CONF:CURR:AC 3");
        }
        #endregion
    }
    class CALCULateFUNCTion
    {
        private Connect Equipment = new Connect();
        public CALCULateFUNCTion(Connect Agu)
        {
            Equipment = Agu;
        }

        #region CALCULateFUNCTion        
        public bool FUNCTionAVERage()
        {
            return Equipment.SendSLICCmd("CALCULate:FUNCTion AVERage");
        }
        public bool CALCULateSTATeOn()
        {
            return Equipment.SendSLICCmd("CALCULate:STATe On");
        }
        #endregion
    }
    class FREQ
    {
        private Connect Equipment = new Connect();
        public FREQ(Connect Agu)
        {
            Equipment = Agu;
        }

        #region FREQ
        public bool FUNCTionFREQ()
        {
            return Equipment.SendSLICCmd("CONF:FREQ");
        }
        #endregion
    }

    class MEASURE
    {
        private Connect Equipment = new Connect();
        public MEASURE(Connect Agu)
        {
            Equipment = Agu;
        }

        public string READ_DC_VOLT()
        {
            return Equipment.QueryString(":MEAS:VOLT:DC?");
        }

        public string READ_DC_CURR()
        {
            return Equipment.QueryString("MEASure:CURRent:DC?");
        }

        public string READ_AC_VOLT()
        {
            return Equipment.QueryString(":MEAS:VOLT:AC?");
        }

        public string READ_AC_CURR()
        {
            //return Equipment.QueryString("MEASure:CURRent:AC?");
            return Equipment.QueryString("READ?");
        }

        public string READ_FREQ()
        {
            return Equipment.QueryString(":MEAS:FREQ?");
        }

        public string READ_RESISTANCE()
        {
            return Equipment.QueryString(":MEAS:RES?");
        }

        public string ReadAVRMaxVal()
        {
            return Equipment.QueryString("CALCulate:AVERage:MAXimum?");
        }

        public string ReadAVRMinVal()
        {
            return Equipment.QueryString("CALCulate:AVERage:MINimum?");
        }
    }

    class SOURCE
    {
        public READ READ;// = new READ();
        public WRITE WRITE;// = new WRITE();
        public SOURCE(Connect Equipment)
        {
            READ = new READ(Equipment);
            WRITE = new WRITE(Equipment);
        }
    }
    class READ
    {
        Connect Equipment = new Connect();
        public READ(Connect Agu)
        {
            Equipment = Agu;
        }
        public string OUTP_STAT()
        {
            return Equipment.QueryString(":OUTP:STAT?");
        }
        public string VOLTAGE()
        {
            return Equipment.QueryString(":SOUR:SENS:VOLT?");
        }
        public string CURRENT()
        {
            return Equipment.QueryString(":SOUR:SENS:CURR?");
        }
    }
    class WRITE
    {
        Connect Equipment = new Connect();
        public WRITE(Connect Agu)
        {
            Equipment = Agu;
        }
        public bool OUTP_ON()
        {
            return Equipment.SendSLICCmd(":OUTP:STAT ON");
        }
        public bool OUTP_OFF()
        {
            return Equipment.SendSLICCmd(":OUTP:STAT OFF");
        }
        public bool OUTP_VOLTAGE(double VoltageLevel)
        {
            string strVoltLevel = ":SOUR:VOLT " + Convert.ToString(VoltageLevel);
            return Equipment.SendSLICCmd(strVoltLevel);
        }
    }
}
