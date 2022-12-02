using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;

using Spire.Xls;
using System.IO.Ports;
using System.Management;
using System.Net.Sockets;
using System.Threading.Tasks;
using System.Globalization;
using System.Reflection;
using System.Linq;
using System.Web.Script.Serialization;
using System.Net;

namespace WindowsFormsApplication1 {
    partial class fMain : Form {
        #region============================================================= Variable =============================================================
        public fMain() {
            setupPay.SelectTab = SetupPay.tabPage.TAB1;
            setupPay.set_nameTab(tester.nameFile);
            setupPay.SelectTab = SetupPay.tabPage.TAB2;
            setupPay.set_nameTab(prismTest.nameFile);
            setupPay.SelectTab = SetupPay.tabPage.TAB3;
            setupPay.set_nameTab(upDataTest.nameFile);
            setupPay.SelectTab = SetupPay.tabPage.TAB4;
            setupPay.set_nameTab(tcpIP.nameFile);
            setupPay.SelectTab = SetupPay.tabPage.TAB5;
            setupPay.set_nameTab(dataLog.nameFile);
            setupPay.setup();
            InitializeComponent();

            functionExcel = Type.GetType(define.formClass);
        }

        public Define define = new Define();
        public Tester tester = new Tester();
        public DataLog dataLog = new DataLog();
        public Excel excel = new Excel();
        public TcpIP tcpIP = new TcpIP();
        public PrismTest prismTest = new PrismTest();
        public UpDataTest upDataTest = new UpDataTest();
        public SetupPay.FormPay setupPay = new SetupPay.FormPay();
        public Automation autoMation = new Automation();

        Type functionExcel;
        public bool[] GlobalTestingFlag = { false, false, false, false, false, false, false, false, false, false,
                                            false, false, false, false, false, false, false, false, false, false,
                                            false, false, false, false, false, false, false, false, false, false,
                                            false, false, false, false, false, false };

        private int[] row = new int[36];
        private string[] time_start = new string[36];
        public bool flag_this_close;
        public bool[] flag_sn_pass = new bool[36];
        public bool[] flagNotReTest = new bool[36];
        
        public string function_timeout(Func<string> function, int timeout) {
            Task<string> task = Task.Run(function);
            if (task.Wait(timeout)) return task.Result;
            else return "over timeout";
        }
        #endregion

        #region================================================== MCP23017 relay cards Cotrol =====================================================
        public static class Matrix {
            public static class Row1 {
                public static readonly int Columns1 = 1;
                public static readonly int Columns2 = 2;
                public static readonly int Columns3 = 3;
                public static readonly int Columns4 = 4;
                public static readonly int Columns5 = 5;
                public static readonly int Columns6 = 6;
                public static readonly int Columns7 = 7;
                public static readonly int Columns8 = 8;
                public static readonly int Columns9 = 9;
                public static readonly int Columns10 = 10;
                public static readonly int Columns11 = 11;
                public static readonly int Columns12 = 12;
                public static readonly int Columns13 = 13;
                public static readonly int Columns14 = 14;
                public static readonly int Columns15 = 15;
                public static readonly int Columns16 = 16;
                public static readonly int Columns17 = 17;
                public static readonly int Columns18 = 18;
                public static readonly int Columns19 = 19;
                public static readonly int Columns20 = 20;
            }
            public static class Row2 {
                public static readonly int Columns1 = 21;
                public static readonly int Columns2 = 22;
                public static readonly int Columns3 = 23;
                public static readonly int Columns4 = 24;
                public static readonly int Columns5 = 25;
                public static readonly int Columns6 = 26;
                public static readonly int Columns7 = 27;
                public static readonly int Columns8 = 28;
                public static readonly int Columns9 = 29;
                public static readonly int Columns10 = 30;
                public static readonly int Columns11 = 31;
                public static readonly int Columns12 = 32;
                public static readonly int Columns13 = 33;
                public static readonly int Columns14 = 34;
                public static readonly int Columns15 = 35;
                public static readonly int Columns16 = 36;
                public static readonly int Columns17 = 37;
                public static readonly int Columns18 = 38;
                public static readonly int Columns19 = 39;
                public static readonly int Columns20 = 40;
            }
            public static class Row3 {
                public static readonly int Columns1 = 41;
                public static readonly int Columns2 = 42;
                public static readonly int Columns3 = 43;
                public static readonly int Columns4 = 44;
                public static readonly int Columns5 = 45;
                public static readonly int Columns6 = 46;
                public static readonly int Columns7 = 47;
                public static readonly int Columns8 = 48;
                public static readonly int Columns9 = 49;
                public static readonly int Columns10 = 50;
                public static readonly int Columns11 = 51;
                public static readonly int Columns12 = 52;
                public static readonly int Columns13 = 53;
                public static readonly int Columns14 = 54;
                public static readonly int Columns15 = 55;
                public static readonly int Columns16 = 56;
                public static readonly int Columns17 = 57;
                public static readonly int Columns18 = 58;
                public static readonly int Columns19 = 59;
                public static readonly int Columns20 = 60;
            }
            public static class Row4 {
                public static readonly int Columns1 = 61;
                public static readonly int Columns2 = 62;
                public static readonly int Columns3 = 63;
                public static readonly int Columns4 = 64;
                public static readonly int Columns5 = 65;
                public static readonly int Columns6 = 66;
                public static readonly int Columns7 = 67;
                public static readonly int Columns8 = 68;
                public static readonly int Columns9 = 69;
                public static readonly int Columns10 = 70;
                public static readonly int Columns11 = 71;
                public static readonly int Columns12 = 72;
                public static readonly int Columns13 = 73;
                public static readonly int Columns14 = 74;
                public static readonly int Columns15 = 75;
                public static readonly int Columns16 = 76;
                public static readonly int Columns17 = 77;
                public static readonly int Columns18 = 78;
                public static readonly int Columns19 = 79;
                public static readonly int Columns20 = 80;
            }
            public static class Multiplex1 {
                public static readonly int bit1 = 101;
                public static readonly int bit2 = 102;
                public static readonly int bit3 = 103;
                public static readonly int bit4 = 104;
                public static readonly int bit5 = 105;
                public static readonly int bit6 = 106;
                public static readonly int bit7 = 107;
                public static readonly int bit8 = 108;
                public static readonly int bit9 = 109;
                public static readonly int bit10 = 110;
                public static readonly int bit11 = 111;
                public static readonly int bit12 = 112;
                public static readonly int bit13 = 113;
                public static readonly int bit14 = 114;
                public static readonly int bit15 = 115;
                public static readonly int bit16 = 116;
            }
            public static class Multiplex2 {
                public static readonly int bit1 = 121;
                public static readonly int bit2 = 122;
                public static readonly int bit3 = 123;
                public static readonly int bit4 = 124;
                public static readonly int bit5 = 125;
                public static readonly int bit6 = 126;
                public static readonly int bit7 = 127;
                public static readonly int bit8 = 128;
                public static readonly int bit9 = 129;
                public static readonly int bit10 = 130;
                public static readonly int bit11 = 131;
                public static readonly int bit12 = 132;
                public static readonly int bit13 = 133;
                public static readonly int bit14 = 134;
                public static readonly int bit15 = 135;
                public static readonly int bit16 = 136;
            }
            public static class Output {
                public static readonly int bit1 = 151;
                public static readonly int bit2 = 152;
                public static readonly int bit3 = 153;
                public static readonly int bit4 = 154;
                public static readonly int bit5 = 155;
                public static readonly int bit6 = 156;
                public static readonly int bit7 = 157;
                public static readonly int bit8 = 158;
                public static readonly int bit9 = 159;
                public static readonly int bit10 = 160;
                public static readonly int bit11 = 161;
                public static readonly int bit12 = 162;
                public static readonly int bit13 = 163;
                public static readonly int bit14 = 164;
                public static readonly int bit15 = 165;
                public static readonly int bit16 = 166;
            }
        }
        public void MatrixOn(int matrix) {
            if (matrix >= 1 && matrix <= 80) {
                int hhh = Matrix.Row1.Columns1;
                switch (matrix) {
                    case 1: Relay_On(1, bit1); break;
                    case 2: Relay_On(1, bit5); break;
                    case 3: Relay_On(1, bit9); break;
                    case 4: Relay_On(1, bit13); break;
                    case 5: Relay_On(2, bit1); break;
                    case 6: Relay_On(2, bit5); break;
                    case 7: Relay_On(2, bit9); break;
                    case 8: Relay_On(2, bit13); break;
                    case 9: Relay_On(3, bit1); break;
                    case 10: Relay_On(3, bit5); break;
                    case 11: Relay_On(3, bit9); break;
                    case 12: Relay_On(3, bit13); break;
                    case 13: Relay_On(4, bit1); break;
                    case 14: Relay_On(4, bit5); break;
                    case 15: Relay_On(4, bit9); break;
                    case 16: Relay_On(4, bit13); break;
                    case 17: Relay_On(5, bit1); break;
                    case 18: Relay_On(5, bit5); break;
                    case 19: Relay_On(5, bit9); break;
                    case 20: Relay_On(5, bit13); break;

                    case 21: Relay_On(1, bit2); break;
                    case 22: Relay_On(1, bit6); break;
                    case 23: Relay_On(1, bit10); break;
                    case 24: Relay_On(1, bit14); break;
                    case 25: Relay_On(2, bit2); break;
                    case 26: Relay_On(2, bit6); break;
                    case 27: Relay_On(2, bit10); break;
                    case 28: Relay_On(2, bit14); break;
                    case 29: Relay_On(3, bit2); break;
                    case 30: Relay_On(3, bit6); break;
                    case 31: Relay_On(3, bit10); break;
                    case 32: Relay_On(3, bit14); break;
                    case 33: Relay_On(4, bit2); break;
                    case 34: Relay_On(4, bit6); break;
                    case 35: Relay_On(4, bit10); break;
                    case 36: Relay_On(4, bit14); break;
                    case 37: Relay_On(5, bit2); break;
                    case 38: Relay_On(5, bit6); break;
                    case 39: Relay_On(5, bit10); break;
                    case 40: Relay_On(5, bit14); break;

                    case 41: Relay_On(1, bit3); break;
                    case 42: Relay_On(1, bit7); break;
                    case 43: Relay_On(1, bit11); break;
                    case 44: Relay_On(1, bit15); break;
                    case 45: Relay_On(2, bit3); break;
                    case 46: Relay_On(2, bit7); break;
                    case 47: Relay_On(2, bit11); break;
                    case 48: Relay_On(2, bit15); break;
                    case 49: Relay_On(3, bit3); break;
                    case 50: Relay_On(3, bit7); break;
                    case 51: Relay_On(3, bit11); break;
                    case 52: Relay_On(3, bit15); break;
                    case 53: Relay_On(4, bit3); break;
                    case 54: Relay_On(4, bit7); break;
                    case 55: Relay_On(4, bit11); break;
                    case 56: Relay_On(4, bit15); break;
                    case 57: Relay_On(5, bit3); break;
                    case 58: Relay_On(5, bit7); break;
                    case 59: Relay_On(5, bit11); break;
                    case 60: Relay_On(5, bit15); break;

                    case 61: Relay_On(1, bit4); break;
                    case 62: Relay_On(1, bit8); break;
                    case 63: Relay_On(1, bit12); break;
                    case 64: Relay_On(1, bit16); break;
                    case 65: Relay_On(2, bit4); break;
                    case 66: Relay_On(2, bit8); break;
                    case 67: Relay_On(2, bit12); break;
                    case 68: Relay_On(2, bit16); break;
                    case 69: Relay_On(3, bit4); break;
                    case 70: Relay_On(3, bit8); break;
                    case 71: Relay_On(3, bit12); break;
                    case 72: Relay_On(3, bit16); break;
                    case 73: Relay_On(4, bit4); break;
                    case 74: Relay_On(4, bit8); break;
                    case 75: Relay_On(4, bit12); break;
                    case 76: Relay_On(4, bit16); break;
                    case 77: Relay_On(5, bit4); break;
                    case 78: Relay_On(5, bit8); break;
                    case 79: Relay_On(5, bit12); break;
                    case 80: Relay_On(5, bit16); break;
                }
                return;
            }
            if (matrix >= 100 && matrix <= 140) {
                switch (matrix) {
                    case 101: Relay_On(6, bit1); break;
                    case 102: Relay_On(6, bit2); break;
                    case 103: Relay_On(6, bit3); break;
                    case 104: Relay_On(6, bit4); break;
                    case 105: Relay_On(6, bit5); break;
                    case 106: Relay_On(6, bit6); break;
                    case 107: Relay_On(6, bit7); break;
                    case 108: Relay_On(6, bit8); break;
                    case 109: Relay_On(6, bit9); break;
                    case 110: Relay_On(6, bit10); break;
                    case 111: Relay_On(6, bit11); break;
                    case 112: Relay_On(6, bit12); break;
                    case 113: Relay_On(6, bit13); break;
                    case 114: Relay_On(6, bit14); break;
                    case 115: Relay_On(6, bit15); break;
                    case 116: Relay_On(6, bit16); break;

                    case 121: Relay_On(7, bit1); break;
                    case 122: Relay_On(7, bit2); break;
                    case 123: Relay_On(7, bit3); break;
                    case 124: Relay_On(7, bit4); break;
                    case 125: Relay_On(7, bit5); break;
                    case 126: Relay_On(7, bit6); break;
                    case 127: Relay_On(7, bit7); break;
                    case 128: Relay_On(7, bit8); break;
                    case 129: Relay_On(7, bit9); break;
                    case 130: Relay_On(7, bit10); break;
                    case 131: Relay_On(7, bit11); break;
                    case 132: Relay_On(7, bit12); break;
                    case 133: Relay_On(7, bit13); break;
                    case 134: Relay_On(7, bit14); break;
                    case 135: Relay_On(7, bit15); break;
                    case 136: Relay_On(7, bit16); break;
                }
                return;
            }
            if (matrix >= 140 && matrix <= 170) {
                switch (matrix) {
                    case 151: Relay_On(8, bit1); break;
                    case 152: Relay_On(8, bit2); break;
                    case 153: Relay_On(8, bit3); break;
                    case 154: Relay_On(8, bit4); break;
                    case 155: Relay_On(8, bit5); break;
                    case 156: Relay_On(8, bit6); break;
                    case 157: Relay_On(8, bit7); break;
                    case 158: Relay_On(8, bit8); break;
                    case 159: Relay_On(8, bit9); break;
                    case 160: Relay_On(8, bit10); break;
                    case 161: Relay_On(8, bit11); break;
                    case 162: Relay_On(8, bit12); break;
                    case 163: Relay_On(8, bit13); break;
                    case 164: Relay_On(8, bit14); break;
                    case 165: Relay_On(8, bit15); break;
                    case 166: Relay_On(8, bit16); break;
                }
                return;
            }
        }
        public void MatrixOff(int matrix) {
            if (matrix >= 1 && matrix <= 80) {
                int hhh = Matrix.Row1.Columns1;
                switch (matrix) {
                    case 1: Relay_Off(1, bit1); break;
                    case 2: Relay_Off(1, bit5); break;
                    case 3: Relay_Off(1, bit9); break;
                    case 4: Relay_Off(1, bit13); break;
                    case 5: Relay_Off(2, bit1); break;
                    case 6: Relay_Off(2, bit5); break;
                    case 7: Relay_Off(2, bit9); break;
                    case 8: Relay_Off(2, bit13); break;
                    case 9: Relay_Off(3, bit1); break;
                    case 10: Relay_Off(3, bit5); break;
                    case 11: Relay_Off(3, bit9); break;
                    case 12: Relay_Off(3, bit13); break;
                    case 13: Relay_Off(4, bit1); break;
                    case 14: Relay_Off(4, bit5); break;
                    case 15: Relay_Off(4, bit9); break;
                    case 16: Relay_Off(4, bit13); break;
                    case 17: Relay_Off(5, bit1); break;
                    case 18: Relay_Off(5, bit5); break;
                    case 19: Relay_Off(5, bit9); break;
                    case 20: Relay_Off(5, bit13); break;

                    case 21: Relay_Off(1, bit2); break;
                    case 22: Relay_Off(1, bit6); break;
                    case 23: Relay_Off(1, bit10); break;
                    case 24: Relay_Off(1, bit14); break;
                    case 25: Relay_Off(2, bit2); break;
                    case 26: Relay_Off(2, bit6); break;
                    case 27: Relay_Off(2, bit10); break;
                    case 28: Relay_Off(2, bit14); break;
                    case 29: Relay_Off(3, bit2); break;
                    case 30: Relay_Off(3, bit6); break;
                    case 31: Relay_Off(3, bit10); break;
                    case 32: Relay_Off(3, bit14); break;
                    case 33: Relay_Off(4, bit2); break;
                    case 34: Relay_Off(4, bit6); break;
                    case 35: Relay_Off(4, bit10); break;
                    case 36: Relay_Off(4, bit14); break;
                    case 37: Relay_Off(5, bit2); break;
                    case 38: Relay_Off(5, bit6); break;
                    case 39: Relay_Off(5, bit10); break;
                    case 40: Relay_Off(5, bit14); break;

                    case 41: Relay_Off(1, bit3); break;
                    case 42: Relay_Off(1, bit7); break;
                    case 43: Relay_Off(1, bit11); break;
                    case 44: Relay_Off(1, bit15); break;
                    case 45: Relay_Off(2, bit3); break;
                    case 46: Relay_Off(2, bit7); break;
                    case 47: Relay_Off(2, bit11); break;
                    case 48: Relay_Off(2, bit15); break;
                    case 49: Relay_Off(3, bit3); break;
                    case 50: Relay_Off(3, bit7); break;
                    case 51: Relay_Off(3, bit11); break;
                    case 52: Relay_Off(3, bit15); break;
                    case 53: Relay_Off(4, bit3); break;
                    case 54: Relay_Off(4, bit7); break;
                    case 55: Relay_Off(4, bit11); break;
                    case 56: Relay_Off(4, bit15); break;
                    case 57: Relay_Off(5, bit3); break;
                    case 58: Relay_Off(5, bit7); break;
                    case 59: Relay_Off(5, bit11); break;
                    case 60: Relay_Off(5, bit15); break;

                    case 61: Relay_Off(1, bit4); break;
                    case 62: Relay_Off(1, bit8); break;
                    case 63: Relay_Off(1, bit12); break;
                    case 64: Relay_Off(1, bit16); break;
                    case 65: Relay_Off(2, bit4); break;
                    case 66: Relay_Off(2, bit8); break;
                    case 67: Relay_Off(2, bit12); break;
                    case 68: Relay_Off(2, bit16); break;
                    case 69: Relay_Off(3, bit4); break;
                    case 70: Relay_Off(3, bit8); break;
                    case 71: Relay_Off(3, bit12); break;
                    case 72: Relay_Off(3, bit16); break;
                    case 73: Relay_Off(4, bit4); break;
                    case 74: Relay_Off(4, bit8); break;
                    case 75: Relay_Off(4, bit12); break;
                    case 76: Relay_Off(4, bit16); break;
                    case 77: Relay_Off(5, bit4); break;
                    case 78: Relay_Off(5, bit8); break;
                    case 79: Relay_Off(5, bit12); break;
                    case 80: Relay_Off(5, bit16); break;
                }
                return;
            }
            if (matrix >= 100 && matrix <= 140) {
                switch (matrix) {
                    case 101: Relay_Off(6, bit1); break;
                    case 102: Relay_Off(6, bit2); break;
                    case 103: Relay_Off(6, bit3); break;
                    case 104: Relay_Off(6, bit4); break;
                    case 105: Relay_Off(6, bit5); break;
                    case 106: Relay_Off(6, bit6); break;
                    case 107: Relay_Off(6, bit7); break;
                    case 108: Relay_Off(6, bit8); break;
                    case 109: Relay_Off(6, bit9); break;
                    case 110: Relay_Off(6, bit10); break;
                    case 111: Relay_Off(6, bit11); break;
                    case 112: Relay_Off(6, bit12); break;
                    case 113: Relay_Off(6, bit13); break;
                    case 114: Relay_Off(6, bit14); break;
                    case 115: Relay_Off(6, bit15); break;
                    case 116: Relay_Off(6, bit16); break;

                    case 121: Relay_Off(7, bit1); break;
                    case 122: Relay_Off(7, bit2); break;
                    case 123: Relay_Off(7, bit3); break;
                    case 124: Relay_Off(7, bit4); break;
                    case 125: Relay_Off(7, bit5); break;
                    case 126: Relay_Off(7, bit6); break;
                    case 127: Relay_Off(7, bit7); break;
                    case 128: Relay_Off(7, bit8); break;
                    case 129: Relay_Off(7, bit9); break;
                    case 130: Relay_Off(7, bit10); break;
                    case 131: Relay_Off(7, bit11); break;
                    case 132: Relay_Off(7, bit12); break;
                    case 133: Relay_Off(7, bit13); break;
                    case 134: Relay_Off(7, bit14); break;
                    case 135: Relay_Off(7, bit15); break;
                    case 136: Relay_Off(7, bit16); break;
                }
                return;
            }
            if (matrix >= 140 && matrix <= 170) {
                switch (matrix) {
                    case 151: Relay_Off(8, bit1); break;
                    case 152: Relay_Off(8, bit2); break;
                    case 153: Relay_Off(8, bit3); break;
                    case 154: Relay_Off(8, bit4); break;
                    case 155: Relay_Off(8, bit5); break;
                    case 156: Relay_Off(8, bit6); break;
                    case 157: Relay_Off(8, bit7); break;
                    case 158: Relay_Off(8, bit8); break;
                    case 159: Relay_Off(8, bit9); break;
                    case 160: Relay_Off(8, bit10); break;
                    case 161: Relay_Off(8, bit11); break;
                    case 162: Relay_Off(8, bit12); break;
                    case 163: Relay_Off(8, bit13); break;
                    case 164: Relay_Off(8, bit14); break;
                    case 165: Relay_Off(8, bit15); break;
                    case 166: Relay_Off(8, bit16); break;
                }
                return;
            }
        }

        private void Matrix_On(string btn) {
            switch (btn) {
                case "R1C01": Relay_On(1, bit1); break;
                case "R1C02": Relay_On(1, bit5); break;
                case "R1C03": Relay_On(1, bit9); break;
                case "R1C04": Relay_On(1, bit13); break;
                case "R1C05": Relay_On(2, bit1); break;
                case "R1C06": Relay_On(2, bit5); break;
                case "R1C07": Relay_On(2, bit9); break;
                case "R1C08": Relay_On(2, bit13); break;
                case "R1C09": Relay_On(3, bit1); break;
                case "R1C10": Relay_On(3, bit5); break;
                case "R1C11": Relay_On(3, bit9); break;
                case "R1C12": Relay_On(3, bit13); break;
                case "R1C13": Relay_On(4, bit1); break;
                case "R1C14": Relay_On(4, bit5); break;
                case "R1C15": Relay_On(4, bit9); break;
                case "R1C16": Relay_On(4, bit13); break;
                case "R1C17": Relay_On(5, bit1); break;
                case "R1C18": Relay_On(5, bit5); break;
                case "R1C19": Relay_On(5, bit9); break;
                case "R1C20": Relay_On(5, bit13); break;

                case "R2C01": Relay_On(1, bit2); break;
                case "R2C02": Relay_On(1, bit6); break;
                case "R2C03": Relay_On(1, bit10); break;
                case "R2C04": Relay_On(1, bit14); break;
                case "R2C05": Relay_On(2, bit2); break;
                case "R2C06": Relay_On(2, bit6); break;
                case "R2C07": Relay_On(2, bit10); break;
                case "R2C08": Relay_On(2, bit14); break;
                case "R2C09": Relay_On(3, bit2); break;
                case "R2C10": Relay_On(3, bit6); break;
                case "R2C11": Relay_On(3, bit10); break;
                case "R2C12": Relay_On(3, bit14); break;
                case "R2C13": Relay_On(4, bit2); break;
                case "R2C14": Relay_On(4, bit6); break;
                case "R2C15": Relay_On(4, bit10); break;
                case "R2C16": Relay_On(4, bit14); break;
                case "R2C17": Relay_On(5, bit2); break;
                case "R2C18": Relay_On(5, bit6); break;
                case "R2C19": Relay_On(5, bit10); break;
                case "R2C20": Relay_On(5, bit14); break;

                case "R3C01": Relay_On(1, bit3); break;
                case "R3C02": Relay_On(1, bit7); break;
                case "R3C03": Relay_On(1, bit11); break;
                case "R3C04": Relay_On(1, bit15); break;
                case "R3C05": Relay_On(2, bit3); break;
                case "R3C06": Relay_On(2, bit7); break;
                case "R3C07": Relay_On(2, bit11); break;
                case "R3C08": Relay_On(2, bit15); break;
                case "R3C09": Relay_On(3, bit3); break;
                case "R3C10": Relay_On(3, bit7); break;
                case "R3C11": Relay_On(3, bit11); break;
                case "R3C12": Relay_On(3, bit15); break;
                case "R3C13": Relay_On(4, bit3); break;
                case "R3C14": Relay_On(4, bit7); break;
                case "R3C15": Relay_On(4, bit11); break;
                case "R3C16": Relay_On(4, bit15); break;
                case "R3C17": Relay_On(5, bit3); break;
                case "R3C18": Relay_On(5, bit7); break;
                case "R3C19": Relay_On(5, bit11); break;
                case "R3C20": Relay_On(5, bit15); break;

                case "R4C01": Relay_On(1, bit4); break;
                case "R4C02": Relay_On(1, bit8); break;
                case "R4C03": Relay_On(1, bit12); break;
                case "R4C04": Relay_On(1, bit16); break;
                case "R4C05": Relay_On(2, bit4); break;
                case "R4C06": Relay_On(2, bit8); break;
                case "R4C07": Relay_On(2, bit12); break;
                case "R4C08": Relay_On(2, bit16); break;
                case "R4C09": Relay_On(3, bit4); break;
                case "R4C10": Relay_On(3, bit8); break;
                case "R4C11": Relay_On(3, bit12); break;
                case "R4C12": Relay_On(3, bit16); break;
                case "R4C13": Relay_On(4, bit4); break;
                case "R4C14": Relay_On(4, bit8); break;
                case "R4C15": Relay_On(4, bit12); break;
                case "R4C16": Relay_On(4, bit16); break;
                case "R4C17": Relay_On(5, bit4); break;
                case "R4C18": Relay_On(5, bit8); break;
                case "R4C19": Relay_On(5, bit12); break;
                case "R4C20": Relay_On(5, bit16); break;
            }
        }
        private void Matrix_Off(string btn) {
            switch (btn) {
                case "R1C01": Relay_Off(1, bit1); break;
                case "R1C02": Relay_Off(1, bit5); break;
                case "R1C03": Relay_Off(1, bit9); break;
                case "R1C04": Relay_Off(1, bit13); break;
                case "R1C05": Relay_Off(2, bit1); break;
                case "R1C06": Relay_Off(2, bit5); break;
                case "R1C07": Relay_Off(2, bit9); break;
                case "R1C08": Relay_Off(2, bit13); break;
                case "R1C09": Relay_Off(3, bit1); break;
                case "R1C10": Relay_Off(3, bit5); break;
                case "R1C11": Relay_Off(3, bit9); break;
                case "R1C12": Relay_Off(3, bit13); break;
                case "R1C13": Relay_Off(4, bit1); break;
                case "R1C14": Relay_Off(4, bit5); break;
                case "R1C15": Relay_Off(4, bit9); break;
                case "R1C16": Relay_Off(4, bit13); break;
                case "R1C17": Relay_Off(5, bit1); break;
                case "R1C18": Relay_Off(5, bit5); break;
                case "R1C19": Relay_Off(5, bit9); break;
                case "R1C20": Relay_Off(5, bit13); break;

                case "R2C01": Relay_Off(1, bit2); break;
                case "R2C02": Relay_Off(1, bit6); break;
                case "R2C03": Relay_Off(1, bit10); break;
                case "R2C04": Relay_Off(1, bit14); break;
                case "R2C05": Relay_Off(2, bit2); break;
                case "R2C06": Relay_Off(2, bit6); break;
                case "R2C07": Relay_Off(2, bit10); break;
                case "R2C08": Relay_Off(2, bit14); break;
                case "R2C09": Relay_Off(3, bit2); break;
                case "R2C10": Relay_Off(3, bit6); break;
                case "R2C11": Relay_Off(3, bit10); break;
                case "R2C12": Relay_Off(3, bit14); break;
                case "R2C13": Relay_Off(4, bit2); break;
                case "R2C14": Relay_Off(4, bit6); break;
                case "R2C15": Relay_Off(4, bit10); break;
                case "R2C16": Relay_Off(4, bit14); break;
                case "R2C17": Relay_Off(5, bit2); break;
                case "R2C18": Relay_Off(5, bit6); break;
                case "R2C19": Relay_Off(5, bit10); break;
                case "R2C20": Relay_Off(5, bit14); break;

                case "R3C01": Relay_Off(1, bit3); break;
                case "R3C02": Relay_Off(1, bit7); break;
                case "R3C03": Relay_Off(1, bit11); break;
                case "R3C04": Relay_Off(1, bit15); break;
                case "R3C05": Relay_Off(2, bit3); break;
                case "R3C06": Relay_Off(2, bit7); break;
                case "R3C07": Relay_Off(2, bit11); break;
                case "R3C08": Relay_Off(2, bit15); break;
                case "R3C09": Relay_Off(3, bit3); break;
                case "R3C10": Relay_Off(3, bit7); break;
                case "R3C11": Relay_Off(3, bit11); break;
                case "R3C12": Relay_Off(3, bit15); break;
                case "R3C13": Relay_Off(4, bit3); break;
                case "R3C14": Relay_Off(4, bit7); break;
                case "R3C15": Relay_Off(4, bit11); break;
                case "R3C16": Relay_Off(4, bit15); break;
                case "R3C17": Relay_Off(5, bit3); break;
                case "R3C18": Relay_Off(5, bit7); break;
                case "R3C19": Relay_Off(5, bit11); break;
                case "R3C20": Relay_Off(5, bit15); break;

                case "R4C01": Relay_Off(1, bit4); break;
                case "R4C02": Relay_Off(1, bit8); break;
                case "R4C03": Relay_Off(1, bit12); break;
                case "R4C04": Relay_Off(1, bit16); break;
                case "R4C05": Relay_Off(2, bit4); break;
                case "R4C06": Relay_Off(2, bit8); break;
                case "R4C07": Relay_Off(2, bit12); break;
                case "R4C08": Relay_Off(2, bit16); break;
                case "R4C09": Relay_Off(3, bit4); break;
                case "R4C10": Relay_Off(3, bit8); break;
                case "R4C11": Relay_Off(3, bit12); break;
                case "R4C12": Relay_Off(3, bit16); break;
                case "R4C13": Relay_Off(4, bit4); break;
                case "R4C14": Relay_Off(4, bit8); break;
                case "R4C15": Relay_Off(4, bit12); break;
                case "R4C16": Relay_Off(4, bit16); break;
                case "R4C17": Relay_Off(5, bit4); break;
                case "R4C18": Relay_Off(5, bit8); break;
                case "R4C19": Relay_Off(5, bit12); break;
                case "R4C20": Relay_Off(5, bit16); break;
            }
        }

        private void Multiplex_On(string btn) {
            switch (btn) {
                case "M1_01": Relay_On(6, bit1); break;
                case "M1_02": Relay_On(6, bit2); break;
                case "M1_03": Relay_On(6, bit3); break;
                case "M1_04": Relay_On(6, bit4); break;
                case "M1_05": Relay_On(6, bit5); break;
                case "M1_06": Relay_On(6, bit6); break;
                case "M1_07": Relay_On(6, bit7); break;
                case "M1_08": Relay_On(6, bit8); break;
                case "M1_09": Relay_On(6, bit9); break;
                case "M1_10": Relay_On(6, bit10); break;
                case "M1_11": Relay_On(6, bit11); break;
                case "M1_12": Relay_On(6, bit12); break;
                case "M1_13": Relay_On(6, bit13); break;
                case "M1_14": Relay_On(6, bit14); break;
                case "M1_15": Relay_On(6, bit15); break;
                case "M1_16": Relay_On(6, bit16); break;

                case "M2_01": Relay_On(7, bit1); break;
                case "M2_02": Relay_On(7, bit2); break;
                case "M2_03": Relay_On(7, bit3); break;
                case "M2_04": Relay_On(7, bit4); break;
                case "M2_05": Relay_On(7, bit5); break;
                case "M2_06": Relay_On(7, bit6); break;
                case "M2_07": Relay_On(7, bit7); break;
                case "M2_08": Relay_On(7, bit8); break;
                case "M2_09": Relay_On(7, bit9); break;
                case "M2_10": Relay_On(7, bit10); break;
                case "M2_11": Relay_On(7, bit11); break;
                case "M2_12": Relay_On(7, bit12); break;
                case "M2_13": Relay_On(7, bit13); break;
                case "M2_14": Relay_On(7, bit14); break;
                case "M2_15": Relay_On(7, bit15); break;
                case "M2_16": Relay_On(7, bit16); break;
            }
        }
        private void Multiplex_Off(string btn) {
            switch (btn) {
                case "M1_01": Relay_Off(6, bit1); break;
                case "M1_02": Relay_Off(6, bit2); break;
                case "M1_03": Relay_Off(6, bit3); break;
                case "M1_04": Relay_Off(6, bit4); break;
                case "M1_05": Relay_Off(6, bit5); break;
                case "M1_06": Relay_Off(6, bit6); break;
                case "M1_07": Relay_Off(6, bit7); break;
                case "M1_08": Relay_Off(6, bit8); break;
                case "M1_09": Relay_Off(6, bit9); break;
                case "M1_10": Relay_Off(6, bit10); break;
                case "M1_11": Relay_Off(6, bit11); break;
                case "M1_12": Relay_Off(6, bit12); break;
                case "M1_13": Relay_Off(6, bit13); break;
                case "M1_14": Relay_Off(6, bit14); break;
                case "M1_15": Relay_Off(6, bit15); break;
                case "M1_16": Relay_Off(6, bit16); break;

                case "M2_01": Relay_Off(7, bit1); break;
                case "M2_02": Relay_Off(7, bit2); break;
                case "M2_03": Relay_Off(7, bit3); break;
                case "M2_04": Relay_Off(7, bit4); break;
                case "M2_05": Relay_Off(7, bit5); break;
                case "M2_06": Relay_Off(7, bit6); break;
                case "M2_07": Relay_Off(7, bit7); break;
                case "M2_08": Relay_Off(7, bit8); break;
                case "M2_09": Relay_Off(7, bit9); break;
                case "M2_10": Relay_Off(7, bit10); break;
                case "M2_11": Relay_Off(7, bit11); break;
                case "M2_12": Relay_Off(7, bit12); break;
                case "M2_13": Relay_Off(7, bit13); break;
                case "M2_14": Relay_Off(7, bit14); break;
                case "M2_15": Relay_Off(7, bit15); break;
                case "M2_16": Relay_Off(7, bit16); break;
            }
        }

        private void Output_On(string btn) {
            switch (btn) {
                case "OP01": Relay_On(8, bit1); break;
                case "OP02": Relay_On(8, bit2); break;
                case "OP03": Relay_On(8, bit3); break;
                case "OP04": Relay_On(8, bit4); break;
                case "OP05": Relay_On(8, bit5); break;
                case "OP06": Relay_On(8, bit6); break;
                case "OP07": Relay_On(8, bit7); break;
                case "OP08": Relay_On(8, bit8); break;
                case "OP09": Relay_On(8, bit9); break;
                case "OP10": Relay_On(8, bit10); break;
                case "OP11": Relay_On(8, bit11); break;
                case "OP12": Relay_On(8, bit12); break;
                case "OP13": Relay_On(8, bit13); break;
                case "OP14": Relay_On(8, bit14); break;
                case "OP15": Relay_On(8, bit15); break;
                case "OP16": Relay_On(8, bit16); break;
            }
        }
        private void Output_Off(string btn) {
            switch (btn) {
                case "OP01": Relay_Off(8, bit1); break;
                case "OP02": Relay_Off(8, bit2); break;
                case "OP03": Relay_Off(8, bit3); break;
                case "OP04": Relay_Off(8, bit4); break;
                case "OP05": Relay_Off(8, bit5); break;
                case "OP06": Relay_Off(8, bit6); break;
                case "OP07": Relay_Off(8, bit7); break;
                case "OP08": Relay_Off(8, bit8); break;
                case "OP09": Relay_Off(8, bit9); break;
                case "OP10": Relay_Off(8, bit10); break;
                case "OP11": Relay_Off(8, bit11); break;
                case "OP12": Relay_Off(8, bit12); break;
                case "OP13": Relay_Off(8, bit13); break;
                case "OP14": Relay_Off(8, bit14); break;
                case "OP15": Relay_Off(8, bit15); break;
                case "OP16": Relay_Off(8, bit16); break;
            }
        }

        public SerialPort RelayPort = new SerialPort();
        private int[] arduino_input = { 2, 4, 7, 8 };
        public const int bit1 = 1;
        public const int bit2 = 2;
        public const int bit3 = 3;
        public const int bit4 = 4;
        public const int bit5 = 5;
        public const int bit6 = 6;
        public const int bit7 = 7;
        public const int bit8 = 8;
        public const int bit9 = 9;
        public const int bit10 = 10;
        public const int bit11 = 11;
        public const int bit12 = 12;
        public const int bit13 = 13;
        public const int bit14 = 14;
        public const int bit15 = 15;
        public const int bit16 = 16;

        private int[] bitcal = { 128, 64, 32, 16, 8, 4, 2, 1, 32768, 16384, 8192, 4096, 2048, 1024, 512, 256 };
        public bool flag_write_23017 = true;
        public string write_23017(string cmd) {
            while (!flag_write_23017) DelaymS(50);
            return write_23017_sup(cmd);
        }
        private string log_cmd_control_arduino = "";
        public string write_23017_sup(string cmd) {
            label_write_23017_sup:
            if (flag_this_close) return "";
            try {
                RelayPort.DiscardInBuffer();
                RelayPort.DiscardOutBuffer();
                RelayPort.Write(cmd);
            } catch { }
            string rx = "";
            Stopwatch s = new Stopwatch();
            s.Restart();
            while (s.ElapsedMilliseconds < 2500) {
                try { rx = RelayPort.ReadExisting(); } catch { }
                if (rx != "") { s.Stop(); break; }
                DelaymS(50);
            }
            if (s.IsRunning) {
                try {
                    RelayPort.DtrEnable = true;
                    RelayPort.RtsEnable = true;
                } catch { }
                DelaymS(5000);
                try {
                    RelayPort.DtrEnable = false;
                    RelayPort.RtsEnable = false;
                } catch { }
                Log(LogMsgType.Error_Red, "\nreset arduino");
                try { RelayPort.Close(); } catch { }
                try { RelayPort.Open(); } catch { }
                goto label_write_23017_sup;
                //return "";
            }
            string sup_rx = "";
            s.Restart();
            while (s.ElapsedMilliseconds < 100) {
                try { sup_rx = RelayPort.ReadExisting(); } catch { }
                if (sup_rx != "") { rx += sup_rx; s.Restart(); continue; }
                DelaymS(10);
            }
            log_cmd_control_arduino = cmd;
            return rx;
        }
        public void connect_relay(bool initial = true) {
            if (!tester.useRelayCard) return; 
            bt_relayCard.BackColor = Color.Red;
            try { RelayPort.Close(); } catch { }
            string[] arrport;
            string portX = "";
            Label lll = new Label();
            lll.Text = "Find prot arduino...";
            lll.Size = new Size(350, 75);
            FontFamily fontFamily = new FontFamily("Arial");
            lll.Font = new Font(fontFamily, 30, FontStyle.Bold, GraphicsUnit.Pixel);
            Form fff = new Form();
            fff.Size = new Size(400, 100);
            fff.ControlBox = false;
            fff.StartPosition = FormStartPosition.CenterScreen;
            fff.Controls.Add(lll);
            fff.Show();
            Stopwatch s = new Stopwatch();

            string nbn = setupPay.read_text(tester.headConfig.arduinoComport, tester.nameFile);
            try { RelayPort = new SerialPort(nbn); } catch { goto connect_relay_lable_1; }
            s.Restart();
            while (s.ElapsedMilliseconds < 2500) {
                try { RelayPort.Close(); } catch { }
                try { RelayPort.Open(); } catch { DelaymS(50); continue; }
                DelaymS(500);
                s.Stop();
                break;
            }
            if (s.IsRunning) { try { RelayPort.Close(); } catch { } goto connect_relay_lable_1; }
            try {
                RelayPort.DiscardInBuffer();
                RelayPort.DiscardOutBuffer();
                RelayPort.Write("23017,CHECKPORT\n");
            } catch { }
            string rxxx = "";
            s.Restart();
            while (s.ElapsedMilliseconds < 2500) {
                DelaymS(1000);
                try { rxxx = RelayPort.ReadExisting(); } catch { }
                if (rxxx != "") { s.Stop(); break; }
            }
            if (s.IsRunning) { RelayPort.Close(); goto connect_relay_lable_1; }
            if (rxxx.Contains("RELAY_PORT_BY_DESIGN")) {
                bt_relayCard.BackColor = Color.LimeGreen;
                Log(LogMsgType.Incoming_Blue, "\n\nConnect Relay Port to " + nbn);
                s.Stop();
                goto connect_relay_lable_2;
            } else { try { RelayPort.Close(); } catch { } }
            connect_relay_lable_1:

            bool flag_arduino = false;
            ManagementObjectSearcher objOSDetails2 = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE Caption like '%(COM%'");
            ManagementObjectCollection osDetailsCollection2 = objOSDetails2.Get();
            foreach (ManagementObject usblist in osDetailsCollection2) {
                if (usblist["Description"].ToString() != "USB-SERIAL CH340" && 
                    usblist["Description"].ToString() != "USB Serial Port" &&
                    usblist["Description"].ToString() != "USB Serial Device" &&
                    usblist["Description"].ToString() != "Arduino Mega 2560") continue;
                arrport = usblist.GetPropertyValue("NAME").ToString().Split('(', ')');
                try { portX = arrport[1]; } catch { Log(LogMsgType.Error_Red, "\nname port not format " + usblist.GetPropertyValue("NAME").ToString()); return; }
                RelayPort = new SerialPort(portX);
                s.Restart();
                while (s.ElapsedMilliseconds < 2500) {
                    try { RelayPort.Close(); } catch { }
                    try { RelayPort.Open(); } catch { DelaymS(50); continue; }
                    DelaymS(500);
                    s.Stop();
                    break;
                }
                if (s.IsRunning) { try { RelayPort.Close(); } catch { } continue; }
                try {
                    RelayPort.DiscardInBuffer();
                    RelayPort.DiscardOutBuffer();
                    RelayPort.Write("23017,CHECKPORT\n");
                } catch { }
                string rx = "";
                s.Restart();
                while (s.ElapsedMilliseconds < 2500) {
                    DelaymS(1000);
                    try { rx = RelayPort.ReadExisting(); } catch { }
                    if (rx != "") { s.Stop(); break; }
                }
                if (s.IsRunning) { RelayPort.Close(); continue; }
                if (rx.Contains("RELAY_PORT_BY_DESIGN")) {
                    bt_relayCard.BackColor = Color.LimeGreen;
                    Log(LogMsgType.Incoming_Blue, "\n\nConnect Relay Port to " + portX);
                    setupPay.write_text(tester.headConfig.arduinoComport, portX, tester.nameFile);
                    s.Stop();
                    flag_arduino = true;
                    break;
                } else { try { RelayPort.Close(); } catch { } continue; }
            }
            if (!flag_arduino) { Log(LogMsgType.Error_Red, "\nCannot connect to Realay Card!!"); fff.Close(); return; }
            connect_relay_lable_2:
            fff.Close();
            if (!RelayPort.IsOpen) { Log(LogMsgType.Error_Red, "\n\"USB-SERIAL CH340\" not have in device"); return; }
            if (initial) write_23017("23017,INITIAL\n");
            string[] buffer = write_23017("23017,SCANI2C\n").Replace("\r\n", "|").Split('|');
            if (buffer.Length != 2) { Log(LogMsgType.Error_Red, "\ncmd \"23017,SCANI2C\" retrun not format"); return; }
            List<string> addr = new List<string>(buffer[1].Split(','));
            addr.RemoveAll(item => item == "");
            List<string> addrarray = new List<string>{ "20", "21", "22", "23", "24", "25", "26", "27" };
            for (int i = 1; i <= 8; i++) {
                bool f = false;
                foreach (string addrX in addr) {
                    if (addrarray.Contains(addrX)) {
                        addr.Remove(addrX);
                        f = true;
                        break;
                    }
                }
                if (!f) Log(LogMsgType.Error_Red, "\ncannot find card " + (i));
                if (i <= Convert.ToInt32(tester.numCardRelay) && !f) { MessageBox.Show("_ใส่การ์ดรีเลย์ไม่ครบ"); break; }
            }
            OffAllRelay();
        }
        public void RelayControl(string card, string bit, string cmd) {
            string[] result = { "" };
            for (int i = 1; i <= 3; i++) {
                result = write_23017("23017," + card + "," + bit + "," + cmd + "\n").Replace("\r\n", "|").Split('|');
                try { string hh = result[2]; } catch { continue; }
                if (result[2] == "65535") continue;//65535
                break;
            }
        }
        public void RelayControl_sup(string card, string bit, string cmd) {
            string[] result = { "" };
            for (int i = 1; i <= 3; i++) {
                result = write_23017_sup("23017," + card + "," + bit + "," + cmd + "\n").Replace("\r\n", "|").Split('|');
                try { string hh = result[2]; } catch { continue; }
                if (result[2] == "65535") continue;//65535
                break;
            }
        }
        public void Relay_On(int card, int bit) {
            RelayControl(card.ToString(), bit.ToString(), "1"); DelaymS(50);
        }
        public void Relay_On_sup(int card, int bit) {
            RelayControl_sup(card.ToString(), bit.ToString(), "1"); DelaymS(50);
        }
        public void Relay_Off(int card, int bit) {
            RelayControl(card.ToString(), bit.ToString(), "0"); DelaymS(50);
        }
        public void Relay_Off_sup(int card, int bit) {
            RelayControl_sup(card.ToString(), bit.ToString(), "0"); DelaymS(50);
        }
        public void OffAllRelay() {
            if (!tester.useRelayCard) return;
            string[] result = write_23017("23017,OFFALLRELAY\n").Replace("\r\n", "|").Split('|');
            Log(LogMsgType.Incoming_Blue, "\n" + result[0].Replace("\n", ""));
            write_23017("rgbset,0,0,0,0,0\n");
        }
        public void OffAllCard(int card)
        {
            if (!tester.useRelayCard) return;
            string[] result = write_23017("23017,OFFALLRELAY," + card + "\n").Replace("\r\n", "|").Split('|');
            Log(LogMsgType.Incoming_Blue, "\n" + result[0].Replace("\n", ""));
        }
        #endregion

        #region====================================================== region Display_Message ======================================================
        private Color[] LogMsgTypeColor = { Color.Blue, Color.Green, Color.Black, Color.Orange, Color.Red };
        public enum LogMsgType { Incoming_Blue, Outgoing_Green, Normal_Black, Warning_Orange, Error_Red };
        public void Log(LogMsgType msgtype, string msg) {
            try {
                rtfTerminal.Invoke(new EventHandler(delegate {
                    rtfTerminal.SelectedText = string.Empty;
                    rtfTerminal.SelectionFont = new Font(rtfTerminal.SelectionFont, FontStyle.Bold);
                    rtfTerminal.SelectionColor = LogMsgTypeColor[(int)msgtype];
                    rtfTerminal.AppendText(msg);
                    rtfTerminal.ScrollToCaret();

                    if (rtfTerminal.TextLength > 50000) {
                        rtfTerminal.Text = string.Empty;
                    }
                }));
            } catch (Exception) { }
        }
        #endregion

        #region ========================================================= Common Function ==========================================================
        private void InitBeforeTest(int head) {
            DataGridView g = getDataGridView(head);
            TextBox t = getTextBoxSN(head);
            Label l = getLabelStatus(head);

            rtfTerminal.Select();
            t.Enabled = false;
            g.FirstDisplayedScrollingRowIndex = 0;
            l.Visible = true;
            l.Text = define.testing;
            l.ForeColor = Color.OrangeRed;
            row[head - 1] = 0;

            ClearDataGridView(head);
        }
        public void ClearDataGridView(int head) {
            DataGridView g = getDataGridView(head);

            int count = 0;

            do //Clear data gridview
            {
                g.Rows[count].Cells[define.dataGrid.columnMeasure].Value = "";
                g.Rows[count].Cells[define.dataGrid.columnResult].Value = "";
                count++;

            } while (count < g.Rows.Count);

            g.Rows.Clear();
            g.Refresh();
            row[head - 1] = 0;
        }
        private string CheckFailure(int head) {
            DataGridView g = getDataGridView(head);

            string result = "";
            string failure = "";

            for (int i = 0; i < g.Rows.Count; i++) {
                try {
                    result = g.Rows[i].Cells[define.dataGrid.columnResult].Value.ToString();

                    if (result == define.fail) {
                        failure = failure + "<>" + g.Rows[i].Cells[define.dataGrid.columnStep].Value.ToString();
                    }

                } catch { }
            }

            return failure;
        }
        private void CreateFolder_datalog() {
            Folder.list.Clear();

            Folder.list.Add(Folder.driveD + this.Text + Folder.dataBase + tb_detail.Text);
            Folder.list.Add(Folder.list[0] + Folder.operationComplete);
            Folder.list.Add(Folder.list[0] + Folder.operationInComplete);
            Folder.list.Add(Folder.list[0] + Folder.debugComplete);
            Folder.list.Add(Folder.list[0] + Folder.debugInComplete);
            Folder.list.Add(Folder.dataMIS);
            Folder.list.Add(Folder.driveD + this.Text + Folder.dataBase + Folder.timeLine);

            for (int i = 1; i <= 4; i++) {

                if (excel.sameStep) {

                    if (!Directory.Exists(Folder.list[i] + "Head All")) {
                        Directory.CreateDirectory(Folder.list[i] + "Head All");
                    }
                        
                } else {

                    for (int j = 1; j <= tester.numHead; j++) {

                        if (!Directory.Exists(Folder.list[i] + "Head " + j)) {
                            Directory.CreateDirectory(Folder.list[i] + "Head " + j);
                        }
                            
                    }
                }
            }

            if (!Directory.Exists(Folder.list[5])) {
                Directory.CreateDirectory(Folder.list[5]);
            }
                
            if (!Directory.Exists(Folder.list[6])) {
                Directory.CreateDirectory(Folder.list[6]);
            }

            if (!Directory.Exists("D:\\Datalog to SQL"))
            {
                Directory.CreateDirectory("D:\\Datalog to SQL");
            }
        }
        private DataGridView getDataGridViewByHeader(string headFile) {
            DataGridView dataGridView = new DataGridView();

            switch (headFile) {//ไอศครีมเชอเบร็ตรสมะนาว
                case "Head 1\\": dataGridView = dataGridView_1; break;
                case "Head 2\\": dataGridView = dataGridView_2; break;
                case "Head 3\\": dataGridView = dataGridView_3; break;
                case "Head 4\\": dataGridView = dataGridView_4; break;
                case "Head 5\\": dataGridView = dataGridView_5; break;
                case "Head 6\\": dataGridView = dataGridView_6; break;
                case "Head 7\\": dataGridView = dataGridView_7; break;
                case "Head 8\\": dataGridView = dataGridView_8; break;
                case "Head 9\\": dataGridView = dataGridView_9; break;
                case "Head 10\\": dataGridView = dataGridView_10; break;
                case "Head 11\\": dataGridView = dataGridView_11; break;
                case "Head 12\\": dataGridView = dataGridView_12; break;
                case "Head 13\\": dataGridView = dataGridView_13; break;
                case "Head 14\\": dataGridView = dataGridView_14; break;
                case "Head 15\\": dataGridView = dataGridView_15; break;
                case "Head 16\\": dataGridView = dataGridView_16; break;
                case "Head 17\\": dataGridView = dataGridView_17; break;
                case "Head 18\\": dataGridView = dataGridView_18; break;
                case "Head 19\\": dataGridView = dataGridView_19; break;
                case "Head 20\\": dataGridView = dataGridView_20; break;
                case "Head 21\\": dataGridView = dataGridView_21; break;
                case "Head 22\\": dataGridView = dataGridView_22; break;
                case "Head 23\\": dataGridView = dataGridView_23; break;
                case "Head 24\\": dataGridView = dataGridView_24; break;
                case "Head 25\\": dataGridView = dataGridView_25; break;
                case "Head 26\\": dataGridView = dataGridView_26; break;
                case "Head 27\\": dataGridView = dataGridView_27; break;
                case "Head 28\\": dataGridView = dataGridView_28; break;
                case "Head 29\\": dataGridView = dataGridView_29; break;
                case "Head 30\\": dataGridView = dataGridView_30; break;
                case "Head 31\\": dataGridView = dataGridView_31; break;
                case "Head 32\\": dataGridView = dataGridView_32; break;
                case "Head 33\\": dataGridView = dataGridView_33; break;
                case "Head 34\\": dataGridView = dataGridView_34; break;
                case "Head 35\\": dataGridView = dataGridView_35; break;
                case "Head 36\\": dataGridView = dataGridView_36; break;
                case "Head All\\": dataGridView = dataGridView_1; break;
            }

            return dataGridView;
        }
        private int getRowArray_CreateHeader(string headFile) {
            int rowArray = 1;

            switch (headFile) {//ไอศครีมเชอเบร็ตรสมะนาว
                case "Head 1\\": rowArray = 1; break;
                case "Head 2\\": rowArray = 2; break;
                case "Head 3\\": rowArray = 3; break;
                case "Head 4\\": rowArray = 4; break;
                case "Head 5\\": rowArray = 5; break;
                case "Head 6\\": rowArray = 6; break;
                case "Head 7\\": rowArray = 7; break;
                case "Head 8\\": rowArray = 8; break;
                case "Head 9\\": rowArray = 9; break;
                case "Head 10\\": rowArray = 10; break;
                case "Head 11\\": rowArray = 11; break;
                case "Head 12\\": rowArray = 12; break;
                case "Head 13\\": rowArray = 13; break;
                case "Head 14\\": rowArray = 14; break;
                case "Head 15\\": rowArray = 15; break;
                case "Head 16\\": rowArray = 16; break;
                case "Head 17\\": rowArray = 17; break;
                case "Head 18\\": rowArray = 18; break;
                case "Head 19\\": rowArray = 19; break;
                case "Head 20\\": rowArray = 20; break;
                case "Head 21\\": rowArray = 21; break;
                case "Head 22\\": rowArray = 22; break;
                case "Head 23\\": rowArray = 23; break;
                case "Head 24\\": rowArray = 24; break;
                case "Head 25\\": rowArray = 25; break;
                case "Head 26\\": rowArray = 26; break;
                case "Head 27\\": rowArray = 27; break;
                case "Head 28\\": rowArray = 28; break;
                case "Head 29\\": rowArray = 29; break;
                case "Head 30\\": rowArray = 30; break;
                case "Head 31\\": rowArray = 31; break;
                case "Head 32\\": rowArray = 32; break;
                case "Head 33\\": rowArray = 33; break;
                case "Head 34\\": rowArray = 34; break;
                case "Head 35\\": rowArray = 35; break;
                case "Head 36\\": rowArray = 36; break;
                case "Head All\\": rowArray = 1; break;
            }

            return rowArray;
        }
        private void CreateHeader_datalog(string fileName, string headFile) {
            string headTimeLine = "";
            string Header = null;
            string csvPath = null;
            int count = 0;
            DataGridView g = getDataGridViewByHeader(headFile);

            if (tester.saveData == tester.saveNormal) {

                do {
                    if (!string.IsNullOrEmpty(g.Rows[count].Cells[2].Value as string)) {

                        Header = Header + "STEP" + g.Rows[count].Cells[0].Value + "(" + 
                            g.Rows[count].Cells[2].Value.ToString().Replace(",", " ") + ")" + 
                            g.Rows[count].Cells[1].Value.ToString().Replace(",", " ") + ",";
                    }
                        
                    count++;
                } while (count < g.Rows.Count);

                row[getRowArray_CreateHeader(headFile) - 1] = 0;
               
                headTimeLine = dataLog.headLog + dataLog.headLog_header + dataLog.headLog_fgAndWo + Header;
                Header = dataLog.headLog_header + Header;
                Header = dataLog.headLog + Header;

            } else {

                //ตรงนี้เป็นกรณีที่ save data แบบพิเศษ ใช้แค่โปรเจก denali ในอนาคตตรงนี้อาจจะเอาออก
                string[] headSpecial = File.ReadAllLines(excel.pathFile + cbb_fg.Text + dataLog.lastNameTXT);

                foreach (string headSplit in headSpecial) {
                    Header += headSplit + ",";
                }

                headTimeLine = Header;
            }


            //get header มันจะต้องทำทั้งหมด 4 file 
            for (int indexFolder = 1; indexFolder <= 4; indexFolder++) {

                bool checkHead = CheckHeader(Header, fileName, headFile, indexFolder);
                csvPath = Folder.list[indexFolder] + headFile + fileName + dataLog.lastNameCSV;

                if (!File.Exists(csvPath) || !checkHead) {

                    StreamWriter streamWriter = new StreamWriter(csvPath, true);
                    streamWriter.WriteLine(Header);
                    streamWriter.Close();
                }
            }

            //อันนี้เป็นการ get header ของ log time line
            bool checkHeadTimeLine = CheckHeader(headTimeLine, dataLog.timeLine.nameFile + dataLog.timeLine.numFile, "", 6);
            string pathDataTimeLine = Folder.list[6] + dataLog.timeLine.nameFile + dataLog.timeLine.numFile + dataLog.lastNameCSV;

            if (!File.Exists(pathDataTimeLine) || !checkHeadTimeLine) {

                StreamWriter streamWriter = new StreamWriter(pathDataTimeLine, true);
                streamWriter.WriteLine(headTimeLine);
                streamWriter.Close();
            }
        }
        public DataGridView getDataGridView(int head) {
            DataGridView dataGridView = new DataGridView();

            switch (head) {//ไอศครีมเชอเบร็ตรสมะนาว
                case 1: dataGridView = dataGridView_1; break;
                case 2: dataGridView = dataGridView_2; break;
                case 3: dataGridView = dataGridView_3; break;
                case 4: dataGridView = dataGridView_4; break;
                case 5: dataGridView = dataGridView_5; break;
                case 6: dataGridView = dataGridView_6; break;
                case 7: dataGridView = dataGridView_7; break;
                case 8: dataGridView = dataGridView_8; break;
                case 9: dataGridView = dataGridView_9; break;
                case 10: dataGridView = dataGridView_10; break;
                case 11: dataGridView = dataGridView_11; break;
                case 12: dataGridView = dataGridView_12; break;
                case 13: dataGridView = dataGridView_13; break;
                case 14: dataGridView = dataGridView_14; break;
                case 15: dataGridView = dataGridView_15; break;
                case 16: dataGridView = dataGridView_16; break;
                case 17: dataGridView = dataGridView_17; break;
                case 18: dataGridView = dataGridView_18; break;
                case 19: dataGridView = dataGridView_19; break;
                case 20: dataGridView = dataGridView_20; break;
                case 21: dataGridView = dataGridView_21; break;
                case 22: dataGridView = dataGridView_22; break;
                case 23: dataGridView = dataGridView_23; break;
                case 24: dataGridView = dataGridView_24; break;
                case 25: dataGridView = dataGridView_25; break;
                case 26: dataGridView = dataGridView_26; break;
                case 27: dataGridView = dataGridView_27; break;
                case 28: dataGridView = dataGridView_28; break;
                case 29: dataGridView = dataGridView_29; break;
                case 30: dataGridView = dataGridView_30; break;
                case 31: dataGridView = dataGridView_31; break;
                case 32: dataGridView = dataGridView_32; break;
                case 33: dataGridView = dataGridView_33; break;
                case 34: dataGridView = dataGridView_34; break;
                case 35: dataGridView = dataGridView_35; break;
                case 36: dataGridView = dataGridView_36; break;
            }

            return dataGridView;
        }
        public TextBox getTextBoxSN(int head) {
            TextBox textBox = new TextBox();

            switch (head) {//ไอศครีมเชอเบร็ตรสมะนาว
                case 1: textBox = txtSNBoard_1; break;
                case 2: textBox = txtSNBoard_2; break;
                case 3: textBox = txtSNBoard_3; break;
                case 4: textBox = txtSNBoard_4; break;
                case 5: textBox = txtSNBoard_5; break;
                case 6: textBox = txtSNBoard_6; break;
                case 7: textBox = txtSNBoard_7; break;
                case 8: textBox = txtSNBoard_8; break;
                case 9: textBox = txtSNBoard_9; break;
                case 10: textBox = txtSNBoard_10; break;
                case 11: textBox = txtSNBoard_11; break;
                case 12: textBox = txtSNBoard_12; break;
                case 13: textBox = txtSNBoard_13; break;
                case 14: textBox = txtSNBoard_14; break;
                case 15: textBox = txtSNBoard_15; break;
                case 16: textBox = txtSNBoard_16; break;
                case 17: textBox = txtSNBoard_17; break;
                case 18: textBox = txtSNBoard_18; break;
                case 19: textBox = txtSNBoard_19; break;
                case 20: textBox = txtSNBoard_20; break;
                case 21: textBox = txtSNBoard_21; break;
                case 22: textBox = txtSNBoard_22; break;
                case 23: textBox = txtSNBoard_23; break;
                case 24: textBox = txtSNBoard_24; break;
                case 25: textBox = txtSNBoard_25; break;
                case 26: textBox = txtSNBoard_26; break;
                case 27: textBox = txtSNBoard_27; break;
                case 28: textBox = txtSNBoard_28; break;
                case 29: textBox = txtSNBoard_29; break;
                case 30: textBox = txtSNBoard_30; break;
                case 31: textBox = txtSNBoard_31; break;
                case 32: textBox = txtSNBoard_32; break;
                case 33: textBox = txtSNBoard_33; break;
                case 34: textBox = txtSNBoard_34; break;
                case 35: textBox = txtSNBoard_35; break;
                case 36: textBox = txtSNBoard_36; break;
            }

            return textBox;
        }
        public Label getLabelStatus(int head) {
            Label label = new Label();

            switch (head) {//ไอศครีมเชอเบร็ตรสมะนาว
                case 1: label = status_1; break;
                case 2: label = status_2; break;
                case 3: label = status_3; break;
                case 4: label = status_4; break;
                case 5: label = status_5; break;
                case 6: label = status_6; break;
                case 7: label = status_7; break;
                case 8: label = status_8; break;
                case 9: label = status_9; break;
                case 10: label = status_10; break;
                case 11: label = status_11; break;
                case 12: label = status_12; break;
                case 13: label = status_13; break;
                case 14: label = status_14; break;
                case 15: label = status_15; break;
                case 16: label = status_16; break;
                case 17: label = status_17; break;
                case 18: label = status_18; break;
                case 19: label = status_19; break;
                case 20: label = status_20; break;
                case 21: label = status_21; break;
                case 22: label = status_22; break;
                case 23: label = status_23; break;
                case 24: label = status_24; break;
                case 25: label = status_25; break;
                case 26: label = status_26; break;
                case 27: label = status_27; break;
                case 28: label = status_28; break;
                case 29: label = status_29; break;
                case 30: label = status_30; break;
                case 31: label = status_31; break;
                case 32: label = status_32; break;
                case 33: label = status_33; break;
                case 34: label = status_34; break;
                case 35: label = status_35; break;
                case 36: label = status_36; break;
            }

            return label;
        }
        public Label getLabelTestTime(int head) {
            Label label = new Label();

            switch (head) {//ไอศครีมเชอเบร็ตรสมะนาว
                case 1: label = lblTestTime_1; break;
                case 2: label = lblTestTime_2; break;
                case 3: label = lblTestTime_3; break;
                case 4: label = lblTestTime_4; break;
                case 5: label = lblTestTime_5; break;
                case 6: label = lblTestTime_6; break;
                case 7: label = lblTestTime_7; break;
                case 8: label = lblTestTime_8; break;
                case 9: label = lblTestTime_9; break;
                case 10: label = lblTestTime_10; break;
                case 11: label = lblTestTime_11; break;
                case 12: label = lblTestTime_12; break;
                case 13: label = lblTestTime_13; break;
                case 14: label = lblTestTime_14; break;
                case 15: label = lblTestTime_15; break;
                case 16: label = lblTestTime_16; break;
                case 17: label = lblTestTime_17; break;
                case 18: label = lblTestTime_18; break;
                case 19: label = lblTestTime_19; break;
                case 20: label = lblTestTime_20; break;
                case 21: label = lblTestTime_21; break;
                case 22: label = lblTestTime_22; break;
                case 23: label = lblTestTime_23; break;
                case 24: label = lblTestTime_24; break;
                case 25: label = lblTestTime_25; break;
                case 26: label = lblTestTime_26; break;
                case 27: label = lblTestTime_27; break;
                case 28: label = lblTestTime_28; break;
                case 29: label = lblTestTime_29; break;
                case 30: label = lblTestTime_30; break;
                case 31: label = lblTestTime_31; break;
                case 32: label = lblTestTime_32; break;
                case 33: label = lblTestTime_33; break;
                case 34: label = lblTestTime_34; break;
                case 35: label = lblTestTime_35; break;
                case 36: label = lblTestTime_36; break;
            }

            return label;
        }
        public Label getLabelInOutTime(int head) {
            Label label = new Label();

            switch (head) {//ไอศครีมเชอเบร็ตรสมะนาว
                case 1: label = lblLoadInOutTime_1; break;
                case 2: label = lblLoadInOutTime_2; break;
                case 3: label = lblLoadInOutTime_3; break;
                case 4: label = lblLoadInOutTime_4; break;
                case 5: label = lblLoadInOutTime_5; break;
                case 6: label = lblLoadInOutTime_6; break;
                case 7: label = lblLoadInOutTime_7; break;
                case 8: label = lblLoadInOutTime_8; break;
                case 9: label = lblLoadInOutTime_9; break;
                case 10: label = lblLoadInOutTime_10; break;
                case 11: label = lblLoadInOutTime_11; break;
                case 12: label = lblLoadInOutTime_12; break;
                case 13: label = lblLoadInOutTime_13; break;
                case 14: label = lblLoadInOutTime_14; break;
                case 15: label = lblLoadInOutTime_15; break;
                case 16: label = lblLoadInOutTime_16; break;
                case 17: label = lblLoadInOutTime_17; break;
                case 18: label = lblLoadInOutTime_18; break;
                case 19: label = lblLoadInOutTime_19; break;
                case 20: label = lblLoadInOutTime_20; break;
                case 21: label = lblLoadInOutTime_21; break;
                case 22: label = lblLoadInOutTime_22; break;
                case 23: label = lblLoadInOutTime_23; break;
                case 24: label = lblLoadInOutTime_24; break;
                case 25: label = lblLoadInOutTime_25; break;
                case 26: label = lblLoadInOutTime_26; break;
                case 27: label = lblLoadInOutTime_27; break;
                case 28: label = lblLoadInOutTime_28; break;
                case 29: label = lblLoadInOutTime_29; break;
                case 30: label = lblLoadInOutTime_30; break;
                case 31: label = lblLoadInOutTime_31; break;
                case 32: label = lblLoadInOutTime_32; break;
                case 33: label = lblLoadInOutTime_33; break;
                case 34: label = lblLoadInOutTime_34; break;
                case 35: label = lblLoadInOutTime_35; break;
                case 36: label = lblLoadInOutTime_36; break;
            }

            return label;
        }
        public string getHeadFile(int head) {
            string headFile = "";

            switch (head) {//ไอศครีมเชอเบร็ตรสมะนาว
                case 1: headFile = "Head 1\\"; break;
                case 2: headFile = "Head 2\\"; break;
                case 3: headFile = "Head 3\\"; break;
                case 4: headFile = "Head 4\\"; break;
                case 5: headFile = "Head 5\\"; break;
                case 6: headFile = "Head 6\\"; break;
                case 7: headFile = "Head 7\\"; break;
                case 8: headFile = "Head 8\\"; break;
                case 9: headFile = "Head 9\\"; break;
                case 10: headFile = "Head 10\\"; break;
                case 11: headFile = "Head 11\\"; break;
                case 12: headFile = "Head 12\\"; break;
                case 13: headFile = "Head 13\\"; break;
                case 14: headFile = "Head 14\\"; break;
                case 15: headFile = "Head 15\\"; break;
                case 16: headFile = "Head 16\\"; break;
                case 17: headFile = "Head 17\\"; break;
                case 18: headFile = "Head 18\\"; break;
                case 19: headFile = "Head 19\\"; break;
                case 20: headFile = "Head 20\\"; break;
                case 21: headFile = "Head 21\\"; break;
                case 22: headFile = "Head 22\\"; break;
                case 23: headFile = "Head 23\\"; break;
                case 24: headFile = "Head 24\\"; break;
                case 25: headFile = "Head 25\\"; break;
                case 26: headFile = "Head 26\\"; break;
                case 27: headFile = "Head 27\\"; break;
                case 28: headFile = "Head 28\\"; break;
                case 29: headFile = "Head 29\\"; break;
                case 30: headFile = "Head 30\\"; break;
                case 31: headFile = "Head 31\\"; break;
                case 32: headFile = "Head 32\\"; break;
                case 33: headFile = "Head 33\\"; break;
                case 34: headFile = "Head 34\\"; break;
                case 35: headFile = "Head 35\\"; break;
                case 36: headFile = "Head 36\\"; break;
            }

            return headFile;
        }
        public Button getButtonTest(int head) {
            Button button = new Button();

            switch (head) {//ไอศครีมเชอเบร็ตรสมะนาว
                case 1: button = btnTEST_1; break;
                case 2: button = btnTEST_2; break;
                case 3: button = btnTEST_3; break;
                case 4: button = btnTEST_4; break;
                case 5: button = btnTEST_5; break;
                case 6: button = btnTEST_6; break;
                case 7: button = btnTEST_7; break;
                case 8: button = btnTEST_8; break;
                case 9: button = btnTEST_9; break;
                case 10: button = btnTEST_10; break;
                case 11: button = btnTEST_11; break;
                case 12: button = btnTEST_12; break;
                case 13: button = btnTEST_13; break;
                case 14: button = btnTEST_14; break;
                case 15: button = btnTEST_15; break;
                case 16: button = btnTEST_16; break;
                case 17: button = btnTEST_17; break;
                case 18: button = btnTEST_18; break;
                case 19: button = btnTEST_19; break;
                case 20: button = btnTEST_20; break;
                case 21: button = btnTEST_21; break;
                case 22: button = btnTEST_22; break;
                case 23: button = btnTEST_23; break;
                case 24: button = btnTEST_24; break;
                case 25: button = btnTEST_25; break;
                case 26: button = btnTEST_26; break;
                case 27: button = btnTEST_27; break;
                case 28: button = btnTEST_28; break;
                case 29: button = btnTEST_29; break;
                case 30: button = btnTEST_30; break;
                case 31: button = btnTEST_31; break;
                case 32: button = btnTEST_32; break;
                case 33: button = btnTEST_33; break;
                case 34: button = btnTEST_34; break;
                case 35: button = btnTEST_35; break;
                case 36: button = btnTEST_36; break;
            }

            return button;
        }
        public Button getButtonExit(int head) {
            Button button = new Button();

            switch (head) { //ไอศครีมเชอเบร็ตรสมะนาว
                case 1: button = btnEXIT_1; break;
                case 2: button = btnEXIT_2; break;
                case 3: button = btnEXIT_3; break;
                case 4: button = btnEXIT_4; break;
                case 5: button = btnEXIT_5; break;
                case 6: button = btnEXIT_6; break;
                case 7: button = btnEXIT_7; break;
                case 8: button = btnEXIT_8; break;
                case 9: button = btnEXIT_9; break;
                case 10: button = btnEXIT_10; break;
                case 11: button = btnEXIT_11; break;
                case 12: button = btnEXIT_12; break;
                case 13: button = btnEXIT_13; break;
                case 14: button = btnEXIT_14; break;
                case 15: button = btnEXIT_15; break;
                case 16: button = btnEXIT_16; break;
                case 17: button = btnEXIT_17; break;
                case 18: button = btnEXIT_18; break;
                case 19: button = btnEXIT_19; break;
                case 20: button = btnEXIT_20; break;
                case 21: button = btnEXIT_21; break;
                case 22: button = btnEXIT_22; break;
                case 23: button = btnEXIT_23; break;
                case 24: button = btnEXIT_24; break;
                case 25: button = btnEXIT_25; break;
                case 26: button = btnEXIT_26; break;
                case 27: button = btnEXIT_27; break;
                case 28: button = btnEXIT_28; break;
                case 29: button = btnEXIT_29; break;
                case 30: button = btnEXIT_30; break;
                case 31: button = btnEXIT_31; break;
                case 32: button = btnEXIT_32; break;
                case 33: button = btnEXIT_33; break;
                case 34: button = btnEXIT_34; break;
                case 35: button = btnEXIT_35; break;
                case 36: button = btnEXIT_36; break;
            }

            return button;
        }
        public ProgressBar getProgressBar(int head) {
            ProgressBar progressBar = new ProgressBar();

            switch (head) {//ไอศครีมเชอเบร็ตรสมะนาว
                case 1: progressBar = progressBar_1; break;
                case 2: progressBar = progressBar_2; break;
                case 3: progressBar = progressBar_3; break;
                case 4: progressBar = progressBar_4; break;
                case 5: progressBar = progressBar_5; break;
                case 6: progressBar = progressBar_6; break;
                case 7: progressBar = progressBar_7; break;
                case 8: progressBar = progressBar_8; break;
                case 9: progressBar = progressBar_9; break;
                case 10: progressBar = progressBar_10; break;
                case 11: progressBar = progressBar_11; break;
                case 12: progressBar = progressBar_12; break;
                case 13: progressBar = progressBar_13; break;
                case 14: progressBar = progressBar_14; break;
                case 15: progressBar = progressBar_15; break;
                case 16: progressBar = progressBar_16; break;
                case 17: progressBar = progressBar_17; break;
                case 18: progressBar = progressBar_18; break;
                case 19: progressBar = progressBar_19; break;
                case 20: progressBar = progressBar_20; break;
                case 21: progressBar = progressBar_21; break;
                case 22: progressBar = progressBar_22; break;
                case 23: progressBar = progressBar_23; break;
                case 24: progressBar = progressBar_24; break;
                case 25: progressBar = progressBar_25; break;
                case 26: progressBar = progressBar_26; break;
                case 27: progressBar = progressBar_27; break;
                case 28: progressBar = progressBar_28; break;
                case 29: progressBar = progressBar_29; break;
                case 30: progressBar = progressBar_30; break;
                case 31: progressBar = progressBar_31; break;
                case 32: progressBar = progressBar_32; break;
                case 33: progressBar = progressBar_33; break;
                case 34: progressBar = progressBar_34; break;
                case 35: progressBar = progressBar_35; break;
                case 36: progressBar = progressBar_36; break;
            }

            return progressBar;
        }
        public ToolStripMenuItem getToolStripMenuItemDebug(int head) {
            ToolStripMenuItem toolStripMenuItem = new ToolStripMenuItem();

            switch (head) {//ไอศครีมเชอเบร็ตรสมะนาว
                case 1: toolStripMenuItem = set_debug_1; break;
                case 2: toolStripMenuItem = set_debug_2; break;
                case 3: toolStripMenuItem = set_debug_3; break;
                case 4: toolStripMenuItem = set_debug_4; break;
                case 5: toolStripMenuItem = set_debug_5; break;
                case 6: toolStripMenuItem = set_debug_6; break;
                case 7: toolStripMenuItem = set_debug_7; break;
                case 8: toolStripMenuItem = set_debug_8; break;
                case 9: toolStripMenuItem = set_debug_9; break;
                case 10: toolStripMenuItem = set_debug_10; break;
                case 11: toolStripMenuItem = set_debug_11; break;
                case 12: toolStripMenuItem = set_debug_12; break;
                case 13: toolStripMenuItem = set_debug_13; break;
                case 14: toolStripMenuItem = set_debug_14; break;
                case 15: toolStripMenuItem = set_debug_15; break;
                case 16: toolStripMenuItem = set_debug_16; break;
                case 17: toolStripMenuItem = set_debug_17; break;
                case 18: toolStripMenuItem = set_debug_18; break;
                case 19: toolStripMenuItem = set_debug_19; break;
                case 20: toolStripMenuItem = set_debug_20; break;
                case 21: toolStripMenuItem = set_debug_21; break;
                case 22: toolStripMenuItem = set_debug_22; break;
                case 23: toolStripMenuItem = set_debug_23; break;
                case 24: toolStripMenuItem = set_debug_24; break;
                case 25: toolStripMenuItem = set_debug_25; break;
                case 26: toolStripMenuItem = set_debug_26; break;
                case 27: toolStripMenuItem = set_debug_27; break;
                case 28: toolStripMenuItem = set_debug_28; break;
                case 29: toolStripMenuItem = set_debug_29; break;
                case 30: toolStripMenuItem = set_debug_30; break;
                case 31: toolStripMenuItem = set_debug_31; break;
                case 32: toolStripMenuItem = set_debug_32; break;
                case 33: toolStripMenuItem = set_debug_33; break;
                case 34: toolStripMenuItem = set_debug_34; break;
                case 35: toolStripMenuItem = set_debug_35; break;
                case 36: toolStripMenuItem = set_debug_36; break;
            }

            return toolStripMenuItem;
        }
        public GroupBox getGroupBoxHead(int head) {
            GroupBox groupBox = new GroupBox();

            switch (head) {//ไอศครีมเชอเบร็ตรสมะนาว
                case 1: groupBox = groupBox_head1; break;
                case 2: groupBox = groupBox_head2; break;
                case 3: groupBox = groupBox_head3; break;
                case 4: groupBox = groupBox_head4; break;
                case 5: groupBox = groupBox_head5; break;
                case 6: groupBox = groupBox_head6; break;
                case 7: groupBox = groupBox_head7; break;
                case 8: groupBox = groupBox_head8; break;
                case 9: groupBox = groupBox_head9; break;
                case 10: groupBox = groupBox_head10; break;
                case 11: groupBox = groupBox_head11; break;
                case 12: groupBox = groupBox_head12; break;
                case 13: groupBox = groupBox_head13; break;
                case 14: groupBox = groupBox_head14; break;
                case 15: groupBox = groupBox_head15; break;
                case 16: groupBox = groupBox_head16; break;
                case 17: groupBox = groupBox_head17; break;
                case 18: groupBox = groupBox_head18; break;
                case 19: groupBox = groupBox_head19; break;
                case 20: groupBox = groupBox_head20; break;
                case 21: groupBox = groupBox_head21; break;
                case 22: groupBox = groupBox_head22; break;
                case 23: groupBox = groupBox_head23; break;
                case 24: groupBox = groupBox_head24; break;
                case 25: groupBox = groupBox_head25; break;
                case 26: groupBox = groupBox_head26; break;
                case 27: groupBox = groupBox_head27; break;
                case 28: groupBox = groupBox_head28; break;
                case 29: groupBox = groupBox_head29; break;
                case 30: groupBox = groupBox_head30; break;
                case 31: groupBox = groupBox_head31; break;
                case 32: groupBox = groupBox_head32; break;
                case 33: groupBox = groupBox_head33; break;
                case 34: groupBox = groupBox_head34; break;
                case 35: groupBox = groupBox_head35; break;
                case 36: groupBox = groupBox_head36; break;
            }

            return groupBox;
        }
        private bool getResultTest(DataGridView gridView) {
            bool results = true;

            if (flagNotReTest[select_test - 1]) {
                return true;
            }

            for (int i = 0; i < gridView.Rows.Count - 1; i++) {
                try {

                    if (gridView.Rows[i].Cells[define.dataGrid.columnResult].Value.ToString() != define.pass &&
                        gridView.Rows[i].Cells[define.dataGrid.columnResult].Value.ToString() != "") {
                        results = false;
                        break;
                    }

                } catch {
                    results = false;
                    break;
                }
            }

            return results;
        }
        private string getDataSummarySpecial(DataGridView gridView, string Detial, string checkFailure, bool results, DateTime dateTime, 
            TextBox textboxSN, string csvPath, string headFile, string fileName, Label labelTestTime, int head) {
            string dataSummary = "";

            //อันนี้เป็นแบบ พิเศษ ใช้แค่ denali ในอนาคตอาจจะเอาออก
            string[] headLog = File.ReadAllLines(excel.pathFile + cbb_fg.Text + "_step.txt");
            List<int> rowDataGridViewList = new List<int>();

            for (int i = 0; i < gridView.Rows.Count; i++) {
                rowDataGridViewList.Add(i);
            }

            foreach (string headLogSplit in headLog) {
                if (headLogSplit == "non") {
                    Detial += "-,";
                    continue;
                }

                if (headLogSplit.Contains("data")) {
                    switch (headLogSplit) {
                        case "data_fail":
                            Detial += checkFailure.Replace(",", "") + ",";
                            break;
                        case "data_Final_Result":
                            if (results) {
                                Detial += "PASS,";
                            } else {
                                Detial += "FAIL,";
                            }
                            break;
                        case "data_DATE_TIME":
                            Detial += dateTime.ToString(DateTimePay.format, CultureInfo.CreateSpecificCulture(DateTimePay.us));
                            break;
                        case "data_FG":
                            Detial += cbb_fg.Text + ",";
                            break;
                        case "data_WO":
                            try {
                                Detial += tb_wo.Text + ",";
                            } catch {
                                Detial += ",";
                            }
                            break;
                        case "data_TESTER_ID":
                            Detial += setupPay.read_text(tester.headConfig.komsonTester, tester.nameFile) + " (head " + head + "),";
                            break;
                        case "data_prism_number":
                            try {
                                Detial += textboxSN.Text + ",";
                            } catch {
                                Detial += ",";
                            }
                            break;
                        case "data_Operator":
                            try {
                                Detial += tb_userID.Text + ",";
                            } catch {
                                Detial += ",";
                            }
                            break;
                        case "data_mode":
                            if (cb_OperationMode.Checked) {
                                Detial += prismTest.OperationMode + ",";
                                if (results) csvPath = Folder.list[1] + headFile + fileName + dataLog.lastNameCSV;
                                else csvPath = Folder.list[2] + headFile + fileName + dataLog.lastNameCSV;
                            } else {
                                Detial += prismTest.DebugMode + ",";
                                if (results == true) csvPath = Folder.list[3] + headFile + fileName + dataLog.lastNameCSV;
                                else csvPath = Folder.list[4] + headFile + fileName + dataLog.lastNameCSV;
                            }
                            break;
                        case "data_Test_Finish_Time":
                            Detial += dateTime.Year.ToString() + "." + dateTime.Month.ToString("00") + "." +
                                dateTime.Day.ToString("00") + " " + dateTime.ToString("T") + ",";
                            break;
                        case "data_Test_Total_Time":
                            Detial += TimeSpan.FromSeconds(Convert.ToInt32(labelTestTime.Text)).ToString(@"hh\:mm\:ss") + ",";
                            break;
                        case "data_Test_Start_Time":
                            Detial += time_start[select_test - 1] + ",";
                            break;
                    }
                    continue;
                }

                string[] headLogArray = headLogSplit.Split(',');

                if (headLogArray.Length == 1) {
                    foreach (int rowSplit in rowDataGridViewList) {
                        string stepGridView = "";

                        try {
                            stepGridView = gridView.Rows[rowSplit].Cells[define.dataGrid.columnStep].Value.ToString();
                        } catch { }

                        if (stepGridView == headLogArray[0]) {
                            try {
                                if (gridView.Rows[rowSplit].Cells[define.dataGrid.columnMeasure].Value.ToString().Length > 12) {
                                    Detial += "'" + gridView.Rows[rowSplit].Cells[define.dataGrid.columnMeasure].Value.ToString() + ",";

                                } else {
                                    Detial += gridView.Rows[rowSplit].Cells[define.dataGrid.columnMeasure].Value.ToString() + ",";
                                }

                            } catch {
                                Detial += ",";
                            }

                            rowDataGridViewList.Remove(rowSplit);
                            break;
                        }
                    }

                } else {
                    foreach (string s_split in headLogArray) {
                        foreach (int nt in rowDataGridViewList) {
                            if (gridView.Rows[nt].Cells[0].Value.ToString() == s_split) {
                                try {
                                    if (gridView.Rows[nt].Cells[3].Value.ToString().Length > 12) Detial += "'" + gridView.Rows[nt].Cells[3].Value.ToString() + " ";
                                    else Detial += gridView.Rows[nt].Cells[3].Value.ToString() + " ";
                                } catch { Detial += ","; }
                                rowDataGridViewList.Remove(nt);
                                break;
                            }
                        }
                    }
                    Detial += ",";
                }
            }
            dataSummary = Detial;

            return dataSummary;
        }
        private void save_data(int head) {
            DataGridView gridView = getDataGridView(head);
            TextBox textboxSN = getTextBoxSN(head);
            Label labelStatus = getLabelStatus(head);
            Label labelTestTime = getLabelTestTime(head);
            Label labelInOutTime = getLabelInOutTime(head);
            string headFile = getHeadFile(head);
            bool results = getResultTest(gridView);

            if (excel.sameStep) {
                headFile = "Head All\\";
            }

            if (results) {
                labelStatus.Text = define.pass;
                labelStatus.ForeColor = Color.Green;

            } else {
                labelStatus.Text = define.fail;
                labelStatus.ForeColor = Color.Red;
            }

            string csvPath = "";
            string dataSummary = "";
            string Detial = "";
            string fileName = "";
            DateTime dateTime = DateTime.Now;

            //แปลงเป็นแบบของไทย ทำเก็บไว้เผื่อมีโอกาสได้ใช้
            ThaiBuddhistCalendar calTime = new ThaiBuddhistCalendar();
            DateTime dateTimeThai = new DateTime(calTime.GetYear(dateTime), calTime.GetMonth(dateTime), dateTime.Day);
            
            if(prismTest.mode == prismTest.Operation) {
                fileName = tb_wo.Text.Replace("/", "_");
            } else {
                fileName = prismTest.Debug;
            }

            CreateHeader_datalog(fileName, headFile);
            string checkFailure = CheckFailure(head);

            if (tester.saveData == tester.saveNormal) {
                Detial = dateTime.ToString(DateTimePay.format, CultureInfo.CreateSpecificCulture(DateTimePay.us));
                Detial += tb_userID.Text + ",";
                Detial += tb_swVersion.Text + ",";
                Detial += tb_fwVersion.Text + ",";
                Detial += tb_spec.Text + ",";
                Detial += labelTestTime.Text + ",";
                Detial += labelInOutTime.Text + ",";

                if (cb_OperationMode.Checked) {
                    Detial = Detial + prismTest.OperationMode + ",";

                    if (results) {
                        csvPath = Folder.list[1] + headFile + fileName + dataLog.lastNameCSV;

                    } else {
                        csvPath = Folder.list[2] + headFile + fileName + dataLog.lastNameCSV;
                    }

                } else {
                    Detial = Detial + prismTest.DebugMode + ",";

                    if (results) {
                        csvPath = Folder.list[3] + headFile + fileName + dataLog.lastNameCSV;

                    } else {
                        csvPath = Folder.list[4] + headFile + fileName + dataLog.lastNameCSV;
                    }
                }

                Detial = Detial + labelStatus.Text + "," + textboxSN.Text + "," + checkFailure;
                Detial += "," + "Head " + head;

                for (int i = 0; i < gridView.Rows.Count; i++) {
                    try {
                        if ((gridView.Rows[i].Cells[2].Value.ToString() != "")) {

                            if (gridView.Rows[i].Cells[3].Value == null) {
                                dataSummary = dataSummary + "," + "";

                            } else {
                                dataSummary = dataSummary + "," + gridView.Rows[i].Cells[3].Value.ToString().Replace(",", " ");
                            }
                                
                        }
                    } catch { }
                }

                dataLog.timeLine.data = Detial + "," + cbb_fg.Text + "," + tb_wo.Text + dataSummary;
                dataSummary = Detial + dataSummary;

            } else {

                string[] sup = File.ReadAllLines("../../TestDescription/" + cbb_fg.Text + "_step.txt");
                List<Int32> row_data = new List<int>();
                for (int i = 0; i < gridView.Rows.Count; i++)
                {
                    row_data.Add(i);
                }
                foreach (string s in sup)
                {
                    if (s == "non") { Detial += "-,"; continue; }
                    if (s.Contains("data"))
                    {
                        switch (s)
                        {
                            case "data_fail": Detial += checkFailure.Replace(",", "") + ","; break;
                            case "data_Final_Result":
                                if (results) Detial += "PASS,";
                                else Detial += "FAIL,";
                                break;
                            case "data_DATE_TIME": Detial += DateTime.Now.ToString("yyyy.MM.dd HH:mm:ss,", System.Globalization.CultureInfo.CreateSpecificCulture("en-US")); break;
                            case "data_FG": Detial += cbb_fg.Text + ","; break;
                            case "data_WO":
                                try { Detial += tb_wo.Text + ","; } catch { Detial += ","; }
                                break;
                            case "data_TESTER_ID": Detial += File.ReadAllText("tester_no.txt") + " (head " + head + "),"; break;
                            case "data_prism_number":
                                try { Detial += textboxSN.Text + ","; } catch { Detial += ","; }
                                break;
                            case "data_Operator":
                                try { Detial += tb_userID.Text + ","; } catch { Detial += ","; }
                                break;
                            case "data_mode":
                                if (cb_OperationMode.Checked)
                                {
                                    Detial += "Operation Mode" + ",";
                                    if (results) csvPath = Folder.list[1] + headFile + fileName + dataLog.lastNameCSV;
                                    else csvPath = Folder.list[2] + headFile + fileName + dataLog.lastNameCSV;
                                }
                                else
                                {
                                    Detial += "Debug Mode" + ",";
                                    if (results == true) csvPath = Folder.list[3] + headFile + fileName + dataLog.lastNameCSV;
                                    else csvPath = Folder.list[4] + headFile + fileName + dataLog.lastNameCSV;
                                }
                                break;
                            case "data_Test_Finish_Time": Detial += dateTime.Year.ToString() + "." + 
                                    dateTime.Month.ToString("00") + "." + dateTime.Day.ToString("00") + " " +
                                    dateTime.ToString("T") + ","; break;
                            case "data_Test_Total_Time": Detial += TimeSpan.FromSeconds(Convert.ToInt32(labelTestTime.Text)).ToString(@"hh\:mm\:ss") + ","; break;
                            case "data_Test_Start_Time": Detial += time_start[select_test - 1] + ","; break;
                        }
                        continue;
                    }
                    string[] split_s = s.Split(',');
                    if (split_s.Length == 1)
                    {
                        foreach (int nt in row_data)
                        {
                            string jj = "";
                            try { jj = gridView.Rows[nt].Cells[0].Value.ToString(); } catch { }
                            if (jj == split_s[0])
                            {
                                try
                                {
                                    if (gridView.Rows[nt].Cells[3].Value.ToString().Length > 12) 
                                        Detial += "'" + gridView.Rows[nt].Cells[3].Value.ToString() + ",";
                                    else 
                                        Detial += gridView.Rows[nt].Cells[3].Value.ToString() + ",";
                                }
                                catch { 
                                    Detial += ","; 
                                }
                                row_data.Remove(nt);
                                break;
                            }
                        }
                    }
                    else
                    {
                        foreach (string s_split in split_s)
                        {
                            foreach (int nt in row_data)
                            {
                                if (gridView.Rows[nt].Cells[0].Value.ToString() == s_split)
                                {
                                    try
                                    {
                                        if (gridView.Rows[nt].Cells[3].Value.ToString().Length > 12) 
                                            Detial += "'" + gridView.Rows[nt].Cells[3].Value.ToString() + " ";
                                        else 
                                            Detial += gridView.Rows[nt].Cells[3].Value.ToString() + " ";
                                    }
                                    catch { }
                                    row_data.Remove(nt);
                                    break;
                                }
                            }
                        }
                        Detial += ",";
                    }
                }

                dataLog.timeLine.data = Detial;
                dataSummary = Detial;
            }

            StreamWriter swOut = new StreamWriter(csvPath, true);
            while (true) {
                try {
                    swOut.WriteLine(dataSummary);
                } catch {
                    MessageBox.Show("_กรุณาปิด log file csv ก่อน");
                    continue;
                }
                break;
            }
            swOut.Close();

            try {
                dataLog.timeLine.rowCSV = Convert.ToDouble(File.ReadAllText(Folder.list[6] + dataLog.timeLine.nameFileRow));
                dataLog.timeLine.rowCSV++;
                File.WriteAllText(Folder.list[6] + dataLog.timeLine.nameFileRow, dataLog.timeLine.rowCSV.ToString());
            } catch {
                dataLog.timeLine.rowCSV = 1;
                File.WriteAllText(Folder.list[6] + dataLog.timeLine.nameFileRow, 1.ToString());
            }
            if(dataLog.timeLine.rowCSV > dataLog.timeLine.maxRow) {
                File.WriteAllText(Folder.list[6] + dataLog.timeLine.nameFileRow, 1.ToString());
                dataLog.timeLine.numFile = (Convert.ToInt32(dataLog.timeLine.numFile) + 1).ToString();
                setupPay.write_text(dataLog.headConfig.fileTimeLine, dataLog.nameFile, tester.nameFile);
            }

            csvPath = Folder.list[6] + dataLog.timeLine.nameFile + dataLog.timeLine.numFile + ".csv";
            StreamWriter StreamWriterTimeLine = new StreamWriter(csvPath, true);
            while (true) {
                try {
                    StreamWriterTimeLine.WriteLine(dataLog.timeLine.data);
                } catch {
                    MessageBox.Show("_กรุณาปิด log file TimeLine.csv ก่อน");
                    continue;
                }
                break;
            }
            StreamWriterTimeLine.Close();

            if (results == true || tester.upFail) {//Update data log to PRISM
                bool digit_snn = true;
                if (textboxSN.Text.Length != Convert.ToInt32(prismTest.digitSN)) {
                    digit_snn = false;
                    Log(LogMsgType.Error_Red, "\n" + "sn prism not " + prismTest.digitSN + " digit");
                }
                if (cb_OperationMode.Checked && prismTest.mode != prismTest.Debug && digit_snn && flag_sn_pass[select_test - 1]){
                    JsonConvert jsonConvert = new JsonConvert();
                    jsonConvert.Date = DateTime.Now.ToString("dd/MM/yyyy", CultureInfo.CreateSpecificCulture("en-US"));
                    jsonConvert.Time = DateTime.Now.ToString("HH:mm:ss", CultureInfo.CreateSpecificCulture("en-US"));
                    jsonConvert.LoginID = tb_userID.Text;
                    jsonConvert.SWVersion = tb_swVersion.Text;
                    jsonConvert.FWVersion = tb_fwVersion.Text;
                    jsonConvert.SpecVersion = tb_spec.Text;
                    jsonConvert.TestTime = labelTestTime.Text;
                    jsonConvert.LoadInOut = labelInOutTime.Text;
                    if (cb_OperationMode.Checked) {
                        jsonConvert.Mode = prismTest.Operation;
                    } else {
                        jsonConvert.Mode = prismTest.Debug;
                    }
                    jsonConvert.FinalResult = labelStatus.Text;
                    jsonConvert.SN = textboxSN.Text;
                    jsonConvert.Failure = checkFailure.Replace(";", string.Empty).Trim();
                    for (int i = 0; i < gridView.Rows.Count; i++) {
                        JsonConvert.ResultString_ resultString = new JsonConvert.ResultString_();
                        try {
                            string trySup = gridView.Rows[i].Cells[2].Value.ToString();
                        } catch { continue; }
                        if (gridView.Rows[i].Cells[2].Value.ToString() == string.Empty) continue;
                        try {
                            resultString.Step = gridView.Rows[i].Cells[0].Value.ToString();
                        } catch { }
                        try {
                            resultString.Description = gridView.Rows[i].Cells[1].Value.ToString();
                        } catch { }
                        try {
                            resultString.Tolerance = gridView.Rows[i].Cells[2].Value.ToString();
                        } catch { }
                        try {
                            resultString.Measured = gridView.Rows[i].Cells[3].Value.ToString().Replace(",", " ");
                        } catch { }
                        try {
                            resultString.Result = gridView.Rows[i].Cells[4].Value.ToString();
                        } catch { }
                        jsonConvert.ResultString.Add(resultString);
                    }
                    string jsonString = new JavaScriptSerializer().Serialize(jsonConvert);
                    string pathMis = Folder.list[5] + "\\" + tb_userID.Text + "_" + textboxSN.Text + "_";
                    pathMis += tb_wo.Text.Replace("/", "-") + "_" + cbb_fg.Text + "_";
                    pathMis += labelStatus.Text + "_" + DateTime.Now.ToString("dd-MM-yyyy-HH-mm-ss", 
                        CultureInfo.CreateSpecificCulture("en-US")) + ".txt";
                    File.WriteAllText(pathMis, jsonString);

                    WaitUpData(pathMis);
                }

                //bool digit_snn = true;
                //if (t.Text.Length != Convert.ToInt32(prismTest.digitSN)) { digit_snn = false; Log(LogMsgType.Error_Red, "\n" + "sn prism not " + prismTest.digitSN + " digit"); }
                //if (radioOperationMode.Checked == true && TeamPrecision.PRISM.cSettingValues.TestingMode != "Debug" && digit_snn && flag_sn_pass[select_test - 1])
                //{
                //    DataSummary = DataSummary.Replace("'", "");
                //    string fctx = TeamPrecision.PRISM.cSettingValues.ProcessName.Trim().ToString();
                //    string MsgPRISM = "";
                //    for (int hgf = 1; hgf <= 5; hgf++)
                //    {
                //        //up_prism_timeout_sn = t.Text;
                //        //up_prism_timeout_result = "FAIL";
                //        //up_prism_timeout_dataSummary = DataSummary;
                //        //if (results) up_prism_timeout_result = "PASS";
                //        //else up_prism_timeout_result = "FAIL";
                //        //MsgPRISM = function_timeout(up_prism_timeout, 2500);
                //        if (results) MsgPRISM = TeamPrecision.PRISM.cResults.SaveTestResult(t.Text, "PASS", DataSummary);
                //        else MsgPRISM = TeamPrecision.PRISM.cResults.SaveTestResult(t.Text, "FAIL", DataSummary);
                //        Log(LogMsgType.Incoming_Blue, "\n" + "up data to prism: " + MsgPRISM);
                //        if (MsgPRISM == "SUCCESS") break;
                //        if (hgf != 5)
                //        {
                //            DelaymS(300);
                //            continue;
                //        }
                //        l.Text = "FAIL";
                //        l.ForeColor = Color.Red;
                //        csvPath = folder[2] + headfer + filename + ".csv";
                //        Log(LogMsgType.Error_Red, "\nPRISM fail while update data");
                //        GlobalTestingFlag[head - 1] = false;
                //    }
                //}
            }
            if (prismTest.mode == prismTest.Operation && prismTest.upDataToKomson) {
                DataGridView d = new DataGridView();
                switch (select_test) {//ไอศครีมเชอเบร็ตรสมะนาว
                    case 1: d = dataGridView_1; break;
                    case 2: d = dataGridView_2; break;
                    case 3: d = dataGridView_3; break;
                    case 4: d = dataGridView_4; break;
                    case 5: d = dataGridView_5; break;
                    case 6: d = dataGridView_6; break;
                    case 7: d = dataGridView_7; break;
                    case 8: d = dataGridView_8; break;
                    case 9: d = dataGridView_9; break;
                    case 10: d = dataGridView_10; break;
                    case 11: d = dataGridView_11; break;
                    case 12: d = dataGridView_12; break;
                    case 13: d = dataGridView_13; break;
                    case 14: d = dataGridView_14; break;
                    case 15: d = dataGridView_15; break;
                    case 16: d = dataGridView_16; break;
                    case 17: d = dataGridView_17; break;
                    case 18: d = dataGridView_18; break;
                    case 19: d = dataGridView_19; break;
                    case 20: d = dataGridView_20; break;
                    case 21: d = dataGridView_21; break;
                    case 22: d = dataGridView_22; break;
                    case 23: d = dataGridView_23; break;
                    case 24: d = dataGridView_24; break;
                    case 25: d = dataGridView_25; break;
                    case 26: d = dataGridView_26; break;
                    case 27: d = dataGridView_27; break;
                    case 28: d = dataGridView_28; break;
                    case 29: d = dataGridView_29; break;
                    case 30: d = dataGridView_30; break;
                    case 31: d = dataGridView_31; break;
                    case 32: d = dataGridView_32; break;
                    case 33: d = dataGridView_33; break;
                    case 34: d = dataGridView_34; break;
                    case 35: d = dataGridView_35; break;
                    case 36: d = dataGridView_36; break;
                }
                string q = cbb_fg.Text + ",";
                q += textboxSN.Text + ",";
                string ddd = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"));
                q += ddd + ",";
                if (results) q += "PASS,";
                else q += "FAIL,";
                q += tb_userID.Text + ",";
                q += tb_wo.Text + ",";
                q += setupPay.read_text(tester.headConfig.komsonTester, tester.nameFile) + ",";
                if (checkFailure != "") {
                    for (int i = 0; i < d.Rows.Count; i++) {
                        if (d.Rows[i].Cells[0].Value.ToString() != checkFailure.Replace("<>", "")) continue;
                        q += checkFailure + " " + d.Rows[i].Cells[1].Value.ToString() + ",";
                        q += d.Rows[i].Cells[3].Value.ToString() + ",";
                        break;
                    }
                } else q += ",,";
                q += head.ToString() + ",";
                File.WriteAllText("D:/Datalog to SQL/" + textboxSN.Text + ddd.Replace("-", "_").Replace(" ", "_").Replace(":", "_") + ".txt", q);
            }

            labelInOutTime.Text = "0000";

            if (prismTest.mode == prismTest.Operation)
            {
                try {
                    string[] strArrGetWO = TeamPrecision.PRISM.cSNs.getWO(tb_wo.Text, TeamPrecision.PRISM.cSettingValues.ProcessName);
                    tb_outputQty.Text = strArrGetWO[4];
                } catch { }
            }
        }
        private void WaitUpData(string path) {
            if (!upDataTest.waitUpData) {
                return;
            }

            while (true) {
                List<string> fileData = new List<string>();

                try {
                    string[] getFile = Directory.GetFiles(Folder.list[5]);
                    fileData = getFile.ToList<string>();
                } catch { }

                if (fileData.Contains(path)) {
                    DelaymS(25);
                    continue;
                }

                break;
            }
        }
        private string up_prism_timeout_sn = "";
        private string up_prism_timeout_result = "";
        private string up_prism_timeout_dataSummary = "";
        private string up_prism_timeout() {
            return TeamPrecision.PRISM.cResults.SaveTestResult(up_prism_timeout_sn, up_prism_timeout_result, up_prism_timeout_dataSummary);
        }

        public bool[] flag_test = { false, false, false, false, false, false, false, false, false, false,
                                    false, false, false, false, false, false, false, false, false, false,
                                    false, false, false, false, false, false, false, false, false, false,
                                    false, false, false, false, false, false};//คือ flag แสดงสถานะของ header ที่กำลังเทสอยู่
        private bool[] flag_head = { true, true, true, true, true, true, true, true, true, true,
                                     true, true, true, true, true, true, true, true, true, true,
                                     true, true, true, true, true, true, true, true, true, true,
                                     true, true, true, true, true, true};//คือ flag ที่แสดงสถานะของการบรรจุบอร์ดลงบน header
        private bool[] flag_loop = { false, false, false, false, false, false, false, false, false, false,
                                     false, false, false, false, false, false, false, false, false, false,
                                     false, false, false, false, false, false, false, false, false, false,
                                     false, false, false, false, false, false};//คือ flag ที่ให้รออ่านค่าจาก txt file ใช้คู่กับ โปรแกรมย่อย exe
        private int[] flag_txt = {  0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
                                    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0};//คือ flag ที่จำค่าลำดับ function ใน 1 ช่อง ของ excel
        public int select_test = 1;//คือ flag ที่แสดงสถานะ header ที่ถูกเทสอยู่ปัจจุบัน, flag นี้จะถูกนับวนลูปตามจำนวน header, เริ่มต้นจาก 1
        private bool flag_lock_select_test = false;//คือ flag ที่ล็อกให้เทสหัวใดหัวหนึ่ง ใช้คู่กับ excel ถ้าหากในตารางมีการทาสี ฟ้า ทับไว้ จะทำให้เทสหัวนั้นจนหมดสีฟ้าก่อน
        public int[] row_test = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
        private int[] flag_solo_step = { 0, 0, };
        private bool flag_goto_after_test = false;
        private void timer_run_Tick(object sender, EventArgs e) {
            bool b = false;
            for (int i = 0; i < tester.numHead; i++) {
                b |= flag_test[i];
            }
            if (b == true) {
                timer_run.Stop();
                if (select_test > tester.numHead) select_test = 1;
                if (flag_goto_after_test) {
                    flag_goto_after_test = false;
                    if (!flag_lock_select_test) {
                        if (select_test != 1) select_test--;
                        else select_test = tester.numHead;
                    }
                }
                run(select_test);
                if (!flag_lock_select_test) select_test++;
                timer_run.Start();
            }
        }
        private void run(int head) {
            DataGridView g = getDataGridView(select_test);
            TextBox t = getTextBoxSN(select_test);
            Label l = getLabelStatus(select_test);
            Button b_test = getButtonTest(select_test);
            ProgressBar p = getProgressBar(select_test);
            
            if (flag_test[select_test - 1] != true) return;
            if (GlobalTestingFlag[select_test - 1] != true) {
                string str = "after_test";
                try { Activator.CreateInstance(functionExcel, this, str); } catch {
                    Log(LogMsgType.Error_Red, "\ncall function " + str + " error");
                }
                flag_lock_select_test = false;
                b_test.Enabled = true;
                save_data(select_test);
                flag_test[select_test - 1] = false;
                return;
            }
            excel.workSheet = excel.workBook.Worksheets["info"];
            excel.sheetTest = excel.workSheet.GetText(44 + select_test, 2);
            excel.workSheet = excel.workBook.Worksheets[excel.sheetTest];
            if (row_test[select_test - 1] == 0) {
                time_start[select_test - 1] = DateTime.Now.Year.ToString() + "." + DateTime.Now.Month.ToString("00") + "." + DateTime.Now.Day.ToString("00") + " " + DateTime.Now.ToString("T");
                ClearDataGridView(select_test);
                excel.row[select_test - 1] = GetRowSequenceExcel();
                p.Maximum = GetRowSteptestExcel();
                GetDescription(select_test);
                p.Value = 0;
                rtfTerminal.Clear();
                flag_sn_pass[select_test - 1] = false;
                flagNotReTest[select_test - 1] = false;
                try {
                    Activator.CreateInstance(functionExcel, this, "intro_test()");
                } catch (Exception) {
                    Log(LogMsgType.Error_Red, "\ncall function intro_test() error");
                    GlobalTestingFlag[select_test - 1] = false;
                    return;
                }
                delete_txt_error();
                flag_txt[select_test - 1] = 0;
                row_test[select_test - 1] += 2;
            }
            if (row_test[select_test - 1] > excel.row[select_test - 1]) {
                string str = "after_test";
                try { Activator.CreateInstance(functionExcel, this, str); } catch (Exception) {
                    Log(LogMsgType.Error_Red, "\ncall function " + str + " error");
                }
                flag_lock_select_test = false;
                b_test.Enabled = true;
                save_data(select_test);
                flag_test[select_test - 1] = false;
                return;
            }
            string string_in_excel;
            string string_in_excel_Formula;
            string[] string_array_in_excel;
            if (excel.workSheet.Range[row_test[select_test - 1], 1].Style.KnownColor.ToString() == excel.color.skyBlue)
                flag_lock_select_test = true;
            else flag_lock_select_test = false;
            if (excel.workSheet.Range[row_test[select_test - 1], 1].Style.KnownColor.ToString() == excel.color.gold) {
                if (flag_solo_step[0] == 0) {
                    flag_solo_step[0] = 1;
                    flag_solo_step[1] = select_test;
                } else {
                    if (select_test != flag_solo_step[1]) return;
                }
            }
            if (excel.workSheet.Range[row_test[select_test - 1], 1].Style.KnownColor.ToString() != excel.color.none && 
                excel.workSheet.Range[row_test[select_test - 1], 1].Style.KnownColor.ToString() != excel.color.skyBlue && 
                excel.workSheet.Range[row_test[select_test - 1], 1].Style.KnownColor.ToString() != excel.color.gold) {
                row_test[select_test - 1]++;
                return;
            }
            string_in_excel = excel.workSheet.GetText(row_test[select_test - 1], 5);
            string_in_excel_Formula = excel.workSheet.GetFormulaStringValue(row_test[select_test - 1], 5);
            if (string_in_excel_Formula != null) string_in_excel = string_in_excel_Formula.Replace("\r\n", "\n");
            if (string_in_excel == "" || string_in_excel == " " || string_in_excel == "  " || string_in_excel == "   " || string_in_excel == null) {
                p.Value += 1;
                row_test[select_test - 1]++;
                return;
            }
            excel.stepTest = excel.workSheet.GetText(row_test[select_test - 1], 1);
            if (excel.stepTest == null) {
                excel.stepTest = excel.workSheet.GetNumber(row_test[select_test - 1], 1).ToString();
            }//if และ for ด้านล่างนี้ ไม่เกี่ยวอะไรกับ จำนวน header มันเป็นการดึงข้อมูลจาก excel ให้ตรงตำแหน่งเท่านั้น แนวแกน y หรือ คอลัม
            excel.alice[0] = excel.workSheet.GetNumber(row_test[select_test - 1], 1).ToString();
            if (excel.alice[0] == "NaN") excel.alice[0] = excel.workSheet.GetText(row_test[select_test - 1], 1);
            if (excel.alice[0] == null) excel.alice[0] = excel.workSheet.GetFormulaStringValue(row_test[select_test - 1], 1);
            excel.alice[1] = excel.workSheet.GetNumber(row_test[select_test - 1], 2).ToString();
            if (excel.alice[1] == "NaN") excel.alice[1] = excel.workSheet.GetText(row_test[select_test - 1], 2);
            if (excel.alice[1] == null) excel.alice[1] = excel.workSheet.GetFormulaStringValue(row_test[select_test - 1], 2);
            excel.alice[2] = excel.workSheet.GetNumber(row_test[select_test - 1], 3).ToString();
            if (excel.alice[2] == "NaN") excel.alice[2] = excel.workSheet.GetFormulaNumberValue(row_test[select_test - 1], 3).ToString();
            if (excel.alice[2] == "NaN") excel.alice[2] = excel.workSheet.GetText(row_test[select_test - 1], 3);
            if (excel.alice[2] == null) excel.alice[2] = excel.workSheet.GetFormulaStringValue(row_test[select_test - 1], 3);
            excel.alice[3] = excel.workSheet.GetNumber(row_test[select_test - 1], 4).ToString();
            if (excel.alice[3] == "NaN") excel.alice[3] = excel.workSheet.GetFormulaNumberValue(row_test[select_test - 1], 4).ToString();
            if (excel.alice[3] == "NaN") excel.alice[3] = excel.workSheet.GetText(row_test[select_test - 1], 4);
            if (excel.alice[3] == null) excel.alice[3] = excel.workSheet.GetFormulaStringValue(row_test[select_test - 1], 4);
            excel.alice[4] = excel.workSheet.GetNumber(row_test[select_test - 1], 5).ToString();
            if (excel.alice[4] == "NaN") excel.alice[4] = excel.workSheet.GetText(row_test[select_test - 1], 5);
            if (excel.alice[4] == null) excel.alice[4] = excel.workSheet.GetFormulaStringValue(row_test[select_test - 1], 5);
            for (int ii = 0; ii <= 4; ii++) {
                if (excel.alice[ii] == null) {
                    excel.alice[ii] = "";
                }
            }
            Log(LogMsgType.Incoming_Blue, "\n");
            if (string_in_excel.Contains("\n")) {
                string_array_in_excel = string_in_excel.Split('\n');
                for (int dsa = flag_txt[select_test - 1]; dsa < string_array_in_excel.Length; dsa++) {
                    if (GlobalTestingFlag[select_test - 1] == false) { flag_goto_after_test = true; break; }
                    string str = string_array_in_excel[dsa];
                    if (str.Substring(0, 1) == "#") continue;
                    Log(LogMsgType.Incoming_Blue, "\n" + select_test + ". " + str);
                    if (str.Substring(0, 2) == ">>") {
                        if (tester.showCMD) {
                            if (!run_call_exe(str, ">>")) continue;
                            return;
                        } else {
                            if (!run_call_exe_black(str, ">>")) continue;
                            return;
                        }
                    }
                    if (str.Substring(0, 3) == ">|>") {
                        if (tester.showCMD) {
                            if (!run_call_exe_black(str, ">|>")) continue;
                            return;
                        } else {
                            if (!run_call_exe(str, ">|>")) continue;
                            return;
                        }
                    }
                    if (str.Substring(0, 2) == "<<") {
                        string[] sup;
                        try {
                            sup = File.ReadAllLines("test_head_" + select_test + "_" + str.Replace("<<", ""));
                            timer_result[select_test - 1].Reset();
                        } catch {
                            if (timer_result[select_test - 1].ElapsedMilliseconds > timeout_result[select_test - 1]) {
                                UpdateResultToDataGrid(excel.alice[0], "*FAIL", "FAIL");
                                return;
                            }
                            flag_loop[select_test - 1] = true;
                            flag_txt[select_test - 1] = dsa;
                            break;
                        }
                        File.Delete("test_head_" + select_test + "_" + str.Replace("<<", ""));
                        flag_solo_step[0] = 0;
                        try {
                            if (sup[1] != "NODISPLAY") UpdateResultToDataGrid(excel.alice[0], sup[0], sup[1]);
                        } catch {
                            Log(LogMsgType.Error_Red, "\nformat error " + "test_head_" + select_test + "_" + str.Replace("<<", ""));
                            UpdateResultToDataGrid(excel.alice[0], "Fail", "FAIL");
                            GlobalTestingFlag[select_test - 1] = false;
                            return;
                        }
                        flag_txt[select_test - 1] = 0;
                        continue;
                    }
                    if (str.Substring(0, 3) == "<|<") {
                        string[] sup = { };
                        while (flag_test[select_test - 1]) {
                            try {
                                sup = File.ReadAllLines("test_head_" + select_test + "_" + str.Replace("<|<", ""));
                                timer_result[select_test - 1].Reset();
                            } catch {
                                if (timer_result[select_test - 1].ElapsedMilliseconds > timeout_result[select_test - 1]) {
                                    UpdateResultToDataGrid(excel.alice[0], "*FAIL", "FAIL");
                                    return;
                                }
                                Log(LogMsgType.Incoming_Blue, "\n" + select_test + ". " + str);
                                DelaymS(50);
                                continue;
                            }
                            break;
                        }
                        File.Delete("test_head_" + select_test + "_" + str.Replace("<|<", ""));
                        try {
                            if (sup[1] != "NODISPLAY") UpdateResultToDataGrid(excel.alice[0], sup[0], sup[1]);
                        } catch (Exception) {
                            Log(LogMsgType.Error_Red, "\nformat error " + "test_head_" + select_test + "_" + str.Replace("<|<", ""));
                            UpdateResultToDataGrid(excel.alice[0], "Fail", "FAIL");
                            GlobalTestingFlag[select_test - 1] = false;
                            return;
                        }
                        flag_txt[select_test - 1] = 0;
                        continue;
                    }
                    try {
                        Activator.CreateInstance(functionExcel, this, str);
                    } catch (Exception) {
                        Log(LogMsgType.Error_Red, "\ncall function " + str + " error");
                        UpdateResultToDataGrid(excel.alice[0], "Fail", "FAIL");
                        GlobalTestingFlag[select_test - 1] = false;
                        return;
                    }
                }
            } else {
                for (int oo = 0; oo < 1; oo++) {
                    if (string_in_excel.Substring(0, 1) == "#") continue;
                    Log(LogMsgType.Incoming_Blue, "\n" + select_test + ". " + string_in_excel);
                    if (string_in_excel.Substring(0, 2) == ">>") {
                        if (tester.showCMD) {
                            if (!run_call_exe(string_in_excel, ">>")) continue;
                            return;
                        } else {
                            if (!run_call_exe_black(string_in_excel, ">>")) continue;
                            return;
                        }
                    }
                    if (string_in_excel.Substring(0, 3) == ">|>") {
                        if (tester.showCMD) {
                            if (!run_call_exe_black(string_in_excel, ">|>")) continue;
                            return;
                        } else {
                            if (!run_call_exe(string_in_excel, ">|>")) continue;
                            return;
                        }
                    }
                    if (string_in_excel.Substring(0, 2) == "<<") {
                        string[] sup;
                        try {
                            sup = File.ReadAllLines("test_head_" + select_test + "_" + string_in_excel.Replace("<<", ""));
                        } catch (Exception) {
                            flag_loop[select_test - 1] = true;
                            continue;
                        }
                        File.Delete("test_head_" + select_test + "_" + string_in_excel.Replace("<<", ""));
                        flag_solo_step[0] = 0;
                        try {
                            if (sup[1] != "NODISPLAY") UpdateResultToDataGrid(excel.alice[0], sup[0], sup[1]);
                        } catch (Exception) {
                            Log(LogMsgType.Error_Red, "\nformat error " + "test_head_" + select_test + "_" + string_in_excel.Replace("<<", ""));
                            UpdateResultToDataGrid(excel.alice[0], "Fail", "FAIL");
                            GlobalTestingFlag[select_test - 1] = false;
                            return;
                        }
                        continue;
                    }
                    if (string_in_excel.Substring(0, 3) == "<|<") {
                        string[] sup = { };
                        while (flag_test[select_test - 1]) {
                            try { sup = File.ReadAllLines("test_head_" + select_test + "_" + string_in_excel.Replace("<|<", "")); } catch (Exception) {
                                Log(LogMsgType.Incoming_Blue, "\n" + select_test + ". " + string_in_excel);
                                DelaymS(50);
                                continue;
                            }
                            break;
                        }
                        File.Delete("test_head_" + select_test + "_" + string_in_excel.Replace("<|<", ""));
                        try {
                            if (sup[1] != "NODISPLAY") UpdateResultToDataGrid(excel.alice[0], sup[0], sup[1]);
                        } catch (Exception) {
                            Log(LogMsgType.Error_Red, "\nformat error " + "test_head_" + select_test + "_" + string_in_excel.Replace("<|<", ""));
                            UpdateResultToDataGrid(excel.alice[0], "Fail", "FAIL");
                            GlobalTestingFlag[select_test - 1] = false;
                            return;
                        }
                        continue;
                    }
                    try {
                        Activator.CreateInstance(functionExcel, this, string_in_excel);
                    } catch (Exception) {
                        Log(LogMsgType.Error_Red, "\ncall function " + string_in_excel + " error");
                        UpdateResultToDataGrid(excel.alice[0], "Fail", "FAIL");
                        GlobalTestingFlag[select_test - 1] = false;
                        return;
                    }
                }
            }
            if (flag_loop[select_test - 1] != true) {
                p.Value += 1;
                row_test[select_test - 1]++;
            } else {
                for (int i = 0; i < tester.numHead; i++) {
                    flag_loop[i] = false;
                }
            }
        }
        private bool run_call_exe(string s, string k) {
            if (call_exe(s.Replace(k, "")) == true) return false;
            Log(LogMsgType.Error_Red, "\ncall function " + s + " error");
            UpdateResultToDataGrid(excel.alice[0], "Fail", "FAIL");
            GlobalTestingFlag[select_test - 1] = false;
            return true;
        }
        private bool run_call_exe_black(string s, string k) {
            if (call_exe_black(s.Replace(k, "")) == true) return false;
            Log(LogMsgType.Error_Red, "\ncall function " + s + " error");
            UpdateResultToDataGrid(excel.alice[0], "Fail", "FAIL");
            GlobalTestingFlag[select_test - 1] = false;
            return true;
        }
        private Stopwatch[] timer_result = { new Stopwatch(), new Stopwatch(), new Stopwatch(), new Stopwatch(), new Stopwatch(),
                                             new Stopwatch(), new Stopwatch(), new Stopwatch(), new Stopwatch(), new Stopwatch(),
                                             new Stopwatch(), new Stopwatch(), new Stopwatch(), new Stopwatch(), new Stopwatch(),
                                             new Stopwatch(), new Stopwatch(), new Stopwatch(), new Stopwatch(), new Stopwatch() };
        private int[] timeout_result = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
        private bool call_exe(string str) {
            timeout_result[select_test - 1] = 99000;
            try {
                string[] zxc = str.Replace(".exe", "#").Split('#');
                timeout_result[select_test - 1] = Convert.ToInt32(zxc[1]) * 1000;
                str = zxc[0] + ".exe";
            } catch { }
            switch (select_test)  //ไอศครีมเชอเบร็ตรสมะนาว
            {
                case 1: if (set_debug_1.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 2: if (set_debug_2.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 3: if (set_debug_3.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 4: if (set_debug_4.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 5: if (set_debug_5.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 6: if (set_debug_6.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 7: if (set_debug_7.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 8: if (set_debug_8.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 9: if (set_debug_9.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 10: if (set_debug_10.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 11: if (set_debug_11.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 12: if (set_debug_12.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 13: if (set_debug_13.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 14: if (set_debug_14.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 15: if (set_debug_15.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 16: if (set_debug_16.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 17: if (set_debug_17.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 18: if (set_debug_18.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 19: if (set_debug_19.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 20: if (set_debug_20.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 21: if (set_debug_21.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 22: if (set_debug_22.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 23: if (set_debug_23.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 24: if (set_debug_24.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 25: if (set_debug_25.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 26: if (set_debug_26.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 27: if (set_debug_27.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 28: if (set_debug_28.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 29: if (set_debug_29.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 30: if (set_debug_30.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 31: if (set_debug_31.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 32: if (set_debug_32.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 33: if (set_debug_33.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 34: if (set_debug_34.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 35: if (set_debug_35.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 36: if (set_debug_36.Checked) timeout_result[select_test - 1] = 99999999; break;
            }
            File.WriteAllText("../../config/head.txt", select_test.ToString());
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.CreateNoWindow = false;
            startInfo.UseShellExecute = false;
            startInfo.FileName = "../../mini_projeck/" + str;
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            File.Delete("call_exe_tric.txt");
            try {
                Process.Start(startInfo);
            } catch {
                return false;
            }
            while (true) {
                try { File.ReadAllText("call_exe_tric.txt"); break; } catch { }
                DelaymS(50);
            }
            timer_result[select_test - 1].Restart();
            return true;
        }
        private bool call_exe_black(string str) {
            timeout_result[select_test - 1] = 99000;
            try {
                string[] zxc = str.Replace(".exe", "#").Split('#');
                timeout_result[select_test - 1] = Convert.ToInt32(zxc[1]) * 1000;
                str = zxc[0] + ".exe";
            } catch { }
            switch (select_test) //ไอศครีมเชอเบร็ตรสมะนาว
            {
                case 1: if (set_debug_1.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 2: if (set_debug_2.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 3: if (set_debug_3.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 4: if (set_debug_4.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 5: if (set_debug_5.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 6: if (set_debug_6.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 7: if (set_debug_7.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 8: if (set_debug_8.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 9: if (set_debug_9.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 10: if (set_debug_10.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 11: if (set_debug_11.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 12: if (set_debug_12.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 13: if (set_debug_13.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 14: if (set_debug_14.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 15: if (set_debug_15.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 16: if (set_debug_16.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 17: if (set_debug_17.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 18: if (set_debug_18.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 19: if (set_debug_19.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 20: if (set_debug_20.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 21: if (set_debug_21.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 22: if (set_debug_22.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 23: if (set_debug_23.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 24: if (set_debug_24.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 25: if (set_debug_25.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 26: if (set_debug_26.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 27: if (set_debug_27.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 28: if (set_debug_28.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 29: if (set_debug_29.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 30: if (set_debug_30.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 31: if (set_debug_31.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 32: if (set_debug_32.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 33: if (set_debug_33.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 34: if (set_debug_34.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 35: if (set_debug_35.Checked) timeout_result[select_test - 1] = 99999999; break;
                case 36: if (set_debug_36.Checked) timeout_result[select_test - 1] = 99999999; break;
            }
            File.WriteAllText("../../config/head.txt", select_test.ToString());
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.CreateNoWindow = true;
            startInfo.UseShellExecute = false;
            startInfo.FileName = "../../mini_projeck/" + str;
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            File.Delete("call_exe_tric.txt");
            try {
                Process.Start(startInfo);
            } catch {
                return false;
            }
            while (true) {
                try { File.ReadAllText("call_exe_tric.txt"); break; } catch { }
                DelaymS(50);
            }
            timer_result[select_test - 1].Restart();
            return true;
        }
        private void delete_txt_error() {
            for (int hh = 1; hh <= excel.row[select_test - 1]; hh++) {
                string string_in_excel;
                string[] string_array_in_excel;
                if (excel.workSheet.Range[hh, 1].Style.KnownColor.ToString() != excel.color.none && 
                    excel.workSheet.Range[hh, 1].Style.KnownColor.ToString() != excel.color.skyBlue && 
                    excel.workSheet.Range[hh, 1].Style.KnownColor.ToString() != excel.color.gold) {
                    continue;
                }
                string_in_excel = excel.workSheet.GetText(hh, 5);
                if (string_in_excel == "" || string_in_excel == " " || string_in_excel == "  " || string_in_excel == "   " || 
                    string_in_excel == null) {
                    continue;
                }
                excel.stepTest = excel.workSheet.GetText(hh, 1);
                if (excel.stepTest == null) {
                    excel.stepTest = excel.workSheet.GetNumber(hh, 1).ToString();
                }
                if (excel.workSheet.GetNumber(hh, 1).ToString() == "NaN") {
                    excel.alice[0] = excel.workSheet.GetText(hh, 1);
                } else {
                    excel.alice[0] = excel.workSheet.GetNumber(hh, 1).ToString();
                }
                if (excel.workSheet.GetNumber(hh, 2).ToString() == "NaN") {
                    excel.alice[1] = excel.workSheet.GetText(hh, 2);
                } else {
                    excel.alice[1] = excel.workSheet.GetNumber(hh, 2).ToString();
                }
                if (excel.workSheet.GetNumber(hh, 3).ToString() == "NaN") {
                    excel.alice[2] = excel.workSheet.GetText(hh, 3);
                } else {
                    excel.alice[2] = excel.workSheet.GetNumber(hh, 3).ToString();
                }
                if (excel.workSheet.GetNumber(hh, 4).ToString() == "NaN") {
                    excel.alice[3] = excel.workSheet.GetText(hh, 4);
                } else {
                    excel.alice[3] = excel.workSheet.GetNumber(hh, 4).ToString();
                }
                if (excel.workSheet.GetNumber(hh, 6).ToString() == "NaN") {
                    excel.alice[4] = excel.workSheet.GetText(hh, 6);
                } else {
                    excel.alice[4] = excel.workSheet.GetNumber(hh, 6).ToString();
                }
                for (int ii = 0; ii <= 4; ii++) {
                    if (excel.alice[ii] == null) {
                        excel.alice[ii] = "";
                    }
                }
                if (string_in_excel.Contains("\n")) {
                    string_in_excel = string_in_excel.Replace("\n", "&");
                    string_array_in_excel = string_in_excel.Split('&');
                    for (int dsa = flag_txt[select_test - 1]; dsa < string_array_in_excel.Length; dsa++) {
                        string str = string_array_in_excel[dsa];
                        if (str.Substring(0, 1) == "#") continue;
                        if (str.Substring(0, 2) == "<<") {
                            File.Delete("test_head_" + select_test + "_" + str.Replace("<<", ""));
                        }
                        if (str.Substring(0, 3) == "<|<") {
                            File.Delete("test_head_" + select_test + "_" + str.Replace("<|<", ""));
                        }
                    }
                } else {
                    for (int oo = 0; oo < 1; oo++) {
                        if (string_in_excel.Substring(0, 1) == "#") continue;
                        if (string_in_excel.Substring(0, 2) == "<<") {
                            File.Delete("test_head_" + select_test + "_" + string_in_excel.Replace("<<", ""));
                        }
                        if (string_in_excel.Substring(0, 3) == "<|<") {
                            File.Delete("test_head_" + select_test + "_" + string_in_excel.Replace("<|<", ""));
                        }
                    }
                }
            }
        }

        bool CheckHeader(string CurrentHeader, string f, string header, int index_folder) {
            string TodayFile = null;
            string PreviousHeader = null;
            string[] CurrentHeader_split = CurrentHeader.Split(',');

            string filename = Folder.list[index_folder] + header + f + ".csv";
            TodayFile = filename;
            const Int32 BufferSize = 500;
            try {
                using (var fileStream = File.OpenRead(TodayFile))//=====================                                               
                using (var streamReader = new StreamReader(fileStream, Encoding.UTF8, true, BufferSize)) {
                    String line;
                    while ((line = streamReader.ReadLine()) != null) {
                        List<string> names = new List<string>(line.Split(','));
                        if (names[0] == "Date" || names[0] == CurrentHeader_split[0]) PreviousHeader = line;
                    }
                    fileStream.Close();
                    if (string.Compare(CurrentHeader, PreviousHeader) != 0)
                        return false;
                }
            } catch (Exception) { }
            return true;
        }
        public static void DelaymS(int mS) {
            Stopwatch stopwatchDelaymS = new Stopwatch();
            stopwatchDelaymS.Restart();
            while (mS > stopwatchDelaymS.ElapsedMilliseconds) {
                if (!stopwatchDelaymS.IsRunning) stopwatchDelaymS.Start();
                Application.DoEvents();
            }
            stopwatchDelaymS.Stop();
        }
        public void UpdateResultToDataGrid(string strStep, string strValue, string strResult = "FAIL") {
            string strBuf = "";
            DataGridView g = getDataGridView(select_test);

            for (int i = 0; i < g.RowCount; i++) {
                try { strBuf = g.Rows[i].Cells[0].Value.ToString(); } catch (Exception) { g.Rows[i].Cells[4].Value = ""; continue; }
                if (strBuf != strStep) {
                    if (g.Rows[i].Cells[4].Value == null) g.Rows[i].Cells[4].Value = "";
                    continue;
                }
                g.Rows[i].Cells[3].Value = strValue;
                if (strResult == "PASS" || strResult == "Pass" || strResult == "pass") {
                    g.Rows[i].Cells[4].Style.ForeColor = Color.Green;
                } else if (strResult == "FAIL" || strResult == "Fail" || strResult == "fail") {
                    g.Rows[i].Cells[4].Style.ForeColor = Color.Red;
                    GlobalTestingFlag[select_test - 1] = false;
                }
                g.Rows[i].Cells[4].Value = strResult;
                if (i > Convert.ToInt32(tester.ScrollDatagrid) && g.Visible) g.FirstDisplayedScrollingRowIndex = i - Convert.ToInt32(tester.ScrollDatagrid);
                i = g.RowCount;
                //DelaymS(5);
                break;
            }
        }
        private void CheckPRISMStatus() {
            rtfTerminal.Clear();
            //lblUserID.Text = TeamPrecision.PRISM.cSettingValues.EmployeeID;
            setupPay.setup();
            tb_userID.Text = setupPay.read_text(prismTest.headConfig.employeeID, prismTest.nameFile);
            Log(LogMsgType.Warning_Orange, "\n========= PRISM Info ========");
            //Log(LogMsgType.Outgoing_Green, "\n- TestingMode = " + TeamPrecision.PRISM.cSettingValues.TestingMode);
            //Log(LogMsgType.Outgoing_Green, "\n- Computer name = " + TeamPrecision.PRISM.cSettingValues.ComputerName);
            //Log(LogMsgType.Outgoing_Green, "\n- Station name = " + TeamPrecision.PRISM.cSettingValues.StationName);
            //Log(LogMsgType.Outgoing_Green, "\n- Process name = " + TeamPrecision.PRISM.cSettingValues.ProcessName);
            //Log(LogMsgType.Outgoing_Green, "\n- Employee ID = " + TeamPrecision.PRISM.cSettingValues.EmployeeID);
            //Log(LogMsgType.Outgoing_Green, "\n- Employee ID = " + TeamPrecision.PRISM.cSettingValues.DatabaseServer);
            Log(LogMsgType.Outgoing_Green, "\n- TestingMode = " + prismTest.mode);
            Log(LogMsgType.Outgoing_Green, "\n- Computer name = " + setupPay.read_text(prismTest.headConfig.computerName, prismTest.nameFile));
            Log(LogMsgType.Outgoing_Green, "\n- Station name = " + setupPay.read_text(prismTest.headConfig.stationName, prismTest.nameFile));
            Log(LogMsgType.Outgoing_Green, "\n- Process name = " + setupPay.read_text(prismTest.headConfig.processName, prismTest.nameFile));
            Log(LogMsgType.Outgoing_Green, "\n- Employee ID = " + setupPay.read_text(prismTest.headConfig.employeeID, prismTest.nameFile));
            Log(LogMsgType.Outgoing_Green, "\n- Database Server = " + setupPay.read_text(prismTest.headConfig.databaseServerTPP, prismTest.nameFile));
            Log(LogMsgType.Warning_Orange, "\n============================");

            if (prismTest.mode == prismTest.Debug)
            {
                cb_DebugMode.Checked = true; cb_OperationMode.Checked = false;
                tb_wo.Enabled = false;
                this.BackColor = Color.Gold;
            }
            else if (prismTest.mode == prismTest.Operation)
            {
                cb_OperationMode.Checked = true; cb_DebugMode.Checked = false;
                cbb_fg.Enabled = false;
                this.BackColor = default(Color);
                this.ForeColor = default(Color);
            }
            else
            {
                Log(LogMsgType.Error_Red, "\n- Selecting PRISM mode Err");
                GlobalTestingFlag[0] = false;
            }
        }
        private void enable_button_test_and_exit() {
            for (int i = 1; i <= tester.numHead; i++) {
                Button b_test = getButtonTest(i);
                Button b_exit = getButtonExit(i);
                DataGridView g = getDataGridView(i);
                
                if (g.Rows[0].Cells[0].Value == null) continue;
                b_test.Enabled = true;
                b_exit.Enabled = true;
            }
            txtSNBoard_1.Focus();
        }
        private void getWorkOrder(string WO) {
            string[] strArrGetWO = { "", "", "", "", "" };
            try {
                strArrGetWO = TeamPrecision.PRISM.cSNs.getWO(tb_wo.Text, setupPay.read_text(prismTest.headConfig.processName, prismTest.nameFile));
            } catch (Exception) {
                rtfTerminal.Clear();
                Log(LogMsgType.Error_Red, "\nPRISM Err at getWorkOrder functiom");
            }
            if (strArrGetWO[0] == "SUCCESS") {
                tb_fwVersion.ForeColor = Color.Green;
                tb_spec.ForeColor = Color.Green;
                tb_detail.ForeColor = Color.Green;
                cbb_fg.Items.Clear();
                cbb_fg.Items.Add(strArrGetWO[1]);
                cbb_fg.SelectedIndex = 0;
                tb_orderQty.Text = strArrGetWO[3]; tb_orderQty.ForeColor = Color.Green;
                tb_outputQty.Text = strArrGetWO[4]; tb_outputQty.ForeColor = Color.Green;
                txtSNBoard_1.Focus();
                string nameFG = strArrGetWO[1];
                try {
                    excel.workBook.LoadFromFile("../../TestDescription/" + nameFG + excel.lastName);
                } catch {
                    Log(LogMsgType.Error_Red, "_ปิด excel ก่อน" + nameFG + excel.lastName + "\n");
                    return;
                }
                LoadDiscription();
                Activator.CreateInstance(functionExcel, this, "LoadTestSpec()");
            } else {
                rtfTerminal.Clear();
                tb_fwVersion.Text = "XXXX"; tb_fwVersion.ForeColor = Color.Red;
                tb_spec.Text = "XXXX"; tb_spec.ForeColor = Color.Red;
                tb_detail.Text = "XXXX"; tb_detail.ForeColor = Color.Red;
                tb_orderQty.Text = "0000"; tb_orderQty.ForeColor = Color.Red;
                tb_outputQty.Text = "0000"; tb_outputQty.ForeColor = Color.Red;
                Log(LogMsgType.Error_Red, "-ไม่พบข้อมูล WO " + tb_wo.Text + " ในระบบ.!");
                for (int i = 0; i < strArrGetWO.Length; i++) Log(LogMsgType.Error_Red, "\n" + strArrGetWO[i]);
                tb_wo.SelectAll();
                tb_wo.Focus();
            }
            enable_button_test_and_exit();
        }
        private void CallUpDataExe()
        {
            if (!upDataTest.callExe)
            {
                return;
            }

            Process[] process = Process.GetProcessesByName("up_data");
            if (process.Length == 0)
            {
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.CreateNoWindow = false;
                startInfo.UseShellExecute = false;
                startInfo.FileName = "up_data.exe";
                startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                try
                {
                    Process.Start(startInfo);
                }
                catch { }
            }
        }
        private void CloseUpDataExe()
        {
            if (!upDataTest.callExe)
            {
                return;
            }

            Process[] process = Process.GetProcessesByName("up_data");
            for (int num = 0; num < process.Count(); num++)
            {
                try
                {
                    process[num].Kill();
                }
                catch { }
            }
        }
        private void set_default() {
            tester.useRelayCard = Convert.ToBoolean(setupPay.read_text(tester.headConfig.useRelayCard, tester.nameFile));
            
            for (int i = 1; i <= tester.numHead; i++) {
                ToolStripMenuItem c = getToolStripMenuItemDebug(i);

                if (!tester.useRelayCard) flag_head[i - 1] = true;
                else flag_head[i - 1] = false;
                c.Visible = true;
                File.WriteAllText("../../config/test_head_" + i + "_debug.txt", c.Checked.ToString());
            }
            tester.ScrollDatagrid = setupPay.read_text(tester.headConfig.ScrollDatagrid, tester.nameFile);

            bool flagAutomation = Convert.ToBoolean(setupPay.read_text(tester.headConfig.automation, tester.nameFile));
            if (flagAutomation) tester.automation = true;
            else {
                tester.automation = false;
                setFlagAutomation1.Enabled = false;
                setFlagAutomation1.Visible = false;
                setFlagAutomation2.Enabled = false;
                setFlagAutomation2.Visible = false;
                tester.cylinder1 = false;
                tester.cylinder2 = false;
            }
            autoTestToolStripMenuItem.Checked = Convert.ToBoolean(setupPay.read_text(tester.headConfig.testAuto, tester.nameFile));
            tcpIP.ip = setupPay.read_text(tcpIP.headConfig.ip, tcpIP.nameFile);
            tcpIP.port = setupPay.read_text(tcpIP.headConfig.port, tcpIP.nameFile);
            tcpIP.timerTric = setupPay.read_text(tcpIP.headConfig.timerTric, tcpIP.nameFile);
            prism_retest.Checked = Convert.ToBoolean(setupPay.read_text(tester.headConfig.allowRetest, tester.nameFile));
            prismTest.processBefore = Convert.ToBoolean(setupPay.read_text(prismTest.headConfig.checkProcessBefore, prismTest.nameFile));
            tcpIP.useRobot = Convert.ToBoolean(setupPay.read_text(tcpIP.headConfig.useRobot, tcpIP.nameFile));
            tester.upFail = Convert.ToBoolean(setupPay.read_text(tester.headConfig.upFail, tester.nameFile));
            tester.saveData = setupPay.read_text(tester.headConfig.saveData, tester.nameFile);
            prismTest.digitSN = setupPay.read_text(prismTest.headConfig.digitSN, prismTest.nameFile);
            prismTest.upDataToKomson = Convert.ToBoolean(setupPay.read_text(prismTest.headConfig.upDataToKomson, prismTest.nameFile));
            try { prism_retest_text_pass.Text = File.ReadAllText("../../config/prism_retest_text_pass.txt"); } catch { }
            try { prism_retest_text_fail.Text = File.ReadAllText("../../config/prism_retest_text_fail.txt"); } catch { }
            prismTest.processBeforeText = setupPay.read_text(prismTest.headConfig.ProcessBefore, prismTest.nameFile);
            tester.click2ClearSN = Convert.ToBoolean(setupPay.read_text(tester.headConfig.click2ClearSN, tester.nameFile));
            tester.showCMD = Convert.ToBoolean(setupPay.read_text(tester.headConfig.showCMD, tester.nameFile));
            ctms_showCmd.Checked = tester.showCMD;
            tester.cylinderRelay1 = setupPay.read_text(tester.headConfig.cylinderRelay1, tester.nameFile); 
            tester.cylinderHead1 = setupPay.read_text(tester.headConfig.cylinderHead1, tester.nameFile);
            tester.cylinderRelay2 = setupPay.read_text(tester.headConfig.cylinderRelay2, tester.nameFile);
            tester.cylinderHead2 = setupPay.read_text(tester.headConfig.cylinderHead2, tester.nameFile); 
            tcpIP.readRobot1 = setupPay.read_text(tcpIP.headConfig.readRobot1, tcpIP.nameFile);
            tcpIP.readRobot2 = setupPay.read_text(tcpIP.headConfig.readRobot2, tcpIP.nameFile);
            tester.numCardRelay = setupPay.read_text(tester.headConfig.numCardRelay, tester.nameFile);
            tester.testPanel = Convert.ToBoolean(setupPay.read_text(tester.headConfig.testPanel, tester.nameFile));
            tester.nameDMM = setupPay.read_text(tester.headConfig.nameDMM, tester.nameFile);
            dataLog.timeLine.numFile = setupPay.read_text(dataLog.headConfig.fileTimeLine, dataLog.nameFile);
            if (setupPay.read_text(tester.headConfig.fileTestDescription, tester.nameFile) == tester.excel) {
                tester.selectExcel = true;
                tester.selectLibre = false;
            } else {
                tester.selectExcel = false;
                tester.selectLibre = true;
            }
            prismTest.mode = setupPay.read_text(prismTest.headConfig.mode, prismTest.nameFile);
            if (prismTest.mode == prismTest.Debug) {
                cb_OperationMode.Checked = false;
                cb_DebugMode.Checked = true;
            } else {
                cb_OperationMode.Checked = true;
                cb_DebugMode.Checked = false;
            }
            upDataTest.callExe = Convert.ToBoolean(setupPay.read_text(upDataTest.headConfig.callExe, upDataTest.nameFile));
            upDataTest.waitUpData = Convert.ToBoolean(setupPay.read_text(upDataTest.headConfig.waitUpData, upDataTest.nameFile));
            ctms_excel_saveFile_sup();
            try { File.Delete("auto_test_trick.txt"); } catch { }
        }
        #endregion

        #region ========================================================== Control Event ===========================================================
        private void Form1_Load(object sender, EventArgs e) {
            Process[] pname = Process.GetProcessesByName("WindowsFormsApplication1");
            if (pname.Length == 2) {
                MessageBox.Show("_โปรแกรมนี้ เปิดใช้งานอยู่");
                Application.Exit();
                return;
            }

            set_default();
            CallUpDataExe();

            File.WriteAllText(upDataTest.reReadConfig, string.Empty);

            if (setupPay.read_text(prismTest.headConfig.mode, prismTest.nameFile) == prismTest.Operation) {
                this.WindowState = FormWindowState.Minimized;
                File.Delete("up_data_login_ok.txt");
                File.Delete("up_data_login_fail.txt");
                File.WriteAllText("up_data_login.txt", "");
                while (true) {
                    try {
                        File.ReadAllText("up_data_login_ok.txt");
                        File.Delete("up_data_login_ok.txt");
                        break;
                    } catch { }
                    try {
                        File.ReadAllText("up_data_login_fail.txt");
                        File.Delete("up_data_login_fail.txt");
                        Application.Exit();
                        return;
                    } catch { }
                    Thread.Sleep(100);
                }
                this.WindowState = FormWindowState.Maximized;
            }
            Cursor.Hide();
            CheckPRISMStatus();
            connect_relay(true);
            background_time.RunWorkerAsync();
            DirectoryInfo lish_excel = new DirectoryInfo("../../TestDescription");
            if (tester.selectExcel) {
                FileInfo[] Files = lish_excel.GetFiles("*.xlsx");
                foreach (FileInfo file in Files) {
                    if (file.Name.Contains("~")) continue;
                    cbb_fg.Items.Add(file.Name.Replace(file.Extension, ""));
                }
            }
            if (tester.selectLibre) {
                FileInfo[] Files = lish_excel.GetFiles("*.ods");
                foreach (FileInfo file in Files) {
                    if (file.Name.Contains("~")) continue;
                    cbb_fg.Items.Add(file.Name.Replace(file.Extension, ""));
                }
            }
            this.Activate();
            tb_wo.Focus();
            Activator.CreateInstance(functionExcel, this, "setup()");
            Cursor.Show();

            flagSizeChanged = true;
            fMain_sizechanged();
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e) {
            flag_this_close = true;
            if (RelayPort.IsOpen) OffAllRelay();
            if (cbb_fg.Text != "") Activator.CreateInstance(functionExcel, this, "OffAllFunction()");
            background_time.Dispose();
            background_relay.Dispose();
            automation_background.Dispose();
            tcpip.Dispose();
            for (int i = 1; i <= tester.numHead; i++) {
                File.Delete("../../config/test_head_" + i + "_debug.txt");
            }
            CloseUpDataExe();
            Application.Exit();
            Application.ExitThread();
            Environment.Exit(0);
        }
        private void Form1_ResizeEnd(object sender, EventArgs e) {
            return;
            ///If from size less than 1155, 625. Program will auto adjust to 1155, 625
            ///and set it to center position of screen
            if ((this.Size.Width < 1155) | (this.Size.Height < 625)) {
                this.Width = 1155;
                this.Height = 625;

                Screen screen = Screen.FromControl(this);
                Rectangle workingArea = screen.WorkingArea;
                this.Location = new Point() {
                    X = Math.Max(workingArea.X, workingArea.X + (workingArea.Width - this.Width) / 2),
                    Y = Math.Max(workingArea.Y, workingArea.Y + (workingArea.Height - this.Height) / 2)
                };
            }
        }
        private void btnTEST_1_Click(object sender, EventArgs e) {
            if (flag_head[0] != true) return;
            btnTEST_1.Enabled = false;
            row_test[0] = 0;
            flag_test[0] = true;
            GlobalTestingFlag[0] = true;
            status_1.Text = "TESTING";
            status_1.ForeColor = Color.Blue;
        }
        private void btnTEST_2_Click(object sender, EventArgs e) {
            if (flag_head[1] != true) return;
            btnTEST_2.Enabled = false;
            row_test[1] = 0;
            flag_test[1] = true;
            GlobalTestingFlag[1] = true;
            status_2.Text = "TESTING";
            status_2.ForeColor = Color.Blue;
        }
        private void btnTEST_3_Click(object sender, EventArgs e) {
            if (flag_head[2] != true) return;
            btnTEST_3.Enabled = false;
            row_test[2] = 0;
            flag_test[2] = true;
            GlobalTestingFlag[2] = true;
            status_3.Text = "TESTING";
            status_3.ForeColor = Color.Blue;
        }
        private void btnTEST_4_Click(object sender, EventArgs e) {
            if (flag_head[3] != true) return;
            btnTEST_4.Enabled = false;
            row_test[3] = 0;
            flag_test[3] = true;
            GlobalTestingFlag[3] = true;
            status_4.Text = "TESTING";
            status_4.ForeColor = Color.Blue;
        }
        private void btnTEST_5_Click(object sender, EventArgs e) {
            if (flag_head[4] != true) return;
            btnTEST_5.Enabled = false;
            row_test[4] = 0;
            flag_test[4] = true;
            GlobalTestingFlag[4] = true;
            status_5.Text = "TESTING";
            status_5.ForeColor = Color.Blue;
        }
        private void btnTEST_6_Click(object sender, EventArgs e) {
            if (flag_head[5] != true) return;
            btnTEST_6.Enabled = false;
            row_test[5] = 0;
            flag_test[5] = true;
            GlobalTestingFlag[5] = true;
            status_6.Text = "TESTING";
            status_6.ForeColor = Color.Blue;
        }
        private void btnTEST_7_Click(object sender, EventArgs e) {
            if (flag_head[6] != true) return;
            btnTEST_7.Enabled = false;
            row_test[6] = 0;
            flag_test[6] = true;
            GlobalTestingFlag[6] = true;
            status_7.Text = "TESTING";
            status_7.ForeColor = Color.Blue;
        }
        private void btnTEST_8_Click(object sender, EventArgs e) {
            if (flag_head[7] != true) return;
            btnTEST_8.Enabled = false;
            row_test[7] = 0;
            flag_test[7] = true;
            GlobalTestingFlag[7] = true;
            status_8.Text = "TESTING";
            status_8.ForeColor = Color.Blue;
        }
        private void btnTEST_9_Click(object sender, EventArgs e) {
            if (flag_head[8] != true) return;
            btnTEST_9.Enabled = false;
            row_test[8] = 0;
            flag_test[8] = true;
            GlobalTestingFlag[8] = true;
            status_9.Text = "TESTING";
            status_9.ForeColor = Color.Blue;
        }
        private void btnTEST_10_Click(object sender, EventArgs e) {
            if (flag_head[9] != true) return;
            btnTEST_10.Enabled = false;
            row_test[9] = 0;
            flag_test[9] = true;
            GlobalTestingFlag[9] = true;
            status_10.Text = "TESTING";
            status_10.ForeColor = Color.Blue;
        }
        private void btnTEST_11_Click(object sender, EventArgs e) {
            if (flag_head[10] != true) return;
            btnTEST_11.Enabled = false;
            row_test[10] = 0;
            flag_test[10] = true;
            GlobalTestingFlag[10] = true;
            status_11.Text = "TESTING";
            status_11.ForeColor = Color.Blue;
        }
        private void btnTEST_12_Click(object sender, EventArgs e) {
            if (flag_head[11] != true) return;
            btnTEST_12.Enabled = false;
            row_test[11] = 0;
            flag_test[11] = true;
            GlobalTestingFlag[11] = true;
            status_12.Text = "TESTING";
            status_12.ForeColor = Color.Blue;
        }
        private void btnTEST_13_Click(object sender, EventArgs e) {
            if (flag_head[12] != true) return;
            btnTEST_13.Enabled = false;
            row_test[12] = 0;
            flag_test[12] = true;
            GlobalTestingFlag[12] = true;
            status_13.Text = "TESTING";
            status_13.ForeColor = Color.Blue;
        }
        private void btnTEST_14_Click(object sender, EventArgs e) {
            if (flag_head[13] != true) return;
            btnTEST_14.Enabled = false;
            row_test[13] = 0;
            flag_test[13] = true;
            GlobalTestingFlag[13] = true;
            status_14.Text = "TESTING";
            status_14.ForeColor = Color.Blue;
        }
        private void btnTEST_15_Click(object sender, EventArgs e) {
            if (flag_head[14] != true) return;
            btnTEST_15.Enabled = false;
            row_test[14] = 0;
            flag_test[14] = true;
            GlobalTestingFlag[14] = true;
            status_15.Text = "TESTING";
            status_15.ForeColor = Color.Blue;
        }
        private void btnTEST_16_Click(object sender, EventArgs e) {
            if (flag_head[15] != true) return;
            btnTEST_16.Enabled = false;
            row_test[15] = 0;
            flag_test[15] = true;
            GlobalTestingFlag[15] = true;
            status_16.Text = "TESTING";
            status_16.ForeColor = Color.Blue;
        }
        private void btnTEST_17_Click(object sender, EventArgs e) {
            if (flag_head[16] != true) return;
            btnTEST_17.Enabled = false;
            row_test[16] = 0;
            flag_test[16] = true;
            GlobalTestingFlag[16] = true;
            status_17.Text = "TESTING";
            status_17.ForeColor = Color.Blue;
        }
        private void btnTEST_18_Click(object sender, EventArgs e) {
            if (flag_head[17] != true) return;
            btnTEST_18.Enabled = false;
            row_test[17] = 0;
            flag_test[17] = true;
            GlobalTestingFlag[17] = true;
            status_18.Text = "TESTING";
            status_18.ForeColor = Color.Blue;
        }
        private void btnTEST_19_Click(object sender, EventArgs e) {
            if (flag_head[18] != true) return;
            btnTEST_19.Enabled = false;
            row_test[18] = 0;
            flag_test[18] = true;
            GlobalTestingFlag[18] = true;
            status_19.Text = "TESTING";
            status_19.ForeColor = Color.Blue;
        }
        private void btnTEST_20_Click(object sender, EventArgs e) {
            if (flag_head[19] != true) return;
            btnTEST_20.Enabled = false;
            row_test[19] = 0;
            flag_test[19] = true;
            GlobalTestingFlag[19] = true;
            status_20.Text = "TESTING";
            status_20.ForeColor = Color.Blue;
        }
        private void btnTEST_21_Click(object sender, EventArgs e) {
            if (flag_head[20] != true) return;
            btnTEST_21.Enabled = false;
            row_test[20] = 0;
            flag_test[20] = true;
            GlobalTestingFlag[20] = true;
            status_21.Text = "TESTING";
            status_21.ForeColor = Color.Blue;
        }
        private void btnTEST_22_Click(object sender, EventArgs e) {
            if (flag_head[21] != true) return;
            btnTEST_22.Enabled = false;
            row_test[21] = 0;
            flag_test[21] = true;
            GlobalTestingFlag[21] = true;
            status_22.Text = "TESTING";
            status_22.ForeColor = Color.Blue;
        }
        private void btnTEST_23_Click(object sender, EventArgs e) {
            if (flag_head[22] != true) return;
            btnTEST_23.Enabled = false;
            row_test[22] = 0;
            flag_test[22] = true;
            GlobalTestingFlag[22] = true;
            status_23.Text = "TESTING";
            status_23.ForeColor = Color.Blue;
        }
        private void btnTEST_24_Click(object sender, EventArgs e) {
            if (flag_head[23] != true) return;
            btnTEST_24.Enabled = false;
            row_test[23] = 0;
            flag_test[23] = true;
            GlobalTestingFlag[23] = true;
            status_24.Text = "TESTING";
            status_24.ForeColor = Color.Blue;
        }
        private void btnTEST_25_Click(object sender, EventArgs e) {
            if (flag_head[24] != true) return;
            btnTEST_25.Enabled = false;
            row_test[24] = 0;
            flag_test[24] = true;
            GlobalTestingFlag[24] = true;
            status_25.Text = "TESTING";
            status_25.ForeColor = Color.Blue;
        }
        private void btnTEST_26_Click(object sender, EventArgs e) {
            if (flag_head[25] != true) return;
            btnTEST_26.Enabled = false;
            row_test[25] = 0;
            flag_test[25] = true;
            GlobalTestingFlag[25] = true;
            status_26.Text = "TESTING";
            status_26.ForeColor = Color.Blue;
        }
        private void btnTEST_27_Click(object sender, EventArgs e) {
            if (flag_head[26] != true) return;
            btnTEST_27.Enabled = false;
            row_test[26] = 0;
            flag_test[26] = true;
            GlobalTestingFlag[26] = true;
            status_27.Text = "TESTING";
            status_27.ForeColor = Color.Blue;
        }
        private void btnTEST_28_Click(object sender, EventArgs e) {
            if (flag_head[27] != true) return;
            btnTEST_28.Enabled = false;
            row_test[27] = 0;
            flag_test[27] = true;
            GlobalTestingFlag[27] = true;
            status_28.Text = "TESTING";
            status_28.ForeColor = Color.Blue;
        }
        private void btnTEST_29_Click(object sender, EventArgs e) {
            if (flag_head[28] != true) return;
            btnTEST_29.Enabled = false;
            row_test[28] = 0;
            flag_test[28] = true;
            GlobalTestingFlag[28] = true;
            status_29.Text = "TESTING";
            status_29.ForeColor = Color.Blue;
        }
        private void btnTEST_30_Click(object sender, EventArgs e) {
            if (flag_head[29] != true) return;
            btnTEST_30.Enabled = false;
            row_test[29] = 0;
            flag_test[29] = true;
            GlobalTestingFlag[29] = true;
            status_30.Text = "TESTING";
            status_30.ForeColor = Color.Blue;
        }
        private void btnTEST_31_Click(object sender, EventArgs e) {
            if (flag_head[30] != true) return;
            btnTEST_31.Enabled = false;
            row_test[30] = 0;
            flag_test[30] = true;
            GlobalTestingFlag[30] = true;
            status_31.Text = "TESTING";
            status_31.ForeColor = Color.Blue;
        }
        private void btnTEST_32_Click(object sender, EventArgs e) {
            if (flag_head[31] != true) return;
            btnTEST_32.Enabled = false;
            row_test[31] = 0;
            flag_test[31] = true;
            GlobalTestingFlag[31] = true;
            status_32.Text = "TESTING";
            status_32.ForeColor = Color.Blue;
        }
        private void btnTEST_33_Click(object sender, EventArgs e) {
            if (flag_head[32] != true) return;
            btnTEST_33.Enabled = false;
            row_test[32] = 0;
            flag_test[32] = true;
            GlobalTestingFlag[32] = true;
            status_33.Text = "TESTING";
            status_33.ForeColor = Color.Blue;
        }
        private void btnTEST_34_Click(object sender, EventArgs e) {
            if (flag_head[33] != true) return;
            btnTEST_34.Enabled = false;
            row_test[33] = 0;
            flag_test[33] = true;
            GlobalTestingFlag[33] = true;
            status_34.Text = "TESTING";
            status_34.ForeColor = Color.Blue;
        }
        private void btnTEST_35_Click(object sender, EventArgs e) {
            if (flag_head[34] != true) return;
            btnTEST_35.Enabled = false;
            row_test[34] = 0;
            flag_test[34] = true;
            GlobalTestingFlag[34] = true;
            status_35.Text = "TESTING";
            status_35.ForeColor = Color.Blue;
        }
        private void btnTEST_36_Click(object sender, EventArgs e) {
            if (flag_head[35] != true) return;
            btnTEST_36.Enabled = false;
            row_test[35] = 0;
            flag_test[35] = true;
            GlobalTestingFlag[35] = true;
            status_36.Text = "TESTING";
            status_36.ForeColor = Color.Blue;
        }
        private void btnEXIT_1_Click(object sender, EventArgs e) {
            GlobalTestingFlag[0] = false;
            btnTEST_1.Enabled = true;
            status_1.Text = "FAIL";
            status_1.ForeColor = Color.Red;
        }
        private void btnEXIT_2_Click(object sender, EventArgs e) {
            GlobalTestingFlag[1] = false;
            btnTEST_2.Enabled = true;
            status_2.Text = "FAIL";
            status_2.ForeColor = Color.Red;
        }
        private void btnEXIT_3_Click(object sender, EventArgs e) {
            GlobalTestingFlag[2] = false;
            btnTEST_3.Enabled = true;
            status_3.Text = "FAIL";
            status_3.ForeColor = Color.Red;
        }
        private void btnEXIT_4_Click(object sender, EventArgs e) {
            GlobalTestingFlag[3] = false;
            btnTEST_4.Enabled = true;
            status_4.Text = "FAIL";
            status_4.ForeColor = Color.Red;
        }
        private void btnEXIT_5_Click(object sender, EventArgs e) {
            GlobalTestingFlag[4] = false;
            btnTEST_5.Enabled = true;
            status_5.Text = "FAIL";
            status_5.ForeColor = Color.Red;
        }
        private void btnEXIT_6_Click(object sender, EventArgs e) {
            GlobalTestingFlag[5] = false;
            btnTEST_6.Enabled = true;
            status_6.Text = "FAIL";
            status_6.ForeColor = Color.Red;
        }
        private void btnEXIT_7_Click(object sender, EventArgs e) {
            GlobalTestingFlag[6] = false;
            btnTEST_7.Enabled = true;
            status_7.Text = "FAIL";
            status_7.ForeColor = Color.Red;
        }
        private void btnEXIT_8_Click(object sender, EventArgs e) {
            GlobalTestingFlag[7] = false;
            btnTEST_8.Enabled = true;
            status_8.Text = "FAIL";
            status_8.ForeColor = Color.Red;
        }
        private void btnEXIT_9_Click(object sender, EventArgs e) {
            GlobalTestingFlag[8] = false;
            btnTEST_9.Enabled = true;
            status_9.Text = "FAIL";
            status_9.ForeColor = Color.Red;
        }
        private void btnEXIT_10_Click(object sender, EventArgs e) {
            GlobalTestingFlag[9] = false;
            btnTEST_10.Enabled = true;
            status_10.Text = "FAIL";
            status_10.ForeColor = Color.Red;
        }
        private void btnEXIT_11_Click(object sender, EventArgs e) {
            GlobalTestingFlag[10] = false;
            btnTEST_11.Enabled = true;
            status_11.Text = "FAIL";
            status_11.ForeColor = Color.Red;
        }
        private void btnEXIT_12_Click(object sender, EventArgs e) {
            GlobalTestingFlag[11] = false;
            btnTEST_12.Enabled = true;
            status_12.Text = "FAIL";
            status_12.ForeColor = Color.Red;
        }
        private void btnEXIT_13_Click(object sender, EventArgs e) {
            GlobalTestingFlag[12] = false;
            btnTEST_13.Enabled = true;
            status_13.Text = "FAIL";
            status_13.ForeColor = Color.Red;
        }
        private void btnEXIT_14_Click(object sender, EventArgs e) {
            GlobalTestingFlag[13] = false;
            btnTEST_14.Enabled = true;
            status_14.Text = "FAIL";
            status_14.ForeColor = Color.Red;
        }
        private void btnEXIT_15_Click(object sender, EventArgs e) {
            GlobalTestingFlag[14] = false;
            btnTEST_15.Enabled = true;
            status_15.Text = "FAIL";
            status_15.ForeColor = Color.Red;
        }
        private void btnEXIT_16_Click(object sender, EventArgs e) {
            GlobalTestingFlag[15] = false;
            btnTEST_16.Enabled = true;
            status_16.Text = "FAIL";
            status_16.ForeColor = Color.Red;
        }
        private void btnEXIT_17_Click(object sender, EventArgs e) {
            GlobalTestingFlag[16] = false;
            btnTEST_17.Enabled = true;
            status_17.Text = "FAIL";
            status_17.ForeColor = Color.Red;
        }
        private void btnEXIT_18_Click(object sender, EventArgs e) {
            GlobalTestingFlag[17] = false;
            btnTEST_18.Enabled = true;
            status_18.Text = "FAIL";
            status_18.ForeColor = Color.Red;
        }
        private void btnEXIT_19_Click(object sender, EventArgs e) {
            GlobalTestingFlag[18] = false;
            btnTEST_19.Enabled = true;
            status_19.Text = "FAIL";
            status_19.ForeColor = Color.Red;
        }
        private void btnEXIT_20_Click(object sender, EventArgs e) {
            GlobalTestingFlag[19] = false;
            btnTEST_20.Enabled = true;
            status_20.Text = "FAIL";
            status_20.ForeColor = Color.Red;
        }
        private void btnEXIT_21_Click(object sender, EventArgs e) {
            GlobalTestingFlag[20] = false;
            btnTEST_21.Enabled = true;
            status_21.Text = "FAIL";
            status_21.ForeColor = Color.Red;
        }
        private void btnEXIT_22_Click(object sender, EventArgs e) {
            GlobalTestingFlag[21] = false;
            btnTEST_22.Enabled = true;
            status_22.Text = "FAIL";
            status_22.ForeColor = Color.Red;
        }
        private void btnEXIT_23_Click(object sender, EventArgs e) {
            GlobalTestingFlag[22] = false;
            btnTEST_23.Enabled = true;
            status_23.Text = "FAIL";
            status_23.ForeColor = Color.Red;
        }
        private void btnEXIT_24_Click(object sender, EventArgs e) {
            GlobalTestingFlag[23] = false;
            btnTEST_24.Enabled = true;
            status_24.Text = "FAIL";
            status_24.ForeColor = Color.Red;
        }
        private void btnEXIT_25_Click(object sender, EventArgs e) {
            GlobalTestingFlag[24] = false;
            btnTEST_25.Enabled = true;
            status_25.Text = "FAIL";
            status_25.ForeColor = Color.Red;
        }
        private void btnEXIT_26_Click(object sender, EventArgs e) {
            GlobalTestingFlag[25] = false;
            btnTEST_26.Enabled = true;
            status_26.Text = "FAIL";
            status_26.ForeColor = Color.Red;
        }
        private void btnEXIT_27_Click(object sender, EventArgs e) {
            GlobalTestingFlag[26] = false;
            btnTEST_27.Enabled = true;
            status_27.Text = "FAIL";
            status_27.ForeColor = Color.Red;
        }
        private void btnEXIT_28_Click(object sender, EventArgs e) {
            GlobalTestingFlag[27] = false;
            btnTEST_28.Enabled = true;
            status_28.Text = "FAIL";
            status_28.ForeColor = Color.Red;
        }
        private void btnEXIT_29_Click(object sender, EventArgs e) {
            GlobalTestingFlag[28] = false;
            btnTEST_29.Enabled = true;
            status_29.Text = "FAIL";
            status_29.ForeColor = Color.Red;
        }
        private void btnEXIT_30_Click(object sender, EventArgs e) {
            GlobalTestingFlag[29] = false;
            btnTEST_30.Enabled = true;
            status_30.Text = "FAIL";
            status_30.ForeColor = Color.Red;
        }
        private void btnEXIT_31_Click(object sender, EventArgs e) {
            GlobalTestingFlag[30] = false;
            btnTEST_31.Enabled = true;
            status_31.Text = "FAIL";
            status_31.ForeColor = Color.Red;
        }
        private void btnEXIT_32_Click(object sender, EventArgs e) {
            GlobalTestingFlag[31] = false;
            btnTEST_32.Enabled = true;
            status_32.Text = "FAIL";
            status_32.ForeColor = Color.Red;
        }
        private void btnEXIT_33_Click(object sender, EventArgs e) {
            GlobalTestingFlag[32] = false;
            btnTEST_33.Enabled = true;
            status_33.Text = "FAIL";
            status_33.ForeColor = Color.Red;
        }
        private void btnEXIT_34_Click(object sender, EventArgs e) {
            GlobalTestingFlag[33] = false;
            btnTEST_34.Enabled = true;
            status_34.Text = "FAIL";
            status_34.ForeColor = Color.Red;
        }
        private void btnEXIT_35_Click(object sender, EventArgs e) {
            GlobalTestingFlag[34] = false;
            btnTEST_35.Enabled = true;
            status_35.Text = "FAIL";
            status_35.ForeColor = Color.Red;
        }
        private void btnEXIT_36_Click(object sender, EventArgs e) {
            GlobalTestingFlag[35] = false;
            btnTEST_36.Enabled = true;
            status_36.Text = "FAIL";
            status_36.ForeColor = Color.Red;
        }
        private void autoTestToolStripMenuItem_Click(object sender, EventArgs e) {
            setupPay.write_text(tester.headConfig.testAuto, autoTestToolStripMenuItem.Checked.ToString().ToUpper(), tester.nameFile);
            setupPay.setup();
            if (autoTestToolStripMenuItem.Checked != true) return;
            start_all_head();
        }
        public void start_all_head()
        {
            bool[] sup1 = flag_test;
            bool[] sup2 = flag_head;
            if (!tester.useRelayCard)
            {
                for (int i = 0; i < tester.numHead; i++)
                {
                    sup2[i] = true;
                }
            }//ไอศครีมเชอเบร็ตรสมะนาว
            Label[] l = { status_1, status_2, status_3, status_4, status_5, status_6, status_7, status_8, status_9, status_10,
                          status_11, status_12, status_13, status_14, status_15, status_16, status_17, status_18, status_19, status_20,
                          status_21, status_22, status_23, status_24, status_25, status_26, status_27, status_28, status_29, status_30,
                          status_31, status_32, status_33, status_34, status_35, status_36};
            Button[] b = { btnTEST_1, btnTEST_2, btnTEST_3, btnTEST_4, btnTEST_5, btnTEST_6, btnTEST_7, btnTEST_8, btnTEST_9, btnTEST_10,
                           btnTEST_11, btnTEST_12, btnTEST_13, btnTEST_14, btnTEST_15, btnTEST_16, btnTEST_17, btnTEST_18, btnTEST_19, btnTEST_20,
                           btnTEST_21, btnTEST_22, btnTEST_23, btnTEST_24, btnTEST_25, btnTEST_26, btnTEST_27, btnTEST_28, btnTEST_29, btnTEST_30,
                           btnTEST_31, btnTEST_32, btnTEST_33, btnTEST_34, btnTEST_35, btnTEST_36};
            for (int i = 0; i < tester.numHead; i++)
            {
                if (sup1[i] != true && sup2[i] == true)
                {
                    b[i].Enabled = false;
                    row_test[i] = 0;
                    flag_test[i] = true;
                    GlobalTestingFlag[i] = true;
                    l[i].Text = "TESTING";
                    l[i].ForeColor = Color.Blue;
                }
            }
        }
        private Stopwatch[] timer_head = { new Stopwatch(), new Stopwatch(), new Stopwatch(), new Stopwatch(), new Stopwatch(),
                                           new Stopwatch(), new Stopwatch(), new Stopwatch(), new Stopwatch(), new Stopwatch(),
                                           new Stopwatch(), new Stopwatch(), new Stopwatch(), new Stopwatch(), new Stopwatch(),
                                           new Stopwatch(), new Stopwatch(), new Stopwatch(), new Stopwatch(), new Stopwatch(),
                                           new Stopwatch(), new Stopwatch(), new Stopwatch(), new Stopwatch(), new Stopwatch(),
                                           new Stopwatch(), new Stopwatch(), new Stopwatch(), new Stopwatch(), new Stopwatch(),
                                           new Stopwatch(), new Stopwatch(), new Stopwatch(), new Stopwatch(), new Stopwatch(),
                                           new Stopwatch()};
        private void background_time_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e) {//ไอศครีมเชอเบร็ตรสมะนาว
            bool[] Status = { flag_test[0], flag_test[1], flag_test[2], flag_test[3], flag_test[4],
                              flag_test[5], flag_test[6], flag_test[7], flag_test[8], flag_test[9],
                              flag_test[10], flag_test[11], flag_test[12], flag_test[13], flag_test[14],
                              flag_test[15], flag_test[16], flag_test[17], flag_test[18], flag_test[19],
                              flag_test[20], flag_test[21], flag_test[22], flag_test[23], flag_test[24],
                              flag_test[25], flag_test[26], flag_test[27], flag_test[28], flag_test[29],
                              flag_test[30], flag_test[31], flag_test[32], flag_test[33], flag_test[34],
                              flag_test[35] };
            for (int i = 0; i < tester.numHead; i++) {
                timer_head[i].Restart();
            }
            bool flag_brink = false;
            bool status_all = false;
            bool status_all_current = false;
            while (true) {
                if (flag_this_close) break;
                for (int i = 0; i < tester.numHead; i++) {
                    if (Status[i] != flag_test[i]) {
                        Status[i] = flag_test[i];
                        timer_head[i].Restart();
                    }
                }
                background_time.ReportProgress(0);
                Thread.Sleep(250);
                if (flag_brink) {
                    background_time.ReportProgress(1);
                    flag_brink = false;
                } else flag_brink = true;

                status_all_current = false;
                for (int i = 0; i < tester.numHead; i++) {
                    if (flag_test[i]) status_all_current = true;
                }
                if (status_all != status_all_current) {
                    status_all = status_all_current;
                    if (status_all) background_time.ReportProgress(100);
                    else { background_time.ReportProgress(101); }
                }

                if (autoTestToolStripMenuItem.Checked)
                {
                    try
                    {
                        File.ReadAllText("auto_test_trick.txt");
                        File.Delete("auto_test_trick.txt");
                        background_time.ReportProgress(9);
                    }
                    catch { }
                }
            }
        }
        private void background_time_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e) {
            if (e.ProgressPercentage == 0) {

                for (int i = 1; i <= tester.numHead; i++) {
                    Label l_testtime = getLabelTestTime(i);
                    Label l_inout = getLabelInOutTime(i);

                    if (flag_test[i - 1]) {
                        l_testtime.Text = (timer_head[i - 1].ElapsedMilliseconds / 1000).ToString("0000");
                    } else {
                        l_inout.Text = (timer_head[i - 1].ElapsedMilliseconds / 1000).ToString("0000");
                    }
                    DateTime now = DateTime.Now;
                    if (lb_time.Text != now.ToString("T")) {
                        lb_time.Text = now.ToString("T");
                        lb_date.Text = now.ToString("d");
                    }
                }
            } else if (e.ProgressPercentage == 1) {
                
                for (int i = 1; i <= tester.numHead; i++) {
                    Label l = getLabelStatus(i);

                    if (flag_test[i - 1] == true) l.Visible = !l.Visible;
                    else l.Visible = true;
                }
            }
            if (e.ProgressPercentage == 100) Activator.CreateInstance(functionExcel, this, "OnAllFunction()");
            if (e.ProgressPercentage == 101) Activator.CreateInstance(functionExcel, this, "OffAllFunction()");
            if (e.ProgressPercentage == 9) start_all_head();
        }
        private void background_relay_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e) {
            Thread.Sleep(500);
            bool[] clearScaner = new bool[4];
            while (true) {
                if (flag_this_close) break;
                flag_write_23017 = false;
                Thread.Sleep(250);
                if (tester.testPanel) {
                    string fg = write_23017_sup("checkinput," + arduino_input[0] + "\n");
                    for (int i = 1; i <= tester.numHead; i++) {
                        if (fg == "0\r\n") background_relay.ReportProgress(Convert.ToInt32("10" + i));
                        else if (fg == "1\r\n") background_relay.ReportProgress(Convert.ToInt32("11" + i));
                    }
                } else {
                    if(tester.numHead <= 4)
                    {
                        for (int i = 1; i <= 4; i++)
                        {
                            string fg = write_23017_sup("checkinput," + arduino_input[i - 1] + "\n");
                            if (fg == "0\r\n")
                            {
                                background_relay.ReportProgress(Convert.ToInt32("10" + i));
                            }
                            else
                            {
                                background_relay.ReportProgress(Convert.ToInt32("11" + i));
                                if (clearScaner[i - 1])
                                {
                                    clearScaner[i - 1] = false;
                                    File.WriteAllText("dryice_scan2d_clear_sn_" + i + ".txt", "");
                                }
                            }
                            if (flag_test[i - 1])
                            {
                                clearScaner[i - 1] = true;
                            }
                        }
                    }
                    else
                    {
                        int headArray = 0;
                        int headPin = 1;
                        string fg = write_23017_sup("checkinput," + arduino_input[headArray] + "\n");
                        if (fg == "0\r\n") background_relay.ReportProgress(Convert.ToInt32("10" + headPin));
                        else if (fg == "1\r\n") background_relay.ReportProgress(Convert.ToInt32("11" + headPin));

                        headArray = 3;
                        headPin = 5;
                        fg = write_23017_sup("checkinput," + arduino_input[headArray] + "\n");
                        if (fg == "0\r\n") background_relay.ReportProgress(Convert.ToInt32("10" + headPin));
                        else if (fg == "1\r\n") background_relay.ReportProgress(Convert.ToInt32("11" + headPin));
                    }
                }
                flag_write_23017 = true;
                Thread.Sleep(5000);
            }
        }
        private void background_relay_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e) {
            if (e.ProgressPercentage == 99999) {
                //bool[] sup1 = flag_test;
                //bool[] sup2 = flag_head;
                //if (!tester.useRelayCard) {
                //    for (int i = 0; i < num_head; i++) {
                //        sup2[i] = true;
                //    }
                //}//ไอศครีมเชอเบร็ตรสมะนาว
                //Label[] l = { status_1, status_2, status_3, status_4, status_5,
                //              status_6, status_7, status_8, status_9, status_10,
                //              status_11, status_12, status_13, status_14, status_15,
                //              status_16, status_17, status_18, status_19, status_20,
                //              status_21, status_22, status_23, status_24, status_25,
                //              status_26, status_27, status_28, status_29, status_30,
                //              status_31, status_32, status_33, status_34, status_35,
                //              status_36};
                //Button[] b = { btnTEST_1, btnTEST_2, btnTEST_3, btnTEST_4, btnTEST_5,
                //               btnTEST_6, btnTEST_7, btnTEST_8, btnTEST_9, btnTEST_10,
                //               btnTEST_11, btnTEST_12, btnTEST_13, btnTEST_14, btnTEST_15,
                //               btnTEST_16, btnTEST_17, btnTEST_18, btnTEST_19, btnTEST_20,
                //               btnTEST_21, btnTEST_22, btnTEST_23, btnTEST_24, btnTEST_25,
                //               btnTEST_26, btnTEST_27, btnTEST_28, btnTEST_29, btnTEST_30,
                //               btnTEST_31, btnTEST_32, btnTEST_33, btnTEST_34, btnTEST_35,
                //               btnTEST_36};
                //for (int i = 0; i < num_head; i++) {
                //    if (sup1[i] != true && sup2[i] == true) {
                //        b[i].Enabled = false;
                //        row_test[i] = 0;
                //        flag_test[i] = true;
                //        GlobalTestingFlag[i] = true;
                //        l[i].Text = "TESTING";
                //        l[i].ForeColor = Color.Blue;
                //    }
                //}
                start_all_head();
                return;
            }
            string s = Convert.ToString(e.ProgressPercentage);
            string s1 = "1";
            if (s.Length == 3) s1 = s.Substring(2, 1);
            else s1 = s.Substring(2, 2);
            string s2 = s.Substring(1, 1);
            if (s2 == "0") sup_backgroud_checkhead(Convert.ToInt32(s1), 0);
            if (s2 == "1") sup_backgroud_checkhead(Convert.ToInt32(s1), 1);
        }
        private void sup_backgroud_checkhead(int head, int get) {
            Button b = getButtonTest(head);
            Label l = getLabelStatus(head);
            GroupBox g = getGroupBoxHead(head);
            
            switch (get) {
                case 0:
                    bt_relayCard.BackColor = Color.LimeGreen;
                    if (flag_head[head - 1] == true) break;
                    flag_head[head - 1] = true;
                    g.ForeColor = Color.Blue;
                    if (autoTestToolStripMenuItem.Checked != true) break;
                    if (flag_test[head - 1] == true) break;
                    b.Enabled = false;
                    row_test[head - 1] = 0;
                    flag_test[head - 1] = true;
                    GlobalTestingFlag[head - 1] = true;
                    l.Text = "TESTING";
                    l.ForeColor = Color.Blue;
                    break;
                case 1:
                    bt_relayCard.BackColor = Color.LimeGreen;
                    flag_head[head - 1] = false;
                    g.ForeColor = Color.Black;
                    flag_test[head - 1] = false;
                    break;
                case -4:
                    bt_relayCard.BackColor = Color.Red;
                    break;
            }
        }
        private void automation_background_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e) {
            Thread.Sleep(500);
            if (tester.cylinder1) tcpip_send_data.Add("0");
            if (tester.cylinder2) tcpip_send_data.Add("0");
            for (int i = 0; i < tester.numHead; i++) tcpip_send_data.Add("2");
            while (true) {
                if (flag_this_close) break;
                #region flag test 1
                if (autoMation.inTric1) {
                    autoMation.inTric1 = false;
                    flag_write_23017 = false;
                    tcpip_send_atk(1);
                    Thread.Sleep(250);
                    string[] bf = tester.cylinderRelay1.Split('&');
                    Relay_On_sup(Convert.ToInt32(bf[0]), Convert.ToInt32(bf[1]));
                    flag_write_23017 = true;
                    Thread.Sleep(1000);
                    automation_background.ReportProgress(1);
                    string[] mnb = tester.cylinderHead1.Split('&');
                    while (true) {
                        bool hj = true;
                        foreach (string nn in mnb) {
                            if (!flag_head[Convert.ToInt32(nn) - 1]) hj = false;
                        }
                        if (hj) break;
                        Thread.Sleep(50);
                    }
                }
                if (autoMation.outTric1) {
                    autoMation.outTric1 = false;
                    flag_write_23017 = false;
                    Thread.Sleep(250);
                    string[] bf = tester.cylinderRelay1.Split('&');
                    Relay_Off_sup(Convert.ToInt32(bf[0]), Convert.ToInt32(bf[1]));
                    flag_write_23017 = true;
                    automation_background.ReportProgress(11);
                    string[] mnb = tester.cylinderHead1.Split('&');
                    while (true) {
                        bool hj = true;
                        foreach (string nn in mnb) {
                            if (flag_head[Convert.ToInt32(nn) - 1]) hj = false;
                        }
                        if (hj) break;
                        Thread.Sleep(50);
                    }
                    tcpip_send();
                }
                #endregion
                #region flag test 2
                if (autoMation.inTric2) {
                    autoMation.inTric2 = false;
                    flag_write_23017 = false;
                    tcpip_send_atk(2);
                    Thread.Sleep(250);
                    string[] bf = tester.cylinderRelay2.Split('&');
                    Relay_On_sup(Convert.ToInt32(bf[0]), Convert.ToInt32(bf[1]));
                    flag_write_23017 = true;
                    Thread.Sleep(1000);
                    automation_background.ReportProgress(2);
                    string[] mnb = tester.cylinderHead2.Split('&');
                    while (true) {
                        bool hj = true;
                        foreach (string nn in mnb) {
                            if (!flag_head[Convert.ToInt32(nn) - 1]) hj = false;
                        }
                        if (hj) break;
                        Thread.Sleep(50);
                    }
                }
                if (autoMation.outTric2) {
                    autoMation.outTric2 = false;
                    flag_write_23017 = false;
                    Thread.Sleep(250);
                    string[] bf = tester.cylinderRelay2.Split('&');
                    Relay_Off_sup(Convert.ToInt32(bf[0]), Convert.ToInt32(bf[1]));
                    flag_write_23017 = true;
                    automation_background.ReportProgress(22);
                    string[] mnb = tester.cylinderHead2.Split('&');
                    while (true) {
                        bool hj = true;
                        foreach (string nn in mnb) {
                            if (flag_head[Convert.ToInt32(nn) - 1]) hj = false;
                        }
                        if (hj) break;
                        Thread.Sleep(50);
                    }
                    tcpip_send();
                }
                #endregion

                if (tcpip_client != null) {
                    if (tcpip_client.Available != 0 && tcpIP.useRobot) {
                        byte[] bytesToRead = new byte[tcpip_client.ReceiveBufferSize];
                        int byteRead = tcpip_stream.Read(bytesToRead, 0, tcpip_client.ReceiveBufferSize);
                        string ss = Encoding.ASCII.GetString(bytesToRead, 0, byteRead);
                        if (ss.Contains(tcpIP.readRobot1.Replace("&", ","))) autoMation.inTric1 = true;
                        if (ss.Contains(tcpIP.readRobot2.Replace("&", ","))) autoMation.inTric2 = true;
                        Log(LogMsgType.Incoming_Blue, "\nread " + ss);
                    }
                }
                Thread.Sleep(250);
            }
        }
        private void automation_background_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e) {
            #region flag test 1
            if (e.ProgressPercentage == 1) {
                setFlagAutomation1.Checked = true;
                string[] mnb = tester.cylinderHead1.Split('&');
                foreach (string nn in mnb) {
                    sup_backgroud_checkhead(Convert.ToInt32(nn), 0);
                }
            }
            if (e.ProgressPercentage == 11) {
                setFlagAutomation1.Checked = false;
                string[] mnb = tester.cylinderHead1.Split('&');
                foreach (string nn in mnb) {
                    sup_backgroud_checkhead(Convert.ToInt32(nn), 1);
                }
            }
            #endregion
            #region flag test 2
            if (e.ProgressPercentage == 2) {
                setFlagAutomation2.Checked = true;
                string[] mnb = tester.cylinderHead2.Split('&');
                foreach (string nn in mnb) {
                    sup_backgroud_checkhead(Convert.ToInt32(nn), 0);
                }
            }
            if (e.ProgressPercentage == 22) {
                setFlagAutomation2.Checked = false;
                string[] mnb = tester.cylinderHead2.Split('&');
                foreach (string nn in mnb) {
                    sup_backgroud_checkhead(Convert.ToInt32(nn), 1);
                }
            }
            #endregion
        }
        private TcpClient tcpip_client = null;
        private List<string> tcpip_send_data = new List<string>();
        private NetworkStream tcpip_stream;
        private void tcpip_send() {
            if (!tcpIP.useRobot) return;
            //tcpip_connect(tcpIP.ip, tcpIP.port);
            int lkj = 0;
            if (tester.cylinder2) lkj++;
            string data = "";
            if (setFlagAutomation1.Checked) tcpip_send_data[0] = "0";
            else tcpip_send_data[0] = "1";
            if (setFlagAutomation2.Enabled) {
                if (setFlagAutomation2.Checked) tcpip_send_data[1] = "0";
                else tcpip_send_data[1] = "1";
            }
            for (int i = 1; i <= tester.numHead; i++) {
                Label l = getLabelStatus(i);
  
                if (l.Text == "PASS") tcpip_send_data[i + lkj] = "1";
                else if (l.Text == "FAIL") tcpip_send_data[i + lkj] = "0";
                else tcpip_send_data[i + lkj] = "2";
            }
            foreach (string g in tcpip_send_data) {
                if (g == "2") data += "-,";
                else data += g + ",";
            }
            data = data.Substring(0, data.Length - 1);
            byte[] bytesToSend = ASCIIEncoding.ASCII.GetBytes(data + "\r\n");
            if (tcpip_client != null && tb_detail.Text != "Rotary Dimmer") {
                //try { int hhj = tcpip_stream.Length; } catch { }
                try { tcpip_stream.Write(bytesToSend, 0, bytesToSend.Length); } catch { }
                Log(LogMsgType.Incoming_Blue, "\nsend " + data);
            }
        }
        private void tcpip_send_atk(int kk) {
            if (!tcpIP.useRobot) return;
            //tcpip_connect(tcpIP.ip, tcpIP.port);
            int lkj = 0;
            if (tester.cylinder2) lkj++;
            string data = "";
            if (kk == 1) {
                tcpip_send_data[0] = "0";
                string[] mnb = tester.cylinderHead1.Split('&');
                foreach (string nn in mnb) {
                    tcpip_send_data[Convert.ToInt32(nn) + lkj] = "2";
                }
            }
            if (kk == 2) {
                tcpip_send_data[1] = "0";
                string[] mnb = tester.cylinderHead1.Split('&');
                foreach (string nn in mnb) {
                    tcpip_send_data[Convert.ToInt32(nn) + lkj] = "2";
                }
            }
            foreach (string g in tcpip_send_data) {
                if (g == "2") data += "-,";
                else data += g + ",";
            }
            data = data.Substring(0, data.Length - 1);
            byte[] bytesToSend = ASCIIEncoding.ASCII.GetBytes(data + "\r\n");
            if (tcpip_client != null && tb_detail.Text != "Rotary Dimmer") {
                try { tcpip_stream.Write(bytesToSend, 0, bytesToSend.Length); } catch { }
                Log(LogMsgType.Incoming_Blue, "\nsend " + data);
            }
        }
        private void tcpip_connect(string ip, string port) {
            try {
                tcpip_client = new TcpClient(ip, Convert.ToInt32(port));
                tcpip_stream = tcpip_client.GetStream();
                tb_tcpip.BackColor = Color.LimeGreen;
            } catch {
                tb_tcpip.BackColor = Color.Red;
                Log(LogMsgType.Error_Red, "\ncan not connect tcpip");
            }
        }
        private void tcpip_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e) {
            tcpip_connect(tcpIP.ip, tcpIP.port);
            while (true) {
                if (flag_this_close) break;
                tcpip_send();
                try { Thread.Sleep(Convert.ToInt32(tcpIP.timerTric)); } catch { }
            }
        }
        private void tcpip_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e) {

        }

        private void txtSNBoard_1_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_1.Clear();
        }
        private void txtSNBoard_2_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_2.Clear();
        }
        private void txtSNBoard_3_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_3.Clear();
        }
        private void txtSNBoard_4_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_4.Clear();
        }
        private void txtSNBoard_5_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_5.Clear();
        }
        private void txtSNBoard_6_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_6.Clear();
        }
        private void txtSNBoard_7_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_7.Clear();
        }
        private void txtSNBoard_8_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_8.Clear();
        }
        private void txtSNBoard_9_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_9.Clear();
        }
        private void txtSNBoard_10_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_10.Clear();
        }
        private void txtSNBoard_11_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_11.Clear();
        }
        private void txtSNBoard_12_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_12.Clear();
        }
        private void txtSNBoard_13_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_13.Clear();
        }
        private void txtSNBoard_14_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_14.Clear();
        }
        private void txtSNBoard_15_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_15.Clear();
        }
        private void txtSNBoard_16_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_16.Clear();
        }
        private void txtSNBoard_17_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_17.Clear();
        }
        private void txtSNBoard_18_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_18.Clear();
        }
        private void txtSNBoard_19_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_19.Clear();
        }
        private void txtSNBoard_20_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_20.Clear();
        }
        private void txtSNBoard_21_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_21.Clear();
        }
        private void txtSNBoard_22_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_22.Clear();
        }
        private void txtSNBoard_23_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_23.Clear();
        }
        private void txtSNBoard_24_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_24.Clear();
        }
        private void txtSNBoard_25_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_25.Clear();
        }
        private void txtSNBoard_26_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_26.Clear();
        }
        private void txtSNBoard_27_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_27.Clear();
        }
        private void txtSNBoard_28_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_28.Clear();
        }
        private void txtSNBoard_29_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_29.Clear();
        }
        private void txtSNBoard_30_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_30.Clear();
        }
        private void txtSNBoard_31_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_31.Clear();
        }
        private void txtSNBoard_32_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_32.Clear();
        }
        private void txtSNBoard_33_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_33.Clear();
        }
        private void txtSNBoard_34_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_34.Clear();
        }
        private void txtSNBoard_35_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_35.Clear();
        }
        private void txtSNBoard_36_Click(object sender, EventArgs e) {
            if (tester.click2ClearSN) txtSNBoard_36.Clear();
        }
        private void txtWO_KeyDown(object sender, KeyEventArgs e) {
            if (e.KeyCode != Keys.Enter || tb_wo.Text == "") return;
            bool b = false;
            for (int i = 0; i < tester.numHead; i++) {
                b |= flag_test[i];
            }
            if (b == true) {
                MessageBox.Show("key WO ไม่ได้ ขณะโปรแกรมกำลังเทสอยู่");
                return;
            }
            tb_wo.Text = tb_wo.Text.ToUpper();
            getWorkOrder(tb_wo.Text);
            //lblUserID.Text = TeamPrecision.PRISM.cSettingValues.EmployeeID;
            tb_userID.Text = setupPay.read_text(prismTest.headConfig.employeeID, prismTest.nameFile);
        }
        private void lblModelName_TextChanged(object sender, EventArgs e) {
            Size size = TextRenderer.MeasureText(tb_detail.Text, tb_detail.Font);
            tb_detail.Width = size.Width;
            lb_fwVersion.Location = new Point(tb_detail.Location.X + tb_detail.Size.Width + 5, lb_fwVersion.Location.Y);
            tb_fwVersion.Location = new Point(lb_fwVersion.Location.X + lb_fwVersion.Size.Width, tb_fwVersion.Location.Y);
            lb_spec.Location = new Point(tb_fwVersion.Location.X + tb_fwVersion.Size.Width + 5, lb_spec.Location.Y);
            tb_spec.Location = new Point(lb_spec.Location.X + lb_spec.Size.Width, tb_spec.Location.Y);
            tb_spec.Size = new Size(groupBox1.Size.Width - (lb_detail.Size.Width + tb_detail.Size.Width + lb_fwVersion.Size.Width + tb_fwVersion.Size.Width + lb_spec.Size.Width + 25), tb_spec.Size.Height);
        }
        private void lblFirmwareVersion_TextChanged(object sender, EventArgs e) {
            Size size = TextRenderer.MeasureText(tb_fwVersion.Text, tb_fwVersion.Font);
            tb_fwVersion.Width = size.Width;
            lb_spec.Location = new Point(tb_fwVersion.Location.X + tb_fwVersion.Size.Width + 5, lb_spec.Location.Y);
            tb_spec.Location = new Point(lb_spec.Location.X + lb_spec.Size.Width, tb_spec.Location.Y);
            tb_spec.Size = new Size(groupBox1.Size.Width - (lb_detail.Size.Width + tb_detail.Size.Width + lb_fwVersion.Size.Width + tb_fwVersion.Size.Width + lb_spec.Size.Width + 25), tb_spec.Size.Height);
        }
        private void head1ToolStripMenuItem_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_1_debug.txt", set_debug_1.Checked.ToString());
        }
        private void head2ToolStripMenuItem_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_2_debug.txt", set_debug_2.Checked.ToString());
        }
        private void head3ToolStripMenuItem_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_3_debug.txt", set_debug_3.Checked.ToString());
        }
        private void head4ToolStripMenuItem_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_4_debug.txt", set_debug_4.Checked.ToString());
        }
        private void head5ToolStripMenuItem_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_5_debug.txt", set_debug_5.Checked.ToString());
        }
        private void head6ToolStripMenuItem_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_6_debug.txt", set_debug_6.Checked.ToString());
        }
        private void head7ToolStripMenuItem_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_7_debug.txt", set_debug_7.Checked.ToString());
        }
        private void head8ToolStripMenuItem_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_8_debug.txt", set_debug_8.Checked.ToString());
        }
        private void head9ToolStripMenuItem_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_9_debug.txt", set_debug_9.Checked.ToString());
        }
        private void head10ToolStripMenuItem_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_10_debug.txt", set_debug_10.Checked.ToString());
        }
        private void set_debug_11_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_11_debug.txt", set_debug_11.Checked.ToString());
        }
        private void set_debug_12_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_12_debug.txt", set_debug_12.Checked.ToString());
        }
        private void set_debug_13_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_13_debug.txt", set_debug_13.Checked.ToString());
        }
        private void set_debug_14_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_14_debug.txt", set_debug_14.Checked.ToString());
        }
        private void set_debug_15_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_15_debug.txt", set_debug_15.Checked.ToString());
        }
        private void set_debug_16_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_16_debug.txt", set_debug_16.Checked.ToString());
        }
        private void set_debug_17_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_17_debug.txt", set_debug_17.Checked.ToString());
        }
        private void set_debug_18_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_18_debug.txt", set_debug_18.Checked.ToString());
        }
        private void set_debug_19_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_19_debug.txt", set_debug_19.Checked.ToString());
        }
        private void set_debug_20_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_20_debug.txt", set_debug_20.Checked.ToString());
        }
        private void set_debug_21_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_21_debug.txt", set_debug_21.Checked.ToString());
        }
        private void set_debug_22_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_22_debug.txt", set_debug_22.Checked.ToString());
        }
        private void set_debug_23_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_23_debug.txt", set_debug_23.Checked.ToString());
        }
        private void set_debug_24_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_24_debug.txt", set_debug_24.Checked.ToString());
        }
        private void set_debug_25_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_25_debug.txt", set_debug_25.Checked.ToString());
        }
        private void set_debug_26_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_26_debug.txt", set_debug_26.Checked.ToString());
        }
        private void set_debug_27_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_27_debug.txt", set_debug_27.Checked.ToString());
        }
        private void set_debug_28_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_28_debug.txt", set_debug_28.Checked.ToString());
        }
        private void set_debug_29_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_29_debug.txt", set_debug_29.Checked.ToString());
        }
        private void set_debug_30_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_30_debug.txt", set_debug_30.Checked.ToString());
        }
        private void set_debug_31_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_31_debug.txt", set_debug_31.Checked.ToString());
        }
        private void set_debug_32_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_32_debug.txt", set_debug_32.Checked.ToString());
        }
        private void set_debug_33_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_33_debug.txt", set_debug_33.Checked.ToString());
        }
        private void set_debug_34_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_34_debug.txt", set_debug_34.Checked.ToString());
        }
        private void set_debug_35_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_35_debug.txt", set_debug_35.Checked.ToString());
        }
        private void set_debug_36_Click(object sender, EventArgs e) {
            File.WriteAllText("../../config/test_head_36_debug.txt", set_debug_36.Checked.ToString());
        }
        private void show_data_grid_Click(object sender, EventArgs e) {
            fMain_sizechanged();
        }
        #endregion

        #region ====================================================== ContextMenuStrip Event ======================================================
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) {
            for (int i = 0; i < tester.numHead; i++) {
                if (flag_test[i] != false) {
                    MessageBox.Show("_ไม่สามารถโหลดไฟล์ได้ ขณะโปรแกรมกำลังเทส");
                    return;
                }
            }
            string nameFG = cbb_fg.Text;
            try {
                excel.workBook.LoadFromFile("../../TestDescription/" + nameFG + excel.lastName);
            } catch {
                Log(LogMsgType.Error_Red, "_ปิด excel ก่อน" + nameFG + excel.lastName + "\n");
                return;
            }
            LoadDiscription();
            Activator.CreateInstance(functionExcel, this, "LoadTestSpec()");
            CreateFolder_datalog();
            enable_button_test_and_exit();
        }
        private void fMain_SizeChanged(object sender, EventArgs e) {
            fMain_sizechanged();
        }
        /// <summary>
        /// Flag for disable enable size changed
        /// </summary>
        private bool flagSizeChanged { get; set; }
        private void fMain_sizechanged() {
            if (!flagSizeChanged) {
                return;
            }
            GroupBox[] g = new GroupBox[36];//ไอศครีมเชอเบร็ตรสมะนาว
            g[0] = groupBox_head1; g[1] = groupBox_head2; g[2] = groupBox_head3; g[3] = groupBox_head4; g[4] = groupBox_head5;
            g[5] = groupBox_head6; g[6] = groupBox_head7; g[7] = groupBox_head8; g[8] = groupBox_head9; g[9] = groupBox_head10;
            g[10] = groupBox_head11; g[11] = groupBox_head12; g[12] = groupBox_head13; g[13] = groupBox_head14; g[14] = groupBox_head15;
            g[15] = groupBox_head16; g[16] = groupBox_head17; g[17] = groupBox_head18; g[18] = groupBox_head19; g[19] = groupBox_head20;
            g[20] = groupBox_head21; g[21] = groupBox_head22; g[22] = groupBox_head23; g[23] = groupBox_head24; g[24] = groupBox_head25;
            g[25] = groupBox_head26; g[26] = groupBox_head27; g[27] = groupBox_head28; g[28] = groupBox_head29; g[29] = groupBox_head30;
            g[30] = groupBox_head31; g[31] = groupBox_head32; g[32] = groupBox_head33; g[33] = groupBox_head34; g[34] = groupBox_head35; g[35] = groupBox_head36;
            if (show_data_grid.Checked) {
                for (int i = 0; i < 36; i++) { g[i].Visible = false; }
                DataGridView gg = getDataGridView(show_datagrid_int);
                GroupBox gb = getGroupBoxHead(show_datagrid_int);

                gg.Visible = true;
                gb.Visible = true;
                gb.Location = new Point(groupBox_head1.Location.X, groupBox_head1.Location.Y);
                gb.Size = new Size(((groupBox1.Width) - 2), ((this.Height - 82) / 1));
                return;
            } else { for (int i = 0; i < 36; i++) { g[i].Visible = true; } }
            tester.numHead = Convert.ToInt32(setupPay.read_text(tester.headConfig.numHead, tester.nameFile));
            if (tester.numHead > 20) {
                int Factor_groupbox = 0;
                int Factor_groupbox_2 = 0;
                int factor = 3;
                switch (tester.numHead) {//ไอศครีมเชอเบร็ตรสมะนาว
                    case 21: Factor_groupbox = 112; Factor_groupbox_2 = 7; break;
                    case 22: Factor_groupbox = 116; Factor_groupbox_2 = 8; groupBox_head23.Visible = false; groupBox_head24.Visible = false; break;
                    case 23: Factor_groupbox = 116; Factor_groupbox_2 = 8; groupBox_head24.Visible = false; break;
                    case 24: Factor_groupbox = 116; Factor_groupbox_2 = 8; break;
                    case 25: Factor_groupbox = 120; Factor_groupbox_2 = 9; groupBox_head26.Visible = false; groupBox_head27.Visible = false; break;
                    case 26: Factor_groupbox = 120; Factor_groupbox_2 = 9; groupBox_head27.Visible = false; break;
                    case 27: Factor_groupbox = 120; Factor_groupbox_2 = 9; break;
                    case 28: Factor_groupbox = 124; Factor_groupbox_2 = 10; groupBox_head29.Visible = false; groupBox_head30.Visible = false; break;
                    case 29: Factor_groupbox = 124; Factor_groupbox_2 = 10; groupBox_head30.Visible = false; break;
                    case 30: Factor_groupbox = 124; Factor_groupbox_2 = 10; break;
                    case 31: Factor_groupbox = 118; Factor_groupbox_2 = 11; groupBox_head32.Visible = false; groupBox_head33.Visible = false; break;
                    case 32: Factor_groupbox = 118; Factor_groupbox_2 = 11; groupBox_head33.Visible = false; break;
                    case 33: Factor_groupbox = 118; Factor_groupbox_2 = 11; break;
                    case 34: Factor_groupbox = 122; Factor_groupbox_2 = 12; groupBox_head35.Visible = false; groupBox_head36.Visible = false; break;
                    case 35: Factor_groupbox = 122; Factor_groupbox_2 = 12; groupBox_head36.Visible = false; break;
                    case 36: Factor_groupbox = 122; Factor_groupbox_2 = 12; break;
                }
                groupBox_head1.Size = new Size(((groupBox1.Width / 3) - 2), ((this.Height - Factor_groupbox) / Factor_groupbox_2));
                groupBox_head2.Location = new Point(groupBox_head1.Location.X + groupBox_head1.Width + 5, groupBox_head2.Location.Y);
                groupBox_head2.Size = new Size(((groupBox1.Width / 3) - 2), ((this.Height - Factor_groupbox) / Factor_groupbox_2));
                groupBox_head3.Location = new Point(groupBox_head2.Location.X + groupBox_head2.Width + 5, groupBox_head2.Location.Y);
                groupBox_head3.Size = new Size(((groupBox1.Width / 3) - 2), ((this.Height - Factor_groupbox) / Factor_groupbox_2));
                groupBox_head4.Location = new Point(groupBox_head1.Location.X, groupBox_head1.Location.Y + groupBox_head1.Size.Height + factor);
                groupBox_head4.Size = new Size(groupBox_head1.Size.Width, groupBox_head1.Size.Height);
                groupBox_head5.Location = new Point(groupBox_head2.Location.X, groupBox_head2.Location.Y + groupBox_head2.Size.Height + factor);
                groupBox_head5.Size = new Size(groupBox_head2.Size.Width, groupBox_head2.Size.Height);
                groupBox_head6.Location = new Point(groupBox_head3.Location.X, groupBox_head3.Location.Y + groupBox_head3.Size.Height + factor);
                groupBox_head6.Size = new Size(groupBox_head3.Size.Width, groupBox_head3.Size.Height);
                groupBox_head7.Location = new Point(groupBox_head4.Location.X, groupBox_head4.Location.Y + groupBox_head4.Size.Height + factor);
                groupBox_head7.Size = new Size(groupBox_head4.Size.Width, groupBox_head4.Size.Height);
                groupBox_head8.Location = new Point(groupBox_head5.Location.X, groupBox_head5.Location.Y + groupBox_head5.Size.Height + factor);
                groupBox_head8.Size = new Size(groupBox_head5.Size.Width, groupBox_head5.Size.Height);
                groupBox_head9.Location = new Point(groupBox_head6.Location.X, groupBox_head6.Location.Y + groupBox_head6.Size.Height + factor);
                groupBox_head9.Size = new Size(groupBox_head6.Size.Width, groupBox_head6.Size.Height);
                groupBox_head10.Location = new Point(groupBox_head7.Location.X, groupBox_head7.Location.Y + groupBox_head7.Size.Height + factor);
                groupBox_head10.Size = new Size(groupBox_head7.Size.Width, groupBox_head7.Size.Height);
                groupBox_head11.Location = new Point(groupBox_head8.Location.X, groupBox_head8.Location.Y + groupBox_head8.Size.Height + factor);
                groupBox_head11.Size = new Size(groupBox_head8.Size.Width, groupBox_head8.Size.Height);
                groupBox_head12.Location = new Point(groupBox_head9.Location.X, groupBox_head9.Location.Y + groupBox_head9.Size.Height + factor);
                groupBox_head12.Size = new Size(groupBox_head9.Size.Width, groupBox_head9.Size.Height);
                groupBox_head13.Location = new Point(groupBox_head10.Location.X, groupBox_head10.Location.Y + groupBox_head10.Size.Height + factor);
                groupBox_head13.Size = new Size(groupBox_head10.Size.Width, groupBox_head10.Size.Height);
                groupBox_head14.Location = new Point(groupBox_head11.Location.X, groupBox_head11.Location.Y + groupBox_head11.Size.Height + factor);
                groupBox_head14.Size = new Size(groupBox_head11.Size.Width, groupBox_head11.Size.Height);
                groupBox_head15.Location = new Point(groupBox_head12.Location.X, groupBox_head12.Location.Y + groupBox_head12.Size.Height + factor);
                groupBox_head15.Size = new Size(groupBox_head12.Size.Width, groupBox_head12.Size.Height);
                groupBox_head16.Location = new Point(groupBox_head13.Location.X, groupBox_head13.Location.Y + groupBox_head13.Size.Height + factor);
                groupBox_head16.Size = new Size(groupBox_head13.Size.Width, groupBox_head13.Size.Height);
                groupBox_head17.Location = new Point(groupBox_head14.Location.X, groupBox_head14.Location.Y + groupBox_head14.Size.Height + factor);
                groupBox_head17.Size = new Size(groupBox_head14.Size.Width, groupBox_head14.Size.Height);
                groupBox_head18.Location = new Point(groupBox_head15.Location.X, groupBox_head15.Location.Y + groupBox_head15.Size.Height + factor);
                groupBox_head18.Size = new Size(groupBox_head15.Size.Width, groupBox_head15.Size.Height);
                groupBox_head19.Location = new Point(groupBox_head16.Location.X, groupBox_head16.Location.Y + groupBox_head16.Size.Height + factor);
                groupBox_head19.Size = new Size(groupBox_head16.Size.Width, groupBox_head16.Size.Height);
                groupBox_head20.Location = new Point(groupBox_head17.Location.X, groupBox_head17.Location.Y + groupBox_head17.Size.Height + factor);
                groupBox_head20.Size = new Size(groupBox_head17.Size.Width, groupBox_head17.Size.Height);
                groupBox_head21.Location = new Point(groupBox_head18.Location.X, groupBox_head18.Location.Y + groupBox_head18.Size.Height + factor);
                groupBox_head21.Size = new Size(groupBox_head18.Size.Width, groupBox_head18.Size.Height);
                groupBox_head22.Location = new Point(groupBox_head19.Location.X, groupBox_head19.Location.Y + groupBox_head19.Size.Height + factor);
                groupBox_head22.Size = new Size(groupBox_head19.Size.Width, groupBox_head19.Size.Height);
                groupBox_head23.Location = new Point(groupBox_head20.Location.X, groupBox_head20.Location.Y + groupBox_head20.Size.Height + factor);
                groupBox_head23.Size = new Size(groupBox_head20.Size.Width, groupBox_head20.Size.Height);
                groupBox_head24.Location = new Point(groupBox_head21.Location.X, groupBox_head21.Location.Y + groupBox_head21.Size.Height + factor);
                groupBox_head24.Size = new Size(groupBox_head21.Size.Width, groupBox_head21.Size.Height);
                groupBox_head25.Location = new Point(groupBox_head22.Location.X, groupBox_head22.Location.Y + groupBox_head22.Size.Height + factor);
                groupBox_head25.Size = new Size(groupBox_head22.Size.Width, groupBox_head22.Size.Height);
                groupBox_head26.Location = new Point(groupBox_head23.Location.X, groupBox_head23.Location.Y + groupBox_head23.Size.Height + factor);
                groupBox_head26.Size = new Size(groupBox_head23.Size.Width, groupBox_head23.Size.Height);
                groupBox_head27.Location = new Point(groupBox_head24.Location.X, groupBox_head24.Location.Y + groupBox_head24.Size.Height + factor);
                groupBox_head27.Size = new Size(groupBox_head24.Size.Width, groupBox_head24.Size.Height);
                groupBox_head28.Location = new Point(groupBox_head25.Location.X, groupBox_head25.Location.Y + groupBox_head25.Size.Height + factor);
                groupBox_head28.Size = new Size(groupBox_head25.Size.Width, groupBox_head25.Size.Height);
                groupBox_head29.Location = new Point(groupBox_head26.Location.X, groupBox_head26.Location.Y + groupBox_head26.Size.Height + factor);
                groupBox_head29.Size = new Size(groupBox_head26.Size.Width, groupBox_head26.Size.Height);
                groupBox_head30.Location = new Point(groupBox_head27.Location.X, groupBox_head27.Location.Y + groupBox_head27.Size.Height + factor);
                groupBox_head30.Size = new Size(groupBox_head27.Size.Width, groupBox_head27.Size.Height);
                groupBox_head31.Location = new Point(groupBox_head28.Location.X, groupBox_head28.Location.Y + groupBox_head28.Size.Height + factor);
                groupBox_head31.Size = new Size(groupBox_head28.Size.Width, groupBox_head28.Size.Height);
                groupBox_head32.Location = new Point(groupBox_head29.Location.X, groupBox_head29.Location.Y + groupBox_head29.Size.Height + factor);
                groupBox_head32.Size = new Size(groupBox_head29.Size.Width, groupBox_head29.Size.Height);
                groupBox_head33.Location = new Point(groupBox_head30.Location.X, groupBox_head30.Location.Y + groupBox_head30.Size.Height + factor);
                groupBox_head33.Size = new Size(groupBox_head30.Size.Width, groupBox_head30.Size.Height);
                groupBox_head34.Location = new Point(groupBox_head31.Location.X, groupBox_head31.Location.Y + groupBox_head31.Size.Height + factor);
                groupBox_head34.Size = new Size(groupBox_head31.Size.Width, groupBox_head31.Size.Height);
                groupBox_head35.Location = new Point(groupBox_head32.Location.X, groupBox_head32.Location.Y + groupBox_head32.Size.Height + factor);
                groupBox_head35.Size = new Size(groupBox_head32.Size.Width, groupBox_head32.Size.Height);
                groupBox_head36.Location = new Point(groupBox_head33.Location.X, groupBox_head33.Location.Y + groupBox_head33.Size.Height + factor);
                groupBox_head36.Size = new Size(groupBox_head33.Size.Width, groupBox_head33.Size.Height);
                for (int i = tester.numHead; i < 36; i++) { g[i].Visible = false; }
            } else {
                int factor_groupbox = 0;
                int factor_groupbox_2 = 0;
                switch (tester.numHead) {//ไอศครีมเชอเบร็ตรสมะนาว
                    case 1: factor_groupbox = 82; factor_groupbox_2 = 1; groupBox_head2.Visible = false; break;
                    case 2: factor_groupbox = 82; factor_groupbox_2 = 1; break;
                    case 3: factor_groupbox = 88; factor_groupbox_2 = 2; groupBox_head4.Visible = false; break;
                    case 4: factor_groupbox = 88; factor_groupbox_2 = 2; break;
                    case 5: factor_groupbox = 92; factor_groupbox_2 = 3; groupBox_head6.Visible = false; break;
                    case 6: factor_groupbox = 92; factor_groupbox_2 = 3; break;
                    case 7: factor_groupbox = 96; factor_groupbox_2 = 4; groupBox_head8.Visible = false; break;
                    case 8: factor_groupbox = 96; factor_groupbox_2 = 4; break;
                    case 9: factor_groupbox = 104; factor_groupbox_2 = 5; groupBox_head10.Visible = false; break;
                    case 10: factor_groupbox = 104; factor_groupbox_2 = 5; break;
                    case 11: factor_groupbox = 108; factor_groupbox_2 = 6; groupBox_head12.Visible = false; break;
                    case 12: factor_groupbox = 108; factor_groupbox_2 = 6; break;
                    case 13: factor_groupbox = 112; factor_groupbox_2 = 7; groupBox_head14.Visible = false; break;
                    case 14: factor_groupbox = 112; factor_groupbox_2 = 7; break;
                    case 15: factor_groupbox = 116; factor_groupbox_2 = 8; groupBox_head16.Visible = false; break;
                    case 16: factor_groupbox = 116; factor_groupbox_2 = 8; break;
                    case 17: factor_groupbox = 120; factor_groupbox_2 = 9; groupBox_head18.Visible = false; break;
                    case 18: factor_groupbox = 120; factor_groupbox_2 = 9; break;
                    case 19: factor_groupbox = 124; factor_groupbox_2 = 10; groupBox_head20.Visible = false; break;
                    case 20: factor_groupbox = 124; factor_groupbox_2 = 10; break;
                }
                for (int i = 20; i < 36; i++) { g[i].Visible = false; }
                if (tester.numHead == 1) groupBox_head1.Size = new Size(((groupBox1.Width) - 2), ((this.Height - factor_groupbox) / factor_groupbox_2));
                else groupBox_head1.Size = new Size(((groupBox1.Width / 2) - 2), ((this.Height - factor_groupbox) / factor_groupbox_2));
                groupBox_head2.Size = new Size(((groupBox1.Width / 2) - 2), ((this.Height - factor_groupbox) / factor_groupbox_2));
                groupBox_head2.Location = new Point(groupBox_head1.Location.X + groupBox_head1.Width + 5, groupBox_head2.Location.Y);
                groupBox_head3.Location = new Point(groupBox_head1.Location.X, groupBox_head1.Location.Y + groupBox_head1.Size.Height + 5);
                groupBox_head3.Size = new Size(groupBox_head1.Size.Width, groupBox_head1.Size.Height);
                groupBox_head4.Location = new Point(groupBox_head2.Location.X, groupBox_head2.Location.Y + groupBox_head2.Size.Height + 5);
                groupBox_head4.Size = new Size(groupBox_head2.Size.Width, groupBox_head2.Size.Height);
                groupBox_head5.Location = new Point(groupBox_head3.Location.X, groupBox_head3.Location.Y + groupBox_head3.Size.Height + 5);
                groupBox_head5.Size = new Size(groupBox_head3.Size.Width, groupBox_head3.Size.Height);
                groupBox_head6.Location = new Point(groupBox_head4.Location.X, groupBox_head4.Location.Y + groupBox_head4.Size.Height + 5);
                groupBox_head6.Size = new Size(groupBox_head4.Size.Width, groupBox_head4.Size.Height);
                groupBox_head7.Location = new Point(groupBox_head5.Location.X, groupBox_head5.Location.Y + groupBox_head5.Size.Height + 5);
                groupBox_head7.Size = new Size(groupBox_head5.Size.Width, groupBox_head5.Size.Height);
                groupBox_head8.Location = new Point(groupBox_head6.Location.X, groupBox_head6.Location.Y + groupBox_head6.Size.Height + 5);
                groupBox_head8.Size = new Size(groupBox_head6.Size.Width, groupBox_head6.Size.Height);
                groupBox_head9.Location = new Point(groupBox_head7.Location.X, groupBox_head7.Location.Y + groupBox_head7.Size.Height + 5);
                groupBox_head9.Size = new Size(groupBox_head7.Size.Width, groupBox_head7.Size.Height);
                groupBox_head10.Location = new Point(groupBox_head8.Location.X, groupBox_head8.Location.Y + groupBox_head8.Size.Height + 5);
                groupBox_head10.Size = new Size(groupBox_head8.Size.Width, groupBox_head8.Size.Height);
                groupBox_head11.Location = new Point(groupBox_head9.Location.X, groupBox_head9.Location.Y + groupBox_head9.Size.Height + 5);
                groupBox_head11.Size = new Size(groupBox_head9.Size.Width, groupBox_head9.Size.Height);
                groupBox_head12.Location = new Point(groupBox_head10.Location.X, groupBox_head10.Location.Y + groupBox_head10.Size.Height + 5);
                groupBox_head12.Size = new Size(groupBox_head10.Size.Width, groupBox_head10.Size.Height);
                groupBox_head13.Location = new Point(groupBox_head11.Location.X, groupBox_head11.Location.Y + groupBox_head11.Size.Height + 5);
                groupBox_head13.Size = new Size(groupBox_head11.Size.Width, groupBox_head11.Size.Height);
                groupBox_head14.Location = new Point(groupBox_head12.Location.X, groupBox_head12.Location.Y + groupBox_head12.Size.Height + 5);
                groupBox_head14.Size = new Size(groupBox_head12.Size.Width, groupBox_head12.Size.Height);
                groupBox_head15.Location = new Point(groupBox_head13.Location.X, groupBox_head13.Location.Y + groupBox_head13.Size.Height + 5);
                groupBox_head15.Size = new Size(groupBox_head13.Size.Width, groupBox_head13.Size.Height);
                groupBox_head16.Location = new Point(groupBox_head14.Location.X, groupBox_head14.Location.Y + groupBox_head14.Size.Height + 5);
                groupBox_head16.Size = new Size(groupBox_head14.Size.Width, groupBox_head14.Size.Height);
                groupBox_head17.Location = new Point(groupBox_head15.Location.X, groupBox_head15.Location.Y + groupBox_head15.Size.Height + 5);
                groupBox_head17.Size = new Size(groupBox_head15.Size.Width, groupBox_head15.Size.Height);
                groupBox_head18.Location = new Point(groupBox_head16.Location.X, groupBox_head16.Location.Y + groupBox_head16.Size.Height + 5);
                groupBox_head18.Size = new Size(groupBox_head16.Size.Width, groupBox_head16.Size.Height);
                groupBox_head19.Location = new Point(groupBox_head17.Location.X, groupBox_head17.Location.Y + groupBox_head17.Size.Height + 5);
                groupBox_head19.Size = new Size(groupBox_head17.Size.Width, groupBox_head17.Size.Height);
                groupBox_head20.Location = new Point(groupBox_head18.Location.X, groupBox_head18.Location.Y + groupBox_head18.Size.Height + 5);
                groupBox_head20.Size = new Size(groupBox_head18.Size.Width, groupBox_head18.Size.Height);
            }
            if (tester.numHead > 10) {
                dataGridView_1.Visible = false; dataGridView_2.Visible = false;
                dataGridView_3.Visible = false; dataGridView_4.Visible = false;
                dataGridView_5.Visible = false; dataGridView_6.Visible = false;
                dataGridView_7.Visible = false; dataGridView_8.Visible = false;
                dataGridView_9.Visible = false; dataGridView_10.Visible = false;
                dataGridView_11.Visible = false; dataGridView_12.Visible = false;
                dataGridView_13.Visible = false; dataGridView_14.Visible = false;
                dataGridView_15.Visible = false; dataGridView_16.Visible = false;
                dataGridView_17.Visible = false; dataGridView_18.Visible = false;
                dataGridView_19.Visible = false; dataGridView_20.Visible = false;
                dataGridView_21.Visible = false; dataGridView_22.Visible = false;
                dataGridView_23.Visible = false; dataGridView_24.Visible = false;
                dataGridView_25.Visible = false; dataGridView_26.Visible = false;
                dataGridView_27.Visible = false; dataGridView_28.Visible = false;
                dataGridView_29.Visible = false; dataGridView_30.Visible = false;
                dataGridView_31.Visible = false; dataGridView_32.Visible = false;
                dataGridView_33.Visible = false; dataGridView_34.Visible = false;
                dataGridView_35.Visible = false; dataGridView_36.Visible = false;
            }
        }
        private int show_datagrid_int = 1;
        private void status_1_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 1;
        }
        private void status_2_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 2;
        }
        private void status_3_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 3;
        }
        private void status_4_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 4;
        }
        private void status_5_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 5;
        }
        private void status_6_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 6;
        }
        private void status_7_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 7;
        }
        private void status_8_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 8;
        }
        private void status_9_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 9;
        }
        private void status_10_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 10;
        }
        private void status_11_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 11;
        }
        private void status_12_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 12;
        }
        private void status_13_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 13;
        }
        private void status_14_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 14;
        }
        private void status_15_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 15;
        }
        private void status_16_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 16;
        }
        private void status_17_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 17;
        }
        private void status_18_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 18;
        }
        private void status_19_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 19;
        }
        private void status_20_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 20;
        }
        private void status_21_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 21;
        }
        private void status_22_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 22;
        }
        private void status_23_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 23;
        }
        private void status_24_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 24;
        }
        private void status_25_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 25;
        }
        private void status_26_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 26;
        }
        private void status_27_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 27;
        }
        private void status_28_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 28;
        }
        private void status_29_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 29;
        }
        private void status_30_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 30;
        }
        private void status_31_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 31;
        }
        private void status_32_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 32;
        }
        private void status_33_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 33;
        }
        private void status_34_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 34;
        }
        private void status_35_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 35;
        }
        private void status_36_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            show_datagrid_int = 36;
        }
        int save_datagrit_excel;
        private void dataGridView_1_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 1;
        }
        private void dataGridView_2_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 2;
        }
        private void dataGridView_3_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 3;
        }
        private void dataGridView_4_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 4;
        }
        private void dataGridView_5_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 5;
        }
        private void dataGridView_6_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 6;
        }
        private void dataGridView_7_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 7;
        }
        private void dataGridView_8_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 8;
        }
        private void dataGridView_9_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 9;
        }
        private void dataGridView_10_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 10;
        }
        private void dataGridView_11_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 11;
        }
        private void dataGridView_12_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 12;
        }
        private void dataGridView_13_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 13;
        }
        private void dataGridView_14_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 14;
        }
        private void dataGridView_15_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 15;
        }
        private void dataGridView_16_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 16;
        }
        private void dataGridView_17_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 17;
        }
        private void dataGridView_18_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 18;
        }
        private void dataGridView_19_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 19;
        }
        private void dataGridView_20_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 20;
        }
        private void dataGridView_21_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 21;
        }
        private void dataGridView_22_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 22;
        }
        private void dataGridView_23_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 23;
        }
        private void dataGridView_24_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 24;
        }
        private void dataGridView_25_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 25;
        }
        private void dataGridView_26_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 26;
        }
        private void dataGridView_27_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 27;
        }
        private void dataGridView_28_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 28;
        }
        private void dataGridView_29_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 29;
        }
        private void dataGridView_30_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 30;
        }
        private void dataGridView_31_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 31;
        }
        private void dataGridView_32_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 32;
        }
        private void dataGridView_33_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 33;
        }
        private void dataGridView_34_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 34;
        }
        private void dataGridView_35_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 35;
        }
        private void dataGridView_36_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Right) return;
            save_datagrit_excel = 36;
        }
        private void saveToolStripMenuItem_Click(object sender, EventArgs e) {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.csv)|*.csv";
            sfd.FileName = "export.csv";
            if (sfd.ShowDialog() != DialogResult.OK) return;
            StreamWriter swOut = new StreamWriter(sfd.FileName, true);
            DataGridView data = getDataGridView(save_datagrit_excel);

            for (int i = 0; i < data.RowCount; i++) {
                string str = "";
                for (int j = 0; j < data.ColumnCount; j++) {
                    str += data.Rows[i].Cells[j].Value + ",";  // "\t" สำหรับ excel และ "," สำหรับ csv 
                }
                swOut.WriteLine(str);
            }
            swOut.Close();
        }
        private void ctms_showCmd_Click(object sender, EventArgs e) {
            tester.showCMD = ctms_showCmd.Checked;
            setupPay.write_text(tester.headConfig.showCMD, tester.showCMD.ToString().ToUpper(), tester.nameFile);
        }
        private void prism_retest_text_pass_Click(object sender, EventArgs e) {
            string asd = "";
            while (true) {
                string input = Microsoft.VisualBasic.Interaction.InputBox("_ใส่ข้อความที่ prism pass", "prism pass", prism_retest_text_pass.Text, 500, 300);
                if (input == "") return;
                asd = input;
                break;
            }
            prism_retest_text_pass.Text = asd;
            File.WriteAllText("../../config/prism_retest_text_pass.txt", asd);
        }
        private void prism_retest_text_fail_Click(object sender, EventArgs e) {
            string asd = "";

            while (true) {
                string input = Microsoft.VisualBasic.Interaction.InputBox("_ใส่ข้อความที่ prism fail", "prism fail", 
                    prism_retest_text_fail.Text, 500, 300);
                if (input == "") return;
                asd = input;
                break;
            }

            prism_retest_text_fail.Text = asd;
            File.WriteAllText("../../config/prism_retest_text_fail.txt", asd);
        }
        private void setFlagAutomation1_Click(object sender, EventArgs e) {
            if (setFlagAutomation1.Checked) {
                autoMation.inTric1 = true;
            } else {

                autoMation.outTric1 = true;
            }
                
        }
        private void setFlagAutomation2_Click(object sender, EventArgs e) {
            if (setFlagAutomation2.Checked) {
                autoMation.inTric2 = true;
            } else {

                autoMation.outTric2 = true;
            }
                
        }
        private void ctms_excel_saveFile_sup() {
            if (tester.selectExcel) {
                excel.lastName = LastNameExcel.excel;
            }

            if (tester.selectLibre) {
                excel.lastName = LastNameExcel.libre;
            }
        }
        #endregion

        #region ============================================================== EXCEL ===============================================================
        private bool CheckSameStep() {
            string nameSheetBefore = "";
            string nameSheetAfter = "";

            if (excel.workSheet.GetNumber(excel.info.rowStart, excel.info.column).ToString() == excel.nan) {
                nameSheetBefore = excel.workSheet.GetText(excel.info.rowStart, excel.info.column);

            } else {
                nameSheetBefore = excel.workSheet.GetNumber(excel.info.rowStart, excel.info.column).ToString();
            }


            for (int hh = 1; hh < tester.numHead; hh++) {

                if (excel.workSheet.GetNumber(excel.info.rowStart + hh, excel.info.column).ToString() == excel.nan) {
                    nameSheetAfter = excel.workSheet.GetText(excel.info.rowStart + hh, excel.info.column);

                } else {
                    nameSheetAfter = excel.workSheet.GetNumber(excel.info.rowStart + hh, excel.info.column).ToString();
                }

                if (nameSheetBefore != nameSheetAfter) {
                    return false;
                }
            }


            if (excel.workSheet.GetNumber(excel.info.rowSequence, excel.info.column).ToString() == excel.nan) {
                nameSheetBefore = excel.workSheet.GetText(excel.info.rowSequence, excel.info.column);   
                                 
            } else {
                nameSheetBefore = excel.workSheet.GetNumber(excel.info.rowSequence, excel.info.column).ToString();
            }


            for (int hh = 1; hh < tester.numHead; hh++) {

                if (excel.workSheet.GetNumber(excel.info.rowSequence + hh, excel.info.column).ToString() == excel.nan) {
                    nameSheetAfter = excel.workSheet.GetText(excel.info.rowSequence + hh, excel.info.column);

                } else {
                    nameSheetAfter = excel.workSheet.GetNumber(excel.info.rowSequence + hh, excel.info.column).ToString();
                }

                if (nameSheetBefore != nameSheetAfter) {
                    return false;
                }
            }
            

            return true;
        }
        private void LoadDiscription() {
            excel.workSheet = excel.workBook.Worksheets[excel.info.nameSheet];

            this.Text = excel.workSheet.GetText(excel.info.rowCustomer, excel.info.column);

            if (excel.workSheet.GetNumber(excel.info.rowDetail, excel.info.column).ToString() == excel.nan) {
                tb_detail.Text = excel.workSheet.GetText(excel.info.rowDetail, excel.info.column) + "_" + cbb_fg.Text;

            } else {
                tb_detail.Text = excel.workSheet.GetText(excel.info.rowDetail, excel.info.column).ToString() + "_" + cbb_fg.Text; ;
            }

            if (excel.workSheet.GetNumber(excel.info.rowFirmware, excel.info.column).ToString() == excel.nan) {
                tb_fwVersion.Text = excel.workSheet.GetText(excel.info.rowFirmware, excel.info.column);

            } else {
                tb_fwVersion.Text = excel.workSheet.GetNumber(excel.info.rowFirmware, excel.info.column).ToString();
            }

            if (excel.workSheet.GetNumber(excel.info.rowSpecVersion, excel.info.column).ToString() == excel.nan) {
                tb_spec.Text = excel.workSheet.GetText(excel.info.rowSpecVersion, excel.info.column);

            } else {
                tb_spec.Text = excel.workSheet.GetNumber(excel.info.rowSpecVersion, excel.info.column).ToString();
            }

            excel.sameStep = CheckSameStep();

            for (int hh = 1; hh <= tester.numHead; hh++) {
                ClearDataGridView(hh);
            }

            for (int hh = 1; hh <= tester.numHead; hh++) {
                GetDescription(hh);
            }

            if (tester.useRelayCard && !tester.automation && !background_relay.IsBusy) {
                background_relay.RunWorkerAsync();
            }

            if (tester.useRelayCard && tester.automation && !automation_background.IsBusy) {
                automation_background.RunWorkerAsync();
                tcpip.RunWorkerAsync();
            }
        }
        private void GetDescription(int head) {
            DataGridView g = getDataGridView(head);

            excel.workSheet = excel.workBook.Worksheets[excel.info.nameSheet];
            int rowDatagrid = 0;
            int rowExcel = 1;
            excel.sheetTest = excel.workSheet.GetText((excel.info.rowStart + head) - 1, excel.info.column);

            try {
                excel.workSheet = excel.workBook.Worksheets[excel.sheetTest];
            } catch {
                MessageBox.Show(excel.info.errHead + head);
                return;
            }

            g.Rows.Add(GetRowExcel());
            excel.workSheet = excel.workBook.Worksheets[excel.sheetTest];

            string numberExcel;
            string detailExcel;
            string minExcel;
            string maxExcel;
            Color color;

            while (true) {
                rowExcel++;
                //string ggg = excel.workSheet.Range[rowExcel, 1].Style.KnownColor.ToString();

                if (excel.workSheet.Range[rowExcel, excel.pcba.columnNumber].Style.KnownColor.ToString() == excel.color.yellow || 
                    excel.workSheet.Range[rowExcel, excel.pcba.columnNumber].Style.KnownColor.ToString() == excel.color.red) {
                    continue;
                }

                numberExcel = excel.workSheet.GetText(rowExcel, excel.pcba.columnNumber);
                detailExcel = excel.workSheet.GetText(rowExcel, excel.pcba.columnDetail);
                minExcel = excel.workSheet.GetNumber(rowExcel, excel.pcba.columnMin).ToString();

                if (minExcel == excel.nan) {
                    minExcel = excel.workSheet.GetFormulaNumberValue(rowExcel, excel.pcba.columnMin).ToString();

                    if (minExcel == excel.nan) {
                        minExcel = excel.workSheet.GetText(rowExcel, excel.pcba.columnMin);
                    }
                }
                    
                if (minExcel == null) {
                    minExcel = excel.workSheet.GetFormulaStringValue(rowExcel, excel.pcba.columnMin);
                }

                if (minExcel != null) {
                    minExcel = minExcel.Trim();
                }

                maxExcel = excel.workSheet.GetNumber(rowExcel, excel.pcba.columnMax).ToString();

                if (maxExcel == excel.nan) {
                    maxExcel = excel.workSheet.GetFormulaNumberValue(rowExcel, excel.pcba.columnMax).ToString();

                    if (maxExcel == excel.nan) {
                        maxExcel = excel.workSheet.GetText(rowExcel, excel.pcba.columnMax);
                    }
                }
                    
                if (maxExcel == null) {
                    maxExcel = excel.workSheet.GetFormulaStringValue(rowExcel, excel.pcba.columnMax);
                }
                    
                if (maxExcel != null) {
                    maxExcel = maxExcel.Trim();
                }
                    
                if (numberExcel == null) {
                    numberExcel = excel.workSheet.GetNumber(rowExcel, excel.pcba.columnNumber).ToString();

                    if (numberExcel == excel.nan) {
                        break;
                    }
                }

                if (excel.workSheet.Range[rowExcel, excel.pcba.columnNumber].Style.KnownColor.ToString() != excel.color.none && 
                    excel.workSheet.Range[rowExcel, excel.pcba.columnNumber].Style.KnownColor.ToString() != excel.color.skyBlue && 
                    excel.workSheet.Range[rowExcel, excel.pcba.columnNumber].Style.KnownColor.ToString() != excel.color.gold) {
                    color = excel.workSheet.Range[rowExcel, excel.pcba.columnNumber].Style.Color;
                    g.Rows[rowDatagrid].DefaultCellStyle.BackColor = Color.FromArgb(color.A, color.R, color.G, color.B);
                }

                g.Rows[rowDatagrid].Cells[0].Value = numberExcel;
                g.Rows[rowDatagrid].Cells[1].Value = detailExcel;

                if (minExcel != "" && minExcel != null) {
                    g.Rows[rowDatagrid].Cells[2].Value = minExcel;
                }

                if (minExcel != "" && minExcel != null && maxExcel != "" && maxExcel != null) {
                    g.Rows[rowDatagrid].Cells[2].Value = minExcel + " - " + maxExcel;
                }

                rowDatagrid++;
            }
        }
        private int GetRowExcel() {
            int rowDataGridView = 0;
            string getExcel;

            for (int i = 2; i < 9999; i++) {

                if (excel.workSheet.Range[i, 1].Style.KnownColor.ToString() == excel.color.yellow || 
                    excel.workSheet.Range[i, 1].Style.KnownColor.ToString() == excel.color.red) {
                    continue;
                }

                getExcel = excel.workSheet.GetText(i, excel.pcba.columnNumber);

                if (getExcel == null) {
                    getExcel = excel.workSheet.GetNumber(i, excel.pcba.columnNumber).ToString();

                    if (getExcel == excel.nan) {
                        if (excel.workSheet.GetText(i, excel.pcba.columnDetail) != null) {

                            MessageBox.Show(excel.pcba.errColumn);
                            rowDataGridView = 1;
                        }

                        break;
                    }
                }

                rowDataGridView += 1;
            }

            return rowDataGridView;
        }
        private int GetRowSequenceExcel() {
            int rowDataGridView = 0;
            string getExcel;

            for (int i = 1; i < 9999; i++) {
                getExcel = excel.workSheet.GetText(i, excel.pcba.columnNumber);

                if (getExcel == null) {
                    getExcel = excel.workSheet.GetNumber(i, excel.pcba.columnNumber).ToString();

                    if (getExcel == excel.nan) {
                        break;
                    }
                }

                rowDataGridView += 1;
            }

            return rowDataGridView;
        }
        private int GetRowSteptestExcel() {
            int rowDataGridView = 0;
            string getExcel;

            for (int i = 2; i < 9999; i++) {

                if (excel.workSheet.Range[i, excel.pcba.columnNumber].Style.KnownColor.ToString() != excel.color.none && 
                    excel.workSheet.Range[i, excel.pcba.columnNumber].Style.KnownColor.ToString() != excel.color.skyBlue && 
                    excel.workSheet.Range[i, excel.pcba.columnNumber].Style.KnownColor.ToString() != excel.color.gold) {
                    continue;
                }

                getExcel = excel.workSheet.GetText(i, excel.pcba.columnNumber);

                if (getExcel == null) {
                    getExcel = excel.workSheet.GetNumber(i, excel.pcba.columnNumber).ToString();

                    if (getExcel == excel.nan) {
                        break;
                    }
                }

                rowDataGridView += 1;
            }

            return rowDataGridView;
        }
        #endregion

        #region ============================================================== Class ==============================================================
        public class Excel {
            public bool sameStep { get; set; }
            public string sheetTest { get; set; }
            public string stepTest { get; set; }
            public int[] row { get; set;}
            public string[] alice { get; set; }
            public Workbook workBook { get; set; }
            public Worksheet workSheet { get; set; }
            public Info info { get; set; }
            /// <summary>nan = "NaN"</summary>
            public string nan { get; set; }
            public Color color { get; set; }
            public string lastName { get; set; }
            public PCBA pcba { get; set; }
            /// <summary>pathFile = "../../TestDescription/"</summary>
            public string pathFile { get; set; }

            public Excel() {
                sameStep = false;
                sheetTest = "";
                stepTest = "";
                row = new int[36];
                alice = new string[5];
                workBook = new Workbook();
                info = new Info();
                nan = "NaN";
                color = new Color();
                pcba = new PCBA();
                pathFile = "../../TestDescription/";
            }
            public class Info {
                /// <summary>nameSheet = "info"</summary>
                public string nameSheet { get; set; }
                /// <summary>column = 2</summary>
                public int column { get; set; }
                /// <summary>rowStart = 7</summary>
                public int rowStart { get; set; }
                /// <summary>rowSequence = 45</summary>
                public int rowSequence { get; set; }
                /// <summary>rowCustomer = 1</summary>
                public int rowCustomer { get; set; }
                /// <summary>rowDetail = 2</summary>
                public int rowDetail { get; set; }
                /// <summary>rowSpecVersion = 4</summary>
                public int rowSpecVersion { get; set; }
                /// <summary>rowFirmware = 5</summary>
                public int rowFirmware { get; set; }
                /// <summary>errHead = "_ใน excel page info ไม่ได้กำหนด head"</summary>
                public string errHead { get; set; }

                public Info() {
                    nameSheet = "info";
                    column = 2;//เลขคอลั่ม หน้า info
                    rowStart = 7;//แถวเริ่มต้น ของหน้า info


                    //ไอศครีมเชอเบร็ตรสมะนาว
                    rowSequence = 45;//เอาเลขใน excel หน้า intro มาใส่ Sequence head 1 อ่ะ


                    rowCustomer = 1;
                    rowDetail = 2;
                    rowSpecVersion = 4;
                    rowFirmware = 5;
                    errHead = "_ใน excel page info ไม่ได้กำหนด head";
                }
            }
            public class Color {
                /// <summary>gold = "Gold"</summary>
                public string gold { get; set; }
                /// <summary>red = "Color2"</summary>
                public string red { get; set; }
                /// <summary>yellow = "Color5"</summary>
                public string yellow { get; set; }
                /// <summary>none = "None"</summary>
                public string none { get; set; }
                /// <summary>skyBlue = "SkyBlue"</summary>
                public string skyBlue { get; set; }

                public Color() {
                    gold = "Gold";
                    red = "Color2";
                    yellow = "Color5";
                    none = "None";
                    skyBlue = "SkyBlue";
                }
            }
            public class PCBA {
                /// <summary>columnNumber = 1</summary>
                public int columnNumber { get; set; }
                /// <summary>columnDetail = 2</summary>
                public int columnDetail { get; set; }
                /// <summary>columnMin = 3</summary>
                public int columnMin { get; set; }
                /// <summary>columnMax = 4</summary>
                public int columnMax { get; set; }
                /// <summary>errColumn = "_ใน excel คอลั่ม No ต้องมีเลขทุกบรรทัด"</summary>
                public string errColumn { get; set; }

                public PCBA() {
                    columnNumber = 1;
                    columnDetail = 2;
                    columnMin = 3;
                    columnMax = 4;
                    errColumn = "_ใน excel คอลั่ม No ต้องมีเลขทุกบรรทัด";
                }
            }
        }
        public class DataLog {
            /// <summary>nameFile = "datalog_config"</summary>
            public string nameFile { get; set; }
            public TimeLine timeLine { get; set; }
            public HeadConfig headConfig { get; set; }
            /// <summary>headLog = "Date,Time,Login ID,SW version,FW version,Spec version,Test Time(Sec),Load In/Out(Sec),
            /// Mode,Result,S/N,Failure,"</summary>
            public string headLog { get; set; }
            /// <summary>headLog_header = "Header,"</summary>
            public string headLog_header { get; set; }
            /// <summary>headLog_fgAndWo = "FG,WO,"</summary>
            public string headLog_fgAndWo { get; set; }
            /// <summary>lastNameTXT = ".txt"</summary>
            public string lastNameTXT { get; set; }
            /// <summary>lastNameCSV = ".csv"</summary>
            public string lastNameCSV { get; set; }

            public DataLog() {
                nameFile = "datalog_config";
                timeLine = new TimeLine();
                headConfig = new HeadConfig();
                headLog = "Date" + "," + "Time" + "," + "Login ID" + "," + "SW version" + "," + "FW version" + "," +
                    "Spec version" + "," + "Test Time(Sec)" + "," + "Load In/Out(Sec)" + "," + "Mode" + "," + "Result" +
                    "," + "S/N" + "," + "Failure" + ",";

                headLog_header = "Header" + ",";
                headLog_fgAndWo = "FG" + "," + "WO" + ",";
                lastNameTXT = ".txt";
                lastNameCSV = ".csv";
            }

            public class TimeLine {
                /// <summary>numFile = "TimeLine#"</summary>
                public string numFile { get; set; }
                public double rowCSV { get; set; }
                public string nameFile { get; set; }
                /// <summary>maxRow = 1000000</summary>
                public double maxRow { get; set; }
                /// <summary>nameFileRow = "row.txt"</summary>
                public string nameFileRow { get; set; }
                public string data { get; set; }


                public TimeLine() {
                    nameFile = "TimeLine#";
                    maxRow = 1000000;
                    nameFileRow = "row.txt";
                }
            }
            public class HeadConfig {
                public string fileTimeLine { get; set; }


                public HeadConfig() {
                    fileTimeLine = "File Time Line";
                }
            }
        }
        public class Tester {
            /// <summary>nameFile = "tester_config"</summary>
            public string nameFile { get; set; }
            public HeadConfig headConfig { get; set; }
            /// <summary>saveNormal = "Normal"</summary>
            public string saveNormal { get; set; }
            /// <summary>excel = "Excel"</summary>
            public string excel { get; set; }
            public bool upFail { get; set; }
            public string ScrollDatagrid { get; set; }
            public bool showCMD { get; set; }
            public bool click2ClearSN { get; set; }
            public string numCardRelay { get; set; }
            public bool useRelayCard { get; set; }
            public bool automation { get; set; }
            public bool testPanel { get; set; }
            public string cylinderRelay1 { get; set; }
            public string cylinderHead1 { get; set; }
            public string cylinderRelay2 { get; set; }
            public string cylinderHead2 { get; set; }
            public bool cylinder1 { get; set; }
            public bool cylinder2 { get; set; }
            public int numHead { get; set; }
            public string saveData { get; set; }
            public string nameDMM { get; set; }
            public bool selectExcel { get; set; }
            public bool selectLibre { get; set; }


            public Tester() {
                nameFile = "tester_config";
                headConfig = new HeadConfig();
                saveNormal = "Normal";
                excel = "Excel";
                cylinder1 = true;
                cylinder2 = true;
            }

            public class HeadConfig {
                public string arduinoComport { get; set; }
                public string komsonTester { get; set; }
                public string useRelayCard { get; set; }
                public string numHead { get; set; }
                public string ScrollDatagrid { get; set; }
                public string automation { get; set; }
                public string testAuto { get; set; }
                public string allowRetest { get; set; }
                public string upFail { get; set; }
                public string saveData { get; set; }
                public string click2ClearSN { get; set; }
                public string showCMD { get; set; }
                public string cylinderRelay1 { get; set; }
                public string cylinderHead1 { get; set; }
                public string cylinderRelay2 { get; set; }
                public string cylinderHead2 { get; set; }
                public string numCardRelay { get; set; }
                public string testPanel { get; set; }
                public string nameDMM { get; set; }
                public string fileTestDescription { get; set; }


                public HeadConfig() {
                    arduinoComport = "Arduino Comport Main";
                    komsonTester = "Komson Tester";
                    useRelayCard = "Use Relay Card";
                    numHead = "Number of Head";
                    ScrollDatagrid = "Scroll Datagrid";
                    automation = "Automation";
                    testAuto = "Test Auto";
                    allowRetest = "Allow Retest";
                    upFail = "Updata Fail to Prism";
                    saveData = "Save Data Type";
                    click2ClearSN = "Click to Clear SN";
                    showCMD = "Show CMD";
                    cylinderRelay1 = "Cylinder 1 Relay";
                    cylinderHead1 = "Cylinder 1 Head";
                    cylinderRelay2 = "Cylinder 2 Relay";
                    cylinderHead2 = "Cylinder 2 Head";
                    numCardRelay = "Num Card Relay";
                    testPanel = "Test Panel";
                    nameDMM = "Name DMM";
                    fileTestDescription = "File TestDescription";
                }
            }
        }
        public class TcpIP {
            /// <summary>nameFile = "tcptp_config"</summary>
            public string nameFile { get; set; }
            public HeadConfig headConfig { get; set; }
            public string ip { get; set; }
            public string port { get; set; }
            public string timerTric { get; set; }
            public bool useRobot { get; set; }
            public string readRobot1 { get; set; }
            public string readRobot2 { get; set; }


            public TcpIP() {
                nameFile = "tcptp_config";
                headConfig = new HeadConfig();
            }

            public class HeadConfig {
                public string ip { get; set; }
                public string port { get; set; }
                public string timerTric { get; set; }
                public string useRobot { get; set; }
                public string readRobot1 { get; set; }
                public string readRobot2 { get; set; }


                public HeadConfig() {
                    ip = "IP";
                    port = "Port";
                    timerTric = "Timer Tric";
                    useRobot = "Use Robot";
                    readRobot1 = "Read Robot 1";
                    readRobot2 = "Read Robot 2";
                }
            }
        }
        public class PrismTest {
            /// <summary>nameFile = "prism_config"</summary>
            public string nameFile { get; set; }
            public HeadConfig headConfig { get; set; }
            public string processBeforeText { get; set; }
            public bool processBefore { get; set; }
            public string digitSN { get; set; }
            /// <summary>Operation = "Operation"</summary>
            public string Operation { get; set; }
            /// <summary>OperationMode = "Operation Mode"</summary>
            public string OperationMode { get; set; }
            /// <summary>Debug = "Debug"</summary>
            public string Debug { get; set; }
            /// <summary>DebugMode = "Debug Mode"</summary>
            public string DebugMode { get; set; }
            /// <summary>success = "SUCCESS"</summary>
            public string success { get; set; }
            public bool upDataToKomson { get; set; }
            public string mode { get; set; }

            public PrismTest() {
                nameFile = "prism_config";
                headConfig = new HeadConfig();
                Operation = "Operation";
                OperationMode = "Operation Mode";
                Debug = "Debug";
                DebugMode = "Debug Mode";
                success = "SUCCESS";
            }

            public class HeadConfig {
                public string mode { get; set; }
                public string databaseName { get; set; }
                public string databaseServerTPP { get; set; }
                public string databaseServerTPR { get; set; }
                public string computerName { get; set; }
                public string stationName { get; set; }
                public string processName { get; set; }
                public string employeeID { get; set; }
                public string checkProcessBefore { get; set; }
                public string ProcessBefore { get; set; }
                public string digitSN { get; set; }
                public string upDataToKomson { get; set; }

                public HeadConfig() {
                    mode = "Mode";
                    databaseName = "Database Name";
                    databaseServerTPP = "Database Server TPP";
                    databaseServerTPR = "Database Server TPR";
                    computerName = "Computer Name";
                    stationName = "Station Name";
                    processName = "Process Name";
                    employeeID = "Employee ID";
                    checkProcessBefore = "Check Process Before";
                    ProcessBefore = "Process Before";
                    digitSN = "Digit SN";
                    upDataToKomson = "Up Data To Komson";
                }
            }
        }
        public class UpDataTest {
            /// <summary>value = "up_data_config"</summary>
            public string nameFile { get; set; }
            /// <summary>value = "up_data_config.txt"</summary>
            public string reReadConfig { get; set; }
            public bool callExe { get; set; }
            public bool waitUpData { get; set; }
            public HeadConfig headConfig { get; set; }

            public UpDataTest() {
                nameFile = "up_data_config";
                reReadConfig = "up_data_config.txt";
                headConfig = new HeadConfig();
            }
            public class HeadConfig
            {
                public string callExe { get; set; }
                public string waitUpData { get; set; }
                public HeadConfig()
                {
                    callExe = "Call Exe";
                    waitUpData = "Wait Up Data";
                }
            }
            public static class Define {
                /// <summary>value = "up_data_getOutPut.txt"</summary>
                public static readonly string getOutPut = "up_data_getOutPut.txt";
                /// <summary>value = "up_data_getOutPutOk.txt"</summary>
                public static readonly string getOutPutOk = "up_data_getOutPutOk.txt";
            }
        }
        public class Define {
            /// <summary>fStartConfig = "fStart_config.txt"</summary>
            public string fStartConfig { get; set; }
            /// <summary>pass = "PASS"</summary>
            public string pass { get; set; }
            /// <summary>fail = "FAIL"</summary>
            public string fail { get; set; }
            /// <summary>formClass = "WindowsFormsApplication1.Function_in_excel"</summary>
            public string formClass { get; set; }
            /// <summary>testing = "TESTING"</summary>
            public string testing { get; set; }
            public DataGrid dataGrid { get; set; }

            public Define() {
                fStartConfig = "fStart_config.txt";
                pass = "PASS";
                fail = "FAIL";
                formClass = "WindowsFormsApplication1.Function_in_excel";
                testing = "TESTING";
                dataGrid = new DataGrid();
            }
            public class DataGrid {
                /// <summary>columnStep = 0</summary>
                public int columnStep { get; set; }
                /// <summary>columnDetail = 1</summary>
                public int columnDetail { get; set; }
                /// <summary>columnSpec = 2</summary>
                public int columnSpec { get; set; }
                /// <summary>columnMeasure = 3</summary>
                public int columnMeasure { get; set; }
                /// <summary>columnResult = 4</summary>
                public int columnResult { get; set; }

                public DataGrid() {
                    columnStep = 0;
                    columnDetail = 1;
                    columnSpec = 2;
                    columnMeasure = 3;
                    columnResult = 4;
                }
            }
        }
        public class Automation {
            public bool inTric1 { get; set; }
            public bool inTric2 { get; set; }
            public bool outTric1 { get; set; }
            public bool outTric2 { get; set; }
        }
        public class Server {
            private Socket socket { get; set; }
            private List<Socket> listSocket { get; set; }
            private byte[] buffer { get; set; }

            public Server() {
                listSocket = new List<Socket>();
                buffer = new byte[65536];
            }
            public void Bind(fMain form, string ip = "127.8.8.8", int port = 2424) {
                socket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);

                while (true) {
                    try {
                        socket.Bind(new IPEndPoint(IPAddress.Parse(ip), port));
                        socket.Listen(0);
                        socket.BeginAccept(AcceptCallback, null);
                        break;

                    } catch {
                        form.Log(LogMsgType.Error_Red, "Bind Error...");
                        DelaymS(1000);
                    }
                }
            }
            private void AcceptCallback(IAsyncResult resultIAsync) {
                Socket socketSup;

                try {
                    socketSup = socket.EndAccept(resultIAsync);
                } catch {
                    return;
                }

                listSocket.Add(socketSup);
                socketSup.BeginReceive(buffer, 0, 65536, SocketFlags.None, ReceiveCallback, socketSup);
                IPEndPoint ipEndPointNewConnect = socketSup.RemoteEndPoint as IPEndPoint;
                socket.BeginAccept(AcceptCallback, null);
            }
            private void ReceiveCallback(IAsyncResult resultIAsync) {
                Thread.Sleep(50);
                Socket socketSup = (Socket)resultIAsync.AsyncState;
                int received;

                try {
                    received = socketSup.EndReceive(resultIAsync);
                    if (received == 0) {
                        ClientDisConnect(socketSup);
                        return;
                    }

                } catch {
                    ClientDisConnect(socketSup);
                    return;
                }

                byte[] buf = new byte[received];
                Array.Copy(buffer, buf, received);
                string text = Encoding.ASCII.GetString(buf);

                byte[] data = Encoding.ASCII.GetBytes(text);
                socketSup.Send(data);

                socketSup.BeginReceive(buffer, 0, 65536, SocketFlags.None, ReceiveCallback, socketSup);
            }
            private void ClientDisConnect(Socket socket) {
                IPEndPoint IPEndPoint = socket.RemoteEndPoint as IPEndPoint;

                socket.Close();
                listSocket.Remove(socket);
            }
        }
        public class Client {
            private Socket socket { get; set; }

            public bool Connect(string ip = "127.1.1.1", int port = 2424) {
                try {
                    socket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                } catch {
                    return false;
                }

                IAsyncResult result = null;
                bool success;
                try {
                    result = socket.BeginConnect(ip, port, null, null);
                    success = result.AsyncWaitHandle.WaitOne(2000, true);
                } catch { }

                if (socket.Connected) {
                    socket.EndConnect(result);
                    return true;
                }

                socket.Close();
                return false;
            }
            public void Close() {
                try {
                    socket.Shutdown(SocketShutdown.Both);
                } catch { }

                try {
                    socket.Close();
                } catch { }

                try {
                    socket.Dispose();
                } catch { }
            }
            public void Send(string data) {
                Connect();
                bool flagSend = true;
                byte[] buffer = Encoding.ASCII.GetBytes(data);

                try {
                    socket.Send(buffer, 0, buffer.Length, SocketFlags.None);
                } catch {
                    flagSend = false;
                }

                Close();
            }
        }

        public static class Folder {
            public static List<string> list = new List<string>();
            /// <summary>driveD = "D:\\"</summary>
            public static readonly string driveD = "D:\\";
            /// <summary>dataBase = " DATA BASE\\"</summary>
            public static readonly string dataBase = " DATA BASE\\";
            /// <summary>operationComplete = "\\Operation Mode\\Complete Test Result\\"</summary>
            public static readonly string operationComplete = "\\Operation Mode\\Complete Test Result\\";
            /// <summary>operationInComplete = "\\Operation Mode\\Incomplete Test Result\\"</summary>
            public static readonly string operationInComplete = "\\Operation Mode\\Incomplete Test Result\\";
            /// <summary>debugComplete = "\\Debug Mode\\Complete Test Result\\"</summary>
            public static readonly string debugComplete = "\\Debug Mode\\Complete Test Result\\";
            /// <summary>debugInComplete = "\\Debug Mode\\Incomplete Test Result\\"</summary>
            public static readonly string debugInComplete = "\\Debug Mode\\Incomplete Test Result\\";
            /// <summary>dataMIS = "D:\\DATA_MIS"</summary>
            public static readonly string dataMIS = "D:\\DATA_MIS";
            /// <summary>timeLine = "TimeLine\\"</summary>
            public static readonly string timeLine = "TimeLine\\";
        }
        public static class DateTimePay {
            /// <summary>format = "dd/MM/yyyy,HH:mm:ss,"</summary>
            public static readonly string format = "dd/MM/yyyy,HH:mm:ss,";
            /// <summary>format = "en-US"</summary>
            public static readonly string us = "en-US";
        }
        public static class LastNameExcel {
            /// <summary>Value = ".xlsx"</summary>
            public static readonly string excel = ".xlsx";
            /// <summary>Value = ".ods"</summary>
            public static readonly string libre = ".ods";
        }

        public class JsonConvert {
            public string Date { get; set; }
            public string Time { get; set; }
            public string LoginID { get; set; }
            public string SWVersion { get; set; }
            public string FWVersion { get; set; }
            public string SpecVersion { get; set; }
            public string TestTime { get; set; }
            public string LoadInOut { get; set; }
            public string Mode { get; set; }
            public string FinalResult { get; set; }
            public string SN { get; set; }
            public object Failure { get; set; }
            public List<ResultString_> ResultString { get; set; }

            public JsonConvert() {
                Date = string.Empty;
                Time = string.Empty;
                LoginID = string.Empty;
                SWVersion = string.Empty;
                FWVersion = string.Empty;
                SpecVersion = string.Empty;
                TestTime = string.Empty;
                LoadInOut = string.Empty;
                Mode = string.Empty;
                FinalResult = string.Empty;
                SN = string.Empty;
                Failure = string.Empty;
                ResultString = new List<ResultString_>();
            }
            public class ResultString_ {
                public string Step { get; set; }
                public string Description { get; set; }
                public string Tolerance { get; set; }
                public string Measured { get; set; }
                public string Result { get; set; }

                public ResultString_() {
                    Step = string.Empty;
                    Description = string.Empty;
                    Tolerance = string.Empty;
                    Measured = string.Empty;
                    Result = string.Empty;
                }
            }
        }
        #endregion
    }
}