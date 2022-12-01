using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using USBClassLibrary;
using Spire.Xls;
using System.Drawing;
using System.Reflection;
using System.Diagnostics;
using System.IO.Ports;
using System.Management;
using System.Text;
using System.Linq;

namespace WindowsFormsApplication1 {
    public class PSU {
        public string nameDMM { get; set; }
        public string nameOSC { get; set; }
        public Define define { get; set; }
        public MessageErr messageErr { get; set; }
        public DMM dmm { get; set; }
        public OSC osc { get; set; }


        public PSU() {
            define = new Define();
            messageErr = new MessageErr();
            dmm = new DMM();
            osc = new OSC();
        }

        public class DMM {
            public string connected { get; set; }
            public string clearErr { get; set; }
            public string readResister { get; set; }
            public string readVoltage { get; set; }
            public string readCurrent { get; set; }
            public string setLocal { get; set; }
            public string selectHold { get; set; }
            public string onHold { get; set; }
            public string offHold { get; set; }
            public string onBeeper { get; set; }
            public string offBeeper { get; set; }
            public string sourceOn { get; set; }
            public string sourceOff { get; set; }
            public string volt { get; set; }

            public DMM() {
                connected = "Connected";
                clearErr = "*CLS\n";
                readResister = ":MEAS:RES? ";
                readVoltage = ":MEAS:VOLT:DC? ";
                readCurrent = "MEAS:CURR:DC? ";
                setLocal = ":SYST:LOC\n";
                selectHold = "CALC:FUNC HOLD";
                onHold = "CALC:STAT ON";
                offHold = "CALC:STAT OFF";
                onBeeper = "SYSTem:BEEPer:STATe 1";
                offBeeper = "SYSTem:BEEPer:STATe 0";
                sourceOn = "OUTP ON";
                sourceOff = "OUTP OFF";
                volt = "VOLT ";
            }
        }
        public class OSC {
            public string readFetc { get; set; }
            public string clearErr { get; set; }
            public string read { get; set; }

            public OSC() {
                readFetc = "INIT;FETC?";
                clearErr = "CLS\n";
                read = ":READ?";
            }
        }
        public class Define {
            public string nameOSC { get; set; }
            public string connect { get; set; }
            public string OSC { get; set; }
            public string DMM { get; set; }
            public string deviceGet { get; set; }
            public string deviceDescription { get; set; }
            public string deviceName { get; set; }
            public string TF960 { get; set; }
            public string pidOSC { get; set; }
            public string vidOSC { get; set; }


            public Define() {
                nameOSC = "0x0957::0x1807";
                connect = "Connected";
                OSC = "Oscilloscope";
                DMM = "DMM";
                deviceGet = "SELECT * FROM Win32_PnPEntity WHERE Caption like '%(COM%'";
                TF960 = "TTi TF960";
                deviceDescription = "Description";
                deviceName = "NAME";
                pidOSC = "0492";
                vidOSC = "103E";
            }
        }
        public class MessageErr {
            public string connectDMM { get; set; }
            public string fineNameDMM { get; set; }
            public string fineNameOSC { get; set; }

            public MessageErr() {
                connectDMM = "\nDMM no connect";
                fineNameDMM = "\nCan't find DMM name";
                fineNameOSC = "\nCan't find OSC name";
            }
        }
    }
    public class SetFile {
        public void cameraList() {
            //เป็นไฟล์ที่กำหนดว่ากล้องตัวไหนจะเปิดก่อนเปิดหลัง ให้มันเรียงตามคิว อันไหนมาก่อนเปิดก่อน มาหลังเปิดหลัง
            File.WriteAllText("camera_show_list.txt", "");
        }
    }
    class Function_in_excel {
        public Function_in_excel(fMain f1, string strCMD) {
            fMain = f1;
            alice = fMain.excel.alice;
            eval(strCMD);
        }

        #region ===============================================Define====================================================
        private fMain fMain;
        private static PSU psu;
        private static Connect DcPSU;
        private SetFile setFile = new SetFile();

        static string[] alice = new string[5];
        private List<USBClassLibrary.USBClass.DeviceProperties> pidList;
        #endregion

        #region ===============================================Function Support====================================================
        private void define_class() {
            psu = new PSU();
            DcPSU = new Connect();
        }
        private void getNameDMM_visa() {
            Ivi.Visa.Interop.ResourceManager visa = new Ivi.Visa.Interop.ResourceManager();

            try {
                string[] visaList = visa.FindRsrc("?*");

                foreach (string list in visaList) {

                    if (list.Contains(fMain.tester.nameDMM)) {
                        psu.nameDMM = list;
                    }
                }

            } catch {
                fMain.Log(fMain.LogMsgType.Incoming_Blue, psu.messageErr.fineNameDMM);
                return;
            }

            if (psu.nameDMM == "" || psu.nameDMM == null) {
                psu.nameDMM = string.Empty;
                fMain.Log(fMain.LogMsgType.Incoming_Blue, psu.messageErr.fineNameDMM);
            }
        }
        private void getNameOSC_visa() {
            Ivi.Visa.Interop.ResourceManager visa = new Ivi.Visa.Interop.ResourceManager();

            try {
                string[] visaList = visa.FindRsrc("?*");

                foreach (string list in visaList) {

                    if (list.Contains(psu.define.nameOSC)) {
                        psu.nameOSC = list;
                    }
                }

            } catch {
                fMain.Log(fMain.LogMsgType.Incoming_Blue, psu.messageErr.fineNameOSC);
                return;
            }

            if (psu.nameOSC == "" || psu.nameOSC == null) {
                psu.nameOSC = string.Empty;
                fMain.Log(fMain.LogMsgType.Incoming_Blue, psu.messageErr.fineNameOSC);
            }
        }
        private void checkDMMconnect_visa() {
            fMain.bt_reserve1.Text = psu.define.DMM;

            if (DcPSU.ConnectInstr(psu.nameDMM) == psu.define.connect) {
                fMain.bt_reserve1.BackColor = Color.LimeGreen;
                DcPSU.DisConnectInstr();

            } else {
                fMain.bt_reserve1.BackColor = Color.Red;
            }
        }
        private void checkOSCconnect_visa() {
            fMain.bt_reserve2.Text = psu.define.OSC;

            if (DcPSU.ConnectInstr(psu.nameOSC) == psu.define.connect) {
                fMain.bt_reserve2.BackColor = Color.LimeGreen;
                DcPSU.DisConnectInstr();

            } else {
                fMain.bt_reserve2.BackColor = Color.Red;
            }
        }
        private void checkOSCconnect_deviceComport() {
            fMain.bt_reserve2.Text = psu.define.OSC;
            fMain.bt_reserve2.BackColor = Color.Red;

            ManagementObjectSearcher getComport = new ManagementObjectSearcher(psu.define.deviceGet);
            ManagementObjectCollection getComportAll = getComport.Get();

            foreach (ManagementObject nameList in getComportAll) {

                if (!nameList[psu.define.deviceDescription].ToString().Contains(psu.define.TF960)) continue;

                string[] nameComport = nameList.GetPropertyValue(psu.define.deviceName).ToString().Split('(', ')');
                psu.nameOSC = nameComport[1];

                fMain.bt_reserve2.BackColor = Color.LimeGreen;
            }
        }
        private void checkOSCconnect_pid() {
            fMain.bt_reserve2.Text = psu.define.OSC;

            pidList = new List<USBClass.DeviceProperties>();

            if (USBClass.GetUSBDevice(uint.Parse(psu.define.vidOSC, System.Globalization.NumberStyles.AllowHexSpecifier),
                uint.Parse(psu.define.pidOSC, System.Globalization.NumberStyles.AllowHexSpecifier), ref pidList, true, null)) {

                psu.nameOSC = pidList[0].COMPort;
                fMain.bt_reserve2.BackColor = Color.LimeGreen;

            } else {
                fMain.bt_reserve2.BackColor = Color.Red;
            }
        }

        private void checkSN_DLL(TextBox sn) {
            if (fMain.prismTest.mode != fMain.prismTest.Debug) {

                string[] status = TeamPrecision.PRISM.cSNs.CheckStatusSNv2(sn.Text, fMain.tb_wo.Text);

                //ตรวจสอบว่ามันตอบกลับด้วย "SUCCESS" ไหม
                if (status[0] == fMain.prismTest.success) {
                    fMain.flag_sn_pass[fMain.select_test - 1] = true;
                    return;
                }

                //ถ้ามันตอบกลับว่า ยังไม่ผ่านการเทส process ก่อนหน้า โปรแกรมจะแสตมป์ fail และหยุดเทสทันที
                if (status[1].Contains(fMain.prismTest.processBeforeText)) {
                    fMain.Log(fMain.LogMsgType.Error_Red, "\n" + status[1]);
                    fMain.UpdateResultToDataGrid(alice[0], status[1], fMain.define.fail);
                    return;
                }

                fMain.Log(fMain.LogMsgType.Incoming_Blue, "\n" + status[1]);

                //ถ้ามันตอบกลับว่า เทสไปแล้วแต่อยู่ใน สถานะ fail มันก็จะให้เทสได้
                if (status[1].Contains(fMain.prism_retest_text_fail.Text)) {
                    fMain.flag_sn_pass[fMain.select_test - 1] = true;
                    return;
                }

                //ถ้ามันตอบกลับว่่า เทสผ่านไปแล้ว มันจะเช็คอีกว่า ยอมให้เทสซ้ำไหมตัวที่ผ่านแล้ว ถ้าได้ก็ให้เทส
                if (status[1].Contains(fMain.prism_retest_text_pass.Text)) {
                    if (fMain.prism_retest.Checked) {
                        fMain.flag_sn_pass[fMain.select_test - 1] = true;
                        return;

                    } else {
                        //ถ้าไม่อยากโชว์ popup ให้เอา if นี้ออก
                        if (CallFormReTest()) {
                            fMain.flag_sn_pass[fMain.select_test - 1] = true;
                            return;
                        }

                        fMain.flagNotReTest[fMain.select_test - 1] = true;
                        fMain.row_test[fMain.select_test - 1] += 10000;
                        return;
                    }
                }

                //ถ้าไม่เข้าเงื่อนไข อะไรเลย จะเพิ่มข้อความ "_SN" ต่อท้ายไปใน sn เดิม และแสตมป์ fail ด้วย
                fMain.UpdateResultToDataGrid(alice[0], status[1], fMain.define.fail);
                sn.Text += "_SN";
            }
        }

        private void offAllRelay_head() {
            fMain.Relay_Off(fMain.select_test, fMain.bit1);
            fMain.Relay_Off(fMain.select_test, fMain.bit2);
            fMain.Relay_Off(fMain.select_test, fMain.bit3);
            fMain.Relay_Off(fMain.select_test, fMain.bit4);
            fMain.Relay_Off(fMain.select_test, fMain.bit5);
            fMain.Relay_Off(fMain.select_test, fMain.bit6);
            fMain.Relay_Off(fMain.select_test, fMain.bit7);
            fMain.Relay_Off(fMain.select_test, fMain.bit8);
            fMain.Relay_Off(fMain.select_test, fMain.bit9);
            fMain.Relay_Off(fMain.select_test, fMain.bit10);
            fMain.Relay_Off(fMain.select_test, fMain.bit11);
            fMain.Relay_Off(fMain.select_test, fMain.bit12);
            fMain.Relay_Off(fMain.select_test, fMain.bit13);
            fMain.Relay_Off(fMain.select_test, fMain.bit14);
            fMain.Relay_Off(fMain.select_test, fMain.bit15);
            fMain.Relay_Off(fMain.select_test, fMain.bit16);
        }
        private void GenFunctionToCSV() {
            List<string> function = new List<string>();

            foreach (MemberInfo memberInfo in this.GetType().GetMembers()) {

                if (memberInfo.Name == ".ctor" || memberInfo.Name == "Equals" || memberInfo.Name == "GetHashCode" ||
                    memberInfo.Name == "GetType" || memberInfo.Name == "ToString" || memberInfo.Name == "LoadTestSpec" ||
                    memberInfo.Name == "OnAllFunction" || memberInfo.Name == "intro_test" || memberInfo.Name == "after_test" ||
                    memberInfo.Name == "OffAllFunction") {
                    continue;
                }

                string nameFunction = memberInfo.Name;
                string parameter = "";

                foreach (ParameterInfo parameterInfo in ((MethodInfo)memberInfo).GetParameters()) {
                    parameter += "," + parameterInfo.Name;
                }

                function.Add(nameFunction + parameter);
            }

            File.WriteAllLines("AllFunctionInClass.csv", function);
        }

        public void camera_set_step(string cmd = "read2d") {
            File.WriteAllText("../../config/test_head_" + fMain.select_test + "_steptest.txt", cmd);
            File.Delete("test_head_" + fMain.select_test + "_result.txt");
        }
        public void set_timeout(string cmd = "1000") {
            File.WriteAllText("../../config/test_head_" + fMain.select_test + "_timeout.txt", cmd);
        }
        public void camera_set_list() {
            File.AppendAllText("camera_show_list.txt", fMain.select_test.ToString());
        }
        #endregion

        #region ===============================================Function Main====================================================
        public void delay_ms(string time) {
            fMain.DelaymS(Convert.ToInt32(time));
        }

        public void test_pass() {
            fMain.UpdateResultToDataGrid(alice[0], "PASS", "PASS");
        }
        public void test_fail() {
            fMain.UpdateResultToDataGrid(alice[0], "FAIL", "FAIL");
        }
        public void test_write(string data) {
            fMain.UpdateResultToDataGrid(alice[0], data, "PASS");
        }

        private TextBox getTextBox() {
            TextBox textBox = fMain.getTextBoxSN(fMain.select_test);

            return textBox;
        }
        private DataGridView getDataGridView() {
            DataGridView dataGridView = fMain.getDataGridView(fMain.select_test);

            return dataGridView;
        }
        public void camera_set_sn_to_textbox() {
            fMain.flag_sn_pass[fMain.select_test - 1] = false;

            TextBox t = fMain.getTextBoxSN(fMain.select_test);
            DataGridView d = fMain.getDataGridView(fMain.select_test);

            for (int i = 0; i < d.Rows.Count; i++) {

                if (d.Rows[i].Cells[0].Value.ToString() != alice[0]) continue;

                t.Text = d.Rows[i].Cells[3].Value.ToString();

                checkSN_DLL(t);

                break;
            }
        }
        public void dryice_scan2d_get() {
            string sn = File.ReadAllText("dryice_scan2d_sn_header_" + fMain.select_test + ".txt");
            File.Delete("dryice_scan2d_sn_header_" + fMain.select_test + ".txt");
            fMain.UpdateResultToDataGrid(alice[0], sn, "PASS");
        }
        private bool CallFormReTest_() {
            DialogResult dialogResult = MessageBox.Show("Want to repeat the test?", "ReTest", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes) {
                return true;
            }

            return false;
        }
        private Form formTest = new Form();
        private bool flagReTestPopUp;
        private bool CallFormReTest() {
            flagReTestPopUp = false;
            formTest.Size = new Size(710, 300);
            formTest.FormBorderStyle = FormBorderStyle.FixedSingle;
            formTest.StartPosition = FormStartPosition.CenterScreen;
            formTest.MaximizeBox = false;
            formTest.Text = fMain.select_test + ".ReTest";
            Label lb_head = new Label();
            lb_head.Text = "Head";
            lb_head.ForeColor = Color.Brown;
            lb_head.Location = new Point(0, 0);
            FontFamily fontFamily = new FontFamily("Arial");
            lb_head.Font = new Font(fontFamily, 70, FontStyle.Bold, GraphicsUnit.Pixel);
            lb_head.AutoSize = true;
            Label lb_number = new Label();
            lb_number.Text = fMain.select_test.ToString();
            lb_number.ForeColor = Color.Blue;
            lb_number.Location = new Point(10, 60);
            lb_number.Font = new Font(fontFamily, 200, FontStyle.Bold, GraphicsUnit.Pixel);
            lb_number.AutoSize = true;
            Label lb_detail = new Label();
            lb_detail.Text = "Want to repeat the test?";
            lb_detail.Font = new Font(fontFamily, 20, FontStyle.Bold, GraphicsUnit.Pixel);
            lb_detail.AutoSize = true;
            lb_detail.Location = new Point(350, 50);
            Button bt_reTest = new Button();
            bt_reTest.Click += ButtonReTestClick;
            bt_reTest.Text = "ReTest";
            bt_reTest.Font = new Font(fontFamily, 30, FontStyle.Bold, GraphicsUnit.Pixel);
            bt_reTest.Location = new Point(250, 150);
            bt_reTest.Size = new Size(200, 60);
            Button bt_cancel = new Button();
            bt_cancel.Click += ButtonCancelClick;
            bt_cancel.Text = "Cancel";
            bt_cancel.Font = new Font(fontFamily, 30, FontStyle.Bold, GraphicsUnit.Pixel);
            bt_cancel.Location = new Point(470, 150);
            bt_cancel.Size = new Size(200, 60);

            formTest.Controls.Add(lb_head);
            formTest.Controls.Add(lb_number);
            formTest.Controls.Add(lb_detail);
            formTest.Controls.Add(bt_reTest);
            formTest.Controls.Add(bt_cancel);
            formTest.ShowDialog();

            if (flagReTestPopUp) {
                return true;
            } else {
                return false;
            }
        }
        private void ButtonReTestClick(object sender, EventArgs e) {
            flagReTestPopUp = true;
            formTest.Close();
        }
        private void ButtonCancelClick(object sender, EventArgs e) {
            formTest.Close();
        }
        #endregion

        #region ================================================EXCEL========================================================
        private void eval(string string_of_function) {
            string name_function = "";
            List<string> parameter = new List<string>();
            string support_parameter = "";
            int num_parameter = 0;
            try {
                for (int i = 0; i < string_of_function.Length; i++) {
                    if (string_of_function.Substring(i, 1) == "(") {
                        break;
                    }
                    name_function += string_of_function.Substring(i, 1);
                }
                for (int i = name_function.Length + 1; i < string_of_function.Length; i++) {
                    if (string_of_function.Substring(i, 1) != "=") {
                        continue;
                    }
                    for (int j = i + 1; j < string_of_function.Length; j++) {
                        if (string_of_function.Substring(j, 1) == "," || string_of_function.Substring(j, 1) == ")") {
                            num_parameter++;
                            i = j;
                            parameter.Add(support_parameter);
                            support_parameter = "";
                            goto label_num_parameter;
                        }
                        support_parameter += string_of_function.Substring(j, 1);
                    }
                    label_num_parameter:;
                }
            } catch (Exception) {
                MessageBox.Show("header " + fMain.select_test + " function error " + string_of_function);
                fMain.UpdateResultToDataGrid(alice[0], "", "FAIL");
                return;
            }
            for (int reHead = 0; reHead <= 1; reHead++) {
                object[] objects = parameter.ConvertAll<object>(item => (object)item).ToArray();
                MethodInfo mi = this.GetType().GetMethod(name_function);
                try {
                    mi.Invoke(this, objects);
                    break;
                } catch {
                    if (reHead == 0) {
                        string headSup = name_function.Substring(0, 1);
                        string tailSup = name_function.Substring(1, name_function.Length - 1);
                        if (headSup.Any(char.IsUpper)) {
                            headSup = headSup.ToLower();
                        } else {
                            headSup = headSup.ToUpper();
                        }
                        name_function = headSup + tailSup;
                        continue;
                    }
                    MessageBox.Show("header " + fMain.select_test + " function error " + string_of_function);
                    fMain.UpdateResultToDataGrid(alice[0], "", "FAIL");
                    return;
                }
            }
        }
        #endregion

        #region ==============================================PSU Function=====================================================
        private void DisConnectPSU() {
            DcPSU.DisConnectInstr();
        }
        private void DmmConnect() {
            if (DcPSU.ConnectInstr(psu.nameDMM) != psu.dmm.connected) {
                fMain.Log(fMain.LogMsgType.Error_Red, "DMM_connect FAIL" + "\n");
            }
        }
        private void OscConnect() {
            if (DcPSU.ConnectInstr(psu.nameOSC) != psu.dmm.connected) {
                fMain.Log(fMain.LogMsgType.Error_Red, "OSC_connect FAIL" + "\n");
            }
        }
        private void DmmClearErr() {
            DcPSU.SendSLICCmd(psu.dmm.clearErr);
        }
        private double DmmReadResister(string rang = "10k") {
            //  100 | 1k | 10k | 100k | 1M | 10M | 100M | AUTO
            double V = 0;
            string result = DcPSU.QueryString(psu.dmm.readResister + rang);
            try {
                V = Convert.ToDouble(result);
            } catch { }
            return V;
        }
        private double DmmReadVoltageDC(string rank = "1V") {
            //  20mV | 100mV | 1V | 10V | 100V | 1000V | AUTO
            double V = 0;
            string result = DcPSU.QueryString(psu.dmm.readVoltage + rank);
            try {
                V = Convert.ToDouble(result);
            } catch { }
            return V;
        }
        private double DmmReadVoltageAC(string rank = "1V") {
            //  100mV | 1V | 10V | 100V | 750V | AUTO
            double V = 0;
            string result = DcPSU.QueryString(psu.dmm.readVoltage + rank);
            try {
                V = Convert.ToDouble(result);
            } catch { }
            return V;
        }
        private double DmmReadCurrent(string rank = "10mA") {
            // 10mA | 100mA | 1000mA | 3A | AUTO
            double V = 0;
            string result = DcPSU.QueryString(psu.dmm.readCurrent + rank);
            try {
                V = Convert.ToDouble(result);
            } catch { }
            return V;
        }
        private double OscReadFrequency() {
            double V = 0;
            DcPSU.SendSLICCmd(psu.osc.clearErr);
            string result = DcPSU.QueryString(psu.osc.readFetc);
            //string result = DcPSU.QueryString(psu.osc.read);
            try {
                V = Convert.ToDouble(result);
            } catch { }
            return V;
        }
        private void DmmLocal() {
            DcPSU.SendSLICCmd(psu.dmm.setLocal);
        }
        private void DmmHold(bool status) {
            DcPSU.SendSLICCmd(psu.dmm.selectHold);
            if (status) {
                DcPSU.SendSLICCmd(psu.dmm.onHold);
            } else {
                DcPSU.SendSLICCmd(psu.dmm.offHold);
            }
        }
        private void DmmBeeper(bool status) {
            if (status) {
                DcPSU.SendSLICCmd(psu.dmm.onBeeper);
            } else {
                DcPSU.SendSLICCmd(psu.dmm.offBeeper);
            }
        }
        private void DmmSourceVoltage(bool status, string voltage = "5V") {
            //  S1  >>  0 - 30V
            //  S2  >>  0 - 8V
            //S1S2  >>  0 - 30V
            if (!status) {
                DcPSU.SendSLICCmd(psu.dmm.sourceOff);
                return;
            }

            DcPSU.SendSLICCmd(psu.dmm.volt + voltage);
            DcPSU.SendSLICCmd(psu.dmm.sourceOn);
        }
        private void DmmSourceCurrent(bool status, string current = "500mA") {
            //  S1  >>  0 - 1A
            //  S2  >>  0 - 3A
            //S1S2  >>  0 - 3A
            if (!status) {
                DcPSU.SendSLICCmd(psu.dmm.sourceOff);
                return;
            }

            DcPSU.SendSLICCmd("CURRent " + current);
            DcPSU.SendSLICCmd(psu.dmm.sourceOn);
        }
        private void DmmSourceLimitCurrent(string current = "50mA") {
            //  S1  >>  0 - 1A
            //  S2  >>  0 - 3A
            DcPSU.SendSLICCmd("CURRent:LIMit " + current);
        }
        private void DmmSourceLimitVolt(string voltage = "5V") {
            //  S1  >>  0 - 30V
            //  S2  >>  0 - 8V
            DcPSU.SendSLICCmd("VOLTage:LIMit " + voltage);
        }
        private double DmmSourceGetCurrent() {
            double V = 0;
            string result = DcPSU.QueryString("SENS:CURR?");
            try {
                V = Convert.ToDouble(result);
            } catch { }
            return V;
        }
        private double DmmSourceGetVolt() {
            double V = 0;
            string result = DcPSU.QueryString("SENS:VOLT?");
            try {
                V = Convert.ToDouble(result);
            } catch { }
            return V;
        }
        #endregion

        public void LoadTestSpec() {
            //getNameDMM_visa();
            //psu.define.nameOSC = fMain.setupPay.read_text("Name OSC", "tester_config");
            //getNameOSC_visa();

            //checkDMMconnect_visa();
            //checkOSCconnect_visa();
            //checkOSCconnect_deviceComport();
            //checkOSCconnect_pid();
        }//ชังก์ชั่นนี้จะทำงานตอน เลือก FG
        public void setup() {
            define_class();
            setFile.cameraList();

            //uart_set_port(fMain.bit16, 2, "0403", "6015", "10C4", "EA60", "USB Serial Port");

        } //ฟังก์ชั่นนี้จะทำงานครั้งเดียวหลังจากเปิดโปรแกรมเทส
        public void OnAllFunction() {

        } //ฟังก์ชั่นนี้จะทำงาน ก่อน เริ่มเทสทั้งพาแนล *แต่ห้ามควบคุม relay card ในนี้
        public void intro_test() {

        } //ฟังก์ชั่นนี้จะทำงานก่อนเริ่มเทสหัวใดหัวหนึ่ง
        public void after_test() {
            if (!fMain.tester.useRelayCard) {
                after_test_sup();
                return;
            }

            offAllRelay_head();

            after_test_user();

            after_test_sup();
        }  //ฟังก์ชั่นนี้จะทำงานหลังจากเทสหัวใดหัวหนึ่งเสร็จ แต่ไม่อนุญาติให้แก้ไขในฟังก์ชั่นนี้
        private void after_test_sup() {
            if (fMain.tester.cylinder1) {

                string[] cylinderSplit = fMain.tester.cylinderHead1.Split('&');
                bool cylinderFlag = true;

                foreach (string cylinder in cylinderSplit) {

                    if (fMain.flag_test[Convert.ToInt32(cylinder) - 1]) {
                        cylinderFlag = false;
                    }
                }

                if (cylinderFlag) fMain.autoMation.outTric1 = true;
            }


            if (fMain.tester.cylinder2) {

                string[] cylinderSplit = fMain.tester.cylinderHead2.Split('&');
                bool cylinderFlag = true;

                foreach (string nn in cylinderSplit) {

                    if (fMain.flag_test[Convert.ToInt32(nn) - 1]) {
                        cylinderFlag = false;
                    }
                }

                if (cylinderFlag) fMain.autoMation.outTric2 = true;
            }
        }//ไม่ต้องยุ่งกับฟังก์ชั่นนี้
        private void after_test_user() {
            //fMain.write_23017("rgbset,0,0,0,0,0\n");

            //fMain.OffAllCard(fMain.select_test);
            //fMain.Relay_Off(fMain.select_test, fMain.bit1);
            //fMain.Relay_Off(fMain.select_test, fMain.bit2);
            //fMain.Relay_Off(fMain.select_test, fMain.bit3);
            //fMain.Relay_Off(fMain.select_test, fMain.bit4);
            //fMain.Relay_Off(fMain.select_test, fMain.bit5);
            //fMain.Relay_Off(fMain.select_test, fMain.bit6);
            //fMain.Relay_Off(fMain.select_test, fMain.bit7);
            //fMain.Relay_Off(fMain.select_test, fMain.bit8);
            //fMain.Relay_Off(fMain.select_test, fMain.bit9);
            //fMain.Relay_Off(fMain.select_test, fMain.bit10);
            //fMain.Relay_Off(fMain.select_test, fMain.bit11);
            //fMain.Relay_Off(fMain.select_test, fMain.bit12);
            //fMain.Relay_Off(fMain.select_test, fMain.bit13);
            //fMain.Relay_Off(fMain.select_test, fMain.bit14);
            //fMain.Relay_Off(fMain.select_test, fMain.bit15);
            //fMain.Relay_Off(fMain.select_test, fMain.bit16);


            after_test_sup();
        }//ให้มาแก้ไขในฟังก์ชั่นนี้แทน
        public void OffAllFunction() {
            if (fMain.tester.useRelayCard) {

            }


        }   //ฟังก์ชั่นนี้จะทำงาน หลัง เทสทั้งพาแนลเสร็จ *แต่ห้ามควบคุม relay card ในนี้


        #region ============================================== Function User =====================================================

        public void TestTest() {
            fMain.MatrixOn(fMain.Matrix.Row1.Columns1);
        }





        #endregion

    }
}