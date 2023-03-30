using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Management;
using Port_Setting;

namespace DPB_Update_tool
{
    public enum _Model
    { 
        DPB_none = 0,
        DPB30A = 30,
        DPB30_5031,
        DPB30_5032,
        DPB30_5033,
        DPB60A = 60,
        DPB60_5051,
        DPB60_5052,
    }
    public partial class Form1 : Form
    {
        public Thread[] MPTest_Thread;
        const string ToolVersion = "0.1.0";
        static int[] drive_name = new int[4];
        static int[] process_status = new int[4];
        static int[] scanportmap_insert = new int[4];
        static string[] scanportmap_insert_InstanceID = new string[4];
        int temp_drive_name = 0;
        int DUT_CMD_Ver = 0;
        static int MPcount = 1;
        _Model _DUT = _Model.DPB_none;
        const int STATUS_READY = 0;
        const int STATUS_PROCESSING = 1;
        const int STATUS_NODEVICE = 2;
        const int STATUS_FINISH = 3;
        const string PID_EMPTY = "EMPTY";   //Support PID Identify
        public Form1()
        {
            InitializeComponent();
            LoadSetting();
        }

        //============= 確認Storage可寫入 ==============//
        private bool CheckWritableForUpgrade(int dev, bool LockStatus, int version)
        {
            int setCount = 0;

            if (LockStatus)
            {
                if (version == 2)
                {
                    setCount = 0;
                    System.Threading.Thread.Sleep(200);
                    while (!NvUSBcmd.SetNOVAData((byte)globalVarManager.USBCMD_SetPassword, globalVarManager.CMD42password, drive_name[dev]))
                    {
                        setCount++;
                        //LogManager.PrintLog(dev, "Fail to send cmd (Set)" + globalVarManager.USBCMD_SetPassword + " setCount:" + setCount + " drive_name: " + drive_name[dev] + " (" + dev + ")");
                        System.Threading.Thread.Sleep(1000);
                        if (setCount > 5)
                        {
                            //ShowInfo(0, dev, "傳送指令失敗: " + globalVarManager.USBCMD_SetPassword, "");
                            return false;
                        }
                    }
                }

                setCount = 0;
                byte[] unlockCMD = new byte[30];
                System.Threading.Thread.Sleep(200);
                while (!NvUSBcmd.SetNOVAData((byte)globalVarManager.USBCMD_UnLock, unlockCMD, drive_name[dev]))
                {
                    setCount++;
                    //LogManager.PrintLog(dev, "Fail to send cmd (Set)" + globalVarManager.USBCMD_UnLock + " setCount:" + setCount + " drive_name: " + drive_name[dev] + " (" + dev + ")");
                    System.Threading.Thread.Sleep(1000);
                    if (setCount > 5)
                    {
                        //ShowInfo(0, dev, "傳送指令失敗: " + globalVarManager.USBCMD_UnLock, "");
                        return false;
                    }
                }
                unlockCMD = null;

                if (!WaitForDiskJustCheckFlag(dev, true, false, 0))
                {
                    return false;
                }

                GC.Collect();
                return true;
            }
            else
            {
                GC.Collect();
                return true;
            }
        }
        //複製
        public bool CopyTestFile(int dev, bool lockstatus, int USBFlow, string filepath)
        {
            Object CopyLock = new Object();
          //  if (!CheckWritableForUpgrade(dev, lockstatus, globalVarManager.USBCMD_Flow)) return false;

            //lock (CopyLock)
            {
                string diskstring = GetDirectory(dev, false);
                string sFilePath = Application.StartupPath + "\\" + filepath;
                string tFileName = filepath.Split('\\').Last();
                string tFilePath = diskstring + ":\\" + tFileName;
                string rename = "";
                Console.WriteLine("1 diskstring= " + diskstring);
                Console.WriteLine("2 sFilePath= " + sFilePath);
                Console.WriteLine("3 tFileName= " + tFileName);
                Console.WriteLine("4 tFilePath= " + tFilePath);
                //複製檔案至裝置
                try
                {
                    if (Directory.Exists(diskstring + ":"))
                    {
                        if (File.Exists(sFilePath))
                        {
                            try
                            {
                                File.Copy(sFilePath, tFilePath, true);
                                switch (_DUT)
                                {
                                    case _Model.DPB60_5052: 
                                        rename = "FWDB5052.bin";
                                        break;
                                    case _Model.DPB60_5051:
                                        rename = "FWDB5051.bin";
                                        break;
                                    case _Model.DPB60A:
                                        rename = "FWDPB60A.bin";
                                        break;
                                    case _Model.DPB30A:
                                        rename = "FWDPB30.bin";
                                        break;
                                    case _Model.DPB30_5031:
                                        rename = "FWDB5031.bin";
                                        break;
                                    case _Model.DPB30_5032:
                                        rename = "FWDB5032.bin";
                                        break;
                                    case _Model.DPB30_5033:
                                        rename = "FWDB5033.bin";
                                        break;
                                }
                                File.Move(tFilePath, diskstring + ":\\" + rename);
                            }
                            catch (Exception ex)
                            {
                                string e = ex.ToString();
                                //LogManager.PrintLog(dev, tFileName + " copy fail");
                                label9.Text = tFilePath + "  1.does not exist";
                                Console.WriteLine("1.does not exist " );
                                return false;
                            }
                        }
                        else
                        {
                            //LogManager.PrintLog(dev, sFilePath + " does not exist");
                            Console.WriteLine("2.does not exist ");
                            return false;
                        }
                    }
                    else
                    {
                        return false;
                    }
                }
                catch { return false; }

                //檢查檔案是否正確複製
                if (!File.Exists(tFilePath))
                {
                    Console.WriteLine("3.does not exist ");
                    //LogManager.PrintLog(dev, tFilePath + " does not exist");
                    return false;
                }
            }
            return true;
        }
        private bool GetSSID(int dev)
        {

            int setCount = 0;
            bool bResult = true;

            USBCmdUnLocker(dev);

            System.Threading.Thread.Sleep(100);
            byte[] ssid = new byte[256];
            string ssid_str = "";
            //string ssidname = globalVar.multi_ConnectionName[WiFiTest_Count - 1];
            //ShowInfo(4, dev, globalVar.multi_ConnectionName[WiFiTest_Count - 1], true);
            //byte[] setssid = Encoding.ASCII.GetBytes(ssidname);


            setCount = 0;
            while (!NvUSBcmd.GetNOVAData((byte)globalVarManager.USBCMD_WiFiSSID, ref ssid, drive_name[dev]))
            {
                setCount++;
                System.Threading.Thread.Sleep(1000);
                if (setCount > 5)
                {
                    bResult = false;
                    //globalFunction.PrintLog(dev, globalVar.USBCMD_WiFiSSID + "(G) WiFiSSID Command send fail");
                    //MessageBox.Show("9(G) Command send fail");
                    break;
                }
            }
            try
            {
                if (globalVarManager.USBCMD_CMDStringVersion == 1)
                {
                    if (ssid[1] >= 48 && ssid[1] <= 57)
                        ssid_str = ASCIIEncoding.ASCII.GetString(ssid, 2, 10 * (ssid[0] - 0x30) + (ssid[1] - 0x30));
                    else
                        ssid_str = ASCIIEncoding.ASCII.GetString(ssid, 1, ssid[0] - 0x30);
                }
                else if (globalVarManager.USBCMD_CMDStringVersion == 2)
                    ssid_str = ASCIIEncoding.ASCII.GetString(ssid, 0, 31);
            }
            catch { bResult = false; }

            if (ssid_str.Contains('\0'))
            {
                ssid_str = ssid_str.Substring(0, ssid_str.IndexOf('\0'));
            }

            //if (!(string.Equals((string)ssidname, (string)ssid_str)))
            //    bResult = false;
            if(bResult)
                label1.Text = ssid_str;
            GC.Collect();
            return bResult;
        }
        private bool GetModel(int dev)
        {
            int setCount = 0;
            bool bResult = true;
            byte[] device_model_name = new byte[30];
            string DeviceModelname = "";

            System.Threading.Thread.Sleep(100);
            setCount = 0;
            while (!NvUSBcmd.GetNOVAData((byte)globalVarManager.USBCMD_DeviceModel, ref device_model_name, drive_name[0]))
            {
                setCount++;
                System.Threading.Thread.Sleep(1000);
                if (setCount > 5)
                {
                    MessageBox.Show("傳送失敗 " + globalVarManager.USBCMD_DeviceModel);
                    bResult = false;
                    break;
                }
            }
            DeviceModelname = System.Text.Encoding.UTF8.GetString(device_model_name, 0, 6);
            label6.Text = DeviceModelname;

            return bResult;
        }
        private bool GetFWver(int dev)
        {
            int setCount = 0;
            bool bResult = true;
            byte[] btFwVer = new byte[30];
            string DeviceFwVer = "";

            System.Threading.Thread.Sleep(100);
            setCount = 0;
            while (!NvUSBcmd.GetNOVAData((byte)globalVarManager.USBCMD_FWVersion, ref btFwVer, drive_name[0]))
            {
                setCount++;
                System.Threading.Thread.Sleep(1000);
                if (setCount > 5)
                {
                    MessageBox.Show("傳送失敗 " + globalVarManager.USBCMD_FWVersion);
                    bResult = false;
                    break;
                }
            }
            DeviceFwVer = UTF8Encoding.UTF8.GetString(btFwVer).Trim('\0');
            label7.Text = DeviceFwVer;
            return bResult;
        }
        private bool Reboot(int dev, bool StorageLockStatus)
        {
            int setCount = 0;

            setCount = 0;
            byte[] rebootCMD = new byte[30];
            System.Threading.Thread.Sleep(200);
            while (!NvUSBcmd.SetNOVAData((byte)globalVarManager.USBCMD_Reboot, rebootCMD, drive_name[dev]))
            {
                setCount++;
                //LogManager.PrintLog(dev, "Fail to send cmd (Set)" + globalVarManager.USBCMD_Reboot + " setCount:" + setCount + " drive_name: " + drive_name[dev] + " (" + dev + ")");
                System.Threading.Thread.Sleep(1000);
                if (setCount > 5)
                {
                    //ShowInfo(0, dev, "傳送指令失敗: " + globalVarManager.USBCMD_Reboot, "");
                    return false;
                }
            }
            rebootCMD = null;
            System.Threading.Thread.Sleep(5000); //Make sure device is left (start rebooting process)

            if (!WaitForDiskJustCheckFlag(dev, StorageLockStatus, false, 0))
                return false;

            return true;
        }
        private bool Reset(int dev)
        {
            int setCount = 0;
            bool bResult = true;

            USBCmdUnLocker(dev);
            int[] data_Reset = new int[1];
            int[] data_Format = new int[1];
            byte[] btCmd = new byte[30];

            while (!NvUSBcmd.SetNOVAState((byte)globalVarManager.USBCMD_Restore, data_Reset, drive_name[dev]))
            {
                setCount++;
                System.Threading.Thread.Sleep(1000);
                if (setCount > 5)
                {
                    bResult = false;
                    //globalFunction.PrintLog(dev, globalVar.USBCMD_SettingReset + "(S) SettingReset Command send fail");
                    //MessageBox.Show("25(S) Command send fail");
                    return bResult;
                }
            }
            Thread.Sleep(500);
            WaitDiskBack(dev, true,false);
            if (!bResult) return false;
            return true;
        }
        private bool USBCmdUnLocker(int dev)
        {
            int[] CmdStatus = new int[1];
            int setCount = 0;
            setCount = 0;
            System.Threading.Thread.Sleep(200);
            while (!NvUSBcmd.GetNOVAState((byte)globalVarManager.USBCMD_QCUnlock, ref CmdStatus, drive_name[dev]))
            {
                setCount++;
                //LogManager.PrintLog(dev, "Port : (" + dev + ") - Get USBCMD_QCUnlock Command" + "(" + globalVarManager.USBCMD_QCUnlock + ")" + "drive_name:" + drive_name[dev] + " send fail, setCount=" + setCount);
                System.Threading.Thread.Sleep(1000);
                if (setCount > 5)
                    return false;
            }
            if (Convert.ToInt32(CmdStatus[0]) == 0)
            {
                setCount = 0;
                System.Threading.Thread.Sleep(200);
                while (!NvUSBcmd.SetNOVAData((byte)globalVarManager.USBCMD_QCUnlock, globalVarManager.QCPassWord, drive_name[dev]))
                {
                    setCount++;
                    //LogManager.PrintLog(dev, "Port : (" + dev + ") - Set USBCMD_QCUnlock Command" + "(" + globalVarManager.USBCMD_QCUnlock + ")" + "drive_name:" + drive_name[dev] + " send fail, setCount=" + setCount);
                    System.Threading.Thread.Sleep(1000);
                    if (setCount > 5)
                        return false;
                }
            }
            return true;
        }
        private void scan_Click(object sender, EventArgs e)
        {
            //程序開始，先鎖上Button
            Scan_btn.Enabled = false;
            Upgrade_btn.Enabled = false;
            Reset_btn.Enabled = false;
            Scan_btn.Text = "waiting...";
            label1.Text = "";
            label5.Text = "";
            label6.Text = "";
            label7.Text = "";
            label9.Text = "";

            //定port數歸零
            MPcount = 0;

            //這裡先掃一遍是否有已定port的機器在PC上了，並把dev_location設定好
            string deviceVIDPIDRaw = "";
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_DiskDrive");
            foreach (ManagementObject WMIObject in searcher.Get())
            {
                string DeviceID = WMIObject["DeviceID"].ToString();
                string PNPDeviceID = WMIObject["PNPDeviceID"].ToString();
                foreach (string USB_VEN_Member in globalVarManager.USB_VEN)
                {
                    int sIndex = PNPDeviceID.IndexOf(USB_VEN_Member, 0);

                    if (sIndex > 0)
                    {
                        try
                        {
                            int subStringIndex = PNPDeviceID.IndexOf("\\", 10);
                            string temp_dev_InstanceID = PNPDeviceID.Substring(subStringIndex + 1, PNPDeviceID.Length - subStringIndex - 1);
                            temp_dev_InstanceID = temp_dev_InstanceID.Substring(0, temp_dev_InstanceID.Length - 2);
                            string temp_dev_location = PortInfo.Get_USBPort(temp_dev_InstanceID, ref deviceVIDPIDRaw);

                            // 判斷機種
                            if (globalVarManager.ModelPID != PID_EMPTY)
                            {
                                if (deviceVIDPIDRaw.Contains("PID_5000"))
                                {
                                    if (PNPDeviceID.Contains("DPB60")) { _DUT = _Model.DPB60A; DUT_CMD_Ver = 2; }
                                    else if (PNPDeviceID.Contains("DPB30")) { _DUT = _Model.DPB30A; DUT_CMD_Ver = 1; }
                                }
                                else if (deviceVIDPIDRaw.Contains("PID_5031")) { _DUT = _Model.DPB30_5031; DUT_CMD_Ver = 2; }
                                else if (deviceVIDPIDRaw.Contains("PID_5032")) { _DUT = _Model.DPB30_5032; DUT_CMD_Ver = 2; }
                                else if (deviceVIDPIDRaw.Contains("PID_5033")) { _DUT = _Model.DPB30_5033; DUT_CMD_Ver = 2; }
                                else if (deviceVIDPIDRaw.Contains("PID_5051")) { _DUT = _Model.DPB60_5051; DUT_CMD_Ver = 2; }
                                else if (deviceVIDPIDRaw.Contains("PID_5052")) { _DUT = _Model.DPB60_5052; DUT_CMD_Ver = 2; }
                            }
                            //MessageBox.Show(_DUT.ToString());
                            //MessageBox.Show(DUT_CMD_Ver.ToString());
                            label6.Text = _DUT.ToString();
                            if (DUT_CMD_Ver == 2) {  //dpb30 Gen2/ Gen3 dpb60 Gen1/ Gen2/ Gen3
                                globalVarManager.USBCMD_QCUnlock = 130;
                                globalVarManager.USBCMD_DeviceModel = 14;
                                globalVarManager.USBCMD_WiFiSSID = 191;
                                globalVarManager.USBCMD_Reboot = 77;
                                globalVarManager.USBCMD_FWVersion = 17;
                                globalVarManager.USBCMD_CMDStringVersion = 2;
                                globalVarManager.USBCMD_Restore = 71;
                            } else {    //dpb30 Gen1
                                globalVarManager.USBCMD_QCUnlock = 150;
                                globalVarManager.USBCMD_DeviceModel = 94;
                                globalVarManager.USBCMD_WiFiSSID = 191;
                                globalVarManager.USBCMD_Reboot = 127;
                                globalVarManager.USBCMD_FWVersion = 97;
                                globalVarManager.USBCMD_CMDStringVersion = 1;
                                globalVarManager.USBCMD_Restore = 121;
                            }
                            ClsDiskInfoEx clsDiskInfoEx = new ClsDiskInfoEx();
                            clsDiskInfoEx.GetPhysicalDisks();
                            temp_drive_name = NumberFromExcelColumn(clsDiskInfoEx.GetDriveInfo(DeviceID).Substring(0, 1));
                            label5.Text = clsDiskInfoEx.GetDriveInfo(DeviceID).Substring(0, 1);
                            for (int index = 0; index < 4; index++)
                            {
                                if (true)//(temp_dev_location == globalVarManager.multi_savedDeviceList[index])
                                {
                                    scanportmap_insert[index] = 1; //Store which port is inserted
                                    scanportmap_insert_InstanceID[index] = temp_dev_InstanceID;
                                    drive_name[index] = temp_drive_name;
                                    process_status[index] = STATUS_READY;
                                }
                            }
                        }
                        catch (Exception)
                        {
                            label9.Text = "InitialState : PortScan Fail";
                        }
                    }
                }
            }
            
            Scan_btn.Enabled = true;
            Scan_btn.Text = "Re-Scan";
            Upgrade_btn.Enabled = true;
            Reset_btn.Enabled = true;

            /////////////////// /////////////////// /////////////////// /////////////////// ///////////////////

            if (false)//(version == 2)
            {
                //先解鎖測試CMD
                if (!USBCmdUnLocker(0))
                {
                    MessageBox.Show("傳送失敗 " + globalVarManager.USBCMD_QCUnlock);
                }
            }

            if (!GetSSID(0)) MessageBox.Show("get ssid fail");
            //if (!GetModel(0)) MessageBox.Show("get Model fail");
            if (!GetFWver(0)) MessageBox.Show("get FWVer fail");

            GC.Collect();
        }

        private void Reset_btn_Click(object sender, EventArgs e)
        {
            Scan_btn.Enabled = false;
            Upgrade_btn.Enabled = false;
            Reset_btn.Enabled = false;
            Reset_btn.Text = "waiting...";

            if (false)//(version == 2)
            {
                //先解鎖測試CMD
                if (!USBCmdUnLocker(0))
                {
                    //ShowInfo(0, dev, "傳送指令失敗: " + globalVarManager.USBCMD_QCUnlock, "");
                    MessageBox.Show("傳送失敗 " + globalVarManager.USBCMD_QCUnlock);
                    return;
                }
            }
            Reset(0);
            Console.WriteLine(4);

            Scan_btn.Enabled = true;
            Upgrade_btn.Enabled = true;
            Reset_btn.Enabled = true;
            Reset_btn.Text = "Reset";
            return;
        }

        private void Upgrade_btn_Click(object sender, EventArgs e)
        {
            //copy
            Scan_btn.Enabled = false;
            Upgrade_btn.Enabled = false;
            Reset_btn.Enabled = false;
            Upgrade_btn.Text = "waiting";
            label9.Text = "waiting";

            switch (_DUT)
            {
                case _Model.DPB60_5052:
                    //globalVarManager.FWBINFile = "Firmware\\DPB60_5052\\v1.1.0_H0324\\FWDB5052.bin";
                    globalVarManager.FWBINFile = "src\\946.bin";
                    break;
                case _Model.DPB60_5051:
                    //globalVarManager.FWBINFile = "Firmware\\DPB60_5051\\v1.1.0_H0324\\FWDB5051.bin";
                    globalVarManager.FWBINFile = "src\\945.bin";
                    break;
                case _Model.DPB60A:
                    //globalVarManager.FWBINFile = "Firmware\\DPB60_5000\\v1.3.0_H0324\\FWDPB60A.bin";
                    globalVarManager.FWBINFile = "src\\944.bin";
                    break;
                case _Model.DPB30A:
                    //globalVarManager.FWBINFile = "Firmware\\DPB30_5000\\v1.13.0_H0324\\FWDPB30.bin";
                    globalVarManager.FWBINFile = "src\\912.bin";
                    break;
                case _Model.DPB30_5031:
                    //globalVarManager.FWBINFile = "Firmware\\DPB30_5031\\v1.3.0_H0324\\FWDB5031.bin";
                    globalVarManager.FWBINFile = "src\\913.bin";
                    break;
                case _Model.DPB30_5032:
                    //globalVarManager.FWBINFile = "Firmware\\DPB30_5032\\v1.2.0_H0324\\FWDB5032.bin";
                    globalVarManager.FWBINFile = "src\\914.bin";
                    break;
                case _Model.DPB30_5033:
                    //globalVarManager.FWBINFile = "Firmware\\DPB30_5033\\v1.3.3_H0324\\FWDB5033.bin";
                    globalVarManager.FWBINFile = "src\\915.bin";
                    break;
            }
            //MessageBox.Show(globalVarManager.FWBINFile);
            if (!CopyTestFile(0,false, 2, globalVarManager.FWBINFile))
            {
                label9.Text = "copy fail";
            }

            //reboot
            //Reboot(0, false);



            //System.Threading.Thread.Sleep(5000); //Make sure device is left (start rebooting process)
            //if (!WaitDiskBack(0, false, false))
            //{
            //    MessageBox.Show("WaitDiskBack fail");
            //    //return false;
            //}

            //Reset(0);
            Scan_btn.Enabled = true;
            Upgrade_btn.Enabled = true;
            Reset_btn.Enabled = true;
            Upgrade_btn.Text = "Upgrade";
            label9.Text = "update done! Press Re-scan to check firmware version";
        }

    }
}
