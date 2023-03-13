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
    public partial class Form1 : Form
    {
        public Thread[] MPTest_Thread;

        static int[] drive_name = new int[4];
        static int[] process_status = new int[4];
        static int[] scanportmap_insert = new int[4];
        static string[] scanportmap_insert_InstanceID = new string[4];

        static int MPcount = 1;

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

        private void scan_Click(object sender, EventArgs e)
        {
            //程序開始，先鎖上Button
            Scan.Enabled = false;
            Scan.Text = "waiting...";
            label1.Text = "";
            label6.Text = "";
            label7.Text = "";
            label9.Text = "";
            //原本有定port資訊不能執行定port程序

            //if (globalVarManager.multi_savedDeviceList[0] != "")
            //{
            //    MessageBox.Show("已存在先前Set Port資訊, 若需重新Set Port, 請先Clear Port");
            //    //程序結束，解開Button
            //    Scan.Enabled = true;
            //    Scan.Text = "Set Port";
            //    return;
            //}

            //清空多port多執行序要用的array
            //for (int index = 0; index < 4; index++)
            //{
            //    drive_name[index] = 0;
            //    process_status[index] = STATUS_NODEVICE;
            //    scanportmap_insert[index] = 0;
            //}

            //定port數歸零
            MPcount = 0;

            //掃USB port，把有接Device的port記下來，並寫到DPBPathPort.ini中
            int port_num = 0;

            //Ini.RWIniFile DPBPathPortINI = new Ini.RWIniFile();

            DriveInfo[] allDrives = DriveInfo.GetDrives();


            foreach (DriveInfo d in allDrives)
            {

                if (d.IsReady == true && d.DriveType.ToString()=="Removable")
                {
                    Console.WriteLine("Drive {0}", d.Name);
                    Console.WriteLine("Drive2 {0}", d.RootDirectory);
                    Console.WriteLine("  Drive type: {0}", d.DriveType);
                    Console.WriteLine("  Volume label: {0}", d.VolumeLabel);
                    Console.WriteLine("  File system: {0}", d.DriveFormat);
                    Console.WriteLine(
                        "  Available space to current user:{0, 15} bytes",
                        d.AvailableFreeSpace);

                    Console.WriteLine(
                        "  Total available space:          {0, 15} bytes",
                        d.TotalFreeSpace);

                    Console.WriteLine(
                        "  Total size of drive:            {0, 15} bytes ",
                        d.TotalSize);
                }
            }


            string[] saveDeviceListTmp = new string[4]; //For Sort Port
            string[] saveDeviceListTmpBefore = new string[4];
            int[] drive_nameTmp = new int[4];
            string[] scanportmap_insert_InstanceIDTmp = new string[4];

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

                            //Support PID Identify
                            if (globalVarManager.ModelPID != PID_EMPTY)
                                if (!deviceVIDPIDRaw.Contains(globalVarManager.ModelPID)) continue;

                            ClsDiskInfoEx clsDiskInfoEx = new ClsDiskInfoEx();
                            clsDiskInfoEx.GetPhysicalDisks();
                            int temp_drive_name = NumberFromExcelColumn(clsDiskInfoEx.GetDriveInfo(DeviceID).Substring(0, 1));

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
                            //LogManager.PrintLog("InitialState : PortScan Fail");
                        }
                    }
                }
            }
            

            Scan.Enabled = true;
            Scan.Text = "Scan";

            /////////////////// /////////////////// /////////////////// /////////////////// ///////////////////
            string DeviceUID = "";
            int setCount = 0;
            byte[] UID_Read = new byte[256];
            while (!NvUSBcmd.GetNOVAData((byte)globalVarManager.USBCMD_SetUID, ref UID_Read, drive_name[0]))
            {
                //MessageBox.Show("1");
                setCount++;
                //LogManager.PrintLog(dev, "Fail to send cmd (Get)" + globalVarManager.USBCMD_SetUID + " setCount:" + setCount + " drive_name: " + drive_name[dev] + " (" + dev + ")");
                System.Threading.Thread.Sleep(1000);
                if (setCount > 5)
                {
                    MessageBox.Show("傳送失敗 "+ globalVarManager.USBCMD_SetUID);
                    GC.Collect();
                    break;
                }
            }
            DeviceUID = ASCIIEncoding.ASCII.GetString(UID_Read, 2, 10 * (UID_Read[0] - 0x30) + (UID_Read[1] - 0x30));
            label1.Text = DeviceUID;


            /////////////////// /////////////////// /////////////////// /////////////////// ///////////////////
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
                    MessageBox.Show("傳送失敗 "+ globalVarManager.USBCMD_FWVersion);
                    break;
                }
            }
            DeviceFwVer = UTF8Encoding.UTF8.GetString(btFwVer).Trim('\0');
            label7.Text = DeviceFwVer;


            /////////////////// /////////////////// /////////////////// /////////////////// ///////////////////
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
                    break;
                }
            }
            DeviceModelname = System.Text.Encoding.UTF8.GetString(device_model_name,0,6);
            label6.Text = DeviceModelname;


            GC.Collect();
        }
    }
}
