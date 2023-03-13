using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management;
using System.Windows.Forms;
using Port_Setting;
using System.Threading;
using System.Runtime.InteropServices;
using System.IO;

namespace DPB_Update_tool
{
    public partial class Form1 : Form
    {
        //============== 等待裝置回PC =================//
        private bool WaitForDisk(int dev, bool LockStatus, bool showProgressBar)
        {
            if (WaitDiskBack(dev, (!LockStatus), showProgressBar))
            {
                GC.Collect();
                return true;
            }
            GC.Collect();
            return false;
        }

        private bool WaitForDiskJustCheckFlag(int dev, bool LockStatus, bool showProgressBar, int DelayTime_sec)
        {
            // Jerry add for reducing the CPU usage 20180802
            Thread.Sleep(500);  // wait for USB remove first, avoiding error drive_name[dev]
            DateTime TimeStart = DateTime.Now;
            DateTime TimeCnt = DateTime.Now;
            int TDiff = 0;
            bool CheckIsBack = false;
            int ReScanCount = 0;

            while (!CheckIsBack && TDiff < globalVarManager.FWUpgtimeout)
            {
                if (scanportmap_insert[dev] == 1 && TDiff > DelayTime_sec)   //device insert
                {
                    DriveInfo[] allDrives = DriveInfo.GetDrives();  // Check all DiskDrive
                    foreach (DriveInfo d in allDrives)
                    {
                        if (NumberFromExcelColumn(d.Name) == drive_name[dev] && drive_name[dev] != 0)
                        {
                            if (!LockStatus)
                            {
                                if (d.IsReady)  // Check storage ready, just like check size before
                                {
                                    if (NvUSBcmd.NvUSB_ConnectToDevice(drive_name[dev]))
                                    {
                                        CheckIsBack = true;
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                CheckIsBack = true;
                                break;
                            }
                        }
                    }
                }
                Thread.Sleep(200);
                TimeCnt = DateTime.Now;
                TDiff = DateDiff(TimeStart, TimeCnt);
                if ((TDiff / globalVarManager.ReScan_USBDetect_sec) == (ReScanCount + 1) && TDiff != 0)
                {
                    ReScanCount++;
                    if (TDiff > DelayTime_sec)
                    {
                        // Rescan the device on PC
                        ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_DiskDrive");

                        foreach (ManagementObject WMIObject in searcher.Get())
                        {
                            try
                            {
                                string DeviceID = WMIObject["DeviceID"].ToString();
                                string PNPDeviceID = WMIObject["PNPDeviceID"].ToString();
                                foreach (string USB_VEN_Member in globalVarManager.USB_VEN)
                                {
                                    int sIndex = PNPDeviceID.IndexOf(USB_VEN_Member, 0);
                                    string deviceVIDPIDRaw = "";

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

                                            if (temp_dev_location == globalVarManager.multi_savedDeviceList[dev])
                                            {
                                                scanportmap_insert[dev] = 1;
                                                scanportmap_insert_InstanceID[dev] = temp_dev_InstanceID;
                                                if (!LockStatus || (LockStatus && WMIObject["Size"] != null && WMIObject["Size"].ToString() != ""))
                                                {
                                                    ClsDiskInfoEx clsDiskInfoEx = new ClsDiskInfoEx();
                                                    clsDiskInfoEx.GetPhysicalDisks();
                                                    string VolumeName = clsDiskInfoEx.GetDriveInfo(DeviceID).Substring(0, 1);
                                                    drive_name[dev] = NumberFromExcelColumn(VolumeName);

                                                    GC.Collect();
                                                }
                                            }
                                        }
                                        catch { }
                                    }
                                }
                            }
                            catch{}
                        }
                    }
                }
                //if (showProgressBar)
                    //ShowUpgradeTime(TDiff, dev);
            }
            //stm8UPBar[dev] = false;
            //ShowUpgradeUI(false, dev);
            if (!CheckIsBack && TDiff >= globalVarManager.FWUpgtimeout)
            {
                //LogManager.PrintLog(dev, "[WaitDiskBack Fail] The process is time out" + " (" + dev + ")");
                GC.Collect();
                return false;
            }
            GC.Collect();
            return true;

        }

        private bool WaitDiskBack(int dev, bool checksize, bool showProgressBar)
        {
            Thread.Sleep(500);
            //LogManager.PrintLog(dev, "waiting the process..." + " (" + dev + ")");
            //ShowUpgradeUI(showProgressBar, dev);

            DateTime TimeStart = DateTime.Now;
            DateTime TimeCnt = DateTime.Now;
            int TDiff = 0;
            string waitdisk = "";

            while (waitdisk == "" && TDiff < globalVarManager.FWUpgtimeout)
            {
                waitdisk = GetDirectory(dev, checksize);
                TimeCnt = DateTime.Now;
                TDiff = DateDiff(TimeStart, TimeCnt);
                //if (showProgressBar)
                    //ShowUpgradeTime(TDiff, dev);
            }
            //stm8UPBar[dev] = false;
            //ShowUpgradeUI(false, dev);
            if (waitdisk == "" && TDiff >= globalVarManager.FWUpgtimeout)
            {
                //LogManager.PrintLog(dev, "[WaitDiskBack Fail] The process is time out" + " (" + dev + ")");
                return false;
            }

            GC.Collect();
            return true;
        }

        //============== 等待裝置離開PC ================//
        private bool WaitDiskLeave(int dev, bool checksize)
        {
            DateTime TimeStart = DateTime.Now;
            DateTime TimeCnt = DateTime.Now;
            int TDiff = 0;
            string waitdisk = GetDirectory(dev, checksize);

            while (waitdisk != "" && TDiff < globalVarManager.FWUpgtimeout)
            {
                waitdisk = GetDirectory(dev, checksize);
                TimeCnt = DateTime.Now;
                TDiff = DateDiff(TimeStart, TimeCnt);
            }
            if (waitdisk != "" && TDiff >= globalVarManager.FWUpgtimeout)
            {
                //LogManager.PrintLog(dev, "[WaitDiskLeave Fail] The process is time out" + " (" + dev + ")");
                return false;
            }

            GC.Collect();
            return true;
        }

        //============= 取得裝置USB訊息 ================//
        private string GetDirectory(int dev, bool checksize)
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_DiskDrive");

            foreach (ManagementObject WMIObject in searcher.Get())
            {
                try
                {
                    string DeviceID = WMIObject["DeviceID"].ToString();
                    string PNPDeviceID = WMIObject["PNPDeviceID"].ToString();
                    foreach (string USB_VEN_Member in globalVarManager.USB_VEN)
                    {
                        int sIndex = PNPDeviceID.IndexOf(USB_VEN_Member, 0);
                        string deviceVIDPIDRaw = "";

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
                                {
                                    //LogManager.PrintLog(dev, "GetDirectory() temp_dev_InstanceID " + temp_dev_InstanceID + " (" + dev + ")");
                                    //LogManager.PrintLog(dev, "GetDirectory() deviceVIDPIDRaw " + deviceVIDPIDRaw + " (" + dev + ")");
                                    //LogManager.PrintLog(dev, "GetDirectory() temp_dev_location " + temp_dev_location + " (" + dev + ")");
                                    if (!deviceVIDPIDRaw.Contains(globalVarManager.ModelPID))
                                    {
                                        //LogManager.PrintLog(dev, "GetDirectory() deviceVIDPIDRaw error " + deviceVIDPIDRaw + " (" + dev + ")");
                                        //return "";
                                        continue;
                                    }
                                }

                                if (temp_dev_location == globalVarManager.multi_savedDeviceList[dev])
                                {
                                    if (!checksize || (checksize && WMIObject["Size"] != null && WMIObject["Size"].ToString() != ""))
                                    {
                                        ClsDiskInfoEx clsDiskInfoEx = new ClsDiskInfoEx();
                                        clsDiskInfoEx.GetPhysicalDisks();
                                        string VolumeName = clsDiskInfoEx.GetDriveInfo(DeviceID).Substring(0, 1);
                                        drive_name[dev] = NumberFromExcelColumn(VolumeName);

                                        GC.Collect();
                                        return VolumeName;
                                    }
                                }
                            }
                            catch (Exception)
                            {
                                //LogManager.PrintLog(dev, "GetDirectory() Error" + " (" + dev + ")");
                                return "";
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    //LogManager.PrintLog(dev, "GetDirectory() Error!!" + " (" + dev + ")");
                    return "";
                }
            }
            //LogManager.PrintLog(dev, "GetDirectory() Disk Not found " + " (" + dev + ")");
            return "";
        }

        //=============== USB Handler ===============//
        private void USBEventHandler(Object sender, EventArrivedEventArgs e)
        {
            bool USB_VEN_Check = false;
            PropertyData devicePropert = e.NewEvent.Properties["TargetInstance"];
            if (devicePropert != null)
            {
                ManagementBaseObject mbo = devicePropert.Value as ManagementBaseObject;
                string DeviceID = mbo.Properties["DeviceID"].Value.ToString();
                string PNPDeviceID = mbo.Properties["PNPDeviceID"].Value.ToString();
                foreach (string USB_VEN_Member in globalVarManager.USB_VEN)
                {
                    if (PNPDeviceID.IndexOf(USB_VEN_Member) > 0)
                    {
                        USB_VEN_Check = true;
                        break;
                    }
                    else
                        USB_VEN_Check = false;
                }
                if (USB_VEN_Check == false)
                    return;
                int subStringIndex = PNPDeviceID.IndexOf("\\", 10);
                string temp_dev_InstanceID = PNPDeviceID.Substring(subStringIndex + 1, PNPDeviceID.Length - subStringIndex - 1);
                temp_dev_InstanceID = temp_dev_InstanceID.Substring(0, temp_dev_InstanceID.Length - 2);

                if (e.NewEvent.ClassPath.ClassName == "__InstanceCreationEvent")
                {
                    try
                    {
                        string deviceVIDPIDRaw = "";
                        string temp_dev_location = PortInfo.Get_USBPort(temp_dev_InstanceID, ref deviceVIDPIDRaw);

						//Support PID Identify
                        if (globalVarManager.ModelPID != PID_EMPTY)
                            if (!deviceVIDPIDRaw.Contains(globalVarManager.ModelPID)) return;


                        ClsDiskInfoEx clsDiskInfoEx = new ClsDiskInfoEx();
                        clsDiskInfoEx.GetPhysicalDisks();
                        //int temp_drive_name = NumberFromExcelColumn(clsDiskInfoEx.GetDriveInfo(DeviceID).Substring(0, 1));
                        string temp_drive_name_str = "";
                        int temp_drive_name;
                        int retry = 0;

                        while (temp_drive_name_str == "" && retry < globalVarManager.USBDetecttimeout)
                        {
                            clsDiskInfoEx.GetPhysicalDisks();
                            temp_drive_name_str = clsDiskInfoEx.GetDriveInfo(DeviceID);
                            retry++;
                            System.Threading.Thread.Sleep(1000);
                        }
                        temp_drive_name = NumberFromExcelColumn(temp_drive_name_str);

                        for (int index = 0; index < MPcount; index++)
                        {
                            if (temp_dev_location == globalVarManager.multi_savedDeviceList[index])
                            {
                                scanportmap_insert[index] = 1;  //Store which port is inserted
                                scanportmap_insert_InstanceID[index] = temp_dev_InstanceID;
                                drive_name[index] = temp_drive_name;
                                //LogManager.PrintLog(index, "[USBEventHandler] CHK Insert, drive_name:" + drive_name[index] + "(" + index + ")");

                                if (process_status[index] == STATUS_NODEVICE)   // New device insert, update UI and start thread
                                {
                                    //LogManager.PrintLog(index, "[USBEventHandler] CHK Insert, waiting for test. drive_name:" + drive_name[index] + "(" + index + ")\r\n");
                                    //ShowInfo(0, index, "等待測試", "");
                                    drive_name[index] = temp_drive_name;
                                    process_status[index] = STATUS_READY;

                                    if (globalVarManager.multi_savedDeviceList[index] != "" && drive_name[index] != 0 && process_status[index] == STATUS_READY )
                                    {
                                        process_status[index] = STATUS_PROCESSING;
                                        //MPTest_Thread[index] = new Thread(new ParameterizedThreadStart(preWork_MP));
                                        MPTest_Thread[index].IsBackground = true;
                                        MPTest_Thread[index].Start(index);
                                        //ShowInfo(2, index, "processing", "");
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception)
                    {
                        //LogManager.PrintLog("USBEventHandler Exception : plug event");
                    }
                }
                else if (e.NewEvent.ClassPath.ClassName == "__InstanceDeletionEvent")
                {
                    for (int index = 0; index < MPcount; index++)
                    {
                        if (string.Compare(temp_dev_InstanceID, scanportmap_insert_InstanceID[index]) == 0)
                        {
                            scanportmap_insert[index] = 0;
                            scanportmap_insert_InstanceID[index] = "";
                            drive_name[index] = 0;

                            //正在測試的不可刪
                            if (MPTest_Thread[index] == null || (MPTest_Thread[index] != null && !MPTest_Thread[index].IsAlive) )
                            {
                                if (process_status[index] == STATUS_FINISH || (process_status[index] == STATUS_READY ))   // MP Fininsh, update UI and parameter setting
                                {
                                    //LogManager.PrintLog(index, "CHK removed : (" + index + ")\r\n");
                                    //ShowInfo(0, index, "請插入 " + globalVarManager.ModelName, "");
                                    //ShowInfo(1, index, "", "ld");
                                    //ShowInfo(1, index, "", "fw");
                                    //ShowInfo(1, index, "", "wo");
                                    //ShowInfo(1, index, "", "lens");
                                    if (globalVarManager.DefaultSTM8Version != -1)
                                    {
                                        //ShowInfo(1, index, "", "stm8");
                                    }
                                    //ShowInfo(1, index, "", "uid");
                                    //ShowInfo(2, index, "ready", "");
                                    drive_name[index] = 0;
                                    process_status[index] = STATUS_NODEVICE;
                                    scanportmap_insert[index] = 0;
                                    globalVarManager.ProduceData_issue[index] = 0;//clear ProduceData_issue flag when unplug

                                    if (MPTest_Thread[index] != null)
                                    {
                                        MPTest_Thread[index].Abort();
                                        MPTest_Thread[index] = null;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            GC.Collect();
        }

        private static int NumberFromExcelColumn(string column)
        {
            string col = column.ToUpper();
            char colPiece = col[0];
            int colNum = colPiece - 64;

            GC.Collect();
            return colNum;
        }

        private int DateDiff(DateTime DateTime1, DateTime DateTime2)
        {
            int dateDiff = 0;
            TimeSpan ts1 = new TimeSpan(DateTime1.Ticks);
            TimeSpan ts2 = new TimeSpan(DateTime2.Ticks);
            TimeSpan ts = ts1.Subtract(ts2).Duration();
            dateDiff = ts.Hours * 3600 + ts.Minutes * 60 + ts.Seconds;

            GC.Collect();
            return dateDiff;
        }
    }

    class NvUSBcmd
    {
        [DllImport("NvUSB.dll", CallingConvention = CallingConvention.Cdecl)]
        public static extern int NvUSB_GetFirstAvailableDevice();

        [DllImport("NvUSB.dll", CallingConvention = CallingConvention.Cdecl)]
        public static extern bool NvUSB_ConnectToDevice(int iDevice);

        [DllImport("NvUSB.dll", CallingConvention = CallingConvention.Cdecl)]
        public static extern bool NvUSB_ConnectIsAvailable(int iDevice);

        [DllImport("NvUSB.dll", CallingConvention = CallingConvention.Cdecl)]
        public static extern bool NvUSB_VenderCmd_GetData(int iDevice, byte iCmd, IntPtr pData, uint nBytes);

        [DllImport("NvUSB.dll", CallingConvention = CallingConvention.Cdecl)]
        public static extern bool NvUSB_VenderCmd_SetData(int iDevice, byte iCmd, IntPtr pData, uint nBytes);

        [DllImport("NvUSB.dll", CallingConvention = CallingConvention.Cdecl)]
        public static extern unsafe bool NvUSB_MemoryWrite(int iDevice, uint dwAddr, uint nBytes, void* pBuf);

        [DllImport("NvUSB.dll", CallingConvention = CallingConvention.Cdecl)]
        public static extern unsafe bool NvUSB_MemoryRead(int iDevice, uint dwAddr, uint nBytes, void* pBuf);

        [DllImport("NvUSB.dll", CallingConvention = CallingConvention.Cdecl)]
        public static extern unsafe bool NvUSB_GetDeviceCount();

        [DllImport("NvUSB.dll", CallingConvention = CallingConvention.Cdecl)]
        public static extern unsafe bool NvUSB_DisconnectFormDevice(int hDevice);

        //Transfer A, B, C to 1, 2, 3
        public static int NumberFromExcelColumn(string column)
        {
            string col = column.ToUpper();
            char colPiece = col[0];
            int colNum = colPiece - 64;
            return colNum;
        }

        public static bool SetNOVAState(byte state, int[] data, int DiskMark)
        {
            IntPtr hFile = Marshal.AllocHGlobal(data.Length * sizeof(int));
            Marshal.Copy(data, 0, hFile, data.Length);
            bool connection_status = NvUSBcmd.NvUSB_ConnectIsAvailable(DiskMark);
            if (connection_status)
            {
                if (NvUSBcmd.NvUSB_VenderCmd_SetData(DiskMark, state, hFile, (uint)data.Length * sizeof(int)))
                {
                    Marshal.FreeHGlobal(hFile);
                    return true;
                }
            }
            else if (NvUSBcmd.NvUSB_ConnectToDevice(DiskMark))
            {
                if (NvUSBcmd.NvUSB_VenderCmd_SetData(DiskMark, state, hFile, (uint)data.Length * sizeof(int)))
                {
                    Marshal.FreeHGlobal(hFile);
                    return true;
                }
            }
            Marshal.FreeHGlobal(hFile);
            return false;
        }

        public static bool SetNOVAData(byte state, byte[] data, int DiskMark)
        {
            IntPtr hFile = Marshal.AllocHGlobal(data.Length);
            Marshal.Copy(data, 0, hFile, data.Length);
            bool connection_status = NvUSBcmd.NvUSB_ConnectIsAvailable(DiskMark);
            if (connection_status)
            {
                if (NvUSBcmd.NvUSB_VenderCmd_SetData(DiskMark, state, hFile, (uint)data.Length * sizeof(byte)))
                {
                    Marshal.FreeHGlobal(hFile);
                    return true;
                }
            }
            else if (NvUSBcmd.NvUSB_ConnectToDevice(DiskMark))
            {
                if (NvUSBcmd.NvUSB_VenderCmd_SetData(DiskMark, state, hFile, (uint)data.Length * sizeof(byte)))
                {
                    Marshal.FreeHGlobal(hFile);
                    return true;
                }
            }
            Marshal.FreeHGlobal(hFile);
            return false;
        }

        public static bool GetNOVAState(byte state, ref int[] data, int DiskMark)
        {
            IntPtr hFile = Marshal.AllocHGlobal(data.Length * sizeof(int));
            bool connection_status = NvUSBcmd.NvUSB_ConnectIsAvailable(DiskMark);
            if (connection_status)
            {
                if (!NvUSBcmd.NvUSB_VenderCmd_GetData(DiskMark, state, hFile, sizeof(int)))
                    return false;
            }
            else if (NvUSBcmd.NvUSB_ConnectToDevice(DiskMark))
            {
                if (!NvUSBcmd.NvUSB_VenderCmd_GetData(DiskMark, state, hFile, sizeof(int)))
                    return false;
            }
            else
            {
                return false;
            }

            Marshal.Copy(hFile, data, 0, 1);
            Marshal.FreeHGlobal(hFile);
            return true;
        }

        public static bool GetNOVAData(byte state, ref byte[] data, int DiskMark)
        {
            IntPtr hFile = Marshal.AllocHGlobal(30);
            bool connection_status = NvUSBcmd.NvUSB_ConnectIsAvailable(DiskMark);
            if (connection_status)
            {
                if (!NvUSBcmd.NvUSB_VenderCmd_GetData(DiskMark, state, hFile, 30))
                    return false;
            }
            else if (NvUSBcmd.NvUSB_ConnectToDevice(DiskMark))
            {
                if (!NvUSBcmd.NvUSB_VenderCmd_GetData(DiskMark, state, hFile, 30))
                    return false;
            }
            else
            {
                return false;
            }

            Marshal.Copy(hFile, data, 0, 30);
            Marshal.FreeHGlobal(hFile);
            return true;
        }

        public static bool GetNOVAData_ADASLicense(byte state, ref byte[] data, int DiskMark)
        {
            IntPtr hFile = Marshal.AllocHGlobal(64);
            bool connection_status = NvUSBcmd.NvUSB_ConnectIsAvailable(DiskMark);
            if (connection_status)
            {
                if (!NvUSBcmd.NvUSB_VenderCmd_GetData(DiskMark, state, hFile, 64))
                    return false;
            }
            else if (NvUSBcmd.NvUSB_ConnectToDevice(DiskMark))
            {
                if (!NvUSBcmd.NvUSB_VenderCmd_GetData(DiskMark, state, hFile, 64))
                    return false;
            }
            else
            {
                return false;
            }

            Marshal.Copy(hFile, data, 0, 64);
            Marshal.FreeHGlobal(hFile);
            return true;
        }

        public static string DataToString(byte[] data)
        {
            try
            {
                string dataString = "";

                if (globalVarManager.USBCMD_CMDStringVersion == 1)
                {
                    ASCIIEncoding ascii = new ASCIIEncoding();
                    dataString = ascii.GetString(data, 1, data[0]);
                }
                else if (globalVarManager.USBCMD_CMDStringVersion == 2)
                {
                    //新的String Version不再傳長度 
                    dataString = Encoding.UTF8.GetString(data).Trim(new Char[] { '\0' });
                }

                return dataString;
            }
            catch
            {
                return "";
            }
        }
    }
}
