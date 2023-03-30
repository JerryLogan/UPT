using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Management;
using System.Windows.Forms;
using Port_Setting;
using System.Threading;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic;
//using System.Data.SQLite;
using System.Data;

namespace DPB_Update_tool
{
    public static class globalVarManager
    {
        public static string CurrentFWVersion = "";
        public static int STM8Version = 0;
        public static string[] multi_savedDeviceList = new string[4];

        //[USBCount]
        public static string devicePathPortCountPWD = "";
        public static int devicePathPortCountType = 0;
        public static int devicePathPortCountMAX1 = 0;
        public static int devicePathPortCountMAX2 = 0;
        public static int[,] devicePathPortCount = new int[4, 2];
        public static string devicePathPortCountPath = "";

        //[WOInfo]
        public static string WO = "";
        public static string PN = "";
        public static string LENS = "";
        public static string TargetLDVersion = "";
        public static string TargetFWVersion = "";
        public static string ModelName = "";
        public static string[] USB_VEN = new string[1];
        public static string TestOrder = "";
        public static string SettingVersion = "";
        public static string ModelPID = "";
        public static int StorageCapacity_Min = 0;
        public static int StorageCapacity_Max = 0;

        //[FW]
        public static string LoaderBINFile = "";
        public static string FWBINFile = "";
        public static string STM8BINFile = "";
        public static string LD_MD5 = "";
        public static string FW_MD5 = "";
        public static string STM8_MD5 = "";
        public static int DefaultSTM8Version = -1;

        //[MPPara]
        public static bool ForceMP = false;
        public static string DevicePathPort = "";
        public static string LogPath = "";
        public static int LogExpiredDays = 0;
        public static int FWUpgtimeout = 180;
        public static int USBDetecttimeout = 0;
        public static int ProtectEnable = 0;
        public static byte[] CMD42password = { };
        public static byte[] QCPassWord = { };
        public static int BarMAXValue = 0;
        public static int BarMAXValue_stm8 = 0;
        public static string TestSuccessString = "";
        public static string TestFailString = "";
        public static string NoteString = "";
        public static string DeviceFileName = "";
        public static string BurninFilePath = "";
        public static string GPSFilePath = "";
        public static string IRFilePath = "";
        public static int USBDetectEvent_ms = 0;
        public static int ReScan_USBDetect_sec = 1;
        public static int UIDCheckTimeOut = 10; //TS_Jeffrey 20220707 Add new setting parameter (UICheckTimeOut)
        public static int FactoryFlag_MP = 0;
        public static string ADASFilePath = "";
        public static string DataBasePath = "";
        public static string ProduceDataBasePath = "";
        public static string DBtable = "";
        public static int[] ProduceData_issue = { 0, 0, 0, 0 };
        //[TestItem]
        public static bool ITEM_CheckStorageAtBeginning = false;
        public static bool ITEM_CheckLockStatus = false;
        public static bool ITEM_RestoreAtBeginning = false;
        public static bool ITEM_SyncTime = false;
        public static bool ITEM_CheckLoader = false;
        public static bool ITEM_UpgradeLoader = false;
        public static bool ITEM_UpgradeFW = false;
        public static bool ITEM_UpgradeSTM8 = false;
        public static bool ITEM_WriteWO = false;
        public static bool ITEM_WriteAPVersion = false;
        public static bool ITEM_WriteUID = false;
        public static bool ITEM_ShowResultLED = false;
        public static bool ITEM_SetFlag_MP = false;
        public static bool ITEM_SetFlag_QC = false;
        public static bool ITEM_SetFlag_WIFI = false;
        public static bool ITEM_SetFlag_PCCam = false;
        public static bool ITEM_SetFlag_SDCheck = false;
        public static bool ITEM_SetFlag_Burnin = false;
        public static bool ITEM_SetFlag_GPS = false;
        public static bool ITEM_SetFlag_ResetBeforeUpgrade = false;
        public static bool ITEM_SetFlag_ResetAfterUpgrade = false;
        public static string ITEM_Format = "";
        public static string ITEM_CopyBurninTest = "";
        public static string ITEM_CopyGPSTest = "";
        public static string ITEM_CopyIRTest = "";
        public static bool ITEM_PStoreFormat_BeforeUpgrade = false;
        public static bool ITEM_PStoreFormat_AfterUpgrade = false;
        public static bool ITEM_CheckWOInfo = false;
        public static bool ITEM_ADASLicense = false;
        public static bool ITEM_CheckStorageCapacity = false;
        public static bool ITEM_SetDualCam = false;
        public static bool ITEM_PRODUCEDATABASE = false;

        //[USBCmdSet]
        public static int USBCMD_Flow = 0;

        public static int USBCMD_SettingReset = -1;
        public static int USBCMD_Format = -1;
        public static int USBCMD_FWUpgrade = -1;
        public static int USBCMD_LDVersion = -1;
        public static int USBCMD_STM8Upgrade = -1;
        public static int USBCMD_STM8Version = -1;
        public static int USBCMD_DateTime = -1;
        public static int USBCMD_Date = -1;
        public static int USBCMD_Time = -1;
        public static int USBCMD_SetWO = -1;
        public static int USBCMD_SetLENS = -1;
        public static int USBCMD_LEDControl = -1;
        public static int USBCMD_LockStatus = -1;
        public static int USBCMD_UnLock = -1;
        public static int USBCMD_SetAversion = -1;
        public static int USBCMD_SetPversion = -1;
        public static int USBCMD_SetPassword = -1;
        public static int USBCMD_FilterStatus = -1;
        public static int USBCMD_FactoryFlag_MP = -1;
        public static int USBCMD_FactoryFlag_QC = -1;
        public static int USBCMD_FactoryFlag_WIFI = -1;
        public static int USBCMD_FactoryFlag_PCCam = -1;
        public static int USBCMD_FactoryFlag_SDCheck = -1;
        public static int USBCMD_FactoryFlag_Burnin = -1;
        public static int USBCMD_FactoryFlag_GPS = -1;
        public static int USBCMD_PStoreFormat = -1;
        public static int USBCMD_ADASLicense = -1;
        public static int USBCMD_ADASLicense_SerialNumber = -1;
        public static int USBCMD_WiFiTest = -1;
        public static int USBCMD_WirelessDualCam = -1;
        public static int USBCMD_SetUID = 162;
        public static int USBCMD_QCUnlock = 130;
        public static int USBCMD_DeviceModel = 14;
        public static int USBCMD_WiFiSSID = 191;
        public static int USBCMD_Reboot = 77;
        public static int USBCMD_FWVersion = 17;
        public static int USBCMD_CMDStringVersion = 2;
        public static int USBCMD_Restore = 72;
        //

    }

    public partial class Form1
    {
        public void LoadSetting()
        {
            //60A
            globalVarManager.USB_VEN[0] = "DPB";
            globalVarManager.ModelPID = "5000";
            globalVarManager.QCPassWord = Encoding.ASCII.GetBytes(" ");

            
            //5060
            //Firmware\DPB60_5000\v1.3.0_H0320\FWDPB60A.bin
        //globalVarManager.USB_VEN[0] = "DPB30";
        //globalVarManager.ModelPID = "5060";
        }

        public void LoadUSBCountSetting()
        {

        }

        public void SaveUSBCountSetting(int type, int dev)
        {

        }

    }
}