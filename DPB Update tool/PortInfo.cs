//PortInfo v1.2.1 (Justin)
using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Management;
using System.IO;
using System.Threading;


namespace Port_Setting
{
    class PortInfo
    {
        #region my_lib Dll setting
        [DllImport("my_lib_x64.dll", EntryPoint = "Port_Reordering_By_Enum", SetLastError = true)]
        static extern IntPtr Port_Reordering_By_Enum_X64(string emun, string info, bool ReorderByPort);

        [DllImport("my_lib.dll", EntryPoint = "Port_Reordering_By_Enum", SetLastError = true, CallingConvention = CallingConvention.Cdecl)]
        static extern IntPtr Port_Reordering_By_Enum(string emun, string info, bool ReorderByPort);

        [DllImport("my_lib_x64.dll", EntryPoint = "Port_Reordering_By_Id", SetLastError = true)]
        static extern IntPtr Port_Reordering_By_Id_X64(string id, string device_relation, string info, bool ReorderByPort, int default_port);

        [DllImport("my_lib.dll", EntryPoint = "Port_Reordering_By_Id", SetLastError = true, CallingConvention = CallingConvention.Cdecl)]
        static extern IntPtr Port_Reordering_By_Id(string id, string device_relation, string info, bool ReorderByPort, int default_port);

        [DllImport("my_lib_x64.dll", EntryPoint = "EnumUSBHostControllers", SetLastError = true)]
        static extern IntPtr EnumUSBHostControllers_X64(string info);

        [DllImport("my_lib.dll", EntryPoint = "EnumUSBHostControllers", SetLastError = true, CallingConvention = CallingConvention.Cdecl)]
        static extern IntPtr EnumUSBHostControllers(string info);

        [DllImport("my_lib_x64.dll", EntryPoint = "EnumUSBHostControllersGUID", SetLastError = true)]
        static extern IntPtr EnumUSBHostControllersGUID_X64(string info);

        [DllImport("my_lib.dll", EntryPoint = "EnumUSBHostControllersGUID", SetLastError = true, CallingConvention = CallingConvention.Cdecl)]
        static extern IntPtr EnumUSBHostControllersGUID(string info);

        [DllImport("my_lib_x64.dll", EntryPoint = "EnumUSBRootHubName", SetLastError = true)]
        static extern IntPtr EnumUSBRootHubName_X64(string HCDName);

        [DllImport("my_lib.dll", EntryPoint = "EnumUSBRootHubName", SetLastError = true, CallingConvention = CallingConvention.Cdecl)]
        static extern IntPtr EnumUSBRootHubName(string HCDName);

        [DllImport("my_lib_x64.dll", EntryPoint = "EnumUSBDevices", SetLastError = true)]
        static extern IntPtr EnumUSBDevices_X64(string HubName, string info);

        [DllImport("my_lib.dll", EntryPoint = "EnumUSBDevices", SetLastError = true, CallingConvention = CallingConvention.Cdecl)]
        static extern IntPtr EnumUSBDevices(string HubName, string info);

        [DllImport("my_lib_x64.dll", EntryPoint = "GetUSBHubMaxPorts", SetLastError = true)]
        static extern int GetUSBHubMaxPorts_X64(string HubName);

        [DllImport("my_lib.dll", EntryPoint = "GetUSBHubMaxPorts", SetLastError = true, CallingConvention = CallingConvention.Cdecl)]
        static extern int GetUSBHubMaxPorts(string HubName);

        [DllImport("my_lib_x64.dll", EntryPoint = "Get_Disk_Size", SetLastError = true)]
        static extern Int64 Get_Disk_Size_X64(string devicename);

        [DllImport("my_lib.dll", EntryPoint = "Get_Disk_Size", SetLastError = true, CallingConvention = CallingConvention.Cdecl)]
        static extern Int64 Get_Disk_Size(string devicename);

        [DllImport("my_lib.dll", EntryPoint = "BufferCompare", SetLastError = true, CallingConvention = CallingConvention.Cdecl)]
        static extern int BufferCompare(IntPtr buff1, IntPtr buff2, int pattern_offset, int size);

        [DllImport("my_lib_x64.dll", EntryPoint = "BufferCompare", SetLastError = true)]
        static extern int BufferCompare_X64(IntPtr buff1, IntPtr buff2, int pattern_offset, int size);

        [DllImport("my_lib.dll", EntryPoint = "FreePointer", SetLastError = true, CallingConvention = CallingConvention.Cdecl)]
        static extern void FreePointer(IntPtr pt);

        [DllImport("my_lib_x64.dll", EntryPoint = "FreePointer", SetLastError = true)]
        static extern void FreePointer_X64(IntPtr pt);

        /*
         * The functions are only used with PDC3 since the drive returns special location information
         */
        [DllImport("my_lib_PDC3_x64.dll", EntryPoint = "Port_Reordering_By_Id", SetLastError = true)]
        static extern IntPtr Port_Reordering_By_Id_PDC3_X64(string id, string device_relation, string info);

        [DllImport("my_lib_PDC3.dll", EntryPoint = "Port_Reordering_By_Id", SetLastError = true)]
        static extern IntPtr Port_Reordering_By_Id_PDC3(string id, string device_relation, string info);
        #endregion

        #region Const parameter setting
        const int max_device = 32;

        private string[] device_port_information = new string[16];
        private string[] dev_path = new string[max_device];
        private string[] dev_instance = new string[max_device];
        private string[] dev_controller = new string[max_device];
        private string[] dev_controllerName = new string[max_device];
        private string[] dev_location = new string[max_device];
        private string[] dev_type = new string[max_device];
        private string[] dev_InstanceID = new string[max_device];
        private string[] dev_PN = new string[max_device];
        private string[] INI_controller = new string[max_device];
        private string[] INI_location = new string[max_device];
        #endregion

        public static string Get_USBPort(string device_instance_ID, ref string deviceVIDPIDRaw)
        {
            //if (device_SetPortBySN)
            //    return device_instance_ID;
            IntPtr ptr = IntPtr.Zero;
            IntPtr ptr2 = IntPtr.Zero;
            IntPtr[] ptr_array = new IntPtr[max_device];
            IntPtr[] ptr_array2 = new IntPtr[max_device];
            string[] dev_desp_Controller = new string[max_device];
            string[] dev_desp_RootHub = new string[max_device];
            string[] dev_path_Controller = new string[max_device];
            string[] dev_path_RootHub_Classes = new string[max_device];
            string[] dev_path_RootHub = new string[max_device];
            string[] dev_path_RootHub_Instance_ID = new string[max_device];
            string[] dev_path2 = new string[max_device];
            string[] dev_pathTemp = new string[max_device];

            int DeviceSaveNumber = 0;
            //string enum_name = "";
            //enum_name = "P\0C\0I\0";
            if (IntPtr.Size == 8)
            {
                ptr = EnumUSBHostControllersGUID_X64("Device_Description");
            }
            else
            {
                ptr = EnumUSBHostControllersGUID("Device_Description");
            }
            for (int i = 0; i < max_device; i++)
            {
                ptr_array[i] = Marshal.ReadIntPtr(ptr, i * IntPtr.Size);
                dev_desp_Controller[i] = Marshal.PtrToStringAnsi(ptr_array[i]);
                if (IntPtr.Size == 8) { FreePointer_X64(ptr_array[i]); }
                else { FreePointer(ptr_array[i]); }
            }
            if (IntPtr.Size == 8) { FreePointer_X64(ptr); }
            else { FreePointer(ptr); }


            if (IntPtr.Size == 8) { ptr = EnumUSBHostControllersGUID_X64("PathName"); }
            else { ptr = EnumUSBHostControllersGUID("PathName"); }
            for (int i = 0; i < max_device; i++)
            {
                ptr_array[i] = Marshal.ReadIntPtr(ptr, i * IntPtr.Size);
                dev_path_Controller[i] = Marshal.PtrToStringAnsi(ptr_array[i]);
                if (IntPtr.Size == 8) { FreePointer_X64(ptr_array[i]); }
                else { FreePointer(ptr_array[i]); }
            }
            if (IntPtr.Size == 8) { FreePointer_X64(ptr); }
            else { FreePointer(ptr); }

            for (int i = 0; i < max_device; i++)
            {
                if (IntPtr.Size == 8) { ptr = EnumUSBRootHubName_X64(dev_path_Controller[i]); }
                else { ptr = EnumUSBRootHubName(dev_path_Controller[i]); }
                if (ptr != IntPtr.Zero)
                {
                    dev_path_RootHub_Classes[i] = Marshal.PtrToStringAnsi(ptr);
                    if (IntPtr.Size == 8) { FreePointer_X64(ptr); }
                    else { FreePointer(ptr); }
                }
            }
            if (IntPtr.Size == 8) { FreePointer_X64(ptr); }
            else { FreePointer(ptr); }

            for (int i = 0; i < max_device; i++)
            {
                if (dev_desp_Controller[i] == "")
                    continue;

                string ChildDveiceLoation = findChildDevice(dev_path_RootHub_Classes[i], DeviceSaveNumber, device_instance_ID, ref deviceVIDPIDRaw);
                if (ChildDveiceLoation != "")
                {
                    return (i+1).ToString() + "_" + ChildDveiceLoation;
                }
            }
            return "Not Found Disk.";
        }

        #region private functions
        private static string findChildDeviceEnhanced(string instance_ID, int DeviceSaveNumber, string device_instance_ID)
        {
            IntPtr ptr = IntPtr.Zero;
            IntPtr[] ptr_array = new IntPtr[max_device];
            int[] usb_class = new int[max_device];
            string childInstanceID;

            if (IntPtr.Size == 8) { ptr = Port_Reordering_By_Id_X64(instance_ID, "Child_Devices", "Instance_ID", true, -1); }
            else { ptr = Port_Reordering_By_Id(instance_ID, "Child_Devices", "Instance_ID", true, -1); }
            for (int i = 0; ptr != IntPtr.Zero && i < max_device; i++)
            {
                ptr_array[i] = Marshal.ReadIntPtr(ptr, i * IntPtr.Size);
            }
            if (IntPtr.Size == 8) { FreePointer_X64(ptr); }
            else { FreePointer(ptr); }

            for (int i = 0; i < max_device; i++)
            {
                childInstanceID = Marshal.PtrToStringAnsi(ptr_array[i]);
                if (childInstanceID != null && childInstanceID != "")
                {
                    // Fucking Rex_Chu
                    //for (int ii = 0; ii < childInstanceID.Length - device_instance_ID.Length + 1; ii++)
                    //{
                    //    if (childInstanceID.Substring(ii, device_instance_ID.Length).CompareTo(device_instance_ID) == 0)
                    //    {
                    //        DeviceSaveNumber++;
                    //        return i.ToString();
                    //    }
                    //}

                    // Smart Justin
                    string childInstanceIDSN = "";
                    int index = childInstanceID.LastIndexOf("\\");
                    if (index > 0)
                        childInstanceIDSN = childInstanceID.Substring(index + 1, childInstanceID.Length - index - 1);

                    if (childInstanceIDSN.Contains(device_instance_ID) && (childInstanceIDSN.Length - device_instance_ID.Length) <= 3)
                    {
                        DeviceSaveNumber++;
                        return i.ToString();
                    }
                }
            }
            return "";
        }

        private static string findChildDevice(string instance_ID, int DeviceSaveNumber, string device_instance_ID, ref string deviceVIDPIDRaw)
        {
            IntPtr ptr = IntPtr.Zero;
            IntPtr[] ptr_array = new IntPtr[max_device];
            int[] usb_class = new int[max_device];
            string childInstanceID, childPathName;

            if (IntPtr.Size == 8) { ptr = EnumUSBDevices_X64(instance_ID, "DeviceClass"); }
            else { ptr = EnumUSBDevices(instance_ID, "DeviceClass"); }
            for (int i = 0; ptr != IntPtr.Zero && i < max_device; i++)
            {
                ptr_array[i] = Marshal.ReadIntPtr(ptr, i * IntPtr.Size);
                usb_class[i] = (int)Marshal.ReadByte(ptr_array[i]);
            }
            if (IntPtr.Size == 8) { FreePointer_X64(ptr); }
            else { FreePointer(ptr); }

            if (IntPtr.Size == 8) { ptr = EnumUSBDevices_X64(instance_ID, "PathName"); }
            else { ptr = EnumUSBDevices(instance_ID, "PathName"); }
            for (int i = 0; ptr != IntPtr.Zero && i < max_device; i++)
            {
                if (usb_class[i] == 9)
                    ptr_array[i] = Marshal.ReadIntPtr(ptr, i * IntPtr.Size);
            }
            if (IntPtr.Size == 8) { FreePointer_X64(ptr); }
            else { FreePointer(ptr); }

            if (IntPtr.Size == 8) { ptr = EnumUSBDevices_X64(instance_ID, "Instance_ID"); }
            else { ptr = EnumUSBDevices(instance_ID, "Instance_ID"); }
            for (int i = 0; ptr != IntPtr.Zero && i < max_device; i++)
            {
                if (usb_class[i] != 9)
                    ptr_array[i] = Marshal.ReadIntPtr(ptr, i * IntPtr.Size);
            }
            if (IntPtr.Size == 8) { FreePointer_X64(ptr); }
            else { FreePointer(ptr); }

            for (int i = 0; ptr != IntPtr.Zero && i < max_device; i++)
            {
                if (usb_class[i] == 9)
                {
                    childPathName = Marshal.PtrToStringAnsi(ptr_array[i]);

                  

                    string ChildDveiceLoation = findChildDevice(childPathName, DeviceSaveNumber, device_instance_ID, ref deviceVIDPIDRaw);


                    if (ChildDveiceLoation != "")
                        return i.ToString() + ":" + ChildDveiceLoation;
                }
                else
                {
                    childInstanceID = Marshal.PtrToStringAnsi(ptr_array[i]);
                    if (childInstanceID != "")
                    {
                        // Fucking Rex_Chu
                        //for (int ii = 0; ii < childInstanceID.Length - device_instance_ID.Length + 1; ii++)
                        //{
                        //    if (childInstanceID.Substring(ii, device_instance_ID.Length).CompareTo(device_instance_ID) == 0)
                        //    {
                        //        DeviceSaveNumber++;
                        //        return i.ToString();
                        //    }
                        //}

                        if (childInstanceID.Contains("VID") && childInstanceID.Contains("PID"))
                            deviceVIDPIDRaw = childInstanceID;

                        // Smart Justin
                        string childInstanceIDSN = "";
                        int index = childInstanceID.LastIndexOf("\\");
                        if (index > 0)
                            childInstanceIDSN = childInstanceID.Substring(index + 1, childInstanceID.Length - index - 1);

                        if (childInstanceIDSN.Contains(device_instance_ID) && (childInstanceIDSN.Length - device_instance_ID.Length) <= 3)
                        {
                            DeviceSaveNumber++;
                            return i.ToString();
                        }

                        string DeviceEnhancedLocation = findChildDeviceEnhanced(childInstanceID, DeviceSaveNumber, device_instance_ID);
                        if (DeviceEnhancedLocation != "")
                            return DeviceEnhancedLocation;
                    }
                }
            }
            return "";
        }

        private bool CheckIfSystemVolume(string DeviceID)
        {
            ManagementObjectSearcher DiskDriveToDiskPartition_Searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_DiskDriveToDiskPartition");
            ManagementObjectSearcher LogicalDiskToPartition_Searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_LogicalDiskToPartition");
            ManagementObjectCollection DiskDrive_Collection = DiskDriveToDiskPartition_Searcher.Get();
            ManagementObjectCollection LogicalDisk_Collection = LogicalDiskToPartition_Searcher.Get();
            foreach (ManagementObject wmi_DriveToVolume in DiskDrive_Collection)
            {
                string targetPhysicalDrive = wmi_DriveToVolume["Antecedent"].ToString();
                string targetVolume = wmi_DriveToVolume["Dependent"].ToString();
                int subStringIndex = targetPhysicalDrive.IndexOf("PHYSICALDRIVE");
                if (DeviceID.Substring(4) == targetPhysicalDrive.Substring(subStringIndex, targetPhysicalDrive.Length - subStringIndex - 1))
                {
                    foreach (ManagementObject wmi_Partition in LogicalDisk_Collection)
                    {
                        if (targetVolume == wmi_Partition["Antecedent"].ToString())
                        {
                            subStringIndex = wmi_Partition["Dependent"].ToString().IndexOf("DeviceID=");
                            string PartitionNumberID = wmi_Partition["Dependent"].ToString().Substring(subStringIndex + 10, wmi_Partition["Dependent"].ToString().Length - subStringIndex - 11);
                            if (PartitionNumberID == "C:")
                                return false;
                        }
                    }
                }
            }
            return true;
        }

        private string GetControllerName(string DeviceController, bool isUSBDevice)
        {
            int subStringIndex = DeviceController.IndexOf("DeviceID=");
            string DeviceControllerName = DeviceController.Substring(subStringIndex + 10, DeviceController.Length - subStringIndex - 11);
            DeviceControllerName = DeviceControllerName.Replace(@"\\", @"\");
            ManagementObjectSearcher USBController_Searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_USBController");
            ManagementObjectCollection USBController_Collection = USBController_Searcher.Get();
            string tempDeviceController;
            if (isUSBDevice)
            {
                foreach (ManagementObject wmi_ControllerDevice in USBController_Collection)
                {
                    tempDeviceController = wmi_ControllerDevice["DeviceID"].ToString();
                    if (DeviceControllerName.CompareTo(tempDeviceController) == 0)
                    {
                        DeviceControllerName = wmi_ControllerDevice["Name"].ToString();
                        break;
                    }
                }
            }
            else
            {
                ManagementObjectSearcher SCSIControllere_Searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_SCSIController");
                ManagementObjectSearcher IDEControllere_Searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_IDEController");
                ManagementObjectCollection SCSIControllere_Collection = SCSIControllere_Searcher.Get();
                ManagementObjectCollection IDEControllere_Collection = IDEControllere_Searcher.Get();
                foreach (ManagementObject wmi_ControllerDevice in SCSIControllere_Collection)  //找SCSIControolerDevice 的上下層
                {
                    tempDeviceController = wmi_ControllerDevice["DeviceID"].ToString();
                    if (DeviceControllerName.CompareTo(tempDeviceController) == 0)
                    {
                        DeviceControllerName = wmi_ControllerDevice["Name"].ToString();
                        break;
                    }
                }
                foreach (ManagementObject wmi_ControllerDevice in IDEControllere_Collection)   //找IDEControolerDevice 的上下層
                {
                    tempDeviceController = wmi_ControllerDevice["DeviceID"].ToString();
                    if (DeviceControllerName.CompareTo(tempDeviceController) == 0)
                    {
                        DeviceControllerName = wmi_ControllerDevice["Name"].ToString();
                        break;
                    }
                }
            }
            return DeviceControllerName;
        }

        private string GetController(string PNPDeviceID, bool isUSBDevice)
        {
            string DeviceController = "";
            PNPDeviceID = "\\" + PNPDeviceID;
            if (isUSBDevice)
            {
                ManagementObjectSearcher USBControllerdevice_Searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_USBControllerDevice");
                ManagementObjectCollection USBControllerDevice_Collection = USBControllerdevice_Searcher.Get();
                string tempDeviceController;
                foreach (ManagementObject wmi_ControllerDevice in USBControllerDevice_Collection)
                {
                    tempDeviceController = wmi_ControllerDevice["Dependent"].ToString();
                    if (PNPDeviceID.Length < tempDeviceController.Length)
                    {
                        for (int i = 0; i < (tempDeviceController.Length - PNPDeviceID.Length); i++)
                        {
                            if (tempDeviceController.Substring(i, PNPDeviceID.Length) == PNPDeviceID)
                            {
                                DeviceController = wmi_ControllerDevice["Antecedent"].ToString();
                                break;
                            }
                        }
                    }
                }
            }
            else
            {
                ManagementObjectSearcher SCSIControllere_Searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_SCSIControllerDevice");
                ManagementObjectSearcher IDEControllere_Searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_IDEControllerDevice");
                ManagementObjectCollection SCSIControllere_Collection = SCSIControllere_Searcher.Get();
                ManagementObjectCollection IDEControllere_Collection = IDEControllere_Searcher.Get();
                string tempDeviceController;
                foreach (ManagementObject wmi_ControllerDevice in SCSIControllere_Collection)  //找SCSIControolerDevice 的上下層
                {
                    tempDeviceController = wmi_ControllerDevice["Dependent"].ToString();
                    if (PNPDeviceID.Length < tempDeviceController.Length)
                    {
                        for (int i = 0; i < (tempDeviceController.Length - PNPDeviceID.Length); i++)
                        {
                            if (tempDeviceController.Substring(i, PNPDeviceID.Length) == PNPDeviceID)
                            {
                                DeviceController = wmi_ControllerDevice["Antecedent"].ToString();
                                break;
                            }
                        }
                    }
                }
                foreach (ManagementObject wmi_ControllerDevice in IDEControllere_Collection)   //找IDEControolerDevice 的上下層
                {
                    tempDeviceController = wmi_ControllerDevice["Dependent"].ToString();
                    if (PNPDeviceID.Length < tempDeviceController.Length)
                    {
                        for (int i = 0; i < (tempDeviceController.Length - PNPDeviceID.Length); i++)
                        {
                            if (tempDeviceController.Substring(i, PNPDeviceID.Length) == PNPDeviceID) //如果不符合不會更改
                            {
                                DeviceController = wmi_ControllerDevice["Antecedent"].ToString();
                                break;
                            }
                        }
                    }
                }
            }

            return DeviceController;
        }
        #endregion

        #region public functions
        public int GetPortNumber(string PhysicalDeviceNumber, ref string dev_desp, ref string deviceVIDPIDRaw)
        {
            ManagementObjectSearcher DiskDrive_Searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_DiskDrive");
            ManagementObjectSearcher USBHub_Searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT *  FROM Win32_USBHub");
            ManagementObjectCollection DiskDrive_Collection = DiskDrive_Searcher.Get();
            ManagementObjectCollection USBHub_Collection = USBHub_Searcher.Get();
            int deviceCount = 31;
            foreach (ManagementObject wmi_Device in DiskDrive_Collection)
            {
                if (PhysicalDeviceNumber.ToUpper() != wmi_Device["DeviceID"].ToString().ToUpper())
                    continue;
                dev_instance[deviceCount] = wmi_Device["PNPDeviceID"].ToString();
                dev_type[deviceCount] = wmi_Device["InterfaceType"].ToString();
                if (wmi_Device["InterfaceType"].ToString() == "USB")
                {
                    int subStringIndex = dev_instance[deviceCount].IndexOf("\\", 10);
                    dev_InstanceID[deviceCount] = dev_instance[deviceCount].Substring(subStringIndex + 1, dev_instance[deviceCount].Length - subStringIndex - 1);
                    dev_InstanceID[deviceCount] = dev_InstanceID[deviceCount].Substring(0, dev_InstanceID[deviceCount].Length - 2);
                    dev_controller[deviceCount] = GetController(dev_InstanceID[deviceCount], true);
                    dev_controllerName[deviceCount] = GetControllerName(dev_controller[deviceCount], true);

                    dev_location[deviceCount] = Get_USBPort(dev_InstanceID[deviceCount], ref deviceVIDPIDRaw);
                    dev_location[deviceCount] = dev_location[deviceCount];
                    dev_desp = dev_controllerName[deviceCount] + " & " + dev_controller[deviceCount].Substring(dev_controller[deviceCount].IndexOf(".DeviceID=") + 11, dev_controller[deviceCount].Length - dev_controller[deviceCount].IndexOf(".DeviceID=") - 12) + " & " + dev_location[deviceCount];
                }
                else
                {
                    int subStringIndex = dev_instance[deviceCount].IndexOf("\\", 10);
                    dev_InstanceID[deviceCount] = dev_instance[deviceCount].Substring(subStringIndex + 1, dev_instance[deviceCount].Length - subStringIndex - 1);
                    dev_location[deviceCount] = wmi_Device["SCSIBus"].ToString() + "." + wmi_Device["SCSILogicalUnit"].ToString() + "." + wmi_Device["SCSITargetId"].ToString();
                    dev_controller[deviceCount] = GetController(dev_InstanceID[deviceCount], false);
                    dev_controllerName[deviceCount] = GetControllerName(dev_controller[deviceCount], false);
                    subStringIndex = dev_controller[deviceCount].IndexOf("DeviceID=");
                    dev_desp = dev_controllerName[deviceCount] + " & " + dev_controller[deviceCount].Substring(dev_controller[deviceCount].IndexOf(".DeviceID=") + 11, dev_controller[deviceCount].Length - dev_controller[deviceCount].IndexOf(".DeviceID=") - 12) + " & " + dev_location[deviceCount];
                }
                for (int i = 0; i < device_port_information.Length; i++)
                {
                    if (dev_desp == device_port_information[i])
                        return i;
                    }
                return -1;
            }
            return -1;
        }

        public int GetPortNumber(string PhysicalDeviceNumber, ref string deviceVIDPIDRaw)
        {
            string dev_desp = "";
            ManagementObjectSearcher DiskDrive_Searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_DiskDrive");
            ManagementObjectSearcher USBHub_Searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT *  FROM Win32_USBHub");
            ManagementObjectCollection DiskDrive_Collection = DiskDrive_Searcher.Get();
            ManagementObjectCollection USBHub_Collection = USBHub_Searcher.Get();
            int deviceCount = 31;
            foreach (ManagementObject wmi_Device in DiskDrive_Collection)
            {
                if (PhysicalDeviceNumber.ToUpper() != wmi_Device["DeviceID"].ToString().ToUpper())
                    continue;
                dev_instance[deviceCount] = wmi_Device["PNPDeviceID"].ToString();
                dev_type[deviceCount] = wmi_Device["InterfaceType"].ToString();
                if (wmi_Device["InterfaceType"].ToString() == "USB")
                {
                    int subStringIndex = dev_instance[deviceCount].IndexOf("\\", 10);
                    dev_InstanceID[deviceCount] = dev_instance[deviceCount].Substring(subStringIndex + 1, dev_instance[deviceCount].Length - subStringIndex - 1);
                    dev_InstanceID[deviceCount] = dev_InstanceID[deviceCount].Substring(0, dev_InstanceID[deviceCount].Length - 2);
                    dev_controller[deviceCount] = GetController(dev_InstanceID[deviceCount], true);
                    dev_controllerName[deviceCount] = GetControllerName(dev_controller[deviceCount], true);

                    dev_location[deviceCount] = Get_USBPort(dev_InstanceID[deviceCount],  ref deviceVIDPIDRaw);
                    dev_location[deviceCount] = dev_location[deviceCount];
                    dev_desp = dev_controllerName[deviceCount] + " & " + dev_controller[deviceCount].Substring(dev_controller[deviceCount].IndexOf(".DeviceID=") + 11, dev_controller[deviceCount].Length - dev_controller[deviceCount].IndexOf(".DeviceID=") - 12) + " & " + dev_location[deviceCount];
                }
                else
                {
                    int subStringIndex = dev_instance[deviceCount].IndexOf("\\", 10);
                    dev_InstanceID[deviceCount] = dev_instance[deviceCount].Substring(subStringIndex + 1, dev_instance[deviceCount].Length - subStringIndex - 1);
                    dev_location[deviceCount] = wmi_Device["SCSIBus"].ToString() + "." + wmi_Device["SCSILogicalUnit"].ToString() + "." + wmi_Device["SCSITargetId"].ToString();
                    dev_controller[deviceCount] = GetController(dev_InstanceID[deviceCount], false);
                    dev_controllerName[deviceCount] = GetControllerName(dev_controller[deviceCount], false);
                    subStringIndex = dev_controller[deviceCount].IndexOf("DeviceID=");
                    dev_desp = dev_controllerName[deviceCount] + " & " + dev_controller[deviceCount].Substring(dev_controller[deviceCount].IndexOf(".DeviceID=") + 11, dev_controller[deviceCount].Length - dev_controller[deviceCount].IndexOf(".DeviceID=") - 12) + " & " + dev_location[deviceCount];
                }
                for (int i = 0; i < device_port_information.Length; i++)
                {
                    if (dev_desp == device_port_information[i])
                        return i;
                }
                return -1;
            }
            return -1;
        }
        
        public bool isPortValid(int portNumber, ref string deviceVIDPIDRaw)
        {
            if (device_port_information.Length < portNumber + 1)
                return false;
            string dev_desp = "";
            ManagementObjectSearcher DiskDrive_Searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_DiskDrive");
            ManagementObjectSearcher USBHub_Searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT *  FROM Win32_USBHub");
            ManagementObjectCollection DiskDrive_Collection = DiskDrive_Searcher.Get();
            ManagementObjectCollection USBHub_Collection = USBHub_Searcher.Get();
            int deviceCount = 31;
            foreach (ManagementObject wmi_Device in DiskDrive_Collection)
            {
                dev_instance[deviceCount] = wmi_Device["PNPDeviceID"].ToString();
                dev_type[deviceCount] = wmi_Device["InterfaceType"].ToString();
                if (wmi_Device["InterfaceType"].ToString() == "USB")
                {
                    int subStringIndex = dev_instance[deviceCount].IndexOf("\\", 10);
                    dev_InstanceID[deviceCount] = dev_instance[deviceCount].Substring(subStringIndex + 1, dev_instance[deviceCount].Length - subStringIndex - 1);
                    dev_InstanceID[deviceCount] = dev_InstanceID[deviceCount].Substring(0, dev_InstanceID[deviceCount].Length - 2);
                    dev_controller[deviceCount] = GetController(dev_InstanceID[deviceCount], true);
                    dev_controllerName[deviceCount] = GetControllerName(dev_controller[deviceCount], true);

                    dev_location[deviceCount] = Get_USBPort(dev_InstanceID[deviceCount], ref deviceVIDPIDRaw);
                    dev_location[deviceCount] = dev_location[deviceCount];
                    dev_desp = dev_controllerName[deviceCount] + " & " + dev_controller[deviceCount].Substring(dev_controller[deviceCount].IndexOf(".DeviceID=") + 11, dev_controller[deviceCount].Length - dev_controller[deviceCount].IndexOf(".DeviceID=") - 12) + " & " + dev_location[deviceCount];
                }
                else
                {
                    int subStringIndex = dev_instance[deviceCount].IndexOf("\\", 10);
                    dev_InstanceID[deviceCount] = dev_instance[deviceCount].Substring(subStringIndex + 1, dev_instance[deviceCount].Length - subStringIndex - 1);
                    dev_location[deviceCount] = wmi_Device["SCSIBus"].ToString() + "." + wmi_Device["SCSILogicalUnit"].ToString() + "." + wmi_Device["SCSITargetId"].ToString();
                    dev_controller[deviceCount] = GetController(dev_InstanceID[deviceCount], false);
                    dev_controllerName[deviceCount] = GetControllerName(dev_controller[deviceCount], false);
                    subStringIndex = dev_controller[deviceCount].IndexOf("DeviceID=");
                    dev_desp = dev_controllerName[deviceCount] + " & " + dev_controller[deviceCount].Substring(dev_controller[deviceCount].IndexOf(".DeviceID=") + 11, dev_controller[deviceCount].Length - dev_controller[deviceCount].IndexOf(".DeviceID=") - 12) + " & " + dev_location[deviceCount];
                }
                if (dev_desp == device_port_information[portNumber])
                {
                    return IsDeviceReady(wmi_Device["DeviceID"].ToString()); 
                }
            }
            return false;
        }

        public void SetDeviceSequence(string[] DeviceSequence)
        {
            Array.Resize(ref device_port_information, 16);
            for (int i = 0; i < DeviceSequence.Length; i++)
            {
                device_port_information[i] = DeviceSequence[i];
            }
            Array.Resize(ref device_port_information, DeviceSequence.Length);
        }

        public string GetLocation(string PNPDeviceID)
        {
            string DiskLocation = "";
            string KeyDevicePath = "SYSTEM\\CurrentControlSet\\Enum\\" + PNPDeviceID;
            Microsoft.Win32.RegistryKey start = Microsoft.Win32.Registry.LocalMachine;
            Microsoft.Win32.RegistryKey programName = start.OpenSubKey(KeyDevicePath);
            if (programName != null)
            {
                DiskLocation = (string)programName.GetValue("LocationInformation");
            }
            else
            {
                DiskLocation = "123456789";
            }
            return DiskLocation;
        }

        public void GetDeviceList(ref string[] dev_desp, ref string deviceVIDPIDRaw)
        {
            Array.Resize(ref dev_desp, 32);
            ManagementObjectSearcher DiskDrive_Searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_DiskDrive");
            ManagementObjectSearcher USBHub_Searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT *  FROM Win32_USBHub");
            ManagementObjectCollection DiskDrive_Collection = DiskDrive_Searcher.Get();
            ManagementObjectCollection USBHub_Collection = USBHub_Searcher.Get();
            int deviceCount = 0;
            foreach (ManagementObject wmi_Device in DiskDrive_Collection)
            {
                if (CheckIfSystemVolume(wmi_Device["DeviceID"].ToString()) == false)
                {
                    continue;
                }
                dev_path[deviceCount] = wmi_Device["DeviceID"].ToString();
                dev_instance[deviceCount] = wmi_Device["PNPDeviceID"].ToString();
                dev_type[deviceCount] = wmi_Device["InterfaceType"].ToString();
                if (wmi_Device["InterfaceType"].ToString() == "USB")
                {
                    int subStringIndex = dev_instance[deviceCount].IndexOf("\\", 10);
                    dev_InstanceID[deviceCount] = dev_instance[deviceCount].Substring(subStringIndex + 1, dev_instance[deviceCount].Length - subStringIndex - 1);
                    dev_InstanceID[deviceCount] = dev_InstanceID[deviceCount].Substring(0, dev_InstanceID[deviceCount].Length - 2);
                    dev_controller[deviceCount] = GetController(dev_InstanceID[deviceCount], true);
                    dev_controllerName[deviceCount] = GetControllerName(dev_controller[deviceCount], true);
                    dev_location[deviceCount] = Get_USBPort(dev_InstanceID[deviceCount], ref deviceVIDPIDRaw);
                    dev_desp[deviceCount] = dev_location[deviceCount] + " & " + dev_controllerName[deviceCount];
                }
                else
                {
                    int subStringIndex = dev_instance[deviceCount].IndexOf("\\", 10);
                    dev_InstanceID[deviceCount] = dev_instance[deviceCount].Substring(subStringIndex + 1, dev_instance[deviceCount].Length - subStringIndex - 1);
                    dev_location[deviceCount] = GetLocation(dev_instance[deviceCount]);
                    dev_controller[deviceCount] = GetController(dev_InstanceID[deviceCount], false);
                    dev_controllerName[deviceCount] = GetControllerName(dev_controller[deviceCount], false);
                    subStringIndex = dev_controller[deviceCount].IndexOf("DeviceID=");
                    dev_desp[deviceCount] = dev_location[deviceCount] + " & " + dev_controllerName[deviceCount];
                }
                deviceCount++;
            }
            Array.Resize(ref dev_desp, deviceCount);
        }

        private bool IsDeviceReady(string deviceID)
        {
            ManagementObjectSearcher DiskDrive_Searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_DiskDrive");
            ManagementObjectCollection DiskDrive_Collection = DiskDrive_Searcher.Get();
            foreach (ManagementObject wmi_Device in DiskDrive_Collection)
            {
                if (wmi_Device["DeviceID"].ToString() == deviceID)
                {
                    if (wmi_Device["TotalSectors"] != null)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            return false;
        }
        #endregion
    }
}
