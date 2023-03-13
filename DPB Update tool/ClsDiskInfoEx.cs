using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32.SafeHandles;
using System.Runtime.InteropServices;

namespace DPB_Update_tool
{
    class ClsDiskInfoEx
    {
        private const int FileShareRead = 1;
        private const int Filesharewrite = 2;
        private const int OpenExisting = 3;
        private const int IoctlVolumeGetVolumeDiskExtents = 0x560000;
        private const int IncorrectFunction = 1;
        private const int ErrorInsufficientBuffer = 122;
        private const int MoreDataIsAvailable = 234;
        private static List<string> currentDriveMappings = new List<string>();

        [StructLayout(LayoutKind.Sequential)]
        private struct DiskExtent
        {
            public int DiskNumber;
            public long StartingOffset;
            public long ExtentLength;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct DiskExtents
        {
            public int numberOfExtents;
            // We can't marhsal an array if we don't know its size.
            public DiskExtent first;
        }
        public enum RESOURCE_SCOPE
        {
            RESOURCE_CONNECTED = 0x1,
            RESOURCE_GLOBALNET = 0x2,
            RESOURCE_REMEMBERED = 0x3,
            RESOURCE_RECENT = 0x4,
            RESOURCE_CONTEXT = 0x5
        }
        public enum RESOURCE_TYPE
        {
            RESOURCETYPE_ANY = 0x0,
            RESOURCETYPE_DISK = 0x1,
            RESOURCETYPE_PRINT = 0x2,
            RESOURCETYPE_RESERVED = 0x8
        }
        public enum RESOURCE_USAGE
        {
            RESOURCEUSAGE_CONNECTABLE = 0x1,
            RESOURCEUSAGE_CONTAINER = 0x2,
            RESOURCEUSAGE_NOLOCALDEVICE = 0x4,
            RESOURCEUSAGE_SIBLING = 0x8,
            RESOURCEUSAGE_ATTACHED = 0x10,
            RESOURCEUSAGE_ALL = (RESOURCEUSAGE_CONNECTABLE | RESOURCEUSAGE_CONTAINER | RESOURCEUSAGE_ATTACHED)
        }
        public enum RESOURCE_DISPLAYTYPE
        {
            RESOURCEDISPLAYTYPE_GENERIC = 0x0,
            RESOURCEDISPLAYTYPE_DOMAIN = 0x1,
            RESOURCEDISPLAYTYPE_SERVER = 0x2,
            RESOURCEDISPLAYTYPE_SHARE = 0x3,
            RESOURCEDISPLAYTYPE_FILE = 0x4,
            RESOURCEDISPLAYTYPE_GROUP = 0x5,
            RESOURCEDISPLAYTYPE_NETWORK = 0x6,
            RESOURCEDISPLAYTYPE_ROOT = 0x7,
            RESOURCEDISPLAYTYPE_SHAREADMIN = 0x8,
            RESOURCEDISPLAYTYPE_DIRECTORY = 0x9,
            RESOURCEDISPLAYTYPE_TREE = 0xa,
            RESOURCEDISPLAYTYPE_NDSCONTAINER = 0xb
        }
        public struct NETRESOURCE
        {
            public RESOURCE_SCOPE dwScope;
            public RESOURCE_TYPE dwType;
            public RESOURCE_DISPLAYTYPE dwDisplayType;
            public RESOURCE_USAGE dwUsage;
            [MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPTStr)]
            public string lpLocalName;
            [MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPTStr)]
            public string lpRemoteName;
            [MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPTStr)]
            public string lpComment;
            [MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPTStr)]
            public string lpProvider;
        }
        private class NativeMethods
        {
            [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
            public static extern SafeFileHandle CreateFile(string fileName, int desiredAccess, int shareMode, IntPtr securityAttributes, int creationDisposition, int flagsAndAttributes, IntPtr hTemplateFile);

            [DllImport("kernel32", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool DeviceIoControl(SafeFileHandle hVol, int controlCode, IntPtr inBuffer, int inBufferSize, ref DiskExtents outBuffer, int outBufferSize, ref int bytesReturned, IntPtr overlapped);

            [DllImport("kernel32", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool DeviceIoControl(SafeFileHandle hVol, int controlCode, IntPtr inBuffer, int inBufferSize, IntPtr outBuffer, int outBufferSize, ref int bytesReturned, IntPtr overlapped);

            [DllImport("mpr.dll", CharSet = CharSet.Auto)]
            public static extern int WNetEnumResource(IntPtr hEnum, ref int lpcCount, IntPtr lpBuffer, ref int lpBufferSize);

            [DllImport("mpr.dll", CharSet = CharSet.Auto)]
            public static extern int WNetOpenEnum(RESOURCE_SCOPE dwScope, RESOURCE_TYPE dwType, RESOURCE_USAGE dwUsage, ref NETRESOURCE lpNetResource, ref IntPtr lphEnum);

            [DllImport("mpr.dll", CharSet = CharSet.Auto)]
            public static extern int WNetCloseEnum(IntPtr hEnum);
        }

        private List<string> GetPhysicalDriveStrings(System.IO.DriveInfo driveInfo)
        {
            SafeFileHandle sfh = null;
            List<string> physicalDrives = new List<string>(1);
            //char[] charsToTrim = {',', '.', '\\'};
            string path = "\\\\.\\" + driveInfo.RootDirectory.ToString().TrimEnd('\\');
            try
            {
                sfh = NativeMethods.CreateFile(path, 0, FileShareRead | Filesharewrite, IntPtr.Zero, OpenExisting, 0, IntPtr.Zero);
                int bytesReturned = 0;
                DiskExtents de1;
                de1.numberOfExtents = 0;
                de1.first.DiskNumber = 0;
                de1.first.StartingOffset = 0;
                de1.first.ExtentLength = 0;
                int numDiskExtents = 0;
                bool result = NativeMethods.DeviceIoControl(sfh, IoctlVolumeGetVolumeDiskExtents, IntPtr.Zero, 0, ref de1, Marshal.SizeOf(de1), ref bytesReturned, IntPtr.Zero);
                if (result == true)
                {
                    // there was only one disk extent. So the volume lies on 1 physical drive.
                    physicalDrives.Add("\\\\.\\PhysicalDrive" + de1.first.DiskNumber.ToString());
                    return physicalDrives;
                }
                if (Marshal.GetLastWin32Error() == IncorrectFunction)
                {
                    // The drive is removable and removed, like a CDRom with nothing in it.
                    return physicalDrives;
                }
                if (Marshal.GetLastWin32Error() == MoreDataIsAvailable)
                {
                    // This drive is part of a mirror or volume - handle it below. 
                }
                else if (Marshal.GetLastWin32Error() != ErrorInsufficientBuffer)
                {
                    throw new System.ComponentModel.Win32Exception();
                }
                // Houston, we have a spanner. The volume is on multiple disks.
                // Untested...
                // We need a blob of memory for the DISK_EXTENTS structure, and all the DISK_EXTENTS
                int blobSize = Marshal.SizeOf(typeof(DiskExtents)) + (de1.numberOfExtents - 1) * Marshal.SizeOf(typeof(DiskExtent));
                IntPtr pBlob = Marshal.AllocHGlobal(blobSize);
                result = NativeMethods.DeviceIoControl(sfh, IoctlVolumeGetVolumeDiskExtents, IntPtr.Zero, 0, pBlob, blobSize, ref bytesReturned, IntPtr.Zero);
                if (result == false)
                    throw new System.ComponentModel.Win32Exception();
                // Read them out one at a time.
                IntPtr pNext = new IntPtr(pBlob.ToInt64() + 8);
                // is this always ok on 64 bit OSes? ToInt64?
                for (int i = 0; i <= de1.numberOfExtents - 1; i++)
                {
                    DiskExtent diskExtentN = (DiskExtent)Marshal.PtrToStructure(pNext, typeof(DiskExtent));
                    physicalDrives.Add("\\\\.\\PhysicalDrive" + diskExtentN.DiskNumber.ToString());
                    pNext = new IntPtr(pNext.ToInt32() + Marshal.SizeOf(typeof(DiskExtent)));
                }
                return physicalDrives;
            }
            finally
            {
                if (sfh != null)
                {
                    if (sfh.IsInvalid == false)
                    {
                        sfh.Close();
                    }
                    sfh.Dispose();
                }
            }
        }

        public void GetPhysicalDisks()
        {
            currentDriveMappings.Clear();
            List<string> drivesList = new List<string>();
            List<string> tmpList = new List<string>();
            StringBuilder drives = new StringBuilder();

            foreach (System.IO.DriveInfo logicalDrive in System.IO.DriveInfo.GetDrives())
            {
                try
                {
                    drives.Remove(0, drives.Length);
                    drives.Append(logicalDrive.RootDirectory.ToString());
                    drives.Append("=");

                    if (logicalDrive.DriveType == System.IO.DriveType.Network)
                    {
                        continue;
                    }
                    else if (logicalDrive.DriveType == System.IO.DriveType.CDRom)
                    {
                        continue;
                    }
                    else
                    {
                        drivesList = GetPhysicalDriveStrings(logicalDrive);

                        if (drivesList.Count > 0)
                        {
                            foreach (string drive in drivesList)
                            {
                                string temp = drive;
                                //temp = temp.Replace("\\\\.\\", "");
                                //temp = temp.Replace("PhysicalDrive", "Physical Drive ");
                                drives.Append(temp);
                                drives.Append(", ");
                            }
                            drives.Remove(drives.Length - 2, 2);
                        }
                        else
                        {
                            drives.Append("n/a");
                        }
                    }
                    currentDriveMappings.Add(drives.ToString());
                }
                catch (Exception)
                {
                    //LogManager.PrintLog("GetPhysicalDisks Exception");
                }
            }
        }

        private string[] GetPhysicalDiskParentFor(string logicalDisk)
        {

            string[] parts = null;

            if (logicalDisk.Length > 0)
            {
                foreach (string driveMapping in currentDriveMappings)
                {
                    if (logicalDisk.Substring(0, 2).ToUpper() == driveMapping.Substring(0, 2).ToUpper())
                    {
                        parts = driveMapping.Split('=');
                        return parts;
                    }
                }
            }

            return null;
        }

        public string GetDriveInfo(string locPhysicalDrive)
        {
            foreach (System.IO.DriveInfo driveInfo in System.IO.DriveInfo.GetDrives())
            {
                string[] parentDrives = GetPhysicalDiskParentFor(driveInfo.RootDirectory.ToString());
                if (parentDrives != null)
                {
                    if (parentDrives[1].ToUpper() == locPhysicalDrive.ToUpper())
                    {
                        return parentDrives[0].TrimEnd('\\'); ;
                    }
                }


            }
            return "";
        }
    }
}
