using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace LyncFellow
{
    class HIDDevices
    {
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        protected struct DeviceInterfaceData
        {
            public int Size;
            public Guid InterfaceClassGuid;
            public int Flags;
            public int Reserved;
        }

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct DeviceInterfaceDetailData
        {
            public int Size;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 0x100)]
            public string DevicePath;
        }

        //public struct HIDBufferSizes
        //{
        //    public short InputReportLength;
        //    public short OutputReportLength;
        //}

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        protected struct HidCaps
        {
            public short Usage;
            public short UsagePage;
            public short InputReportByteLength;
            public short OutputReportByteLength;
            public short FeatureReportByteLength;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 0x11)]
            public short[] Reserved;
            public short NumberLinkCollectionNodes;
            public short NumberInputButtonCaps;
            public short NumberInputValueCaps;
            public short NumberInputDataIndices;
            public short NumberOutputButtonCaps;
            public short NumberOutputValueCaps;
            public short NumberOutputDataIndices;
            public short NumberFeatureButtonCaps;
            public short NumberFeatureValueCaps;
            public short NumberFeatureDataIndices;
        }


        [DllImport("hid.dll", SetLastError = true)]
        protected static extern void HidD_GetHidGuid(out Guid gHid);

        [DllImport("setupapi.dll", SetLastError = true)]
        protected static extern IntPtr SetupDiGetClassDevs(ref Guid gClass, [MarshalAs(UnmanagedType.LPStr)] string strEnumerator, IntPtr hParent, uint nFlags);

        [DllImport("setupapi.dll", SetLastError = true)]
        protected static extern bool SetupDiEnumDeviceInterfaces(IntPtr lpDeviceInfoSet, uint nDeviceInfoData, ref Guid gClass, uint nIndex, ref DeviceInterfaceData oInterfaceData);

        [DllImport("setupapi.dll", SetLastError = true)]
        protected static extern bool SetupDiGetDeviceInterfaceDetail(IntPtr lpDeviceInfoSet, ref DeviceInterfaceData oInterfaceData, ref DeviceInterfaceDetailData oDetailData, uint nDeviceInterfaceDetailDataSize, ref uint nRequiredSize, IntPtr lpDeviceInfoData);

        [DllImport("setupapi.dll", SetLastError = true)]
        protected static extern bool SetupDiGetDeviceInterfaceDetail(IntPtr lpDeviceInfoSet, ref DeviceInterfaceData oInterfaceData, IntPtr lpDeviceInterfaceDetailData, uint nDeviceInterfaceDetailDataSize, ref uint nRequiredSize, IntPtr lpDeviceInfoData);

        [DllImport("setupapi.dll", SetLastError = true)]
        protected static extern int SetupDiDestroyDeviceInfoList(IntPtr lpInfoSet);

        [DllImport("hid.dll", SetLastError = true)]
        protected static extern bool HidD_GetPreparsedData(int hFile, out IntPtr lpData);

        [DllImport("hid.dll", SetLastError = true)]
        protected static extern int HidP_GetCaps(IntPtr lpData, out HidCaps oCaps);

        [DllImport("hid.dll", SetLastError = true)]
        protected static extern bool HidD_FreePreparsedData(ref IntPtr pData);


        private static string GetDevicePath(IntPtr hInfoSet, ref DeviceInterfaceData oInterface)
        {
            uint nRequiredSize = 0;
            if (!SetupDiGetDeviceInterfaceDetail(hInfoSet, ref oInterface, IntPtr.Zero, 0, ref nRequiredSize, IntPtr.Zero))
            {
                DeviceInterfaceDetailData oDetailData = new DeviceInterfaceDetailData();
                oDetailData.Size = 5;
                if (SetupDiGetDeviceInterfaceDetail(hInfoSet, ref oInterface, ref oDetailData, nRequiredSize, ref nRequiredSize, IntPtr.Zero))
                {
                    return oDetailData.DevicePath;
                }
            }
            return null;
        }

        public static string[] Find(int nVid, int nPid)
        {
            Guid guid;
            List<string> list = new List<string>();
            string str = string.Format("vid_{0:x4}&pid_{1:x4}", nVid, nPid);
            HidD_GetHidGuid(out guid);
            IntPtr lpDeviceInfoSet = SetupDiGetClassDevs(ref guid, null, IntPtr.Zero, 0x12);
            try
            {
                DeviceInterfaceData structure = new DeviceInterfaceData();
                structure.Size = Marshal.SizeOf(structure);
                for (int i = 0; SetupDiEnumDeviceInterfaces(lpDeviceInfoSet, 0, ref guid, (uint)i, ref structure); i++)
                {
                    string devicePath = GetDevicePath(lpDeviceInfoSet, ref structure);
                    if (devicePath.IndexOf(str) >= 0)
                    {
                        list.Add(devicePath);
                    }
                }
            }
            finally
            {
                SetupDiDestroyDeviceInfoList(lpDeviceInfoSet);
            }
            return list.ToArray();
        }

        //public static HIDBufferSizes GetBufferSizes(int hBuddy)
        //{
        //    IntPtr ptr;
        //    HIDBufferSizes buffertsizes = new HIDBufferSizes();
        //    buffertsizes.InputReportLength = buffertsizes.OutputReportLength = 0;

        //    if (HidD_GetPreparsedData(hBuddy, out ptr))
        //    {
        //        try
        //        {
        //            HidCaps caps;
        //            HidP_GetCaps(ptr, out caps);
        //            buffertsizes.InputReportLength = caps.InputReportByteLength;
        //            buffertsizes.OutputReportLength = caps.OutputReportByteLength;

        //        }
        //        finally
        //        {
        //            HidD_FreePreparsedData(ref ptr);
        //        }
        //    }
        //    return buffertsizes;
        //}
    }
}
