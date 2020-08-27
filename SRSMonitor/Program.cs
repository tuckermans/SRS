using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management;
using System.Diagnostics;
using System.IO;
using System.Threading;
using FTD2XX_NET;
using System.Runtime.InteropServices;


 
namespace SRSMonitor
{
    class Program
    {
        [DllImport(@"FTD2XX.dll")]
        private static extern UInt32 FT_CreateDeviceInfoList(ref UInt32 NumDevs);

        [DllImport(@"FTD2XX.dll")]
        private static extern UInt32 FT_GetDeviceInfoDetail(
             UInt32 dwIndex,
             ref UInt32 lpdwFlags,
             ref UInt32 lpdwType,
             ref UInt32 lpdwID,
             ref UInt32 lpdwLocId,
             byte [] pcSerialNumber,
             byte[] pcDescription,
             ref UInt32 Handle);
    


        static ManagementEventWatcher Connect;
        static ManagementEventWatcher Disconnect;
        //static readonly ushort EP_VENDOR_ID = 0x0403;
       // static readonly ushort EP_PRODUCT_ID = 0x6015;
        static readonly string VIDString = "VID_0403";
        static readonly string PIDString = "PID_6015";
        static readonly string DeviceIDString = "FTDIBUS";
        static AutoResetEvent WaitForEver = new AutoResetEvent(false);
        static string ExecName = "EasyProOne.exe";
        static string ExecPath = string.Empty;
        static FTDI deviceCom = new FTDI();

       
        //private static extern int methodName(int b);
        static void Main(string[] args)
        {
            devicelist = new FTDI.FT_DEVICE_INFO_NODE[8];

#if !DEBUG
            ExecPath = Path.Combine("C:\\Users\\Jack\\Documents\\BAT\\SRS Medical\\WVPM\\EasyProOne\\ZedGraphSample\\bin", ExecName);
#else
            ExecPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "SRS Medical");

            ExecPath = Path.Combine(ExecPath, "VPM");
            ExecPath = Path.Combine(ExecPath, ExecName);
#endif
            // setup watch on usb PnP devices attached
            WqlEventQuery ConnectQuery = new WqlEventQuery("SELECT * FROM __InstanceCreationEvent " +
                                                    "WITHIN 2 "
                                                     + "WHERE TargetInstance ISA 'Win32_PnPEntity'"); // AND TargetInstance.DeviceID = 'USB\\VID_0403&PID_6015\\DM01AX7E' ");
            Connect = new ManagementEventWatcher(ConnectQuery);
            Connect.EventArrived += new EventArrivedEventHandler(Connect_UsbDeviceAttached);


            // setup watch on usb device removed
            WqlEventQuery DisconnectQuery = new WqlEventQuery("SELECT * FROM __InstanceDeletionEvent " +
                                                    "WITHIN 2 "
                                                     + "WHERE TargetInstance ISA 'Win32_PnPEntity'");

            Disconnect = new ManagementEventWatcher(DisconnectQuery);
            Disconnect.EventArrived += new EventArrivedEventHandler(Disconnect_UsbDeviceRemoved);

            Connect.Start();
            Disconnect.Start();

            // wait forever
            WaitForEver.WaitOne();
        }

        static void Connect_UsbDeviceAttached(object sender, EventArrivedEventArgs e)
        {
            ManagementBaseObject instance = (ManagementBaseObject)e.NewEvent["TargetInstance"];

            // make sure this is our device

            string DeviceID = (string)instance.Properties["PNPDeviceID"].Value;
            string DeviceDescription = (string)instance.Properties["Description"].Value;
            

            var properties = instance.Properties;
            Thread.Sleep(2000);

            // If the PID/VID matches then check the 
            if (Match(instance.Properties))
            {
                Console.WriteLine("USB Device Attached: " + DeviceDescription + " : " + DeviceID);
                
               

                // if the EP One app is not already running then start iy
                Process[] procs = Process.GetProcessesByName("EasyProOne");
                if (procs.Length == 0)
                {
                    
                    
                    if (File.Exists(ExecPath))
                    {
                        Process EasyProOne = new Process();
                        EasyProOne.StartInfo.FileName = ExecPath;
                        EasyProOne.StartInfo.UseShellExecute = true;
                        EasyProOne.StartInfo.RedirectStandardOutput = false;
                        EasyProOne.StartInfo.CreateNoWindow = true;
                        EasyProOne.Start();
                    }
                }
                else
                {
                    Console.WriteLine("EasyProOne already running.");
                }

            }
        }

        static void Disconnect_UsbDeviceRemoved(object sender, EventArrivedEventArgs e)
        {
            ManagementBaseObject instance = (ManagementBaseObject)e.NewEvent["TargetInstance"];

            // make sure this is our device

            string DeviceID = (string)instance.Properties["PNPDeviceID"].Value;
            string DeviceDescription = (string)instance.Properties["Description"].Value;

            if (Match(instance.Properties))
            {
                Console.WriteLine("USB Device Removed: " + DeviceDescription + " : " + DeviceID);

                Console.WriteLine("EasyPro Device Removed.. ");

            }
        }

        static bool Match(PropertyDataCollection properties)
        {
            // look for Device ID and device description
           // string DeviceID = (string)properties["PNPDeviceID"].Value;
           // string DeviceDescription = (string)properties["Description"].Value;

            for (int j = 0; j < 8; j++)
            {
                devicelist[j] = null;
            }

            UInt32 NumDevices = 0;
            FT_CreateDeviceInfoList(ref NumDevices);
            if (NumDevices > 0)
            {
                if (deviceCom.GetDeviceList(devicelist) == FTDI.FT_STATUS.FT_OK)
                {
                    for (UInt32 i = 0; i < NumDevices; i++)
                    {
                        Console.WriteLine("Description:" + devicelist[i].Description);

                        if (devicelist[i].Description != string.Empty && devicelist[i].Description != "CT3KPRO")
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
            /*
            if (DeviceDescription.Equals("USB Serial Port"))
            {
                return (DeviceID.StartsWith(DeviceIDString)
                                        && DeviceID.Contains(VIDString)
                                        && DeviceID.Contains(PIDString));
            }

            return false;
             * */
        }




        public static FTDI.FT_DEVICE_INFO_NODE[] devicelist { get; set; }
    }
}
