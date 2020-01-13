using System.Runtime.InteropServices;
using System;

namespace PirnterUtility.Tool
{
    /// <summary>
    /// To get USBBroadcastinterface
    /// </summary>
    //Sequential,顺序布局,内存排列是按成员的先后顺序排列.
    [StructLayout(LayoutKind.Sequential)]
    public struct USBBroadcastinterface
    {
        /// <summary>
        /// The size
        /// </summary>
        internal int Size;

        /// <summary>
        /// The device type
        /// </summary>
        internal int USBType;

        /// <summary>
        /// The reserved
        /// </summary>
        internal int Reserved;

        /// <summary>
        /// The class unique identifier
        /// </summary>
        internal Guid ClassGuid;

        /// <summary>
        /// The name
        /// </summary>
        internal short Name;
    }

    /// <summary>
    /// To get DeviceDiscoveryManager
    /// </summary>
    public class USBDetector
    {
        /// <summary>
        /// The new usb device connected
        /// </summary>
        //A device or piece of media has been inserted and is now available.
        //message from Windows
        public const int NewUsbDeviceConnected = 0x8000;

        /// <summary>
        /// The usb device removed
        /// </summary>
        public const int UsbDeviceRemoved = 0x8004;

        /// <summary>
        /// The usb devicechange
        /// </summary>
        public const int UsbDevicechange = 0x0219;

        /// <summary>
        /// The DBT devtyp deviceinterface
        /// </summary>
        private const int DbtDevtypDeviceinterface = 5;

        /// <summary>
        /// The unique identifier devinterface usb device
        /// </summary>
        private static readonly Guid GuidDevinterfaceUSBDevice = new Guid("A5DCBF10-6530-11D2-901F-00C04FB951ED"); // USB devices

        /// <summary>
        /// The notification handle
        /// </summary>
        private static IntPtr notificationHandle;

        /// <summary>
        /// Registers a window to receive notifications when USB devices are plugged or unplugged.
        /// </summary>
        /// <param name="windowHandle">Handle to the window receiving notifications.</param>
        public static void RegisterUsbDeviceNotification(IntPtr windowHandle)
        {
            USBBroadcastinterface dbi = new USBBroadcastinterface
            {
                USBType = DbtDevtypDeviceinterface,
                Reserved = 0,
                ClassGuid = GuidDevinterfaceUSBDevice,
                Name = 0
            };

            dbi.Size = Marshal.SizeOf(dbi); //Returns the unmanaged size, in bytes, of a class.
            //AllocHGlobal:使用指定的位元組數目，從處理序的 Unmanaged 記憶體中配置記憶體。
            //傳回新配置的記憶體的指標。
            IntPtr buffer = Marshal.AllocHGlobal(dbi.Size);

            //從 Managed(託管) 物件封送處理資料到 Unmanaged (非託管)記憶體區塊。
            Marshal.StructureToPtr(dbi, buffer, true);

            notificationHandle = RegisterDeviceNotification(windowHandle, buffer, 0);
        }

        /// <summary>
        /// Unregisters the window for USB device notifications
        /// </summary>
        public static void UnregisterUsbDeviceNotification()
        {
            UnregisterDeviceNotification(notificationHandle);
        }

        /// <summary>
        /// Registers the device notification.
        /// </summary>
        /// <param name="recipient">The recipient.</param>
        /// <param name="notificationFilter">The notification filter.</param>
        /// <param name="flags">The flags.</param>
        /// <returns>returns IntPtr</returns>
        [DllImport("user32.dll", CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern IntPtr RegisterDeviceNotification(IntPtr recipient, IntPtr notificationFilter, int flags);

        /// <summary>
        /// Unregisters the device notification.
        /// </summary>
        /// <param name="handle">The handle.</param>
        /// <returns>returns bool</returns>
        //直接使用DllImport外部Dll的方法
        [DllImport("user32.dll")]
        private static extern bool UnregisterDeviceNotification(IntPtr handle);
    }
}
