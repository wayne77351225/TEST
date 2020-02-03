using System;
using System.Collections.Generic;
using System.Linq;

using System.Collections.ObjectModel;
using G80Utility.Models;

namespace G80Utility.ViewModels
{
    class DeviceViewModel : ViewModelBase
    {
        //ObservableCollection<Device> deviceObserve;
        ObservableCollection<Device> USBDeviceObserve;
        ObservableCollection<Device> RS232DeviceObserve;


        public DeviceViewModel()
        {
            //因為預設沒有資料，所以建構式不帶入property資料
            //deviceObserve = new ObservableCollection<Device> { };
            USBDeviceObserve = new ObservableCollection<Device> { };
            RS232DeviceObserve = new ObservableCollection<Device> { };
        }

        public ObservableCollection<Device> RS232Device //這邊是xaml要binding的對象
        {
            get
            {
                return RS232DeviceObserve;
            }
            set
            {
                if (RS232DeviceObserve != null)
                {
                    RS232DeviceObserve = value;
                    OnPropertyChanged();
                }
            }
        }

        public ObservableCollection<Device> USBDevice //這邊是xaml要binding的對象
        {
            get
            {
                return USBDeviceObserve;
            }
            set
            {
                if (USBDeviceObserve != null)
                {
                    USBDeviceObserve = value;
                    OnPropertyChanged();
                }
            }
        }

        public Device addDevice(Device device)
        {
            USBDeviceObserve.Add(device);
            RS232DeviceObserve.Add(device);
            return device;
        }

        //刪去不符合選定type的device
        public void getDeviceObserve(string type)
        {

            if (type == "usb")
            {
                foreach (Device device in USBDeviceObserve.ToArray()) //在這邊使用toArray避免在remove時出現excepiton
                {
                    if (device.DeviceType != "usb") USBDeviceObserve.Remove(device);
                }
            }

            if (type == "rs232")
            {
                foreach (Device device in RS232DeviceObserve.ToArray()) //在這邊使用toArray避免在remove時出現excepiton
                {
                    if (device.DeviceType != "rs232") RS232DeviceObserve.Remove(device);
                }
            }
            //switch (type)
            //{
            //    case "usb":
            //        foreach (Device device in deviceObserve.ToArray()) //在這邊使用toArray避免在remove時出現excepiton
            //        {
            //            if (device.DeviceType != "usb") deviceObserve.Remove(device);
            //        }              
            //        break;

            //    case "ethernet":
            //        foreach (Device device in deviceObserve.ToArray())
            //        {
            //            if (device.DeviceType != "ethernet") deviceObserve.Remove(device);
            //        }
            //        break;
            //    case "rs232":
            //        foreach (Device device in deviceObserve.ToArray())
            //        {
            //            if (device.DeviceType != "rs232") deviceObserve.Remove(device);
            //        }
            //        break;

            //}

        }

        //移除combobox所有usb資料
        public void removePort(List<Device> devicelist)
        {
            foreach (Device device in devicelist)
            {

                if (RS232DeviceObserve.Contains(device))
                {
                    RS232DeviceObserve.Remove(device);
                }
            }
        }

    }
}
