using System;
using System.Collections.Generic;
using System.Linq;

using System.Collections.ObjectModel;
using PirnterUtility.Models;

namespace PirnterUtility.ViewModels
{
    class DeviceViewModel : ViewModelBase
    {
        ObservableCollection<Device> deviceObserve;

        public DeviceViewModel()
        {
            //因為預設沒有資料，所以建構式不帶入property資料
            deviceObserve = new ObservableCollection<Device>
            {

            };
        }

        public ObservableCollection<Device> Device //這邊是xaml要binding的對象
        {
            get
            {
                return deviceObserve;
            }
            set
            {
                if (deviceObserve != null)
                {
                    deviceObserve = value;
                    OnPropertyChanged();
                }
            }
        }

        public Device addDevice(Device device)
        {
            deviceObserve.Add(device);

            return device;
        }

        //刪去不符合選定type的device
        public void getDeviceObserve(string type)
        {
            switch (type)
            {
                case "usb":
                    foreach (Device device in deviceObserve.ToArray()) //在這邊使用toArray避免在remove時出現excepiton
                    {
                        if (device.DeviceType != "usb") deviceObserve.Remove(device);
                    }              
                    break;

                case "ethernet":
                    foreach (Device device in deviceObserve.ToArray())
                    {
                        if (device.DeviceType != "ethernet") deviceObserve.Remove(device);
                    }
                    break;
                case "rs232":
                    foreach (Device device in deviceObserve.ToArray())
                    {
                        if (device.DeviceType != "rs232") deviceObserve.Remove(device);
                    }
                    break;

            }

        }

        //移除combobox所有usb資料
        public void removePort(List<Device> devicelist)
        {
            foreach (Device device in devicelist)
            {

                if (deviceObserve.Contains(device))
                {
                    deviceObserve.Remove(device);
                }
            }
        }

    }
}
