using USBPOC;

UsbService usbService = new();
//usbService.RunAllManagementObjectSearches();
//usbService.CheckPorts();
usbService.ExplorePnPEntities();