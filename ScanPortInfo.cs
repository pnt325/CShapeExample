using System.Management;
using System.IO.Ports;
using System;

/*
PS U:\> Get-WmiObject -Query "select *from Win32_PnPEntity where PNPClass='Ports'"


__GENUS                     : 2
__CLASS                     : Win32_PnPEntity
__SUPERCLASS                : CIM_LogicalDevice
__DYNASTY                   : CIM_ManagedSystemElement
__RELPATH                   : Win32_PnPEntity.DeviceID="USB\\VID_10C4&PID_EA60\\0001"
__PROPERTY_COUNT            : 26
__DERIVATION                : {CIM_LogicalDevice, CIM_LogicalElement, CIM_ManagedSystemElement}
__SERVER                    : FVN-NB-054
__NAMESPACE                 : root\cimv2
__PATH                      : \\FVN-NB-054\root\cimv2:Win32_PnPEntity.DeviceID="USB\\VID_10C4&PID_EA60\\0001"
Availability                :
Caption                     : Silicon Labs CP210x USB to UART Bridge (COM18)
ClassGuid                   : {4d36e978-e325-11ce-bfc1-08002be10318}
CompatibleID                : {USB\Class_FF&SubClass_00&Prot_00, USB\Class_FF&SubClass_00, USB\Class_FF}
ConfigManagerErrorCode      : 0
ConfigManagerUserConfig     : False
CreationClassName           : Win32_PnPEntity
Description                 : Silicon Labs CP210x USB to UART Bridge
DeviceID                    : USB\VID_10C4&PID_EA60\0001
ErrorCleared                :
ErrorDescription            :
HardwareID                  : {USB\VID_10C4&PID_EA60&REV_0100, USB\VID_10C4&PID_EA60}
InstallDate                 :
LastErrorCode               :
Manufacturer                : Silicon Labs
Name                        : Silicon Labs CP210x USB to UART Bridge (COM18)
PNPClass                    : Ports
PNPDeviceID                 : USB\VID_10C4&PID_EA60\0001
PowerManagementCapabilities :
PowerManagementSupported    :
Present                     : True
Service                     : silabser
Status                      : OK
StatusInfo                  :
SystemCreationClassName     : Win32_ComputerSystem
SystemName                  : FVN-NB-054
PSComputerName              : FVN-NB-054
*/


class ScanPortInfo
{
  public SerialPortInfo()
  {
  
  }
  
  public PortInfo[] Scan()
  {
  string[] ports = SerialPort.GetPortNames();
  List<PortInfo> portInfos = new List<PortInfo>();
  ManagementObjectSearcher seacher = new ManagementObjectSearcher("SELECT *FROM Win32_PnPEntity WHERE PNPClass='Ports'");
  var queryObj = seacher.Get();
    foreach (string port in ports)
  {
      PortInfo pi = new PortInfo();
      pi.PortName = port;
      foreach (var obj in queryObj)
      {
          if (obj != null)
          {
              string _name = obj["Name"].ToString();
              if (_name.Contains(port))
              {
                  pi.Description = obj["Description"].ToString();
                  portInfos.Add(pi);
                  break;
              }

          }
      }
  }

  return (portInfos.Count > 0) ? portInfos.ToArray() : null;
  }
}

public class PortInfo
{
  public string PortName;     // COM_x
  public string Description;  // Port detail information
  public PortInfo()
  {
  }
}
