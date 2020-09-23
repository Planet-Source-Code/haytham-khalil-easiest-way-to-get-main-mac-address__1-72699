<div align="center">

## Easiest way to get main MAC address


</div>

### Description

A simple function to get PC's MAC address
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Haytham Khalil](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/haytham-khalil.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/haytham-khalil-easiest-way-to-get-main-mac-address__1-72699/archive/master.zip)





### Source Code

```
Private Function GetMACAddress() As String
Dim Devices As Object
Dim Device As Object
Dim Temp As Variant
Dim Info As String
  Set Devices = GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapter")
  For Each Device In Devices
     For Each Temp In Device.Properties_
       If Temp.Name = "MACAddress" Then GetMACAddress = CStr(Temp): Exit Function
     Next
  Next Device
End Function
```

