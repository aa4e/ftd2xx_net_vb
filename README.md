# ftd2xx_net_vb
**FTD2XX** wrapper on `Visual Basic .NET`.

## Description

It is a Visual Basic branch of an official FTDI's FTD2XX_NET C# wrapper for the `ftd2xx` library with more object-oriented manner.

## Usage

- Get number of FTD2XX devices and common info:

```
Console.WriteLine($"Library version: {FTD2XX_NET.Ftdi.LibraryVersion:X6}")
Console.WriteLine($"Found {FTD2XX_NET.Ftdi.GetNumberOfDevices()} FTDI devices:")
For Each d In FTD2XX_NET.Ftdi.GetDeviceList()  
  With d
    Console.WriteLine($"- {.Description}")    
    Console.WriteLine($"-- ID: {.DeviceID:X8}")
    Console.WriteLine($"-- Type: {.DeviceType}")
    Console.WriteLine($"-- Serial: {.SerialNumber}")
    Console.WriteLine($"-- Driver: {.DriverVersion:X6}")
    Console.WriteLine($"-- Descr: {.Description}")
  End With
Next
```

- Send byte buffer:

```
Using ftd As New FTD2XX_NET.Ftdi(0)
  ftd.ResetPort()  
  Dim buf As Byte() = {1,2,3,4,5}
  Dim wrote = ftd.Write(buf, buf.Length)
End Using
```

- Finally, unload library:

```
FTD2XX_NET.Ftdi.UnloadLibrary()
```
