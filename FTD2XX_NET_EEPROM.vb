Imports System.Runtime.InteropServices
Imports System.Text
Imports FTD2XX_NET.Ftdi

Namespace FTD2XX_NET

    ''' <summary>
    ''' Работа с ППЗУ микросхем FTDI.
    ''' </summary>
    Public Class Eeprom

#Region "CTOR"

        Private ReadOnly FtHandle As Integer
        Private ReadOnly DeviceType As FT_DEVICE

        Friend Sub New(ftDllHandle As Integer, ftHandle As Integer, devType As FT_DEVICE)
            Me.FtHandle = ftHandle
            Me.DeviceType = devType
            FindEeFunctionPointers(ftDllHandle)
        End Sub

#End Region '/CTOR

#Region "ЧТЕНИЕ ППЗУ"

        ''' <summary>
        ''' Читает из ПЗУ и возвращает одно слово данных по адресу <paramref name="address"/>.
        ''' </summary>
        ''' <param name="address">Адрес ячейки памяти.</param>
        Public Function ReadEeprom(address As UInteger) As UShort
            Dim value As UShort
            Dim dlg As ReadEeDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_ReadEE, IntPtr), GetType(ReadEeDelegate)), ReadEeDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, address, value)
            CheckErrors(status)
            Return value
        End Function

        ''' <summary>
        ''' Читает данные из ПЗУ устройства.
        ''' </summary>
        Public Function ReadEeprom() As FT_EEPROM_DATA
            Select Case DeviceType
                Case FT_DEVICE.FT_DEVICE_BM
                    Return ReadFt232bEeprom()
                Case FT_DEVICE.FT_DEVICE_2232
                    Return ReadFt2232Eeprom()
                Case FT_DEVICE.FT_DEVICE_232R
                    Return ReadFt232rEeprom()
                Case FT_DEVICE.FT_DEVICE_2232H
                    Return ReadFt2232hEeprom()
                Case FT_DEVICE.FT_DEVICE_4232H
                    Return ReadFt4232hEeprom()
                Case FT_DEVICE.FT_DEVICE_232H
                    Return ReadFt232hEeprom()
                Case FT_DEVICE.FT_DEVICE_X_SERIES
                    Return ReadXSeriesEeprom()
            End Select
            Throw New FT_EXCEPTION("Ошибка чтения ППЗУ.")
        End Function

        Private Function ReadFt232bEeprom() As FT232B_EEPROM_STRUCTURE
            Dim ee232b As New FT232B_EEPROM_STRUCTURE()
            Dim programData As New FT_PROGRAM_DATA() With {
                .Signature1 = 0UI,
                .Signature2 = UInteger.MaxValue,
                .Version = 2UI,
                .Manufacturer = Marshal.AllocHGlobal(32),
                .ManufacturerID = Marshal.AllocHGlobal(16),
                .Description = Marshal.AllocHGlobal(64),
                .SerialNumber = Marshal.AllocHGlobal(16)
            }
            Dim rdDlg As EeReadDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EE_Read, IntPtr), GetType(EeReadDelegate)), EeReadDelegate)
            Dim status As FT_STATUS = rdDlg(FtHandle, programData)
            ee232b.Manufacturer = Marshal.PtrToStringAnsi(programData.Manufacturer)
            ee232b.ManufacturerID = Marshal.PtrToStringAnsi(programData.ManufacturerID)
            ee232b.Description = Marshal.PtrToStringAnsi(programData.Description)
            ee232b.SerialNumber = Marshal.PtrToStringAnsi(programData.SerialNumber)
            Marshal.FreeHGlobal(programData.Manufacturer)
            Marshal.FreeHGlobal(programData.ManufacturerID)
            Marshal.FreeHGlobal(programData.Description)
            Marshal.FreeHGlobal(programData.SerialNumber)
            ee232b.VendorID = programData.VendorID
            ee232b.ProductID = programData.ProductID
            ee232b.MaxPower = programData.MaxPower
            ee232b.SelfPowered = Convert.ToBoolean(programData.SelfPowered)
            ee232b.RemoteWakeup = Convert.ToBoolean(programData.RemoteWakeup)
            ee232b.PullDownEnable = Convert.ToBoolean(programData.PullDownEnable)
            ee232b.SerNumEnable = Convert.ToBoolean(programData.SerNumEnable)
            ee232b.USBVersionEnable = Convert.ToBoolean(programData.USBVersionEnable)
            ee232b.USBVersion = programData.USBVersion
            CheckErrors(status)
            Return ee232b
        End Function

        Private Function ReadFt2232Eeprom() As FT2232_EEPROM_STRUCTURE
            Dim ee2232 As New FT2232_EEPROM_STRUCTURE()
            Dim programData As New FT_PROGRAM_DATA() With {
                .Signature1 = 0UI,
                .Signature2 = UInteger.MaxValue,
                .Version = 2UI,
                .Manufacturer = Marshal.AllocHGlobal(32),
                .ManufacturerID = Marshal.AllocHGlobal(16),
                .Description = Marshal.AllocHGlobal(64),
                .SerialNumber = Marshal.AllocHGlobal(16)
            }
            Dim rdDlg As EeReadDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EE_Read, IntPtr), GetType(EeReadDelegate)), EeReadDelegate)
            Dim status As FT_STATUS = rdDlg(FtHandle, programData)
            ee2232.Manufacturer = Marshal.PtrToStringAnsi(programData.Manufacturer)
            ee2232.ManufacturerID = Marshal.PtrToStringAnsi(programData.ManufacturerID)
            ee2232.Description = Marshal.PtrToStringAnsi(programData.Description)
            ee2232.SerialNumber = Marshal.PtrToStringAnsi(programData.SerialNumber)
            Marshal.FreeHGlobal(programData.Manufacturer)
            Marshal.FreeHGlobal(programData.ManufacturerID)
            Marshal.FreeHGlobal(programData.Description)
            Marshal.FreeHGlobal(programData.SerialNumber)
            ee2232.VendorID = programData.VendorID
            ee2232.ProductID = programData.ProductID
            ee2232.MaxPower = programData.MaxPower
            ee2232.SelfPowered = Convert.ToBoolean(programData.SelfPowered)
            ee2232.RemoteWakeup = Convert.ToBoolean(programData.RemoteWakeup)
            ee2232.PullDownEnable = Convert.ToBoolean(programData.PullDownEnable5)
            ee2232.SerNumEnable = Convert.ToBoolean(programData.SerNumEnable5)
            ee2232.USBVersionEnable = Convert.ToBoolean(programData.USBVersionEnable5)
            ee2232.USBVersion = programData.USBVersion5
            ee2232.AIsHighCurrent = Convert.ToBoolean(programData.AIsHighCurrent)
            ee2232.BIsHighCurrent = Convert.ToBoolean(programData.BIsHighCurrent)
            ee2232.IFAIsFifo = Convert.ToBoolean(programData.IFAIsFifo)
            ee2232.IFAIsFifoTar = Convert.ToBoolean(programData.IFAIsFifoTar)
            ee2232.IFAIsFastSer = Convert.ToBoolean(programData.IFAIsFastSer)
            ee2232.AIsVCP = Convert.ToBoolean(programData.AIsVCP)
            ee2232.IFBIsFifo = Convert.ToBoolean(programData.IFBIsFifo)
            ee2232.IFBIsFifoTar = Convert.ToBoolean(programData.IFBIsFifoTar)
            ee2232.IFBIsFastSer = Convert.ToBoolean(programData.IFBIsFastSer)
            ee2232.BIsVCP = Convert.ToBoolean(programData.BIsVCP)
            CheckErrors(status)
            Return ee2232
        End Function

        Private Function ReadFt232rEeprom() As FT232R_EEPROM_STRUCTURE
            Dim ee232r As New FT232R_EEPROM_STRUCTURE()
            Dim programData As New FT_PROGRAM_DATA() With {
                .Signature1 = 0UI,
                .Signature2 = UInteger.MaxValue,
                .Version = 2UI,
                .Manufacturer = Marshal.AllocHGlobal(32),
                .ManufacturerID = Marshal.AllocHGlobal(16),
                .Description = Marshal.AllocHGlobal(64),
                .SerialNumber = Marshal.AllocHGlobal(16)
            }
            Dim rdDlg As EeReadDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EE_Read, IntPtr), GetType(EeReadDelegate)), EeReadDelegate)
            Dim status As FT_STATUS = rdDlg(FtHandle, programData)
            ee232r.Manufacturer = Marshal.PtrToStringAnsi(programData.Manufacturer)
            ee232r.ManufacturerID = Marshal.PtrToStringAnsi(programData.ManufacturerID)
            ee232r.Description = Marshal.PtrToStringAnsi(programData.Description)
            ee232r.SerialNumber = Marshal.PtrToStringAnsi(programData.SerialNumber)
            Marshal.FreeHGlobal(programData.Manufacturer)
            Marshal.FreeHGlobal(programData.ManufacturerID)
            Marshal.FreeHGlobal(programData.Description)
            Marshal.FreeHGlobal(programData.SerialNumber)
            ee232r.VendorID = programData.VendorID
            ee232r.ProductID = programData.ProductID
            ee232r.MaxPower = programData.MaxPower
            ee232r.SelfPowered = Convert.ToBoolean(programData.SelfPowered)
            ee232r.RemoteWakeup = Convert.ToBoolean(programData.RemoteWakeup)
            ee232r.UseExtOsc = Convert.ToBoolean(programData.UseExtOsc)
            ee232r.HighDriveIOs = Convert.ToBoolean(programData.HighDriveIOs)
            ee232r.EndpointSize = programData.EndpointSize
            ee232r.PullDownEnable = Convert.ToBoolean(programData.PullDownEnableR)
            ee232r.SerNumEnable = Convert.ToBoolean(programData.SerNumEnableR)
            ee232r.InvertTXD = Convert.ToBoolean(programData.InvertTXD)
            ee232r.InvertRXD = Convert.ToBoolean(programData.InvertRXD)
            ee232r.InvertRTS = Convert.ToBoolean(programData.InvertRTS)
            ee232r.InvertCTS = Convert.ToBoolean(programData.InvertCTS)
            ee232r.InvertDTR = Convert.ToBoolean(programData.InvertDTR)
            ee232r.InvertDSR = Convert.ToBoolean(programData.InvertDSR)
            ee232r.InvertDCD = Convert.ToBoolean(programData.InvertDCD)
            ee232r.InvertRI = Convert.ToBoolean(programData.InvertRI)
            ee232r.Cbus0 = programData.Cbus0
            ee232r.Cbus1 = programData.Cbus1
            ee232r.Cbus2 = programData.Cbus2
            ee232r.Cbus3 = programData.Cbus3
            ee232r.Cbus4 = programData.Cbus4
            ee232r.RIsD2XX = Convert.ToBoolean(programData.RIsD2XX)
            CheckErrors(status)
            Return ee232r
        End Function

        Private Function ReadFt2232hEeprom() As FT2232H_EEPROM_STRUCTURE
            Dim ee2232h As New FT2232H_EEPROM_STRUCTURE()
            Dim programData As New FT_PROGRAM_DATA() With {
                .Signature1 = 0UI,
                .Signature2 = UInteger.MaxValue,
                .Version = 3UI,
                .Manufacturer = Marshal.AllocHGlobal(32),
                .ManufacturerID = Marshal.AllocHGlobal(16),
                .Description = Marshal.AllocHGlobal(64),
                .SerialNumber = Marshal.AllocHGlobal(16)
            }
            Dim rdDlg As EeReadDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EE_Read, IntPtr), GetType(EeReadDelegate)), EeReadDelegate)
            Dim status As FT_STATUS = rdDlg(FtHandle, programData)
            ee2232h.Manufacturer = Marshal.PtrToStringAnsi(programData.Manufacturer)
            ee2232h.ManufacturerID = Marshal.PtrToStringAnsi(programData.ManufacturerID)
            ee2232h.Description = Marshal.PtrToStringAnsi(programData.Description)
            ee2232h.SerialNumber = Marshal.PtrToStringAnsi(programData.SerialNumber)
            Marshal.FreeHGlobal(programData.Manufacturer)
            Marshal.FreeHGlobal(programData.ManufacturerID)
            Marshal.FreeHGlobal(programData.Description)
            Marshal.FreeHGlobal(programData.SerialNumber)
            ee2232h.VendorID = programData.VendorID
            ee2232h.ProductID = programData.ProductID
            ee2232h.MaxPower = programData.MaxPower
            ee2232h.SelfPowered = Convert.ToBoolean(programData.SelfPowered)
            ee2232h.RemoteWakeup = Convert.ToBoolean(programData.RemoteWakeup)
            ee2232h.PullDownEnable = Convert.ToBoolean(programData.PullDownEnable7)
            ee2232h.SerNumEnable = Convert.ToBoolean(programData.SerNumEnable7)
            ee2232h.ALSlowSlew = Convert.ToBoolean(programData.ALSlowSlew)
            ee2232h.ALSchmittInput = Convert.ToBoolean(programData.ALSchmittInput)
            ee2232h.ALDriveCurrent = programData.ALDriveCurrent
            ee2232h.AHSlowSlew = Convert.ToBoolean(programData.AHSlowSlew)
            ee2232h.AHSchmittInput = Convert.ToBoolean(programData.AHSchmittInput)
            ee2232h.AHDriveCurrent = programData.AHDriveCurrent
            ee2232h.BLSlowSlew = Convert.ToBoolean(programData.BLSlowSlew)
            ee2232h.BLSchmittInput = Convert.ToBoolean(programData.BLSchmittInput)
            ee2232h.BLDriveCurrent = programData.BLDriveCurrent
            ee2232h.BHSlowSlew = Convert.ToBoolean(programData.BHSlowSlew)
            ee2232h.BHSchmittInput = Convert.ToBoolean(programData.BHSchmittInput)
            ee2232h.BHDriveCurrent = programData.BHDriveCurrent
            ee2232h.IFAIsFifo = Convert.ToBoolean(programData.IFAIsFifo7)
            ee2232h.IFAIsFifoTar = Convert.ToBoolean(programData.IFAIsFifoTar7)
            ee2232h.IFAIsFastSer = Convert.ToBoolean(programData.IFAIsFastSer7)
            ee2232h.AIsVCP = Convert.ToBoolean(programData.AIsVCP7)
            ee2232h.IFBIsFifo = Convert.ToBoolean(programData.IFBIsFifo7)
            ee2232h.IFBIsFifoTar = Convert.ToBoolean(programData.IFBIsFifoTar7)
            ee2232h.IFBIsFastSer = Convert.ToBoolean(programData.IFBIsFastSer7)
            ee2232h.BIsVCP = Convert.ToBoolean(programData.BIsVCP7)
            ee2232h.PowerSaveEnable = Convert.ToBoolean(programData.PowerSaveEnable)
            CheckErrors(status)
            Return ee2232h
        End Function

        Private Function ReadFt4232hEeprom() As FT4232H_EEPROM_STRUCTURE
            Dim ee4232h As New FT4232H_EEPROM_STRUCTURE()
            Dim programData As New FT_PROGRAM_DATA() With {
                .Signature1 = 0UI,
                .Signature2 = UInteger.MaxValue,
                .Version = 4UI,
                .Manufacturer = Marshal.AllocHGlobal(32),
                .ManufacturerID = Marshal.AllocHGlobal(16),
                .Description = Marshal.AllocHGlobal(64),
                .SerialNumber = Marshal.AllocHGlobal(16)
            }
            Dim rdDlg As EeReadDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EE_Read, IntPtr), GetType(EeReadDelegate)), EeReadDelegate)
            Dim status As FT_STATUS = rdDlg(FtHandle, programData)
            ee4232h.Manufacturer = Marshal.PtrToStringAnsi(programData.Manufacturer)
            ee4232h.ManufacturerID = Marshal.PtrToStringAnsi(programData.ManufacturerID)
            ee4232h.Description = Marshal.PtrToStringAnsi(programData.Description)
            ee4232h.SerialNumber = Marshal.PtrToStringAnsi(programData.SerialNumber)
            Marshal.FreeHGlobal(programData.Manufacturer)
            Marshal.FreeHGlobal(programData.ManufacturerID)
            Marshal.FreeHGlobal(programData.Description)
            Marshal.FreeHGlobal(programData.SerialNumber)
            ee4232h.VendorID = programData.VendorID
            ee4232h.ProductID = programData.ProductID
            ee4232h.MaxPower = programData.MaxPower
            ee4232h.SelfPowered = Convert.ToBoolean(programData.SelfPowered)
            ee4232h.RemoteWakeup = Convert.ToBoolean(programData.RemoteWakeup)
            ee4232h.PullDownEnable = Convert.ToBoolean(programData.PullDownEnable8)
            ee4232h.SerNumEnable = Convert.ToBoolean(programData.SerNumEnable8)
            ee4232h.ASlowSlew = Convert.ToBoolean(programData.ASlowSlew)
            ee4232h.ASchmittInput = Convert.ToBoolean(programData.ASchmittInput)
            ee4232h.ADriveCurrent = programData.ADriveCurrent
            ee4232h.BSlowSlew = Convert.ToBoolean(programData.BSlowSlew)
            ee4232h.BSchmittInput = Convert.ToBoolean(programData.BSchmittInput)
            ee4232h.BDriveCurrent = programData.BDriveCurrent
            ee4232h.CSlowSlew = Convert.ToBoolean(programData.CSlowSlew)
            ee4232h.CSchmittInput = Convert.ToBoolean(programData.CSchmittInput)
            ee4232h.CDriveCurrent = programData.CDriveCurrent
            ee4232h.DSlowSlew = Convert.ToBoolean(programData.DSlowSlew)
            ee4232h.DSchmittInput = Convert.ToBoolean(programData.DSchmittInput)
            ee4232h.DDriveCurrent = programData.DDriveCurrent
            ee4232h.ARIIsTXDEN = Convert.ToBoolean(programData.ARIIsTXDEN)
            ee4232h.BRIIsTXDEN = Convert.ToBoolean(programData.BRIIsTXDEN)
            ee4232h.CRIIsTXDEN = Convert.ToBoolean(programData.CRIIsTXDEN)
            ee4232h.DRIIsTXDEN = Convert.ToBoolean(programData.DRIIsTXDEN)
            ee4232h.AIsVCP = Convert.ToBoolean(programData.AIsVCP8)
            ee4232h.BIsVCP = Convert.ToBoolean(programData.BIsVCP8)
            ee4232h.CIsVCP = Convert.ToBoolean(programData.CIsVCP8)
            ee4232h.DIsVCP = Convert.ToBoolean(programData.DIsVCP8)
            CheckErrors(status)
            Return ee4232h
        End Function

        Private Function ReadFt232hEeprom() As FT232H_EEPROM_STRUCTURE
            Dim ee232h As New FT232H_EEPROM_STRUCTURE()
            Dim programData As New FT_PROGRAM_DATA() With {
                .Signature1 = 0UI,
                .Signature2 = UInteger.MaxValue,
                .Version = 5UI,
                .Manufacturer = Marshal.AllocHGlobal(32),
                .ManufacturerID = Marshal.AllocHGlobal(16),
                .Description = Marshal.AllocHGlobal(64),
                .SerialNumber = Marshal.AllocHGlobal(16)
            }
            Dim rdDlg As EeReadDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(PFt_EE_Read), GetType(EeReadDelegate)), EeReadDelegate)
            Dim status As FT_STATUS = rdDlg(FtHandle, programData)
            ee232h.Manufacturer = Marshal.PtrToStringAnsi(programData.Manufacturer)
            ee232h.ManufacturerID = Marshal.PtrToStringAnsi(programData.ManufacturerID)
            ee232h.Description = Marshal.PtrToStringAnsi(programData.Description)
            ee232h.SerialNumber = Marshal.PtrToStringAnsi(programData.SerialNumber)
            Marshal.FreeHGlobal(programData.Manufacturer)
            Marshal.FreeHGlobal(programData.ManufacturerID)
            Marshal.FreeHGlobal(programData.Description)
            Marshal.FreeHGlobal(programData.SerialNumber)
            ee232h.VendorID = programData.VendorID
            ee232h.ProductID = programData.ProductID
            ee232h.MaxPower = programData.MaxPower
            ee232h.SelfPowered = Convert.ToBoolean(programData.SelfPowered)
            ee232h.RemoteWakeup = Convert.ToBoolean(programData.RemoteWakeup)
            ee232h.PullDownEnable = Convert.ToBoolean(programData.PullDownEnableH)
            ee232h.SerNumEnable = Convert.ToBoolean(programData.SerNumEnableH)
            ee232h.ACSlowSlew = Convert.ToBoolean(programData.ACSlowSlewH)
            ee232h.ACSchmittInput = Convert.ToBoolean(programData.ACSchmittInputH)
            ee232h.ACDriveCurrent = programData.ACDriveCurrentH
            ee232h.ADSlowSlew = Convert.ToBoolean(programData.ADSlowSlewH)
            ee232h.ADSchmittInput = Convert.ToBoolean(programData.ADSchmittInputH)
            ee232h.ADDriveCurrent = programData.ADDriveCurrentH
            ee232h.Cbus0 = programData.Cbus0H
            ee232h.Cbus1 = programData.Cbus1H
            ee232h.Cbus2 = programData.Cbus2H
            ee232h.Cbus3 = programData.Cbus3H
            ee232h.Cbus4 = programData.Cbus4H
            ee232h.Cbus5 = programData.Cbus5H
            ee232h.Cbus6 = programData.Cbus6H
            ee232h.Cbus7 = programData.Cbus7H
            ee232h.Cbus8 = programData.Cbus8H
            ee232h.Cbus9 = programData.Cbus9H
            ee232h.IsFifo = Convert.ToBoolean(programData.IsFifoH)
            ee232h.IsFifoTar = Convert.ToBoolean(programData.IsFifoTarH)
            ee232h.IsFastSer = Convert.ToBoolean(programData.IsFastSerH)
            ee232h.IsFT1248 = Convert.ToBoolean(programData.IsFT1248H)
            ee232h.FT1248Cpol = Convert.ToBoolean(programData.FT1248CpolH)
            ee232h.FT1248Lsb = Convert.ToBoolean(programData.FT1248LsbH)
            ee232h.FT1248FlowControl = Convert.ToBoolean(programData.FT1248FlowControlH)
            ee232h.IsVCP = Convert.ToBoolean(programData.IsVCPH)
            ee232h.PowerSaveEnable = Convert.ToBoolean(programData.PowerSaveEnableH)
            CheckErrors(status)
            Return ee232h
        End Function

        Private Function ReadXSeriesEeprom() As FT_XSERIES_EEPROM_STRUCTURE
            Dim eeX As New FT_XSERIES_EEPROM_STRUCTURE()
            Dim header As New FT_EEPROM_HEADER() With {.deviceType = 9UI}
            Dim programData As New FT_XSERIES_DATA() With {.common = header}
            Dim strSize As Integer = Marshal.SizeOf(programData)
            Dim intPtr As IntPtr = Marshal.AllocHGlobal(strSize)
            Marshal.StructureToPtr(programData, intPtr, False)
            Dim man As Byte() = New Byte(31) {}
            Dim id As Byte() = New Byte(15) {}
            Dim descr As Byte() = New Byte(63) {}
            Dim sn As Byte() = New Byte(15) {}
            Dim rdDlg As EepromReadDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(PFt_EEPROM_Read), GetType(EepromReadDelegate)), EepromReadDelegate)
            Dim status As FT_STATUS = rdDlg(FtHandle, intPtr, CUInt(strSize), man, id, descr, sn)
            If (status = FT_STATUS.FT_OK) Then
                programData = CType(Marshal.PtrToStructure(intPtr, GetType(FT_XSERIES_DATA)), FT_XSERIES_DATA)
                Dim utf8Encoding As New UTF8Encoding()
                eeX.Manufacturer = utf8Encoding.GetString(man)
                eeX.ManufacturerID = utf8Encoding.GetString(id)
                eeX.Description = utf8Encoding.GetString(descr)
                eeX.SerialNumber = utf8Encoding.GetString(sn)
                eeX.VendorID = programData.common.VendorId
                eeX.ProductID = programData.common.ProductId
                eeX.MaxPower = programData.common.MaxPower
                eeX.SelfPowered = Convert.ToBoolean(programData.common.SelfPowered)
                eeX.RemoteWakeup = Convert.ToBoolean(programData.common.RemoteWakeup)
                eeX.SerNumEnable = Convert.ToBoolean(programData.common.SerNumEnable)
                eeX.PullDownEnable = Convert.ToBoolean(programData.common.PullDownEnable)
                eeX.Cbus0 = programData.Cbus0
                eeX.Cbus1 = programData.Cbus1
                eeX.Cbus2 = programData.Cbus2
                eeX.Cbus3 = programData.Cbus3
                eeX.Cbus4 = programData.Cbus4
                eeX.Cbus5 = programData.Cbus5
                eeX.Cbus6 = programData.Cbus6
                eeX.ACDriveCurrent = programData.ACDriveCurrent
                eeX.ACSchmittInput = programData.ACSchmittInput
                eeX.ACSlowSlew = programData.ACSlowSlew
                eeX.ADDriveCurrent = programData.ADDriveCurrent
                eeX.ADSchmittInput = programData.ADSchmittInput
                eeX.ADSlowSlew = programData.ADSlowSlew
                eeX.BCDDisableSleep = programData.BCDDisableSleep
                eeX.BCDEnable = programData.BCDEnable
                eeX.BCDForceCbusPWREN = programData.BCDForceCbusPWREN
                eeX.FT1248Cpol = programData.FT1248Cpol
                eeX.FT1248FlowControl = programData.FT1248FlowControl
                eeX.FT1248Lsb = programData.FT1248Lsb
                eeX.I2CDeviceId = programData.I2CDeviceId
                eeX.I2CDisableSchmitt = programData.I2CDisableSchmitt
                eeX.I2CSlaveAddress = programData.I2CSlaveAddress
                eeX.InvertCTS = programData.InvertCTS
                eeX.InvertDCD = programData.InvertDCD
                eeX.InvertDSR = programData.InvertDSR
                eeX.InvertDTR = programData.InvertDTR
                eeX.InvertRI = programData.InvertRI
                eeX.InvertRTS = programData.InvertRTS
                eeX.InvertRXD = programData.InvertRXD
                eeX.InvertTXD = programData.InvertTXD
                eeX.PowerSaveEnable = programData.PowerSaveEnable
                eeX.RS485EchoSuppress = programData.RS485EchoSuppress
                eeX.IsVCP = programData.DriverType
                Marshal.DestroyStructure(intPtr, GetType(FT_XSERIES_DATA))
                Marshal.FreeHGlobal(intPtr)
            End If
            CheckErrors(status)
            Return eeX
        End Function

#End Region '/ЧТЕНИЕ ППЗУ

#Region "ЗАПИСЬ ППЗУ"

        ''' <summary>
        ''' Стирает ПЗУ устройства.
        ''' </summary>
        ''' <remarks>
        ''' Для устройств FT232R и FT245R функция недоступна.
        ''' </remarks>
        Public Sub EraseEeprom()
            If (DeviceType = FT_DEVICE.FT_DEVICE_232R) Then
                CheckErrors(FT_STATUS.FT_INVALID_PARAMETER, FT_ERROR.FT_INCORRECT_DEVICE)
            End If
            Dim dlg As EraseEeDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EraseEE, IntPtr), GetType(EraseEeDelegate)), EraseEeDelegate)
            Dim status As FT_STATUS = dlg(FtHandle)
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Записывает в ПЗУ по заданному адресу <paramref name="address"/> слово данных <paramref name="value"/>.
        ''' </summary>
        Public Sub WriteEeprom(address As UInteger, value As UShort)
            Dim dlg As WriteEeDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_WriteEE, IntPtr), GetType(WriteEeDelegate)), WriteEeDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, address, value)
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Записывает заданную конфигурацию <paramref name="data"/> в ПЗУ микросхемы FTDI.
        ''' </summary>
        Public Sub WriteEeprom(data As FT_EEPROM_DATA)
            Select Case DeviceType
                Case FT_DEVICE.FT_DEVICE_BM
                    WriteFt232bEeprom(CType(data, FT232B_EEPROM_STRUCTURE))
                Case FT_DEVICE.FT_DEVICE_2232
                    WriteFt2232Eeprom(CType(data, FT2232_EEPROM_STRUCTURE))
                Case FT_DEVICE.FT_DEVICE_232R
                    WriteFt232rEeprom(CType(data, FT232R_EEPROM_STRUCTURE))
                Case FT_DEVICE.FT_DEVICE_2232H
                    WriteFt2232hEeprom(CType(data, FT2232H_EEPROM_STRUCTURE))
                Case FT_DEVICE.FT_DEVICE_4232H
                    WriteFt4232hEeprom(CType(data, FT4232H_EEPROM_STRUCTURE))
                Case FT_DEVICE.FT_DEVICE_232H
                    WriteFt232hEeprom(CType(data, FT232H_EEPROM_STRUCTURE))
                Case FT_DEVICE.FT_DEVICE_X_SERIES
                    WriteXSeriesEeprom(CType(data, FT_XSERIES_EEPROM_STRUCTURE))
                Case Else
                    CheckErrors(FT_STATUS.FT_OTHER_ERROR, FT_ERROR.FT_INCORRECT_DEVICE)
            End Select
        End Sub

        Private Sub WriteFt232bEeprom(ee232b As FT232B_EEPROM_STRUCTURE)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If (ee232b.VendorID = 0US) OrElse (ee232b.ProductID = 0US) Then
                CheckErrors(FT_STATUS.FT_INVALID_PARAMETER)
            End If
            If (ee232b.Manufacturer.Length > 32) Then
                ee232b.Manufacturer = ee232b.Manufacturer.Substring(0, 32)
            End If
            If (ee232b.ManufacturerID.Length > 16) Then
                ee232b.ManufacturerID = ee232b.ManufacturerID.Substring(0, 16)
            End If
            If (ee232b.Description.Length > 64) Then
                ee232b.Description = ee232b.Description.Substring(0, 64)
            End If
            If (ee232b.SerialNumber.Length > 16) Then
                ee232b.SerialNumber = ee232b.SerialNumber.Substring(0, 16)
            End If
            Dim programData As New FT_PROGRAM_DATA() With {
                .Signature1 = 0UI,
                .Signature2 = UInteger.MaxValue,
                .Version = 2UI,
                .Manufacturer = Marshal.AllocHGlobal(32),
                .ManufacturerID = Marshal.AllocHGlobal(16),
                .Description = Marshal.AllocHGlobal(64),
                .SerialNumber = Marshal.AllocHGlobal(16),
                .VendorID = ee232b.VendorID,
                .ProductID = ee232b.ProductID,
                .MaxPower = ee232b.MaxPower,
                .SelfPowered = Convert.ToUInt16(ee232b.SelfPowered),
                .RemoteWakeup = Convert.ToUInt16(ee232b.RemoteWakeup),
                .Rev4 = Convert.ToByte(True),
                .PullDownEnable = Convert.ToByte(ee232b.PullDownEnable),
                .SerNumEnable = Convert.ToByte(ee232b.SerNumEnable),
                .USBVersionEnable = Convert.ToByte(ee232b.USBVersionEnable),
                .USBVersion = ee232b.USBVersion
            }
            programData.Manufacturer = Marshal.StringToHGlobalAnsi(ee232b.Manufacturer)
            programData.ManufacturerID = Marshal.StringToHGlobalAnsi(ee232b.ManufacturerID)
            programData.Description = Marshal.StringToHGlobalAnsi(ee232b.Description)
            programData.SerialNumber = Marshal.StringToHGlobalAnsi(ee232b.SerialNumber)
            Dim prgDlg As EeProgramDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EE_Program, IntPtr), GetType(EeProgramDelegate)), EeProgramDelegate)
            status = prgDlg(FtHandle, programData)
            Marshal.FreeHGlobal(programData.Manufacturer)
            Marshal.FreeHGlobal(programData.ManufacturerID)
            Marshal.FreeHGlobal(programData.Description)
            Marshal.FreeHGlobal(programData.SerialNumber)
            CheckErrors(status)
        End Sub

        Private Sub WriteFt2232Eeprom(ee2232 As FT2232_EEPROM_STRUCTURE)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If (ee2232.VendorID = 0US) OrElse (ee2232.ProductID = 0US) Then
                CheckErrors(FT_STATUS.FT_INVALID_PARAMETER)
            End If
            If (ee2232.Manufacturer.Length > 32) Then
                ee2232.Manufacturer = ee2232.Manufacturer.Substring(0, 32)
            End If
            If (ee2232.ManufacturerID.Length > 16) Then
                ee2232.ManufacturerID = ee2232.ManufacturerID.Substring(0, 16)
            End If
            If (ee2232.Description.Length > 64) Then
                ee2232.Description = ee2232.Description.Substring(0, 64)
            End If
            If (ee2232.SerialNumber.Length > 16) Then
                ee2232.SerialNumber = ee2232.SerialNumber.Substring(0, 16)
            End If
            Dim programData As New FT_PROGRAM_DATA() With {
                .Signature1 = 0UI,
                .Signature2 = UInteger.MaxValue,
                .Version = 2UI,
                .Manufacturer = Marshal.AllocHGlobal(32),
                .ManufacturerID = Marshal.AllocHGlobal(16),
                .Description = Marshal.AllocHGlobal(64),
                .SerialNumber = Marshal.AllocHGlobal(16),
                .VendorID = ee2232.VendorID,
                .ProductID = ee2232.ProductID,
                .MaxPower = ee2232.MaxPower,
                .SelfPowered = Convert.ToUInt16(ee2232.SelfPowered),
                .RemoteWakeup = Convert.ToUInt16(ee2232.RemoteWakeup),
                .Rev5 = Convert.ToByte(True),
                .PullDownEnable5 = Convert.ToByte(ee2232.PullDownEnable),
                .SerNumEnable5 = Convert.ToByte(ee2232.SerNumEnable),
                .USBVersionEnable5 = Convert.ToByte(ee2232.USBVersionEnable),
                .USBVersion5 = ee2232.USBVersion,
                .AIsHighCurrent = Convert.ToByte(ee2232.AIsHighCurrent),
                .BIsHighCurrent = Convert.ToByte(ee2232.BIsHighCurrent),
                .IFAIsFifo = Convert.ToByte(ee2232.IFAIsFifo),
                .IFAIsFifoTar = Convert.ToByte(ee2232.IFAIsFifoTar),
                .IFAIsFastSer = Convert.ToByte(ee2232.IFAIsFastSer),
                .AIsVCP = Convert.ToByte(ee2232.AIsVCP),
                .IFBIsFifo = Convert.ToByte(ee2232.IFBIsFifo),
                .IFBIsFifoTar = Convert.ToByte(ee2232.IFBIsFifoTar),
                .IFBIsFastSer = Convert.ToByte(ee2232.IFBIsFastSer),
                .BIsVCP = Convert.ToByte(ee2232.BIsVCP)
            }
            programData.Manufacturer = Marshal.StringToHGlobalAnsi(ee2232.Manufacturer)
            programData.ManufacturerID = Marshal.StringToHGlobalAnsi(ee2232.ManufacturerID)
            programData.Description = Marshal.StringToHGlobalAnsi(ee2232.Description)
            programData.SerialNumber = Marshal.StringToHGlobalAnsi(ee2232.SerialNumber)
            Dim prgDlg As EeProgramDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EE_Program, IntPtr), GetType(EeProgramDelegate)), EeProgramDelegate)
            status = prgDlg(FtHandle, programData)
            Marshal.FreeHGlobal(programData.Manufacturer)
            Marshal.FreeHGlobal(programData.ManufacturerID)
            Marshal.FreeHGlobal(programData.Description)
            Marshal.FreeHGlobal(programData.SerialNumber)
            CheckErrors(status)
        End Sub

        Private Sub WriteFt232rEeprom(ee232r As FT232R_EEPROM_STRUCTURE)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If (ee232r.VendorID = 0US) OrElse (ee232r.ProductID = 0US) Then
                CheckErrors(FT_STATUS.FT_INVALID_PARAMETER)
            End If
            If (ee232r.Manufacturer.Length > 32) Then
                ee232r.Manufacturer = ee232r.Manufacturer.Substring(0, 32)
            End If
            If (ee232r.ManufacturerID.Length > 16) Then
                ee232r.ManufacturerID = ee232r.ManufacturerID.Substring(0, 16)
            End If
            If (ee232r.Description.Length > 64) Then
                ee232r.Description = ee232r.Description.Substring(0, 64)
            End If
            If (ee232r.SerialNumber.Length > 16) Then
                ee232r.SerialNumber = ee232r.SerialNumber.Substring(0, 16)
            End If
            Dim programData As New FT_PROGRAM_DATA() With {
                .Signature1 = 0UI,
                .Signature2 = UInteger.MaxValue,
                .Version = 2UI,
                .Manufacturer = Marshal.AllocHGlobal(32),
                .ManufacturerID = Marshal.AllocHGlobal(16),
                .Description = Marshal.AllocHGlobal(64),
                .SerialNumber = Marshal.AllocHGlobal(16),
                .VendorID = ee232r.VendorID,
                .ProductID = ee232r.ProductID,
                .MaxPower = ee232r.MaxPower,
                .SelfPowered = Convert.ToUInt16(ee232r.SelfPowered),
                .RemoteWakeup = Convert.ToUInt16(ee232r.RemoteWakeup),
                .PullDownEnableR = Convert.ToByte(ee232r.PullDownEnable),
                .SerNumEnableR = Convert.ToByte(ee232r.SerNumEnable),
                .UseExtOsc = Convert.ToByte(ee232r.UseExtOsc),
                .HighDriveIOs = Convert.ToByte(ee232r.HighDriveIOs),
                .EndpointSize = 64,
                .InvertTXD = Convert.ToByte(ee232r.InvertTXD),
                .InvertRXD = Convert.ToByte(ee232r.InvertRXD),
                .InvertRTS = Convert.ToByte(ee232r.InvertRTS),
                .InvertCTS = Convert.ToByte(ee232r.InvertCTS),
                .InvertDTR = Convert.ToByte(ee232r.InvertDTR),
                .InvertDSR = Convert.ToByte(ee232r.InvertDSR),
                .InvertDCD = Convert.ToByte(ee232r.InvertDCD),
                .InvertRI = Convert.ToByte(ee232r.InvertRI),
                .Cbus0 = ee232r.Cbus0,
                .Cbus1 = ee232r.Cbus1,
                .Cbus2 = ee232r.Cbus2,
                .Cbus3 = ee232r.Cbus3,
                .Cbus4 = ee232r.Cbus4,
                .RIsD2XX = Convert.ToByte(ee232r.RIsD2XX)
            }
            programData.Manufacturer = Marshal.StringToHGlobalAnsi(ee232r.Manufacturer)
            programData.ManufacturerID = Marshal.StringToHGlobalAnsi(ee232r.ManufacturerID)
            programData.Description = Marshal.StringToHGlobalAnsi(ee232r.Description)
            programData.SerialNumber = Marshal.StringToHGlobalAnsi(ee232r.SerialNumber)
            Dim prgDlg As EeProgramDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EE_Program, IntPtr), GetType(EeProgramDelegate)), EeProgramDelegate)
            status = prgDlg(FtHandle, programData)
            Marshal.FreeHGlobal(programData.Manufacturer)
            Marshal.FreeHGlobal(programData.ManufacturerID)
            Marshal.FreeHGlobal(programData.Description)
            Marshal.FreeHGlobal(programData.SerialNumber)
            CheckErrors(status)
        End Sub

        Private Sub WriteFt2232hEeprom(ee2232h As FT2232H_EEPROM_STRUCTURE)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If (ee2232h.VendorID = 0US) OrElse (ee2232h.ProductID = 0US) Then
                CheckErrors(FT_STATUS.FT_INVALID_PARAMETER)
            End If
            If (ee2232h.Manufacturer.Length > 32) Then
                ee2232h.Manufacturer = ee2232h.Manufacturer.Substring(0, 32)
            End If
            If (ee2232h.ManufacturerID.Length > 16) Then
                ee2232h.ManufacturerID = ee2232h.ManufacturerID.Substring(0, 16)
            End If
            If (ee2232h.Description.Length > 64) Then
                ee2232h.Description = ee2232h.Description.Substring(0, 64)
            End If
            If (ee2232h.SerialNumber.Length > 16) Then
                ee2232h.SerialNumber = ee2232h.SerialNumber.Substring(0, 16)
            End If
            Dim programData As New FT_PROGRAM_DATA() With {
                .Signature1 = 0UI,
                .Signature2 = UInteger.MaxValue,
                .Version = 3UI,
                .Manufacturer = Marshal.AllocHGlobal(32),
                .ManufacturerID = Marshal.AllocHGlobal(16),
                .Description = Marshal.AllocHGlobal(64),
                .SerialNumber = Marshal.AllocHGlobal(16),
                .VendorID = ee2232h.VendorID,
                .ProductID = ee2232h.ProductID,
                .MaxPower = ee2232h.MaxPower,
                .SelfPowered = Convert.ToUInt16(ee2232h.SelfPowered),
                .RemoteWakeup = Convert.ToUInt16(ee2232h.RemoteWakeup),
                .PullDownEnable7 = Convert.ToByte(ee2232h.PullDownEnable),
                .SerNumEnable7 = Convert.ToByte(ee2232h.SerNumEnable),
                .ALSlowSlew = Convert.ToByte(ee2232h.ALSlowSlew),
                .ALSchmittInput = Convert.ToByte(ee2232h.ALSchmittInput),
                .ALDriveCurrent = ee2232h.ALDriveCurrent,
                .AHSlowSlew = Convert.ToByte(ee2232h.AHSlowSlew),
                .AHSchmittInput = Convert.ToByte(ee2232h.AHSchmittInput),
                .AHDriveCurrent = ee2232h.AHDriveCurrent,
                .BLSlowSlew = Convert.ToByte(ee2232h.BLSlowSlew),
                .BLSchmittInput = Convert.ToByte(ee2232h.BLSchmittInput),
                .BLDriveCurrent = ee2232h.BLDriveCurrent,
                .BHSlowSlew = Convert.ToByte(ee2232h.BHSlowSlew),
                .BHSchmittInput = Convert.ToByte(ee2232h.BHSchmittInput),
                .BHDriveCurrent = ee2232h.BHDriveCurrent,
                .IFAIsFifo7 = Convert.ToByte(ee2232h.IFAIsFifo),
                .IFAIsFifoTar7 = Convert.ToByte(ee2232h.IFAIsFifoTar),
                .IFAIsFastSer7 = Convert.ToByte(ee2232h.IFAIsFastSer),
                .AIsVCP7 = Convert.ToByte(ee2232h.AIsVCP),
                .IFBIsFifo7 = Convert.ToByte(ee2232h.IFBIsFifo),
                .IFBIsFifoTar7 = Convert.ToByte(ee2232h.IFBIsFifoTar),
                .IFBIsFastSer7 = Convert.ToByte(ee2232h.IFBIsFastSer),
                .BIsVCP7 = Convert.ToByte(ee2232h.BIsVCP),
                .PowerSaveEnable = Convert.ToByte(ee2232h.PowerSaveEnable)
            }
            programData.Manufacturer = Marshal.StringToHGlobalAnsi(ee2232h.Manufacturer)
            programData.ManufacturerID = Marshal.StringToHGlobalAnsi(ee2232h.ManufacturerID)
            programData.Description = Marshal.StringToHGlobalAnsi(ee2232h.Description)
            programData.SerialNumber = Marshal.StringToHGlobalAnsi(ee2232h.SerialNumber)
            Dim prgDlg As EeProgramDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EE_Program, IntPtr), GetType(EeProgramDelegate)), EeProgramDelegate)
            status = prgDlg(FtHandle, programData)
            Marshal.FreeHGlobal(programData.Manufacturer)
            Marshal.FreeHGlobal(programData.ManufacturerID)
            Marshal.FreeHGlobal(programData.Description)
            Marshal.FreeHGlobal(programData.SerialNumber)
            CheckErrors(status)
        End Sub

        Private Sub WriteFt4232hEeprom(ee4232h As FT4232H_EEPROM_STRUCTURE)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If (ee4232h.VendorID = 0US) OrElse (ee4232h.ProductID = 0US) Then
                CheckErrors(FT_STATUS.FT_INVALID_PARAMETER)
            End If
            If (ee4232h.Manufacturer.Length > 32) Then
                ee4232h.Manufacturer = ee4232h.Manufacturer.Substring(0, 32)
            End If
            If (ee4232h.ManufacturerID.Length > 16) Then
                ee4232h.ManufacturerID = ee4232h.ManufacturerID.Substring(0, 16)
            End If
            If (ee4232h.Description.Length > 64) Then
                ee4232h.Description = ee4232h.Description.Substring(0, 64)
            End If
            If (ee4232h.SerialNumber.Length > 16) Then
                ee4232h.SerialNumber = ee4232h.SerialNumber.Substring(0, 16)
            End If
            Dim programData As New FT_PROGRAM_DATA() With {
                .Signature1 = 0UI,
                .Signature2 = UInteger.MaxValue,
                .Version = 4UI,
                .Manufacturer = Marshal.AllocHGlobal(32),
                .ManufacturerID = Marshal.AllocHGlobal(16),
                .Description = Marshal.AllocHGlobal(64),
                .SerialNumber = Marshal.AllocHGlobal(16),
                .VendorID = ee4232h.VendorID,
                .ProductID = ee4232h.ProductID,
                .MaxPower = ee4232h.MaxPower,
                .SelfPowered = Convert.ToUInt16(ee4232h.SelfPowered),
                .RemoteWakeup = Convert.ToUInt16(ee4232h.RemoteWakeup),
                .PullDownEnable8 = Convert.ToByte(ee4232h.PullDownEnable),
                .SerNumEnable8 = Convert.ToByte(ee4232h.SerNumEnable),
                .ASlowSlew = Convert.ToByte(ee4232h.ASlowSlew),
                .ASchmittInput = Convert.ToByte(ee4232h.ASchmittInput),
                .ADriveCurrent = ee4232h.ADriveCurrent,
                .BSlowSlew = Convert.ToByte(ee4232h.BSlowSlew),
                .BSchmittInput = Convert.ToByte(ee4232h.BSchmittInput),
                .BDriveCurrent = ee4232h.BDriveCurrent,
                .CSlowSlew = Convert.ToByte(ee4232h.CSlowSlew),
                .CSchmittInput = Convert.ToByte(ee4232h.CSchmittInput),
                .CDriveCurrent = ee4232h.CDriveCurrent,
                .DSlowSlew = Convert.ToByte(ee4232h.DSlowSlew),
                .DSchmittInput = Convert.ToByte(ee4232h.DSchmittInput),
                .DDriveCurrent = ee4232h.DDriveCurrent,
                .ARIIsTXDEN = Convert.ToByte(ee4232h.ARIIsTXDEN),
                .BRIIsTXDEN = Convert.ToByte(ee4232h.BRIIsTXDEN),
                .CRIIsTXDEN = Convert.ToByte(ee4232h.CRIIsTXDEN),
                .DRIIsTXDEN = Convert.ToByte(ee4232h.DRIIsTXDEN),
                .AIsVCP8 = Convert.ToByte(ee4232h.AIsVCP),
                .BIsVCP8 = Convert.ToByte(ee4232h.BIsVCP),
                .CIsVCP8 = Convert.ToByte(ee4232h.CIsVCP),
                .DIsVCP8 = Convert.ToByte(ee4232h.DIsVCP)
            }
            programData.Manufacturer = Marshal.StringToHGlobalAnsi(ee4232h.Manufacturer)
            programData.ManufacturerID = Marshal.StringToHGlobalAnsi(ee4232h.ManufacturerID)
            programData.Description = Marshal.StringToHGlobalAnsi(ee4232h.Description)
            programData.SerialNumber = Marshal.StringToHGlobalAnsi(ee4232h.SerialNumber)
            Dim prgDlg As EeProgramDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EE_Program, IntPtr), GetType(EeProgramDelegate)), EeProgramDelegate)
            status = prgDlg(FtHandle, programData)
            Marshal.FreeHGlobal(programData.Manufacturer)
            Marshal.FreeHGlobal(programData.ManufacturerID)
            Marshal.FreeHGlobal(programData.Description)
            Marshal.FreeHGlobal(programData.SerialNumber)
            CheckErrors(status)
        End Sub

        Private Sub WriteFt232hEeprom(ee232h As FT232H_EEPROM_STRUCTURE)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If (ee232h.VendorID = 0US) OrElse (ee232h.ProductID = 0US) Then
                CheckErrors(FT_STATUS.FT_INVALID_PARAMETER)
            End If
            If (ee232h.Manufacturer.Length > 32) Then
                ee232h.Manufacturer = ee232h.Manufacturer.Substring(0, 32)
            End If
            If (ee232h.ManufacturerID.Length > 16) Then
                ee232h.ManufacturerID = ee232h.ManufacturerID.Substring(0, 16)
            End If
            If (ee232h.Description.Length > 64) Then
                ee232h.Description = ee232h.Description.Substring(0, 64)
            End If
            If (ee232h.SerialNumber.Length > 16) Then
                ee232h.SerialNumber = ee232h.SerialNumber.Substring(0, 16)
            End If
            Dim programData As New FT_PROGRAM_DATA() With {
                .Signature1 = 0UI,
                .Signature2 = UInteger.MaxValue,
                .Version = 5UI,
                .Manufacturer = Marshal.AllocHGlobal(32),
                .ManufacturerID = Marshal.AllocHGlobal(16),
                .Description = Marshal.AllocHGlobal(64),
                .SerialNumber = Marshal.AllocHGlobal(16),
                .VendorID = ee232h.VendorID,
                .ProductID = ee232h.ProductID,
                .MaxPower = ee232h.MaxPower,
                .SelfPowered = Convert.ToUInt16(ee232h.SelfPowered),
                .RemoteWakeup = Convert.ToUInt16(ee232h.RemoteWakeup),
                .PullDownEnableH = Convert.ToByte(ee232h.PullDownEnable),
                .SerNumEnableH = Convert.ToByte(ee232h.SerNumEnable),
                .ACSlowSlewH = Convert.ToByte(ee232h.ACSlowSlew),
                .ACSchmittInputH = Convert.ToByte(ee232h.ACSchmittInput),
                .ACDriveCurrentH = ee232h.ACDriveCurrent,
                .ADSlowSlewH = Convert.ToByte(ee232h.ADSlowSlew),
                .ADSchmittInputH = Convert.ToByte(ee232h.ADSchmittInput),
                .ADDriveCurrentH = ee232h.ADDriveCurrent,
                .Cbus0H = Convert.ToByte(ee232h.Cbus0),
                .Cbus1H = Convert.ToByte(ee232h.Cbus1),
                .Cbus2H = Convert.ToByte(ee232h.Cbus2),
                .Cbus3H = Convert.ToByte(ee232h.Cbus3),
                .Cbus4H = Convert.ToByte(ee232h.Cbus4),
                .Cbus5H = Convert.ToByte(ee232h.Cbus5),
                .Cbus6H = Convert.ToByte(ee232h.Cbus6),
                .Cbus7H = Convert.ToByte(ee232h.Cbus7),
                .Cbus8H = Convert.ToByte(ee232h.Cbus8),
                .Cbus9H = Convert.ToByte(ee232h.Cbus9),
                .IsFifoH = Convert.ToByte(ee232h.IsFifo),
                .IsFifoTarH = Convert.ToByte(ee232h.IsFifoTar),
                .IsFastSerH = Convert.ToByte(ee232h.IsFastSer),
                .IsFT1248H = Convert.ToByte(ee232h.IsFT1248),
                .FT1248CpolH = Convert.ToByte(ee232h.FT1248Cpol),
                .FT1248LsbH = Convert.ToByte(ee232h.FT1248Lsb),
                .FT1248FlowControlH = Convert.ToByte(ee232h.FT1248FlowControl),
                .IsVCPH = Convert.ToByte(ee232h.IsVCP),
                .PowerSaveEnableH = Convert.ToByte(ee232h.PowerSaveEnable)
            }
            programData.Manufacturer = Marshal.StringToHGlobalAnsi(ee232h.Manufacturer)
            programData.ManufacturerID = Marshal.StringToHGlobalAnsi(ee232h.ManufacturerID)
            programData.Description = Marshal.StringToHGlobalAnsi(ee232h.Description)
            programData.SerialNumber = Marshal.StringToHGlobalAnsi(ee232h.SerialNumber)
            Dim prgDlg As EeProgramDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EE_Program, IntPtr), GetType(EeProgramDelegate)), EeProgramDelegate)
            status = prgDlg(FtHandle, programData)
            Marshal.FreeHGlobal(programData.Manufacturer)
            Marshal.FreeHGlobal(programData.ManufacturerID)
            Marshal.FreeHGlobal(programData.Description)
            Marshal.FreeHGlobal(programData.SerialNumber)
            CheckErrors(status)
        End Sub

        Private Sub WriteXSeriesEeprom(eeX As FT_XSERIES_EEPROM_STRUCTURE)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If (eeX.VendorID = 0US) OrElse (eeX.ProductID = 0US) Then
                CheckErrors(FT_STATUS.FT_INVALID_PARAMETER)
            End If
            If (eeX.Manufacturer.Length > 32) Then
                eeX.Manufacturer = eeX.Manufacturer.Substring(0, 32)
            End If
            If (eeX.ManufacturerID.Length > 16) Then
                eeX.ManufacturerID = eeX.ManufacturerID.Substring(0, 16)
            End If
            If (eeX.Description.Length > 64) Then
                eeX.Description = eeX.Description.Substring(0, 64)
            End If
            If (eeX.SerialNumber.Length > 16) Then
                eeX.SerialNumber = eeX.SerialNumber.Substring(0, 16)
            End If
            Dim programData As New FT_XSERIES_DATA() With {
                .Cbus0 = eeX.Cbus0,
                .Cbus1 = eeX.Cbus1,
                .Cbus2 = eeX.Cbus2,
                .Cbus3 = eeX.Cbus3,
                .Cbus4 = eeX.Cbus4,
                .Cbus5 = eeX.Cbus5,
                .Cbus6 = eeX.Cbus6,
                .ACDriveCurrent = eeX.ACDriveCurrent,
                .ACSchmittInput = eeX.ACSchmittInput,
                .ACSlowSlew = eeX.ACSlowSlew,
                .ADDriveCurrent = eeX.ADDriveCurrent,
                .ADSchmittInput = eeX.ADSchmittInput,
                .ADSlowSlew = eeX.ADSlowSlew,
                .BCDDisableSleep = eeX.BCDDisableSleep,
                .BCDEnable = eeX.BCDEnable,
                .BCDForceCbusPWREN = eeX.BCDForceCbusPWREN,
                .FT1248Cpol = eeX.FT1248Cpol,
                .FT1248FlowControl = eeX.FT1248FlowControl,
                .FT1248Lsb = eeX.FT1248Lsb,
                .I2CDeviceId = eeX.I2CDeviceId,
                .I2CDisableSchmitt = eeX.I2CDisableSchmitt,
                .I2CSlaveAddress = eeX.I2CSlaveAddress,
                .InvertCTS = eeX.InvertCTS,
                .InvertDCD = eeX.InvertDCD,
                .InvertDSR = eeX.InvertDSR,
                .InvertDTR = eeX.InvertDTR,
                .InvertRI = eeX.InvertRI,
                .InvertRTS = eeX.InvertRTS,
                .InvertRXD = eeX.InvertRXD,
                .InvertTXD = eeX.InvertTXD,
                .PowerSaveEnable = eeX.PowerSaveEnable,
                .RS485EchoSuppress = eeX.RS485EchoSuppress,
                .DriverType = eeX.IsVCP
            }
            programData.common.deviceType = 9UI
            programData.common.VendorId = eeX.VendorID
            programData.common.ProductId = eeX.ProductID
            programData.common.MaxPower = eeX.MaxPower
            programData.common.SelfPowered = Convert.ToByte(eeX.SelfPowered)
            programData.common.RemoteWakeup = Convert.ToByte(eeX.RemoteWakeup)
            programData.common.SerNumEnable = Convert.ToByte(eeX.SerNumEnable)
            programData.common.PullDownEnable = Convert.ToByte(eeX.PullDownEnable)
            Dim utf8Encoding As UTF8Encoding = New UTF8Encoding()
            Dim manufacturer As Byte() = utf8Encoding.GetBytes(eeX.Manufacturer)
            Dim manufacturerID As Byte() = utf8Encoding.GetBytes(eeX.ManufacturerID)
            Dim description As Byte() = utf8Encoding.GetBytes(eeX.Description)
            Dim serialnumber As Byte() = utf8Encoding.GetBytes(eeX.SerialNumber)
            Dim programDataSize As Integer = Marshal.SizeOf(programData)
            Dim ptr As IntPtr = Marshal.AllocHGlobal(programDataSize)
            Marshal.StructureToPtr(programData, ptr, False)
            Dim prgDlg As EepromProgramDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Me.PFt_EEPROM_Program, IntPtr), GetType(EepromProgramDelegate)), EepromProgramDelegate)
            status = prgDlg(FtHandle, ptr, CUInt(programDataSize), manufacturer, manufacturerID, description, serialnumber)
            CheckErrors(status)
        End Sub

#End Region '/ЗАПИСЬ ППЗУ

#Region "ПОЛЬЗОВАТЕЛЬСКАЯ ОБЛАСТЬ ППЗУ FTDI"

        ''' <summary>
        ''' Размер пользовательской области ПЗУ в байтах.
        ''' </summary>
        Public ReadOnly Property UserAreaSize As Integer
            Get
                If (_UserAreaSize = NOT_INIT) Then
                    Dim uaSize As UInteger
                    Dim dlg As EeUaSizeDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EE_UASize, IntPtr), GetType(EeUaSizeDelegate)), EeUaSizeDelegate)
                    Dim status As FT_STATUS = dlg(FtHandle, uaSize)
                    CheckErrors(status)
                    _UserAreaSize = CInt(uaSize)
                End If
                Return _UserAreaSize
            End Get
        End Property
        Private _UserAreaSize As Integer = NOT_INIT

        ''' <summary>
        ''' Читает данные пользовательской области ПЗУ.
        ''' </summary>
        Public Function ReadUserArea() As Byte()
            Dim userAreaDataBuffer As Byte() = New Byte() {}
            Dim bufSize As UInteger = 0UI
            Dim dlg As EeUaSizeDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EE_UASize, IntPtr), GetType(EeUaSizeDelegate)), EeUaSizeDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, bufSize)
            ReDim userAreaDataBuffer(CInt(bufSize - 1))
            If (userAreaDataBuffer.Length >= bufSize) Then
                Dim numBytesWereRead As UInteger = 0
                Dim rdDlg As EeUaReadDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EE_UARead, IntPtr), GetType(EeUaReadDelegate)), EeUaReadDelegate)
                status = status Or rdDlg(FtHandle, userAreaDataBuffer, userAreaDataBuffer.Length, numBytesWereRead)
            End If
            CheckErrors(status)
            Return userAreaDataBuffer
        End Function

        ''' <summary>
        ''' Записывает данные в пользовательскую область ПЗУ.
        ''' </summary>
        Public Sub WriteUserArea(data As Byte())
            Dim bufSize As UInteger = 0UI
            Dim dlg As EeUaSizeDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EE_UASize, IntPtr), GetType(EeUaSizeDelegate)), EeUaSizeDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, bufSize)
            If (data.Length <= bufSize) Then
                Dim wDlg As EeUaWriteDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EE_UAWrite, IntPtr), GetType(EeUaWriteDelegate)), EeUaWriteDelegate)
                status = status Or wDlg(FtHandle, data, data.Length)
            End If
            CheckErrors(status)
        End Sub

#End Region '/ПОЛЬЗОВАТЕЛЬСКАЯ ОБЛАСТЬ ППЗУ FTDI

#Region "СТРУКТУРЫ ОБЪЕКТОВ В ППЗУ FTDI"

        <StructLayout(LayoutKind.Sequential, Pack:=4)>
        Private Structure FT_PROGRAM_DATA
            Public Signature1 As UInteger
            Public Signature2 As UInteger
            Public Version As UInteger
            Public VendorID As UShort
            Public ProductID As UShort
            Public Manufacturer As IntPtr
            Public ManufacturerID As IntPtr
            Public Description As IntPtr
            Public SerialNumber As IntPtr
            Public MaxPower As UShort
            Public PnP As UShort
            Public SelfPowered As UShort
            Public RemoteWakeup As UShort
            Public Rev4 As Byte
            Public IsoIn As Byte
            Public IsoOut As Byte
            Public PullDownEnable As Byte
            Public SerNumEnable As Byte
            Public USBVersionEnable As Byte
            Public USBVersion As UShort
            Public Rev5 As Byte
            Public IsoInA As Byte
            Public IsoInB As Byte
            Public IsoOutA As Byte
            Public IsoOutB As Byte
            Public PullDownEnable5 As Byte
            Public SerNumEnable5 As Byte
            Public USBVersionEnable5 As Byte
            Public USBVersion5 As UShort
            Public AIsHighCurrent As Byte
            Public BIsHighCurrent As Byte
            Public IFAIsFifo As Byte
            Public IFAIsFifoTar As Byte
            Public IFAIsFastSer As Byte
            Public AIsVCP As Byte
            Public IFBIsFifo As Byte
            Public IFBIsFifoTar As Byte
            Public IFBIsFastSer As Byte
            Public BIsVCP As Byte
            Public UseExtOsc As Byte
            Public HighDriveIOs As Byte
            Public EndpointSize As Byte
            Public PullDownEnableR As Byte
            Public SerNumEnableR As Byte
            Public InvertTXD As Byte
            Public InvertRXD As Byte
            Public InvertRTS As Byte
            Public InvertCTS As Byte
            Public InvertDTR As Byte
            Public InvertDSR As Byte
            Public InvertDCD As Byte
            Public InvertRI As Byte
            Public Cbus0 As Byte
            Public Cbus1 As Byte
            Public Cbus2 As Byte
            Public Cbus3 As Byte
            Public Cbus4 As Byte
            Public RIsD2XX As Byte
            Public PullDownEnable7 As Byte
            Public SerNumEnable7 As Byte
            Public ALSlowSlew As Byte
            Public ALSchmittInput As Byte
            Public ALDriveCurrent As FT_DRIVE_CURRENT
            Public AHSlowSlew As Byte
            Public AHSchmittInput As Byte
            Public AHDriveCurrent As FT_DRIVE_CURRENT
            Public BLSlowSlew As Byte
            Public BLSchmittInput As Byte
            Public BLDriveCurrent As FT_DRIVE_CURRENT
            Public BHSlowSlew As Byte
            Public BHSchmittInput As Byte
            Public BHDriveCurrent As FT_DRIVE_CURRENT
            Public IFAIsFifo7 As Byte
            Public IFAIsFifoTar7 As Byte
            Public IFAIsFastSer7 As Byte
            Public AIsVCP7 As Byte
            Public IFBIsFifo7 As Byte
            Public IFBIsFifoTar7 As Byte
            Public IFBIsFastSer7 As Byte
            Public BIsVCP7 As Byte
            Public PowerSaveEnable As Byte
            Public PullDownEnable8 As Byte
            Public SerNumEnable8 As Byte
            Public ASlowSlew As Byte
            Public ASchmittInput As Byte
            Public ADriveCurrent As FT_DRIVE_CURRENT
            Public BSlowSlew As Byte
            Public BSchmittInput As Byte
            Public BDriveCurrent As FT_DRIVE_CURRENT
            Public CSlowSlew As Byte
            Public CSchmittInput As Byte
            Public CDriveCurrent As FT_DRIVE_CURRENT
            Public DSlowSlew As Byte
            Public DSchmittInput As Byte
            Public DDriveCurrent As FT_DRIVE_CURRENT
            Public ARIIsTXDEN As Byte
            Public BRIIsTXDEN As Byte
            Public CRIIsTXDEN As Byte
            Public DRIIsTXDEN As Byte
            Public AIsVCP8 As Byte
            Public BIsVCP8 As Byte
            Public CIsVCP8 As Byte
            Public DIsVCP8 As Byte
            Public PullDownEnableH As Byte
            Public SerNumEnableH As Byte
            Public ACSlowSlewH As Byte
            Public ACSchmittInputH As Byte
            Public ACDriveCurrentH As FT_DRIVE_CURRENT
            Public ADSlowSlewH As Byte
            Public ADSchmittInputH As Byte
            Public ADDriveCurrentH As FT_DRIVE_CURRENT
            Public Cbus0H As Byte
            Public Cbus1H As Byte
            Public Cbus2H As Byte
            Public Cbus3H As Byte
            Public Cbus4H As Byte
            Public Cbus5H As Byte
            Public Cbus6H As Byte
            Public Cbus7H As Byte
            Public Cbus8H As Byte
            Public Cbus9H As Byte
            Public IsFifoH As Byte
            Public IsFifoTarH As Byte
            Public IsFastSerH As Byte
            Public IsFT1248H As Byte
            Public FT1248CpolH As Byte
            Public FT1248LsbH As Byte
            Public FT1248FlowControlH As Byte
            Public IsVCPH As Byte
            Public PowerSaveEnableH As Byte
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=4)>
        Private Structure FT_EEPROM_HEADER
            Public deviceType As UInteger
            Public VendorId As UShort
            Public ProductId As UShort
            Public SerNumEnable As Byte
            Public MaxPower As UShort
            Public SelfPowered As Byte
            Public RemoteWakeup As Byte
            Public PullDownEnable As Byte
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=4)>
        Private Structure FT_XSERIES_DATA
            Public common As FT_EEPROM_HEADER
            Public ACSlowSlew As Byte
            Public ACSchmittInput As Byte
            Public ACDriveCurrent As FT_DRIVE_CURRENT
            Public ADSlowSlew As Byte
            Public ADSchmittInput As Byte
            Public ADDriveCurrent As FT_DRIVE_CURRENT
            Public Cbus0 As Byte
            Public Cbus1 As Byte
            Public Cbus2 As Byte
            Public Cbus3 As Byte
            Public Cbus4 As Byte
            Public Cbus5 As Byte
            Public Cbus6 As Byte
            Public InvertTXD As Byte
            Public InvertRXD As Byte
            Public InvertRTS As Byte
            Public InvertCTS As Byte
            Public InvertDTR As Byte
            Public InvertDSR As Byte
            Public InvertDCD As Byte
            Public InvertRI As Byte
            Public BCDEnable As Byte
            Public BCDForceCbusPWREN As Byte
            Public BCDDisableSleep As Byte
            Public I2CSlaveAddress As UShort
            Public I2CDeviceId As UInteger
            Public I2CDisableSchmitt As Byte
            Public FT1248Cpol As Byte
            Public FT1248Lsb As Byte
            Public FT1248FlowControl As Byte
            Public RS485EchoSuppress As Byte
            Public PowerSaveEnable As Byte
            Public DriverType As Byte
        End Structure

        Public Class FT_EEPROM_DATA
            Public VendorID As UShort = &H403US
            Public ProductID As UShort = &H6001US
            Public Manufacturer As String = "FTDI"
            Public ManufacturerID As String = "FT"
            Public Description As String = "USB-Serial Converter"
            Public SerialNumber As String = ""
            Public MaxPower As UShort = 144US
            Public SelfPowered As Boolean
            Public RemoteWakeup As Boolean
        End Class

        Public Class FT232B_EEPROM_STRUCTURE : Inherits FT_EEPROM_DATA
            Public PullDownEnable As Boolean
            Public SerNumEnable As Boolean = True
            Public USBVersionEnable As Boolean = True
            Public USBVersion As UShort = 512US
        End Class

        Public Class FT2232_EEPROM_STRUCTURE : Inherits FT_EEPROM_DATA
            Public PullDownEnable As Boolean
            Public SerNumEnable As Boolean = True
            Public USBVersionEnable As Boolean = True
            Public USBVersion As UShort = 512US
            Public AIsHighCurrent As Boolean
            Public BIsHighCurrent As Boolean
            Public IFAIsFifo As Boolean
            Public IFAIsFifoTar As Boolean
            Public IFAIsFastSer As Boolean
            Public AIsVCP As Boolean = True
            Public IFBIsFifo As Boolean
            Public IFBIsFifoTar As Boolean
            Public IFBIsFastSer As Boolean
            Public BIsVCP As Boolean = True
        End Class

        Public Class FT232R_EEPROM_STRUCTURE : Inherits FT_EEPROM_DATA
            Public UseExtOsc As Boolean
            Public HighDriveIOs As Boolean
            Public EndpointSize As Byte = 64
            Public PullDownEnable As Boolean
            Public SerNumEnable As Boolean = True
            Public InvertTXD As Boolean
            Public InvertRXD As Boolean
            Public InvertRTS As Boolean
            Public InvertCTS As Boolean
            Public InvertDTR As Boolean
            Public InvertDSR As Boolean
            Public InvertDCD As Boolean
            Public InvertRI As Boolean
            Public Cbus0 As Byte = 5
            Public Cbus1 As Byte = 5
            Public Cbus2 As Byte = 5
            Public Cbus3 As Byte = 5
            Public Cbus4 As Byte = 5
            Public RIsD2XX As Boolean
        End Class

        Public Class FT2232H_EEPROM_STRUCTURE : Inherits FT_EEPROM_DATA
            Public PullDownEnable As Boolean
            Public SerNumEnable As Boolean = True
            Public ALSlowSlew As Boolean
            Public ALSchmittInput As Boolean
            Public ALDriveCurrent As FT_DRIVE_CURRENT = FT_DRIVE_CURRENT.FT_DRIVE_CURRENT_4MA
            Public AHSlowSlew As Boolean
            Public AHSchmittInput As Boolean
            Public AHDriveCurrent As FT_DRIVE_CURRENT = FT_DRIVE_CURRENT.FT_DRIVE_CURRENT_4MA
            Public BLSlowSlew As Boolean
            Public BLSchmittInput As Boolean
            Public BLDriveCurrent As FT_DRIVE_CURRENT = FT_DRIVE_CURRENT.FT_DRIVE_CURRENT_4MA
            Public BHSlowSlew As Boolean
            Public BHSchmittInput As Boolean
            Public BHDriveCurrent As FT_DRIVE_CURRENT = FT_DRIVE_CURRENT.FT_DRIVE_CURRENT_4MA
            Public IFAIsFifo As Boolean
            Public IFAIsFifoTar As Boolean
            Public IFAIsFastSer As Boolean
            Public AIsVCP As Boolean = True
            Public IFBIsFifo As Boolean
            Public IFBIsFifoTar As Boolean
            Public IFBIsFastSer As Boolean
            Public BIsVCP As Boolean = True
            Public PowerSaveEnable As Boolean
        End Class

        Public Class FT4232H_EEPROM_STRUCTURE : Inherits FT_EEPROM_DATA
            Public PullDownEnable As Boolean
            Public SerNumEnable As Boolean = True
            Public ASlowSlew As Boolean
            Public ASchmittInput As Boolean
            Public ADriveCurrent As FT_DRIVE_CURRENT = FT_DRIVE_CURRENT.FT_DRIVE_CURRENT_4MA
            Public BSlowSlew As Boolean
            Public BSchmittInput As Boolean
            Public BDriveCurrent As FT_DRIVE_CURRENT = FT_DRIVE_CURRENT.FT_DRIVE_CURRENT_4MA
            Public CSlowSlew As Boolean
            Public CSchmittInput As Boolean
            Public CDriveCurrent As FT_DRIVE_CURRENT = FT_DRIVE_CURRENT.FT_DRIVE_CURRENT_4MA
            Public DSlowSlew As Boolean
            Public DSchmittInput As Boolean
            Public DDriveCurrent As FT_DRIVE_CURRENT = FT_DRIVE_CURRENT.FT_DRIVE_CURRENT_4MA
            Public ARIIsTXDEN As Boolean
            Public BRIIsTXDEN As Boolean
            Public CRIIsTXDEN As Boolean
            Public DRIIsTXDEN As Boolean
            Public AIsVCP As Boolean = True
            Public BIsVCP As Boolean = True
            Public CIsVCP As Boolean = True
            Public DIsVCP As Boolean = True
        End Class

        Public Class FT232H_EEPROM_STRUCTURE : Inherits FT_EEPROM_DATA
            Public PullDownEnable As Boolean
            Public SerNumEnable As Boolean = True
            Public ACSlowSlew As Boolean
            Public ACSchmittInput As Boolean
            Public ACDriveCurrent As FT_DRIVE_CURRENT = FT_DRIVE_CURRENT.FT_DRIVE_CURRENT_4MA
            Public ADSlowSlew As Boolean
            Public ADSchmittInput As Boolean
            Public ADDriveCurrent As FT_DRIVE_CURRENT = FT_DRIVE_CURRENT.FT_DRIVE_CURRENT_4MA
            Public Cbus0 As Byte
            Public Cbus1 As Byte
            Public Cbus2 As Byte
            Public Cbus3 As Byte
            Public Cbus4 As Byte
            Public Cbus5 As Byte
            Public Cbus6 As Byte
            Public Cbus7 As Byte
            Public Cbus8 As Byte
            Public Cbus9 As Byte
            Public IsFifo As Boolean
            Public IsFifoTar As Boolean
            Public IsFastSer As Boolean
            Public IsFT1248 As Boolean
            Public FT1248Cpol As Boolean
            Public FT1248Lsb As Boolean
            Public FT1248FlowControl As Boolean
            Public IsVCP As Boolean = True
            Public PowerSaveEnable As Boolean
        End Class

        Public Class FT_XSERIES_EEPROM_STRUCTURE : Inherits FT_EEPROM_DATA
            Public PullDownEnable As Boolean
            Public SerNumEnable As Boolean = True
            Public USBVersionEnable As Boolean = True
            Public USBVersion As UShort = 512US
            Public ACSlowSlew As Byte
            Public ACSchmittInput As Byte
            Public ACDriveCurrent As FT_DRIVE_CURRENT
            Public ADSlowSlew As Byte
            Public ADSchmittInput As Byte
            Public ADDriveCurrent As FT_DRIVE_CURRENT
            Public Cbus0 As Byte
            Public Cbus1 As Byte
            Public Cbus2 As Byte
            Public Cbus3 As Byte
            Public Cbus4 As Byte
            Public Cbus5 As Byte
            Public Cbus6 As Byte
            Public InvertTXD As Byte
            Public InvertRXD As Byte
            Public InvertRTS As Byte
            Public InvertCTS As Byte
            Public InvertDTR As Byte
            Public InvertDSR As Byte
            Public InvertDCD As Byte
            Public InvertRI As Byte
            Public BCDEnable As Byte
            Public BCDForceCbusPWREN As Byte
            Public BCDDisableSleep As Byte
            Public I2CSlaveAddress As UShort
            Public I2CDeviceId As UInteger
            Public I2CDisableSchmitt As Byte
            Public FT1248Cpol As Byte
            Public FT1248Lsb As Byte
            Public FT1248FlowControl As Byte
            Public RS485EchoSuppress As Byte
            Public PowerSaveEnable As Byte
            Public IsVCP As Byte
        End Class

#End Region '/СТРУКТУРЫ ОБЪЕКТОВ В ППЗУ FTDI

#Region "УКАЗАТЕЛИ НА ФУНКЦИИ БИБЛИОТЕКИ ДЛЯ РАБОТЫ С ППЗУ"

        Private Const NOT_INIT As Integer = -1

        Private Shared PFt_ReadEE As Integer = NOT_INIT
        Private Shared PFt_WriteEE As Integer = NOT_INIT
        Private Shared PFt_EraseEE As Integer = NOT_INIT
        Private Shared PFt_EE_UASize As Integer = NOT_INIT
        Private Shared PFt_EE_UARead As Integer = NOT_INIT
        Private Shared PFt_EE_UAWrite As Integer = NOT_INIT
        Private Shared PFt_EE_Read As Integer = NOT_INIT
        Private Shared PFt_EE_Program As Integer = NOT_INIT
        Private Shared PFt_EEPROM_Read As Integer = NOT_INIT
        Private Shared PFt_EEPROM_Program As Integer = NOT_INIT

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function EeReadDelegate(ftHandle As Integer, ByRef pData As FT_PROGRAM_DATA) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function EeProgramDelegate(ftHandle As Integer, ByRef pData As FT_PROGRAM_DATA) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function EepromReadDelegate(ftHandle As Integer, eepromData As IntPtr, eepromDataSize As UInteger, manufacturer As Byte(), manufacturerID As Byte(), description As Byte(), serialnumber As Byte()) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function EepromProgramDelegate(ftHandle As Integer, eepromData As IntPtr, eepromDataSize As UInteger, manufacturer As Byte(), manufacturerID As Byte(), description As Byte(), serialnumber As Byte()) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function WriteEeDelegate(ftHandle As Integer, dwWordOffset As UInteger, wValue As UShort) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function EraseEeDelegate(ftHandle As Integer) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function EeUaSizeDelegate(ftHandle As Integer, ByRef dwSize As UInteger) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function EeUaReadDelegate(ftHandle As Integer, pucData As Byte(), dwDataLen As Integer, ByRef lpdwDataRead As UInteger) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function EeUaWriteDelegate(ftHandle As Integer, pucData As Byte(), dwDataLen As Integer) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function ReadEeDelegate(ftHandle As Integer, dwWordOffset As UInteger, ByRef lpwValue As UShort) As FT_STATUS

        Private Shared Sub FindEeFunctionPointers(ftd2xxDllHandle As Integer)
            If (PFt_ReadEE = NOT_INIT) Then
                PFt_ReadEE = GetProcAddress(ftd2xxDllHandle, "FT_ReadEE")
            End If
            If (PFt_WriteEE = NOT_INIT) Then
                PFt_WriteEE = GetProcAddress(ftd2xxDllHandle, "FT_WriteEE")
            End If
            If (PFt_EraseEE = NOT_INIT) Then
                PFt_EraseEE = GetProcAddress(ftd2xxDllHandle, "FT_EraseEE")
            End If
            If (PFt_EE_UASize = NOT_INIT) Then
                PFt_EE_UASize = GetProcAddress(ftd2xxDllHandle, "FT_EE_UASize")
            End If
            If (PFt_EE_UARead = NOT_INIT) Then
                PFt_EE_UARead = GetProcAddress(ftd2xxDllHandle, "FT_EE_UARead")
            End If
            If (PFt_EE_UAWrite = NOT_INIT) Then
                PFt_EE_UAWrite = GetProcAddress(ftd2xxDllHandle, "FT_EE_UAWrite")
            End If
            If (PFt_EE_Read = NOT_INIT) Then
                PFt_EE_Read = GetProcAddress(ftd2xxDllHandle, "FT_EE_Read")
            End If
            If (PFt_EE_Program = NOT_INIT) Then
                PFt_EE_Program = GetProcAddress(ftd2xxDllHandle, "FT_EE_Program")
            End If
            If (PFt_EEPROM_Read = NOT_INIT) Then
                PFt_EEPROM_Read = GetProcAddress(ftd2xxDllHandle, "FT_EEPROM_Read")
            End If
            If (PFt_EEPROM_Program = NOT_INIT) Then
                PFt_EEPROM_Program = GetProcAddress(ftd2xxDllHandle, "FT_EEPROM_Program")
            End If
        End Sub

        <DllImport("kernel32.dll", SetLastError:=True)>
        Private Shared Function GetProcAddress(hModule As Integer, procedureName As String) As Integer
        End Function

#End Region '/УКАЗАТЕЛИ НА ФУНКЦИИ БИБЛИОТЕКИ ДЛЯ РАБОТЫ С ППЗУ

#Region "ПЕРЕЧИСЛЕНИЯ"

        Public Enum FT_DRIVE_CURRENT As Byte
            FT_DRIVE_CURRENT_4MA = 4
            FT_DRIVE_CURRENT_8MA = 8
            FT_DRIVE_CURRENT_12MA = 12
            FT_DRIVE_CURRENT_16MA = 16
        End Enum

#End Region '/ПЕРЕЧИСЛЕНИЯ

    End Class '/Eeprom

End Namespace
