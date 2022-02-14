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
            Dim dlg As EeReadDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EE_Read, IntPtr), GetType(EeReadDelegate)), EeReadDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, programData)
            ee232b.Manufacturer = Marshal.PtrToStringAnsi(programData.Manufacturer)
            ee232b.ManufacturerID = Marshal.PtrToStringAnsi(programData.ManufacturerID)
            ee232b.Description = Marshal.PtrToStringAnsi(programData.Description)
            ee232b.SerialNumber = Marshal.PtrToStringAnsi(programData.SerialNumber)
            Marshal.FreeHGlobal(programData.Manufacturer)
            Marshal.FreeHGlobal(programData.ManufacturerID)
            Marshal.FreeHGlobal(programData.Description)
            Marshal.FreeHGlobal(programData.SerialNumber)
            CheckErrors(status) 'статус проверяем после выгрузки всех неуправляемых ресурсов
            ee232b.VendorID = programData.VendorID
            ee232b.ProductID = programData.ProductID
            ee232b.MaxPower = programData.MaxPower
            ee232b.SelfPowered = Convert.ToBoolean(programData.SelfPowered)
            ee232b.RemoteWakeup = Convert.ToBoolean(programData.RemoteWakeup)
            ee232b.PullDownEnable = Convert.ToBoolean(programData.PullDownEnable)
            ee232b.SerNumEnable = Convert.ToBoolean(programData.SerNumEnable)
            ee232b.USBVersionEnable = Convert.ToBoolean(programData.USBVersionEnable)
            ee232b.USBVersion = programData.USBVersion
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
            Dim dlg As EeReadDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EE_Read, IntPtr), GetType(EeReadDelegate)), EeReadDelegate)
            Dim status As FT_STATUS = dlg.Invoke(FtHandle, programData)
            ee2232.Manufacturer = Marshal.PtrToStringAnsi(programData.Manufacturer)
            ee2232.ManufacturerID = Marshal.PtrToStringAnsi(programData.ManufacturerID)
            ee2232.Description = Marshal.PtrToStringAnsi(programData.Description)
            ee2232.SerialNumber = Marshal.PtrToStringAnsi(programData.SerialNumber)
            Marshal.FreeHGlobal(programData.Manufacturer)
            Marshal.FreeHGlobal(programData.ManufacturerID)
            Marshal.FreeHGlobal(programData.Description)
            Marshal.FreeHGlobal(programData.SerialNumber)
            CheckErrors(status) 'статус проверяем после выгрузки всех неуправляемых ресурсов
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
            CheckErrors(status) 'статус проверяем после выгрузки всех неуправляемых ресурсов
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
            Dim dlg As EeReadDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EE_Read, IntPtr), GetType(EeReadDelegate)), EeReadDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, programData)
            ee2232h.Manufacturer = Marshal.PtrToStringAnsi(programData.Manufacturer)
            ee2232h.ManufacturerID = Marshal.PtrToStringAnsi(programData.ManufacturerID)
            ee2232h.Description = Marshal.PtrToStringAnsi(programData.Description)
            ee2232h.SerialNumber = Marshal.PtrToStringAnsi(programData.SerialNumber)
            Marshal.FreeHGlobal(programData.Manufacturer)
            Marshal.FreeHGlobal(programData.ManufacturerID)
            Marshal.FreeHGlobal(programData.Description)
            Marshal.FreeHGlobal(programData.SerialNumber)
            CheckErrors(status) 'статус проверяем после выгрузки всех неуправляемых ресурсов
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
            Dim dlg As EeReadDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EE_Read, IntPtr), GetType(EeReadDelegate)), EeReadDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, programData)
            ee4232h.Manufacturer = Marshal.PtrToStringAnsi(programData.Manufacturer)
            ee4232h.ManufacturerID = Marshal.PtrToStringAnsi(programData.ManufacturerID)
            ee4232h.Description = Marshal.PtrToStringAnsi(programData.Description)
            ee4232h.SerialNumber = Marshal.PtrToStringAnsi(programData.SerialNumber)
            Marshal.FreeHGlobal(programData.Manufacturer)
            Marshal.FreeHGlobal(programData.ManufacturerID)
            Marshal.FreeHGlobal(programData.Description)
            Marshal.FreeHGlobal(programData.SerialNumber)
            CheckErrors(status) 'статус проверяем после выгрузки всех неуправляемых ресурсов
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
            Dim dlg As EeReadDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(PFt_EE_Read), GetType(EeReadDelegate)), EeReadDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, programData)
            ee232h.Manufacturer = Marshal.PtrToStringAnsi(programData.Manufacturer)
            ee232h.ManufacturerID = Marshal.PtrToStringAnsi(programData.ManufacturerID)
            ee232h.Description = Marshal.PtrToStringAnsi(programData.Description)
            ee232h.SerialNumber = Marshal.PtrToStringAnsi(programData.SerialNumber)
            Marshal.FreeHGlobal(programData.Manufacturer)
            Marshal.FreeHGlobal(programData.ManufacturerID)
            Marshal.FreeHGlobal(programData.Description)
            Marshal.FreeHGlobal(programData.SerialNumber)
            CheckErrors(status) 'статус проверяем после выгрузки всех неуправляемых ресурсов
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
            ee232h.Cbus0 = CType(programData.Cbus0H, FT_232H_CBUS_OPTIONS)
            ee232h.Cbus1 = CType(programData.Cbus1H, FT_232H_CBUS_OPTIONS)
            ee232h.Cbus2 = CType(programData.Cbus2H, FT_232H_CBUS_OPTIONS)
            ee232h.Cbus3 = CType(programData.Cbus3H, FT_232H_CBUS_OPTIONS)
            ee232h.Cbus4 = CType(programData.Cbus4H, FT_232H_CBUS_OPTIONS)
            ee232h.Cbus5 = CType(programData.Cbus5H, FT_232H_CBUS_OPTIONS)
            ee232h.Cbus6 = CType(programData.Cbus6H, FT_232H_CBUS_OPTIONS)
            ee232h.Cbus7 = CType(programData.Cbus7H, FT_232H_CBUS_OPTIONS)
            ee232h.Cbus8 = CType(programData.Cbus8H, FT_232H_CBUS_OPTIONS)
            ee232h.Cbus9 = CType(programData.Cbus9H, FT_232H_CBUS_OPTIONS)
            ee232h.IsFifo = Convert.ToBoolean(programData.IsFifoH)
            ee232h.IsFifoTar = Convert.ToBoolean(programData.IsFifoTarH)
            ee232h.IsFastSer = Convert.ToBoolean(programData.IsFastSerH)
            ee232h.IsFT1248 = Convert.ToBoolean(programData.IsFT1248H)
            ee232h.FT1248Cpol = Convert.ToBoolean(programData.FT1248CpolH)
            ee232h.FT1248Lsb = Convert.ToBoolean(programData.FT1248LsbH)
            ee232h.FT1248FlowControl = Convert.ToBoolean(programData.FT1248FlowControlH)
            ee232h.IsVCP = Convert.ToBoolean(programData.IsVCPH)
            ee232h.PowerSaveEnable = Convert.ToBoolean(programData.PowerSaveEnableH)
            Return ee232h
        End Function

        Private Function ReadXSeriesEeprom() As FT_XSERIES_EEPROM_STRUCTURE
            Dim eeX As New FT_XSERIES_EEPROM_STRUCTURE()
            Dim header As New FT_EEPROM_HEADER() With {.DeviceType = FT_DEVICE.FT_DEVICE_X_SERIES}
            Dim programData As New FT_XSERIES_DATA() With {.common = header}
            Dim strSize As Integer = Marshal.SizeOf(programData)
            Dim intPtr As IntPtr = Marshal.AllocHGlobal(strSize)
            Marshal.StructureToPtr(programData, intPtr, False)
            Dim man As Byte() = New Byte(31) {}
            Dim id As Byte() = New Byte(15) {}
            Dim descr As Byte() = New Byte(63) {}
            Dim sn As Byte() = New Byte(15) {}
            Dim dlg As EepromReadDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(PFt_EEPROM_Read), GetType(EepromReadDelegate)), EepromReadDelegate)
            Dim status As FT_STATUS = dlg.Invoke(FtHandle, intPtr, CUInt(strSize), man, id, descr, sn)
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
            eeX.Cbus0 = CType(programData.Cbus0, FT_XSERIES_CBUS_OPTIONS)
            eeX.Cbus1 = CType(programData.Cbus1, FT_XSERIES_CBUS_OPTIONS)
            eeX.Cbus2 = CType(programData.Cbus2, FT_XSERIES_CBUS_OPTIONS)
            eeX.Cbus3 = CType(programData.Cbus3, FT_XSERIES_CBUS_OPTIONS)
            eeX.Cbus4 = CType(programData.Cbus4, FT_XSERIES_CBUS_OPTIONS)
            eeX.Cbus5 = CType(programData.Cbus5, FT_XSERIES_CBUS_OPTIONS)
            eeX.Cbus6 = CType(programData.Cbus6, FT_XSERIES_CBUS_OPTIONS)
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
            CheckErrors(status) 'статус проверяем после выгрузки всех неуправляемых ресурсов
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
            Dim dlg As EeProgramDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EE_Program, IntPtr), GetType(EeProgramDelegate)), EeProgramDelegate)
            status = dlg(FtHandle, programData)
            Marshal.FreeHGlobal(programData.Manufacturer)
            Marshal.FreeHGlobal(programData.ManufacturerID)
            Marshal.FreeHGlobal(programData.Description)
            Marshal.FreeHGlobal(programData.SerialNumber)
            CheckErrors(status) 'статус проверяем после освобождения всех неуправляемых ресурсов
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
            Dim dlg As EeProgramDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EE_Program, IntPtr), GetType(EeProgramDelegate)), EeProgramDelegate)
            status = dlg(FtHandle, programData)
            Marshal.FreeHGlobal(programData.Manufacturer)
            Marshal.FreeHGlobal(programData.ManufacturerID)
            Marshal.FreeHGlobal(programData.Description)
            Marshal.FreeHGlobal(programData.SerialNumber)
            CheckErrors(status) 'статус проверяем после освобождения всех неуправляемых ресурсов
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
            Dim dlg As EeProgramDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EE_Program, IntPtr), GetType(EeProgramDelegate)), EeProgramDelegate)
            status = dlg(FtHandle, programData)
            Marshal.FreeHGlobal(programData.Manufacturer)
            Marshal.FreeHGlobal(programData.ManufacturerID)
            Marshal.FreeHGlobal(programData.Description)
            Marshal.FreeHGlobal(programData.SerialNumber)
            CheckErrors(status) 'статус проверяем после освобождения всех неуправляемых ресурсов
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
            Dim dlg As EeProgramDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EE_Program, IntPtr), GetType(EeProgramDelegate)), EeProgramDelegate)
            status = dlg(FtHandle, programData)
            Marshal.FreeHGlobal(programData.Manufacturer)
            Marshal.FreeHGlobal(programData.ManufacturerID)
            Marshal.FreeHGlobal(programData.Description)
            Marshal.FreeHGlobal(programData.SerialNumber)
            CheckErrors(status) 'статус проверяем после освобождения всех неуправляемых ресурсов
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
            Dim dlg As EeProgramDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EE_Program, IntPtr), GetType(EeProgramDelegate)), EeProgramDelegate)
            status = dlg(FtHandle, programData)
            Marshal.FreeHGlobal(programData.Manufacturer)
            Marshal.FreeHGlobal(programData.ManufacturerID)
            Marshal.FreeHGlobal(programData.Description)
            Marshal.FreeHGlobal(programData.SerialNumber)
            CheckErrors(status) 'статус проверяем после освобождения всех неуправляемых ресурсов
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
                .Cbus0H = CType(ee232h.Cbus0, FT_CBUS_OPTIONS),
                .Cbus1H = CType(ee232h.Cbus1, FT_CBUS_OPTIONS),
                .Cbus2H = CType(ee232h.Cbus2, FT_CBUS_OPTIONS),
                .Cbus3H = CType(ee232h.Cbus3, FT_CBUS_OPTIONS),
                .Cbus4H = CType(ee232h.Cbus4, FT_CBUS_OPTIONS),
                .Cbus5H = CType(ee232h.Cbus5, FT_CBUS_OPTIONS),
                .Cbus6H = CType(ee232h.Cbus6, FT_CBUS_OPTIONS),
                .Cbus7H = CType(ee232h.Cbus7, FT_CBUS_OPTIONS),
                .Cbus8H = CType(ee232h.Cbus8, FT_CBUS_OPTIONS),
                .Cbus9H = CType(ee232h.Cbus9, FT_CBUS_OPTIONS),
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
            Dim dlg As EeProgramDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EE_Program, IntPtr), GetType(EeProgramDelegate)), EeProgramDelegate)
            status = dlg(FtHandle, programData)
            Marshal.FreeHGlobal(programData.Manufacturer)
            Marshal.FreeHGlobal(programData.ManufacturerID)
            Marshal.FreeHGlobal(programData.Description)
            Marshal.FreeHGlobal(programData.SerialNumber)
            CheckErrors(status) 'статус проверяем после освобождения всех неуправляемых ресурсов
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
                .Cbus0 = CType(eeX.Cbus0, FT_CBUS_OPTIONS),
                .Cbus1 = CType(eeX.Cbus1, FT_CBUS_OPTIONS),
                .Cbus2 = CType(eeX.Cbus2, FT_CBUS_OPTIONS),
                .Cbus3 = CType(eeX.Cbus3, FT_CBUS_OPTIONS),
                .Cbus4 = CType(eeX.Cbus4, FT_CBUS_OPTIONS),
                .Cbus5 = CType(eeX.Cbus5, FT_CBUS_OPTIONS),
                .Cbus6 = CType(eeX.Cbus6, FT_CBUS_OPTIONS),
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
            programData.Common.DeviceType = FT_DEVICE.FT_DEVICE_X_SERIES
            programData.Common.VendorId = eeX.VendorID
            programData.Common.ProductId = eeX.ProductID
            programData.Common.MaxPower = eeX.MaxPower
            programData.Common.SelfPowered = Convert.ToByte(eeX.SelfPowered)
            programData.Common.RemoteWakeup = Convert.ToByte(eeX.RemoteWakeup)
            programData.Common.SerNumEnable = Convert.ToByte(eeX.SerNumEnable)
            programData.Common.PullDownEnable = Convert.ToByte(eeX.PullDownEnable)
            Dim utf8Encoding As UTF8Encoding = New UTF8Encoding()
            Dim manufacturer As Byte() = utf8Encoding.GetBytes(eeX.Manufacturer)
            Dim manufacturerID As Byte() = utf8Encoding.GetBytes(eeX.ManufacturerID)
            Dim description As Byte() = utf8Encoding.GetBytes(eeX.Description)
            Dim serialnumber As Byte() = utf8Encoding.GetBytes(eeX.SerialNumber)
            Dim programDataSize As Integer = Marshal.SizeOf(programData)
            Dim ptr As IntPtr = Marshal.AllocHGlobal(programDataSize)
            Marshal.StructureToPtr(programData, ptr, False)
            Dim dlg As EepromProgramDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(PFt_EEPROM_Program, IntPtr), GetType(EepromProgramDelegate)), EepromProgramDelegate)
            status = dlg(FtHandle, ptr, CUInt(programDataSize), manufacturer, manufacturerID, description, serialnumber)
            CheckErrors(status) 'статус проверяем после освобождения всех неуправляемых ресурсов
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

            'FT232B extensions
            Public Rev4 As Byte
            Public IsoIn As Byte
            Public IsoOut As Byte
            Public PullDownEnable As Byte
            Public SerNumEnable As Byte
            Public USBVersionEnable As Byte
            Public USBVersion As USB_VERSION

            'FT2232D extensions
            Public Rev5 As Byte
            Public IsoInA As Byte
            Public IsoInB As Byte
            Public IsoOutA As Byte
            Public IsoOutB As Byte
            Public PullDownEnable5 As Byte
            Public SerNumEnable5 As Byte
            Public USBVersionEnable5 As Byte
            Public USBVersion5 As USB_VERSION
            Public AIsHighCurrent As Byte
            Public BIsHighCurrent As Byte
            Public IFAIsFifo As Byte
            Public IFAIsFifoTar As Byte
            Public IFAIsFastSer As Byte
            Public AIsVCP As Byte
            Public IFBIsFifo As Byte
            Public IFBIsFifoTar As Byte
            Public IFBIsFastSer As Byte

            'FT232R extensions
            Public BIsVCP As Byte
            Public UseExtOsc As Byte
            Public HighDriveIOs As Byte
            Public EndpointSize As Byte
            Public PullDownEnableR As Byte
            Public SerNumEnableR As Byte
            Public InvertTXD As Byte 'non-zero if invert TXD
            Public InvertRXD As Byte 'non-zero if invert RXD
            Public InvertRTS As Byte 'non-zero if invert RTS
            Public InvertCTS As Byte 'non-zero if invert CTS
            Public InvertDTR As Byte 'non-zero if invert DTR
            Public InvertDSR As Byte 'non-zero if invert DSR
            Public InvertDCD As Byte 'non-zero if invert DCD
            Public InvertRI As Byte  'non-zero if invert RI
            Public Cbus0 As FT_CBUS_OPTIONS 'Cbus Mux control - Ignored for FT245R
            Public Cbus1 As FT_CBUS_OPTIONS 'Cbus Mux control - Ignored for FT245R
            Public Cbus2 As FT_CBUS_OPTIONS 'Cbus Mux control - Ignored for FT245R
            Public Cbus3 As FT_CBUS_OPTIONS 'Cbus Mux control - Ignored for FT245R
            Public Cbus4 As FT_CBUS_OPTIONS 'Cbus Mux control - Ignored for FT245R
            Public RIsD2XX As Byte 'Default to loading VCP

            'FT2232H extensions
            Public PullDownEnable7 As Byte
            Public SerNumEnable7 As Byte
            Public ALSlowSlew As Byte    'non-zero if AL pins have slow slew
            Public ALSchmittInput As Byte 'non-zero if AL pins are Schmitt input
            Public ALDriveCurrent As FT_DRIVE_CURRENT
            Public AHSlowSlew As Byte    'non-zero if AH pins have slow slew
            Public AHSchmittInput As Byte 'non-zero if AH pins are Schmitt input
            Public AHDriveCurrent As FT_DRIVE_CURRENT
            Public BLSlowSlew As Byte    'non-zero if BL pins have slow slew
            Public BLSchmittInput As Byte 'non-zero if BL pins are Schmitt input
            Public BLDriveCurrent As FT_DRIVE_CURRENT
            Public BHSlowSlew As Byte    'non-zero if BH pins have slow slew
            Public BHSchmittInput As Byte 'non-zero if BH pins are Schmitt input
            Public BHDriveCurrent As FT_DRIVE_CURRENT
            Public IFAIsFifo7 As Byte     'non-zero if interface is 245 FIFO
            Public IFAIsFifoTar7 As Byte  'non-zero if interface is 245 FIFO CPU target
            Public IFAIsFastSer7 As Byte  'non-zero if interface is Fast serial
            Public AIsVCP7 As Byte        'non-zero if interface is to use VCP drivers
            Public IFBIsFifo7 As Byte     'non-zero if interface is 245 FIFO
            Public IFBIsFifoTar7 As Byte  'non-zero if interface is 245 FIFO CPU target
            Public IFBIsFastSer7 As Byte  'non-zero if interface is Fast serial
            Public BIsVCP7 As Byte        'non-zero if interface is to use VCP drivers
            Public PowerSaveEnable As Byte 'non-zero if using BCBUS7 to save power for self-powered designs

            'FT4232H extensions
            Public PullDownEnable8 As Byte
            Public SerNumEnable8 As Byte
            Public ASlowSlew As Byte    'non-zero if AL pins have slow slew
            Public ASchmittInput As Byte 'non-zero if AL pins are Schmitt input
            Public ADriveCurrent As FT_DRIVE_CURRENT
            Public BSlowSlew As Byte    'non-zero if AH pins have slow slew
            Public BSchmittInput As Byte 'non-zero if AH pins are Schmitt input
            Public BDriveCurrent As FT_DRIVE_CURRENT
            Public CSlowSlew As Byte    'non-zero if BL pins have slow slew
            Public CSchmittInput As Byte 'non-zero if BL pins are Schmitt input
            Public CDriveCurrent As FT_DRIVE_CURRENT
            Public DSlowSlew As Byte    'non-zero if BH pins have slow slew
            Public DSchmittInput As Byte 'non-zero if BH pins are Schmitt input
            Public DDriveCurrent As FT_DRIVE_CURRENT
            Public ARIIsTXDEN As Byte
            Public BRIIsTXDEN As Byte
            Public CRIIsTXDEN As Byte
            Public DRIIsTXDEN As Byte
            Public AIsVCP8 As Byte 'non-zero if interface is to use VCP drivers
            Public BIsVCP8 As Byte 'non-zero if interface is to use VCP drivers
            Public CIsVCP8 As Byte 'non-zero if interface is to use VCP drivers
            Public DIsVCP8 As Byte 'non-zero if interface is to use VCP drivers

            'FT232H extensions
            Public PullDownEnableH As Byte 'non-zero if pull down enabled
            Public SerNumEnableH As Byte  'non-zero if serial number to be used
            Public ACSlowSlewH As Byte    'non-zero if AC pins have slow slew
            Public ACSchmittInputH As Byte 'non-zero if AC pins are Schmitt input
            Public ACDriveCurrentH As FT_DRIVE_CURRENT
            Public ADSlowSlewH As Byte    'non-zero if AD pins have slow slew
            Public ADSchmittInputH As Byte 'non-zero if AD pins are Schmitt input
            Public ADDriveCurrentH As FT_DRIVE_CURRENT
            Public Cbus0H As FT_CBUS_OPTIONS 'Cbus Mux control
            Public Cbus1H As FT_CBUS_OPTIONS 'Cbus Mux control
            Public Cbus2H As FT_CBUS_OPTIONS 'Cbus Mux control
            Public Cbus3H As FT_CBUS_OPTIONS 'Cbus Mux control
            Public Cbus4H As FT_CBUS_OPTIONS 'Cbus Mux control
            Public Cbus5H As FT_CBUS_OPTIONS 'Cbus Mux control
            Public Cbus6H As FT_CBUS_OPTIONS 'Cbus Mux control
            Public Cbus7H As FT_CBUS_OPTIONS 'Cbus Mux control
            Public Cbus8H As FT_CBUS_OPTIONS 'Cbus Mux control
            Public Cbus9H As FT_CBUS_OPTIONS 'Cbus Mux control
            Public IsFifoH As Byte     'non-zero if interface is 245 FIFO
            Public IsFifoTarH As Byte  'non-zero if interface is 245 FIFO CPU target
            Public IsFastSerH As Byte  'non-zero if interface is Fast serial
            Public IsFT1248H As Byte   'non-zero if interface is FT1248
            Public FT1248CpolH As Byte 'FT1248 clock polarity
            Public FT1248LsbH As Byte  'FT1248 data is LSB (1) or MSB (0)
            Public FT1248FlowControlH As Byte 'FT1248 flow control enable
            Public IsVCPH As Byte          'non-zero if interface is to use VCP drivers
            Public PowerSaveEnableH As Byte 'non-zero if using ACBUS7 to save power for self-powered designs
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=4)>
        Private Structure FT_EEPROM_HEADER
            Public DeviceType As FT_DEVICE 'FTxxxx device type to be programmed

            'Device descriptor options
            Public VendorId As UShort  '0x0403
            Public ProductId As UShort '0x6001
            Public SerNumEnable As Byte 'non-zero if serial number to be used

            'Config descriptor options
            ''' <summary>
            ''' 0...500.
            ''' </summary>
            Public MaxPower As UShort
            ''' <summary>
            ''' 1 = self powered, 0 = bus powered.
            ''' </summary>
            Public SelfPowered As Byte
            ''' <summary>
            ''' 1 = capable, 0 = not capable.
            ''' </summary>
            Public RemoteWakeup As Byte

            'Hardware options
            Public PullDownEnable As Byte 'non-zero if pull down in suspend enabled
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=4)>
        Private Structure FT_XSERIES_DATA
            Public Common As FT_EEPROM_HEADER
            Public ACSlowSlew As Byte    'non-zero if AC bus pins have slow slew
            Public ACSchmittInput As Byte 'non-zero if AC bus pins are Schmitt input
            Public ACDriveCurrent As FT_DRIVE_CURRENT
            Public ADSlowSlew As Byte    'non-zero if AD bus pins have slow slew
            Public ADSchmittInput As Byte 'non-zero if AD bus pins are Schmitt input
            Public ADDriveCurrent As FT_DRIVE_CURRENT

            'CBUS options
            Public Cbus0 As FT_CBUS_OPTIONS 'Cbus Mux control
            Public Cbus1 As FT_CBUS_OPTIONS 'Cbus Mux control
            Public Cbus2 As FT_CBUS_OPTIONS 'Cbus Mux control
            Public Cbus3 As FT_CBUS_OPTIONS 'Cbus Mux control
            Public Cbus4 As FT_CBUS_OPTIONS 'Cbus Mux control
            Public Cbus5 As FT_CBUS_OPTIONS 'Cbus Mux control
            Public Cbus6 As FT_CBUS_OPTIONS 'Cbus Mux control

            'UART signal options
            Public InvertTXD As Byte 'non-zero if invert TXD
            Public InvertRXD As Byte 'non-zero if invert RXD
            Public InvertRTS As Byte 'non-zero if invert RTS
            Public InvertCTS As Byte 'non-zero if invert CTS
            Public InvertDTR As Byte 'non-zero if invert DTR
            Public InvertDSR As Byte 'non-zero if invert DSR
            Public InvertDCD As Byte 'non-zero if invert DCD
            Public InvertRI As Byte ' non-zero if invert RI

            ' Battery Charge Detect options
            ''' <summary>
            ''' Enable Battery Charger Detection.
            ''' </summary>
            Public BCDEnable As Byte
            ''' <summary>
            ''' Asserts the power enable signal on CBUS when charging port detected.
            ''' </summary>
            Public BCDForceCbusPWREN As Byte
            ''' <summary>
            ''' Forces the device never to go into sleep mode.
            ''' </summary>
            Public BCDDisableSleep As Byte

            ' I2C options

            ''' <summary>
            ''' I2C slave device address.
            ''' </summary>
            Public I2CSlaveAddress As UShort
            ''' <summary>
            ''' I2C device ID.
            ''' </summary>
            Public I2CDeviceId As UInteger
            ''' <summary>
            ''' Disable I2C Schmitt trigger.
            ''' </summary>
            Public I2CDisableSchmitt As Byte

            ' FT1248 options

            ''' <summary>
            ''' FT1248 clock polarity - clock idle high (1) or clock idle low (0)
            ''' </summary>
            Public FT1248Cpol As Byte
            ''' <summary>
            ''' FT1248 data is LSB (1) or MSB (0).
            ''' </summary>
            Public FT1248Lsb As Byte
            ''' <summary>
            ''' FT1248 flow control enable.
            ''' </summary>
            Public FT1248FlowControl As Byte

            'Hardware options
            Public RS485EchoSuppress As Byte
            Public PowerSaveEnable As Byte

            'Driver option
            Public DriverType As Byte
        End Structure

        Public Class FT_EEPROM_DATA
            Public VendorID As UShort = &H403US
            Public ProductID As UShort = &H6001US
            Public Manufacturer As String = "FTDI"
            ''' <summary>
            ''' Аббревиатура производителя, которая используется как префикс в автоматически генерируемых серийных номерах.
            ''' </summary>
            Public ManufacturerID As String = "FT"
            Public Description As String = "USB-Serial Converter"
            Public SerialNumber As String = ""
            ''' <summary>
            ''' Максимальная мощность, требующаяся устройству.
            ''' </summary>
            Public MaxPower As UShort = &H90
            ''' <summary>
            ''' Показывает, устройство имеет свой собственный источник питания (self-powered) или берёт питание от USB порта (bus-powered).
            ''' </summary>
            Public SelfPowered As Boolean
            ''' <summary>
            ''' Определяет, может ли устрйоство выводить ПК из режима ожидания, переключая линию RI.
            ''' </summary>
            Public RemoteWakeup As Boolean
        End Class

        Public Class FT232B_EEPROM_STRUCTURE : Inherits FT_EEPROM_DATA
            ''' <summary>
            ''' Определяет, притянуты ли выводы IO, когда устройство в режиме ожидания.
            ''' </summary>
            Public PullDownEnable As Boolean
            ''' <summary>
            ''' Determines if the serial number is enabled.
            ''' </summary>
            Public SerNumEnable As Boolean = True
            ''' <summary>
            ''' Determines if the USB version number is enabled.
            ''' </summary>
            Public USBVersionEnable As Boolean = True
            ''' <summary>
            ''' The USB version number.
            ''' </summary>
            Public USBVersion As USB_VERSION = USB_VERSION.VER_20
        End Class

        Public Class FT2232_EEPROM_STRUCTURE : Inherits FT_EEPROM_DATA
            ''' <summary>
            ''' Определяет, притянуты ли выводы IO, когда устройство в режиме ожидания.
            ''' </summary>
            Public PullDownEnable As Boolean
            ''' <summary>
            ''' Determines if the serial number is enabled.
            ''' </summary>
            Public SerNumEnable As Boolean = True
            ''' <summary>
            ''' Determines if the USB version number is enabled.
            ''' </summary>
            Public USBVersionEnable As Boolean = True
            ''' <summary>
            ''' The USB version number.
            ''' </summary>
            Public USBVersion As USB_VERSION = USB_VERSION.VER_20
            ''' <summary>
            ''' Enables high current IOs on channel A.
            ''' </summary>
            Public AIsHighCurrent As Boolean
            ''' <summary>
            ''' Enables high current IOs on channel B.
            ''' </summary>
            Public BIsHighCurrent As Boolean
            ''' <summary>
            ''' Determines if channel A is in FIFO mode.
            ''' </summary>
            Public IFAIsFifo As Boolean
            ''' <summary>
            ''' Determines if channel A is in FIFO target mode.
            ''' </summary>
            Public IFAIsFifoTar As Boolean
            ''' <summary>
            ''' Determines if channel A is in fast serial mode.
            ''' </summary>
            Public IFAIsFastSer As Boolean
            ''' <summary>
            ''' Determines if channel A loads the VCP driver.
            ''' </summary>
            Public AIsVCP As Boolean = True
            ''' <summary>
            ''' Determines if channel B is in FIFO mode.
            ''' </summary>
            Public IFBIsFifo As Boolean
            ''' <summary>
            ''' Determines if channel B is in FIFO target mode.
            ''' </summary>
            Public IFBIsFifoTar As Boolean
            ''' <summary>
            ''' Determines if channel B is in fast serial mode.
            ''' </summary>
            Public IFBIsFastSer As Boolean
            ''' <summary>
            ''' Determines if channel B loads the VCP driver.
            ''' </summary>
            Public BIsVCP As Boolean = True
        End Class

        ''' <summary>
        ''' EEPROM для FT232R и FT245R.
        ''' </summary>
        Public Class FT232R_EEPROM_STRUCTURE : Inherits FT_EEPROM_DATA
            ''' <summary>
            ''' Disables the FT232R internal clock source.
            ''' If the device has external oscillator enabled it must have an external oscillator fitted to function.
            ''' </summary>
            Public UseExtOsc As Boolean
            ''' <summary>
            ''' Enables high current IOs.
            ''' </summary>
            Public HighDriveIOs As Boolean
            ''' <summary>
            ''' Sets the endpoint size. This should always be set to 64.
            ''' </summary>
            Public EndpointSize As Byte = 64
            ''' <summary>
            ''' Determines if IOs are pulled down when the device is in suspend.
            ''' </summary>
            Public PullDownEnable As Boolean
            ''' <summary>
            ''' Determines if the serial number is enabled.
            ''' </summary>
            Public SerNumEnable As Boolean = True
            ''' <summary>
            ''' Inverts the sense of the TXD line.
            ''' </summary>
            Public InvertTXD As Boolean
            ''' <summary>
            ''' Inverts the sense of the RXD line.
            ''' </summary>
            Public InvertRXD As Boolean
            ''' <summary>
            ''' Inverts the sense of the RTS line.
            ''' </summary>
            Public InvertRTS As Boolean
            ''' <summary>
            ''' Inverts the sense of the CTS line.
            ''' </summary>
            Public InvertCTS As Boolean
            ''' <summary>
            ''' Inverts the sense of the DTR line.
            ''' </summary>
            Public InvertDTR As Boolean
            ''' <summary>
            ''' Inverts the sense of the DSR line.
            ''' </summary>
            Public InvertDSR As Boolean
            ''' <summary>
            ''' Inverts the sense of the DCD line.
            ''' </summary>
            Public InvertDCD As Boolean
            ''' <summary>
            ''' Inverts the sense of the RI line.
            ''' </summary>
            Public InvertRI As Boolean
            Public Cbus0 As FT_CBUS_OPTIONS = FT_CBUS_OPTIONS.FT_CBUS_SLEEP
            Public Cbus1 As FT_CBUS_OPTIONS = FT_CBUS_OPTIONS.FT_CBUS_SLEEP
            Public Cbus2 As FT_CBUS_OPTIONS = FT_CBUS_OPTIONS.FT_CBUS_SLEEP
            Public Cbus3 As FT_CBUS_OPTIONS = FT_CBUS_OPTIONS.FT_CBUS_SLEEP
            Public Cbus4 As FT_CBUS_OPTIONS = FT_CBUS_OPTIONS.FT_CBUS_SLEEP
            ''' <summary>
            ''' Determines if the VCP driver is loaded.
            ''' </summary>
            Public RIsD2XX As Boolean
        End Class

        Public Class FT2232H_EEPROM_STRUCTURE : Inherits FT_EEPROM_DATA
            ''' <summary>
            ''' Determines if IOs are pulled down when the device is in suspend.
            ''' </summary>
            Public PullDownEnable As Boolean
            ''' <summary>
            ''' Determines if the serial number is enabled.
            ''' </summary>
            Public SerNumEnable As Boolean = True
            ''' <summary>
            ''' Determines if AL pins have a slow slew rate.
            ''' </summary>
            Public ALSlowSlew As Boolean
            ''' <summary>
            ''' Determines if the AL pins have a Schmitt input.
            ''' </summary>
            Public ALSchmittInput As Boolean
            ''' <summary>
            ''' Determines the AL pins drive current in mA.
            ''' </summary>
            Public ALDriveCurrent As FT_DRIVE_CURRENT = FT_DRIVE_CURRENT.FT_DRIVE_CURRENT_4MA
            ''' <summary>
            ''' Determines if AH pins have a slow slew rate.
            ''' </summary>
            Public AHSlowSlew As Boolean
            ''' <summary>
            ''' Determines if the AH pins have a Schmitt input.
            ''' </summary>
            Public AHSchmittInput As Boolean
            ''' <summary>
            ''' Determines the AH pins drive current in mA.
            ''' </summary>
            Public AHDriveCurrent As FT_DRIVE_CURRENT = FT_DRIVE_CURRENT.FT_DRIVE_CURRENT_4MA
            ''' <summary>
            ''' Determines if BL pins have a slow slew rate.
            ''' </summary>
            Public BLSlowSlew As Boolean
            ''' <summary>
            ''' Determines if the BL pins have a Schmitt input.
            ''' </summary>
            Public BLSchmittInput As Boolean
            ''' <summary>
            ''' Determines the BL pins drive current in mA.
            ''' </summary>
            Public BLDriveCurrent As FT_DRIVE_CURRENT = FT_DRIVE_CURRENT.FT_DRIVE_CURRENT_4MA
            ''' <summary>
            ''' Determines if BH pins have a slow slew rate.
            ''' </summary>
            Public BHSlowSlew As Boolean
            ''' <summary>
            ''' Determines if the BH pins have a Schmitt input.
            ''' </summary>
            Public BHSchmittInput As Boolean
            ''' <summary>
            ''' Determines the BH pins drive current in mA.
            ''' </summary>
            Public BHDriveCurrent As FT_DRIVE_CURRENT = FT_DRIVE_CURRENT.FT_DRIVE_CURRENT_4MA
            ''' <summary>
            ''' Determines if channel A is in FIFO mode.
            ''' </summary>
            Public IFAIsFifo As Boolean
            ''' <summary>
            ''' Determines if channel A is in FIFO target mode.
            ''' </summary>
            Public IFAIsFifoTar As Boolean
            ''' <summary>
            ''' Determines if channel A is in fast serial mode.
            ''' </summary>
            Public IFAIsFastSer As Boolean
            ''' <summary>
            ''' Determines if channel A loads the VCP driver.
            ''' </summary>
            Public AIsVCP As Boolean = True
            ''' <summary>
            ''' Determines if channel B is in FIFO mode.
            ''' </summary>
            Public IFBIsFifo As Boolean
            ''' <summary>
            ''' Determines if channel B is in FIFO target mode.
            ''' </summary>
            Public IFBIsFifoTar As Boolean
            ''' <summary>
            ''' Determines if channel B is in fast serial mode.
            ''' </summary>
            Public IFBIsFastSer As Boolean
            ''' <summary>
            ''' Determines if channel B loads the VCP driver.
            ''' </summary>
            Public BIsVCP As Boolean = True
            ''' <summary>
            ''' For self-powered designs, keeps the FT2232H in low power state until BCBUS7 is high.
            ''' </summary>
            Public PowerSaveEnable As Boolean
        End Class

        Public Class FT4232H_EEPROM_STRUCTURE : Inherits FT_EEPROM_DATA
            ''' <summary>
            ''' Determines if IOs are pulled down when the device is in suspend.
            ''' </summary>
            Public PullDownEnable As Boolean
            ''' <summary>
            ''' Determines if the serial number is enabled.
            ''' </summary>
            Public SerNumEnable As Boolean = True
            ''' <summary>
            ''' Determines if A pins have a slow slew rate.
            ''' </summary>
            Public ASlowSlew As Boolean
            ''' <summary>
            ''' Determines if the A pins have a Schmitt input.
            ''' </summary>
            Public ASchmittInput As Boolean
            ''' <summary>
            ''' Determines the A pins drive current in mA.
            ''' </summary>
            Public ADriveCurrent As FT_DRIVE_CURRENT = FT_DRIVE_CURRENT.FT_DRIVE_CURRENT_4MA
            ''' <summary>
            ''' Determines if B pins have a slow slew rate.
            ''' </summary>
            Public BSlowSlew As Boolean
            ''' <summary>
            ''' Determines if the B pins have a Schmitt input.
            ''' </summary>
            Public BSchmittInput As Boolean
            ''' <summary>
            ''' Determines the B pins drive current in mA.
            ''' </summary>
            Public BDriveCurrent As FT_DRIVE_CURRENT = FT_DRIVE_CURRENT.FT_DRIVE_CURRENT_4MA
            ''' <summary>
            ''' Determines if C pins have a slow slew rate.
            ''' </summary>
            Public CSlowSlew As Boolean
            ''' <summary>
            ''' Determines if the C pins have a Schmitt input.
            ''' </summary>
            Public CSchmittInput As Boolean
            ''' <summary>
            ''' Determines the C pins drive current in mA.
            ''' </summary>
            Public CDriveCurrent As FT_DRIVE_CURRENT = FT_DRIVE_CURRENT.FT_DRIVE_CURRENT_4MA
            ''' <summary>
            ''' Determines if D pins have a slow slew rate.
            ''' </summary>
            Public DSlowSlew As Boolean
            ''' <summary>
            ''' Determines if the D pins have a Schmitt input.
            ''' </summary>
            Public DSchmittInput As Boolean
            ''' <summary>
            ''' Determines the D pins drive current in mA.
            ''' </summary>
            Public DDriveCurrent As FT_DRIVE_CURRENT = FT_DRIVE_CURRENT.FT_DRIVE_CURRENT_4MA
            ''' <summary>
            ''' RI of port A acts as RS485 transmit enable (TXDEN).
            ''' </summary>
            Public ARIIsTXDEN As Boolean
            ''' <summary>
            ''' RI of port B acts as RS485 transmit enable (TXDEN).
            ''' </summary>
            Public BRIIsTXDEN As Boolean
            ''' <summary>
            ''' RI of port C acts as RS485 transmit enable (TXDEN).
            ''' </summary>
            Public CRIIsTXDEN As Boolean
            ''' <summary>
            ''' RI of port D acts as RS485 transmit enable (TXDEN).
            ''' </summary>
            Public DRIIsTXDEN As Boolean
            ''' <summary>
            ''' Determines if channel A loads the VCP driver.
            ''' </summary>
            Public AIsVCP As Boolean = True
            ''' <summary>
            ''' Determines if channel B loads the VCP driver.
            ''' </summary>
            Public BIsVCP As Boolean = True
            ''' <summary>
            ''' Determines if channel C loads the VCP driver.
            ''' </summary>
            Public CIsVCP As Boolean = True
            ''' <summary>
            ''' Determines if channel D loads the VCP driver.
            ''' </summary>
            Public DIsVCP As Boolean = True
        End Class

        Public Class FT232H_EEPROM_STRUCTURE : Inherits FT_EEPROM_DATA
            ''' <summary>
            ''' Determines if IOs are pulled down when the device is in suspend.
            ''' </summary>
            Public PullDownEnable As Boolean
            ''' <summary>
            ''' Determines if the serial number is enabled.
            ''' </summary>
            Public SerNumEnable As Boolean = True
            ''' <summary>
            ''' Determines if AC pins have a slow slew rate.
            ''' </summary>
            Public ACSlowSlew As Boolean
            ''' <summary>
            ''' Determines if the AC pins have a Schmitt input.
            ''' </summary>
            Public ACSchmittInput As Boolean
            ''' <summary>
            ''' Determines the AC pins drive current in mA.
            ''' </summary>
            Public ACDriveCurrent As FT_DRIVE_CURRENT = FT_DRIVE_CURRENT.FT_DRIVE_CURRENT_4MA
            ''' <summary>
            ''' Determines if AD pins have a slow slew rate.
            ''' </summary>
            Public ADSlowSlew As Boolean
            ''' <summary>
            ''' Determines if the AD pins have a Schmitt input.
            ''' </summary>
            Public ADSchmittInput As Boolean
            ''' <summary>
            ''' Determines the AD pins drive current in mA.
            ''' </summary>
            Public ADDriveCurrent As FT_DRIVE_CURRENT = FT_DRIVE_CURRENT.FT_DRIVE_CURRENT_4MA
            ''' <summary>
            ''' Sets the function of the CBUS0 pin for FT232H devices.
            ''' </summary>
            Public Cbus0 As FT_232H_CBUS_OPTIONS = FT_232H_CBUS_OPTIONS.FT_CBUS_TRISTATE
            ''' <summary>
            ''' Sets the function of the CBUS1 pin for FT232H devices.
            ''' </summary>
            Public Cbus1 As FT_232H_CBUS_OPTIONS = FT_232H_CBUS_OPTIONS.FT_CBUS_TRISTATE
            ''' <summary>
            ''' Sets the function of the CBUS2 pin for FT232H devices.
            ''' </summary>
            Public Cbus2 As FT_232H_CBUS_OPTIONS = FT_232H_CBUS_OPTIONS.FT_CBUS_TRISTATE
            ''' <summary>
            ''' Sets the function of the CBUS3 pin for FT232H devices.
            ''' </summary>
            Public Cbus3 As FT_232H_CBUS_OPTIONS = FT_232H_CBUS_OPTIONS.FT_CBUS_TRISTATE
            ''' <summary>
            ''' Sets the function of the CBUS4 pin for FT232H devices.
            ''' </summary>
            Public Cbus4 As FT_232H_CBUS_OPTIONS = FT_232H_CBUS_OPTIONS.FT_CBUS_TRISTATE
            ''' <summary>
            ''' Sets the function of the CBUS5 pin for FT232H devices.
            ''' </summary>
            Public Cbus5 As FT_232H_CBUS_OPTIONS = FT_232H_CBUS_OPTIONS.FT_CBUS_TRISTATE
            ''' <summary>
            ''' Sets the function of the CBUS6 pin for FT232H devices.
            ''' </summary>
            Public Cbus6 As FT_232H_CBUS_OPTIONS = FT_232H_CBUS_OPTIONS.FT_CBUS_TRISTATE
            ''' <summary>
            ''' Sets the function of the CBUS7 pin for FT232H devices.
            ''' </summary>
            Public Cbus7 As FT_232H_CBUS_OPTIONS = FT_232H_CBUS_OPTIONS.FT_CBUS_TRISTATE
            ''' <summary>
            ''' Sets the function of the CBUS8 pin for FT232H devices.
            ''' </summary>
            Public Cbus8 As FT_232H_CBUS_OPTIONS = FT_232H_CBUS_OPTIONS.FT_CBUS_TRISTATE
            ''' <summary>
            ''' Sets the function of the CBUS9 pin for FT232H devices.
            ''' </summary>
            Public Cbus9 As FT_232H_CBUS_OPTIONS = FT_232H_CBUS_OPTIONS.FT_CBUS_TRISTATE
            ''' <summary>
            ''' Determines if the device is in FIFO mode.
            ''' </summary>
            Public IsFifo As Boolean
            ''' <summary>
            ''' Determines if the device is in FIFO target mode.
            ''' </summary>
            Public IsFifoTar As Boolean
            ''' <summary>
            ''' Determines if the device is in fast serial mode.
            ''' </summary>
            Public IsFastSer As Boolean
            ''' <summary>
            ''' Determines if the device is in FT1248 mode.
            ''' </summary>
            Public IsFT1248 As Boolean
            ''' <summary>
            ''' Determines FT1248 mode clock polarity.
            ''' </summary>
            Public FT1248Cpol As Boolean
            ''' <summary>
            ''' Determines if data is ent MSB (0) or LSB (1) in FT1248 mode.
            ''' </summary>
            Public FT1248Lsb As Boolean
            ''' <summary>
            ''' Determines if FT1248 mode uses flow control.
            ''' </summary>
            Public FT1248FlowControl As Boolean
            ''' <summary>
            ''' Determines if the VCP driver is loaded.
            ''' </summary>
            Public IsVCP As Boolean = True
            ''' <summary>
            ''' For self-powered designs, keeps the FT232H in low power state until ACBUS7 is high.
            ''' </summary>
            Public PowerSaveEnable As Boolean
        End Class

        Public Class FT_XSERIES_EEPROM_STRUCTURE : Inherits FT_EEPROM_DATA
            ''' <summary>
            ''' Determines if IOs are pulled down when the device Is in suspend.
            ''' </summary>
            Public PullDownEnable As Boolean
            ''' <summary>
            ''' Determines if the serial number is enabled.
            ''' </summary>
            Public SerNumEnable As Boolean = True
            ''' <summary>
            ''' Determines if the USB version number is enabled.
            ''' </summary>
            Public USBVersionEnable As Boolean = True
            ''' <summary>
            ''' The USB version number.
            ''' </summary>
            Public USBVersion As USB_VERSION = USB_VERSION.VER_20
            ''' <summary>
            ''' Determines if AC pins have a slow slew rate.
            ''' </summary>
            Public ACSlowSlew As Byte
            ''' <summary>
            ''' Determines if the AC pins have a Schmitt input.
            ''' </summary>
            Public ACSchmittInput As Byte
            ''' <summary>
            ''' Determines the AC pins drive current in mA.
            ''' </summary>
            Public ACDriveCurrent As FT_DRIVE_CURRENT = FT_DRIVE_CURRENT.FT_DRIVE_CURRENT_4MA
            ''' <summary>
            ''' Determines if AD pins have a slow slew rate.
            ''' </summary>
            Public ADSlowSlew As Byte
            ''' <summary>
            ''' Determines if AD pins have a schmitt input.
            ''' </summary>
            Public ADSchmittInput As Byte
            ''' <summary>
            ''' Determines the AD pins drive current in mA.
            ''' </summary>
            Public ADDriveCurrent As FT_DRIVE_CURRENT = FT_DRIVE_CURRENT.FT_DRIVE_CURRENT_4MA
            ''' <summary>
            ''' Sets the function of the CBUS0 pin for FT232H devices.
            ''' </summary>
            Public Cbus0 As FT_XSERIES_CBUS_OPTIONS = FT_XSERIES_CBUS_OPTIONS.FT_CBUS_TRISTATE
            ''' <summary>
            ''' Sets the function of the CBUS1 pin for FT232H devices.
            ''' </summary>
            Public Cbus1 As FT_XSERIES_CBUS_OPTIONS = FT_XSERIES_CBUS_OPTIONS.FT_CBUS_TRISTATE
            ''' <summary>
            ''' Sets the function of the CBUS2 pin for FT232H devices.
            ''' </summary>
            Public Cbus2 As FT_XSERIES_CBUS_OPTIONS = FT_XSERIES_CBUS_OPTIONS.FT_CBUS_TRISTATE
            ''' <summary>
            ''' Sets the function of the CBUS3 pin for FT232H devices.
            ''' </summary>
            Public Cbus3 As FT_XSERIES_CBUS_OPTIONS = FT_XSERIES_CBUS_OPTIONS.FT_CBUS_TRISTATE
            ''' <summary>
            ''' Sets the function of the CBUS4 pin for FT232H devices.
            ''' </summary>
            Public Cbus4 As FT_XSERIES_CBUS_OPTIONS = FT_XSERIES_CBUS_OPTIONS.FT_CBUS_TRISTATE
            ''' <summary>
            ''' Sets the function of the CBUS5 pin for FT232H devices.
            ''' </summary>
            Public Cbus5 As FT_XSERIES_CBUS_OPTIONS = FT_XSERIES_CBUS_OPTIONS.FT_CBUS_TRISTATE
            ''' <summary>
            ''' Sets the function of the CBUS6 pin for FT232H devices.
            ''' </summary>
            Public Cbus6 As FT_XSERIES_CBUS_OPTIONS = FT_XSERIES_CBUS_OPTIONS.FT_CBUS_TRISTATE
            ''' <summary>
            ''' Inverts the sense of the TXD line.
            ''' </summary>
            Public InvertTXD As Byte
            ''' <summary>
            ''' Inverts the sense of the RXD line.
            ''' </summary>
            Public InvertRXD As Byte
            ''' <summary>
            ''' Inverts the sense of the RTS line.
            ''' </summary>
            Public InvertRTS As Byte
            ''' <summary>
            ''' Inverts the sense of the CTS line.
            ''' </summary>
            Public InvertCTS As Byte
            ''' <summary>
            ''' Inverts the sense of the DTR line.
            ''' </summary>
            Public InvertDTR As Byte
            ''' <summary>
            ''' Inverts the sense of the DSR line.
            ''' </summary>
            Public InvertDSR As Byte
            ''' <summary>
            ''' Inverts the sense of the DCD line.
            ''' </summary>
            Public InvertDCD As Byte
            ''' <summary>
            ''' Inverts the sense of the RI line.
            ''' </summary>
            Public InvertRI As Byte
            ''' <summary>
            ''' Determines whether the Battery Charge Detection option is enabled.
            ''' </summary>
            Public BCDEnable As Byte
            ''' <summary>
            ''' Asserts the power enable signal on CBUS when charging port detected.
            ''' </summary>
            Public BCDForceCbusPWREN As Byte
            ''' <summary>
            ''' Forces the device never to go into sleep mode.
            ''' </summary>
            Public BCDDisableSleep As Byte
            ''' <summary>
            ''' I2C slave device address.
            ''' </summary>
            Public I2CSlaveAddress As UShort
            ''' <summary>
            ''' I2C device ID.
            ''' </summary>
            Public I2CDeviceId As UInteger
            ''' <summary>
            ''' Disable I2C Schmitt trigger.
            ''' </summary>
            Public I2CDisableSchmitt As Byte
            ''' <summary>
            ''' FT1248 clock polarity - clock idle high (1) or clock idle low (0).
            ''' </summary>
            Public FT1248Cpol As Byte
            ''' <summary>
            ''' FT1248 data is LSB (1) or MSB (0).
            ''' </summary>
            Public FT1248Lsb As Byte
            ''' <summary>
            ''' FT1248 flow control enable.
            ''' </summary>
            Public FT1248FlowControl As Byte
            ''' <summary>
            ''' Enable RS485 Echo Suppression.
            ''' </summary>
            Public RS485EchoSuppress As Byte
            ''' <summary>
            ''' Enable Power Save mode.
            ''' </summary>
            Public PowerSaveEnable As Byte
            ''' <summary>
            ''' Determines whether the VCP driver is loaded.
            ''' </summary>
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

        Public Enum FT_232H_CBUS_OPTIONS As Byte
            FT_CBUS_TRISTATE = 0
            FT_CBUS_RXLED = 1
            FT_CBUS_TXLED = &H2
            FT_CBUS_TXRXLED = &H3
            FT_CBUS_PWREN = &H4
            FT_CBUS_SLEEP = &H5
            FT_CBUS_DRIVE_0 = &H6
            FT_CBUS_DRIVE_1 = &H7
            FT_CBUS_IOMODE = &H8
            FT_CBUS_TXDEN = &H9
            FT_CBUS_CLK30 = &HA
            FT_CBUS_CLK15 = &HB
            FT_CBUS_CLK7_5 = &HC
        End Enum

        Public Enum FT_XSERIES_CBUS_OPTIONS As Byte
            FT_CBUS_TRISTATE = &H0
            FT_CBUS_RXLED = &H1
            FT_CBUS_TXLED = &H2
            FT_CBUS_TXRXLED = &H3
            FT_CBUS_PWREN = &H4
            FT_CBUS_SLEEP = &H5
            FT_CBUS_Drive_0 = &H6
            FT_CBUS_Drive_1 = &H7
            FT_CBUS_GPIO = &H8
            FT_CBUS_TXDEN = &H9
            FT_CBUS_CLK24MHz = &HA
            FT_CBUS_CLK12MHz = &HB
            FT_CBUS_CLK6MHz = &HC
            FT_CBUS_BCD_Charger = &HD
            FT_CBUS_BCD_Charger_N = &HE
            FT_CBUS_I2C_TXE = &HF
            FT_CBUS_I2C_RXF = &H10
            FT_CBUS_VBUS_Sense = &H11
            FT_CBUS_BitBang_WR = &H12
            FT_CBUS_BitBang_RD = &H13
            FT_CBUS_Time_Stamp = &H14
            FT_CBUS_Keep_Awake = &H15
        End Enum

        Public Enum USB_VERSION As UShort
            VER_11 = &H110
            VER_20 = &H200
        End Enum

#End Region '/ПЕРЕЧИСЛЕНИЯ

    End Class '/Eeprom

End Namespace
