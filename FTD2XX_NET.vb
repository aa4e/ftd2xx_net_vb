Imports System.Runtime.InteropServices
Imports System.Runtime.Serialization
Imports System.Threading
Imports System.Text

Namespace FTD2XX_NET

    ''' <summary>
    ''' Класс для работы с устройствами FTDI.
    ''' </summary>
    Public Class Ftdi
        Implements IDisposable

#Region "CTORs"

        ''' <summary>
        ''' Конструктор по умолчанию, который не открывает устройство.
        ''' </summary>
        Public Sub New()
            'ничего не делает
        End Sub

        ''' <summary>
        ''' Конструктор, который открывает заданное устройство по индексу.
        ''' </summary>
        ''' <param name="index">Индекс устройства.</param>
        Public Sub New(index As Integer)
            Me.New()
            OpenByIndex(CUInt(index))
        End Sub

        ''' <summary>
        ''' Конструктор, который открывает заданное устройство по серийному номеру.
        ''' </summary>
        ''' <param name="serial">Серийный номер. Должен точно соответствовать даже регистром.</param>
        Public Sub New(serial As String)
            Me.New()
            OpenBySerialNumber(serial)
        End Sub

        ''' <summary>
        ''' Конструктор, который сразу открывает заданное устройство по размещению.
        ''' </summary>
        ''' <param name="location ">Размещение устройства.</param>
        Public Sub New(location As UInteger)
            Me.New()
            OpenByLocation(location)
        End Sub

        ''' <summary>
        ''' Освобождает библиотеку ftd2xx.
        ''' </summary>
        Public Shared Sub UnloadLibrary()
            If FreeLibrary(Ftd2xxDllHandle) Then
                _Ftd2xxDllHandle = CLOSED_HANDLE
                Debug.WriteLine("Library unloaded.")
            End If
        End Sub

#End Region '/CTORs

#Region "КОНСТАНТЫ"

        Public Const FT_OPEN_BY_SERIAL_NUMBER As UInteger = 1UI
        Public Const FT_OPEN_BY_DESCRIPTION As UInteger = 2UI
        Public Const FT_OPEN_BY_LOCATION As UInteger = 4UI
        Public Const FT_DEFAULT_BAUD_RATE As UInteger = 9600UI
        Public Const FT_COM_PORT_NOT_ASSIGNED As Integer = -1
        Public Const FT_DEFAULT_DEADMAN_TIMEOUT As UInteger = 5000UI
        Public Const FT_DEFAULT_LATENCY As Byte = 16
        Public Const FT_DEFAULT_IN_TRANSFER_SIZE As UInteger = 4096UI
        Public Const FT_DEFAULT_OUT_TRANSFER_SIZE As UInteger = 4096UI

        Private Const CLOSED_HANDLE As Integer = 0
        Private Const NOT_INIT As Integer = -1

#End Region '/КОНСТАНТЫ

#Region "READ-ONLY PROPS"

        ''' <summary>
        ''' Путь к библиотеке ftd2xx.
        ''' </summary>
        Public Shared ReadOnly Property LibPath As String
            Get
                Return _LibPath
            End Get
        End Property
        Private Shared _LibPath As String = "C:\Temp\ftd2xx.dll"

        ''' <summary>
        ''' Объект для работы с ПЗУ.
        ''' </summary>
        Public ReadOnly Property Eeprom As FTD2XX_NET.Eeprom
            Get
                If (_Eeprom Is Nothing) Then
                    _Eeprom = New Eeprom(Ftd2xxDllHandle, FtHandle, DeviceType)
                End If
                Return _Eeprom
            End Get
        End Property
        Private _Eeprom As FTD2XX_NET.Eeprom

        ''' <summary>
        ''' Загружена ли библиотека.
        ''' </summary>
        Public Shared ReadOnly Property IsLibraryLoaded As Boolean
            Get
                Return (Ftd2xxDllHandle <> CLOSED_HANDLE)
            End Get
        End Property

        ''' <summary>
        ''' Открыто ли устройство.
        ''' </summary>
        Public ReadOnly Property IsDeviceOpen As Boolean
            Get
                Return (FtHandle <> CLOSED_HANDLE)
            End Get
        End Property

        ''' <summary>
        ''' Признак канала (A, B и т.д.).
        ''' </summary>
        Private ReadOnly Property InterfaceIdentifier As String
            Get
                If (_InterfaceIdentifier.Length = 0) AndAlso IsLibraryLoaded AndAlso IsDeviceOpen Then
                    Dim devType As FT_DEVICE = DeviceType
                    If (devType = FT_DEVICE.FT_DEVICE_2232) OrElse (devType = FT_DEVICE.FT_DEVICE_2232H) OrElse (devType = FT_DEVICE.FT_DEVICE_4232H) Then
                        _InterfaceIdentifier = Description().Substring(Description.Length - 1)
                    End If
                End If
                Return _InterfaceIdentifier
            End Get
        End Property
        Private _InterfaceIdentifier As String = ""

        ''' <summary>
        ''' Тип текущего чипа FTDI.
        ''' </summary>
        Public ReadOnly Property DeviceType As FT_DEVICE
            Get
                If (_DeviceType = FT_DEVICE.FT_DEVICE_UNKNOWN) Then
                    Dim devType As FT_DEVICE = FT_DEVICE.FT_DEVICE_UNKNOWN
                    Dim num As UInteger = 0
                    Dim sn As Byte() = New Byte(15) {}
                    Dim descr As Byte() = New Byte(63) {}
                    Dim dlg As GetDeviceInfoDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_GetDeviceInfo, IntPtr), GetType(GetDeviceInfoDelegate)), GetDeviceInfoDelegate)
                    Dim status As FT_STATUS = dlg(FtHandle, devType, num, sn, descr, 0)
                    _DeviceType = devType
                End If
                Return _DeviceType
            End Get
        End Property
        Private _DeviceType As FT_DEVICE = FT_DEVICE.FT_DEVICE_UNKNOWN

        ''' <summary>
        ''' Vendor ID и Product ID текущего устройства.
        ''' </summary>
        Public ReadOnly Property DeviceID As UInteger
            Get
                If (_DeviceID = 0) AndAlso IsLibraryLoaded AndAlso IsDeviceOpen Then
                    Dim devid As UInteger
                    Dim dev As FT_DEVICE = FT_DEVICE.FT_DEVICE_UNKNOWN
                    Dim sn As Byte() = New Byte(15) {}
                    Dim descr As Byte() = New Byte(63) {}
                    Dim dlg As GetDeviceInfoDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_GetDeviceInfo, IntPtr), GetType(GetDeviceInfoDelegate)), GetDeviceInfoDelegate)
                    Dim status As FT_STATUS = dlg(FtHandle, dev, devid, sn, descr, 0)
                    _DeviceID = devid
                End If
                Return _DeviceID
            End Get
        End Property
        Private _DeviceID As UInteger = 0

        ''' <summary>
        ''' Описание текущего устройства.
        ''' </summary>
        Public ReadOnly Property Description As String
            Get
                If (_Description.Length = 0) AndAlso IsLibraryLoaded AndAlso IsDeviceOpen Then
                    Dim sn As Byte() = New Byte(15) {}
                    Dim descr As Byte() = New Byte(63) {}
                    Dim num As UInteger = 0UI
                    Dim dev As FT_DEVICE = FT_DEVICE.FT_DEVICE_UNKNOWN
                    Dim dlg As GetDeviceInfoDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_GetDeviceInfo, IntPtr), GetType(GetDeviceInfoDelegate)), GetDeviceInfoDelegate)
                    Dim status As FT_STATUS = dlg(FtHandle, dev, num, sn, descr, 0)
                    Dim d As String = Encoding.ASCII.GetString(descr)
                    _Description = d.Substring(0, d.IndexOf(vbNullChar))
                End If
                Return _Description
            End Get
        End Property
        Private _Description As String = String.Empty

        ''' <summary>
        ''' Серийный номер устройства.
        ''' </summary>
        Public ReadOnly Property SerialNumber As String
            Get
                If (_SerialNumber.Length = 0) AndAlso IsLibraryLoaded AndAlso IsDeviceOpen Then
                    Dim sn As Byte() = New Byte(15) {}
                    Dim descr As Byte() = New Byte(63) {}
                    Dim num As UInteger = 0UI
                    Dim dev As FT_DEVICE = FT_DEVICE.FT_DEVICE_UNKNOWN
                    Dim dlg As GetDeviceInfoDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_GetDeviceInfo, IntPtr), GetType(GetDeviceInfoDelegate)), GetDeviceInfoDelegate)
                    Dim status As FT_STATUS = dlg(FtHandle, dev, num, sn, descr, 0)
                    Dim serNum As String = Encoding.ASCII.GetString(sn)
                    _SerialNumber = serNum.Substring(0, serNum.IndexOf(vbNullChar))
                End If
                Return _SerialNumber
            End Get
        End Property
        Private _SerialNumber As String = String.Empty

        ''' <summary>
        ''' Текущая версия драйвера FTDIBUS.SYS.
        ''' </summary>
        Public ReadOnly Property DriverVersion As UInteger
            Get
                If (_DriverVersion = 0) AndAlso (Pft_GetDriverVersion <> NOT_INIT) Then
                    Dim dlg As GetDriverVersionDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_GetDriverVersion), GetType(GetDriverVersionDelegate)), GetDriverVersionDelegate)
                    Dim status As FT_STATUS = dlg(FtHandle, _DriverVersion)
                End If
                Return _DriverVersion
            End Get
        End Property
        Private _DriverVersion As UInteger = 0

        ''' <summary>
        ''' Текущая версия драйвера FTD2XX.DLL.
        ''' </summary>
        Public Shared ReadOnly Property LibraryVersion As UInteger
            Get
                If (_LibraryVersion = 0) Then
                    Dim pft_GetLibraryVersion As Integer = GetProcAddress(Ftd2xxDllHandle, "FT_GetLibraryVersion")
                    Dim dlg As GetLibraryVersionDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pft_GetLibraryVersion), GetType(GetLibraryVersionDelegate)), GetLibraryVersionDelegate)
                    Dim status As FT_STATUS = dlg.Invoke(_LibraryVersion)
                End If
                Return _LibraryVersion
            End Get
        End Property
        Private Shared _LibraryVersion As UInteger = 0

#End Region '/READ-ONLY PROPS

#Region "ОТКРЫТИЕ, ЗАКРЫТИЕ"

        'Дескриптор устройства:
        Private FtHandle As Integer = CLOSED_HANDLE

        ''' <summary>
        ''' Открывает устройство FTDI по индексу <paramref name="index"/>.
        ''' </summary>
        ''' <remarks>
        ''' Это не гарантирует открытие заданного устройства.
        ''' </remarks>
        Public Sub OpenByIndex(index As UInteger)
            Dim dlg As OpenDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_Open), GetType(OpenDelegate)), OpenDelegate)
            Dim status As FT_STATUS = dlg(index, FtHandle)
            If (status = FT_STATUS.FT_OK) Then
                Dim uWordLength As Byte = 8
                Dim uStopBits As Byte = 0
                Dim uParity As Byte = 0
                Dim dDlg As SetDataCharacteristicsDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_SetDataCharacteristics), GetType(SetDataCharacteristicsDelegate)), SetDataCharacteristicsDelegate)
                status = dDlg(FtHandle, uWordLength, uStopBits, uParity)
                CheckErrors(status)

                Dim usFlowControl As UShort = 0US
                Dim uXon As Byte = 17
                Dim uXoff As Byte = 19
                Dim fDlg As SetFlowControlDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_SetFlowControl), GetType(SetFlowControlDelegate)), SetFlowControlDelegate)
                status = fDlg(FtHandle, usFlowControl, uXon, uXoff)
                CheckErrors(status)

                Dim dwBaudRate As UInteger = FT_DEFAULT_BAUD_RATE
                Dim bDlg As SetBaudRateDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_SetBaudRate), GetType(SetBaudRateDelegate)), SetBaudRateDelegate)
                status = bDlg(FtHandle, dwBaudRate)
            Else
                FtHandle = CLOSED_HANDLE
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Открывает устройство FTDI по серийному номеру <paramref name="serialnumber"/>.
        ''' </summary>
        Public Sub OpenBySerialNumber(serialnumber As String)
            Dim dlg As OpenExDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_OpenEx, IntPtr), GetType(OpenExDelegate)), OpenExDelegate)
            Dim status As FT_STATUS = dlg(serialnumber, FT_OPEN_BY_SERIAL_NUMBER, FtHandle)
            If (status = FT_STATUS.FT_OK) Then
                Dim uWordLength As Byte = 8
                Dim uStopBits As Byte = 0
                Dim uParity As Byte = 0
                Dim dDlg As SetDataCharacteristicsDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_SetDataCharacteristics, IntPtr), GetType(SetDataCharacteristicsDelegate)), SetDataCharacteristicsDelegate)
                status = dDlg(FtHandle, uWordLength, uStopBits, uParity)
                CheckErrors(status)

                Dim usFlowControl As UShort = 0US
                Dim uXon As Byte = 17
                Dim uXoff As Byte = 19
                Dim fDlg As SetFlowControlDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_SetFlowControl, IntPtr), GetType(SetFlowControlDelegate)), SetFlowControlDelegate)
                status = fDlg(FtHandle, usFlowControl, uXon, uXoff)
                CheckErrors(status)

                Dim dwBaudRate As UInteger = FT_DEFAULT_BAUD_RATE
                Dim bDlg As SetBaudRateDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_SetBaudRate, IntPtr), GetType(SetBaudRateDelegate)), SetBaudRateDelegate)
                status = bDlg(FtHandle, dwBaudRate)
            Else
                FtHandle = CLOSED_HANDLE
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Открывает устройство FTDI по описанию <paramref name="description"/>.
        ''' </summary>
        Public Sub OpenByDescription(description As String)
            Dim dlg As OpenExDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_OpenEx, IntPtr), GetType(OpenExDelegate)), OpenExDelegate)
            Dim status As FT_STATUS = dlg(description, FT_OPEN_BY_DESCRIPTION, FtHandle)
            If (status = FT_STATUS.FT_OK) Then
                Dim uWordLength As Byte = 8
                Dim uStopBits As Byte = 0
                Dim uParity As Byte = 0
                Dim dDlg As SetDataCharacteristicsDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_SetDataCharacteristics, IntPtr), GetType(SetDataCharacteristicsDelegate)), SetDataCharacteristicsDelegate)
                status = dDlg(FtHandle, uWordLength, uStopBits, uParity)
                CheckErrors(status)

                Dim usFlowControl As UShort = 0US
                Dim uXon As Byte = 17
                Dim uXoff As Byte = 19
                Dim fDlg As SetFlowControlDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_SetFlowControl, IntPtr), GetType(SetFlowControlDelegate)), SetFlowControlDelegate)
                status = fDlg(FtHandle, usFlowControl, uXon, uXoff)
                CheckErrors(status)

                Dim dwBaudRate As UInteger = FT_DEFAULT_BAUD_RATE
                Dim bDlg As SetBaudRateDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_SetBaudRate, IntPtr), GetType(SetBaudRateDelegate)), SetBaudRateDelegate)
                status = bDlg(FtHandle, dwBaudRate)
            Else
                FtHandle = CLOSED_HANDLE
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Открывает устройство FTDI по физическому расположению <paramref name="location"/>.
        ''' </summary>
        Public Sub OpenByLocation(location As UInteger)
            Dim dlg As OpenExLocDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_OpenEx, IntPtr), GetType(OpenExLocDelegate)), OpenExLocDelegate)
            Dim status As FT_STATUS = dlg(location, FT_OPEN_BY_LOCATION, FtHandle)
            If (status = FT_STATUS.FT_OK) Then
                Dim uWordLength As Byte = 8
                Dim uStopBits As Byte = 0
                Dim uParity As Byte = 0
                Dim dDlg As SetDataCharacteristicsDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_SetDataCharacteristics, IntPtr), GetType(SetDataCharacteristicsDelegate)), SetDataCharacteristicsDelegate)
                status = dDlg(FtHandle, uWordLength, uStopBits, uParity)
                CheckErrors(status)

                Dim usFlowControl As UShort = 0US
                Dim uXon As Byte = 17
                Dim uXoff As Byte = 19
                Dim fDlg As SetFlowControlDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_SetFlowControl, IntPtr), GetType(SetFlowControlDelegate)), SetFlowControlDelegate)
                status = fDlg(FtHandle, usFlowControl, uXon, uXoff)
                CheckErrors(status)

                Dim dwBaudRate As UInteger = FT_DEFAULT_BAUD_RATE
                Dim bDlg As SetBaudRateDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_SetBaudRate, IntPtr), GetType(SetBaudRateDelegate)), SetBaudRateDelegate)
                status = bDlg(FtHandle, dwBaudRate)
            Else
                FtHandle = CLOSED_HANDLE
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Закрывает открытое устройство.
        ''' </summary>
        Public Sub Close()
            Dim dlg As CloseDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_Close), GetType(CloseDelegate)), CloseDelegate)
            Dim status As FT_STATUS = dlg.Invoke(FtHandle)
            CheckErrors(status)
            If (status = FT_STATUS.FT_OK) Then
                FtHandle = CLOSED_HANDLE
            End If
        End Sub

#End Region '/ОТКРЫТИЕ, ЗАКРЫТИЕ

#Region "ИНФО"

        ''' <summary>
        ''' Возвращает число доступных устройств FTDI.
        ''' </summary>
        Public Shared Function GetNumberOfDevices() As Integer
            Dim devCount As UInteger = 0
            Dim dlg As CreateDeviceInfoListDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_CreateDeviceInfoList), GetType(CreateDeviceInfoListDelegate)), CreateDeviceInfoListDelegate)
            Dim status As FT_STATUS = dlg.Invoke(devCount)
            CheckErrors(status)
            Return CInt(devCount)
        End Function

        ''' <summary>
        ''' Возвращает информацию обо всех доступных устройствах FTDI.
        ''' </summary>
        Public Shared Function GetDeviceList() As FT_DEVICE_INFO_NODE()
            Dim numDevices As UInteger = 0UI
            Dim dlg As CreateDeviceInfoListDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_CreateDeviceInfoList), GetType(CreateDeviceInfoListDelegate)), CreateDeviceInfoListDelegate)
            Dim status As FT_STATUS = dlg(numDevices)
            CheckErrors(status)
            If (numDevices > 0) Then
                Dim devicelist(CInt(numDevices - 1)) As FT_DEVICE_INFO_NODE
                Dim gdiDeleg As GetDeviceInfoDetailDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_GetDeviceInfoDetail), GetType(GetDeviceInfoDetailDelegate)), GetDeviceInfoDetailDelegate)
                For i As Integer = 0 To CInt(numDevices - 1)
                    devicelist(i) = New FT_DEVICE_INFO_NODE()
                    Dim sn As Byte() = New Byte(15) {}
                    Dim descr As Byte() = New Byte(63) {}
                    status = gdiDeleg(CUInt(i), devicelist(i).Flags, devicelist(i).Type, devicelist(i).ID, devicelist(i).LocId, sn, descr, devicelist(i).FtHandle)
                    CheckErrors(status)
                    devicelist(i).SerialNumber = Encoding.ASCII.GetString(sn)
                    devicelist(i).Description = Encoding.ASCII.GetString(descr)
                    Dim endOfStringIndex As Integer = devicelist(i).SerialNumber.IndexOf(vbNullChar)
                    If (endOfStringIndex <> -1) Then
                        devicelist(i).SerialNumber = devicelist(i).SerialNumber.Substring(0, endOfStringIndex)
                    End If
                    endOfStringIndex = devicelist(i).Description.IndexOf(vbNullChar)
                    If (endOfStringIndex <> -1) Then
                        devicelist(i).Description = devicelist(i).Description.Substring(0, endOfStringIndex)
                    End If
                Next
                Return devicelist
            End If
            Return New FT_DEVICE_INFO_NODE() {}
        End Function

        ''' <summary>
        ''' Возвращает число доступных байтов в приёмном буфере.
        ''' </summary>
        Public Function GetRxBytesAvailable() As Integer
            Dim rxQueue As UInteger
            Dim dlg As GetQueueStatusDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_GetQueueStatus, IntPtr), GetType(GetQueueStatusDelegate)), GetQueueStatusDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, rxQueue)
            CheckErrors(status)
            Return CInt(rxQueue)
        End Function

        ''' <summary>
        ''' Возвращает ожидающее число байтов в передающем буфере.
        ''' </summary>
        Public Function GetTxBytesWaiting() As Integer
            Dim txQueue As UInteger
            Dim inTxLen As UInteger = 0UI
            Dim eventType As UInteger = 0UI
            Dim dlg As GetStatusDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_GetStatus, IntPtr), GetType(GetStatusDelegate)), GetStatusDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, inTxLen, txQueue, eventType)
            CheckErrors(status)
            Return CInt(txQueue)
        End Function

        ''' <summary>
        ''' Возвращает тип события после его возникновения.
        ''' </summary>
        ''' <remarks>
        ''' Используется чтобы определить, какое событие сработало, когда ожидаются осбытия разных типов.
        ''' </remarks>
        Public Function GetEventType() As UInteger
            Dim eventType As UInteger
            Dim inTxLen As UInteger = 0UI
            Dim txQueue As UInteger = 0UI
            Dim dlg As GetStatusDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_GetStatus, IntPtr), GetType(GetStatusDelegate)), GetStatusDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, inTxLen, txQueue, eventType)
            CheckErrors(status)
            Return eventType
        End Function

        ''' <summary>
        ''' Возвращает текущее состояние модема (битовую маску).
        ''' </summary>
        Public Function GetModemStatus() As Byte
            Dim s As UInteger = 0UI
            Dim dlg As GetModemStatusDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_GetModemStatus, IntPtr), GetType(GetModemStatusDelegate)), GetModemStatusDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, s)
            CheckErrors(status)
            Dim modemStatus As Byte = Convert.ToByte(s And &HFF)
            Return modemStatus
        End Function

        ''' <summary>
        ''' Возвращает текущее состояние линии (битовую маску).
        ''' </summary>
        Public Function GetLineStatus() As Byte
            Dim s As UInteger = 0UI
            Dim dlg As GetModemStatusDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_GetModemStatus, IntPtr), GetType(GetModemStatusDelegate)), GetModemStatusDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, s)
            CheckErrors(status)
            Dim lineStatus As Byte = Convert.ToByte((s >> 8) And &HFF)
            Return lineStatus
        End Function

        ''' <summary>
        ''' Возвращает назначенный порт или пустую строку, если порт не был назначен.
        ''' </summary>
        Public Function GetComPort() As String
            Dim portNum As Integer = FT_COM_PORT_NOT_ASSIGNED
            Dim dlg As GetComPortNumberDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_GetComPortNumber, IntPtr), GetType(GetComPortNumberDelegate)), GetComPortNumberDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, portNum)
            CheckErrors(status)
            If (portNum <> FT_COM_PORT_NOT_ASSIGNED) Then
                Return $"COM{portNum}"
            End If
            Return String.Empty
        End Function

        ''' <summary>
        ''' Возвращает состояние выводов входа/выхода (битовую маску).
        ''' </summary>
        Public Function GetPinStates() As Byte
            Dim bitMode As Byte
            Dim dlg As GetBitModeDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_GetBitMode, IntPtr), GetType(GetBitModeDelegate)), GetBitModeDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, bitMode)
            CheckErrors(status)
            Return bitMode
        End Function

#End Region '/ИНФО

#Region "ПАРАМЕТРЫ"

        ''' <summary>
        ''' Задаёт битрейт.
        ''' </summary>
        ''' <param name="baudRate">Желаемый битрейт устройства, бит/сек.</param>
        Public Sub SetBaudRate(Optional baudRate As Integer = FT_DEFAULT_BAUD_RATE)
            Dim dlg As SetBaudRateDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_SetBaudRate, IntPtr), GetType(SetBaudRateDelegate)), SetBaudRateDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, CUInt(baudRate))
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Устанавливает число битов данных, стоповых и тип проверки чётности.
        ''' </summary>
        ''' <param name="dataBits">Число битов данных.</param>
        ''' <param name="stopBits">Число стоповых бит.</param>
        ''' <param name="parity">Тип проверки чётности.</param>
        Public Sub SetDataCharacteristics(dataBits As FT_DATA_BITS, stopBits As FT_STOP_BITS, parity As FT_PARITY)
            Dim dlg As SetDataCharacteristicsDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_SetDataCharacteristics, IntPtr), GetType(SetDataCharacteristicsDelegate)), SetDataCharacteristicsDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, dataBits, stopBits, parity)
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Устанавливает тип управления потоком данных.
        ''' </summary>
        ''' <param name="flowControl">Тип управления потоком.</param>
        ''' <param name="xOn">Символ Xon для Xon/Xoff. Игнорируется, если Xon/XOff не используется.</param>
        ''' <param name="xOff">Символ Xoff для Xon/Xoff. Игнорируется, если Xon/XOff не используется.</param>
        Public Sub SetFlowControl(flowControl As FT_FLOW_CONTROL, Optional xOn As Byte = 0, Optional xOff As Byte = 0)
            Dim dlg As SetFlowControlDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_SetFlowControl, IntPtr), GetType(SetFlowControlDelegate)), SetFlowControlDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, flowControl, xOn, xOff)
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Выставляет или сбрасывает линию Request To Send (RTS).
        ''' </summary>
        ''' <param name="enable">True выставляет, False сбрасывает RTS.</param>
        Public Sub SetRts(enable As Boolean)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If enable Then
                Dim dlg As SetRtsDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_SetRts, IntPtr), GetType(SetRtsDelegate)), SetRtsDelegate)
                status = dlg.Invoke(FtHandle)
            Else
                Dim dlg As ClrRtsDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_ClrRts, IntPtr), GetType(ClrRtsDelegate)), ClrRtsDelegate)
                status = dlg.Invoke(FtHandle)
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Выставляет или сбрасывает линию Data Terminal Ready (DTR).
        ''' </summary>
        ''' <param name="enable">True выставляет, False сбрасывает DTR.</param>
        Public Sub SetDtr(enable As Boolean)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If enable Then
                Dim dlg As SetDtrDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_SetDtr, IntPtr), GetType(SetDtrDelegate)), SetDtrDelegate)
                status = dlg.Invoke(FtHandle)
            Else
                Dim dlg As ClrDtrDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_ClrDtr, IntPtr), GetType(ClrDtrDelegate)), ClrDtrDelegate)
                status = dlg.Invoke(FtHandle)
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Устанавливает время ожидания чтения и записи.
        ''' </summary>
        ''' <param name="readTimeout">Время ожидания при чтении, мс. Значение "0" означает бесконечное ожидание.</param>
        ''' <param name="writeTimeout">Время ожидания при записи, мс. Значение "0" означает бесконечное ожидание.</param>
        Public Sub SetTimeouts(readTimeout As UInteger, writeTimeout As UInteger)
            Dim dlg As SetTimeoutsDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Pft_SetTimeouts, IntPtr), GetType(SetTimeoutsDelegate)), SetTimeoutsDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, readTimeout, writeTimeout)
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Устанавливает или сбрасывает состояние прерывания.
        ''' </summary>
        ''' <param name="enable">True включает, False выключает.</param>
        Public Sub SetBreak(enable As Boolean)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If enable Then
                Dim dlg As SetBreakOnDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_SetBreakOn), GetType(SetBreakOnDelegate)), SetBreakOnDelegate)
                status = dlg(FtHandle)
            Else
                Dim dlg As SetBreakOffDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_SetBreakOff), GetType(SetBreakOffDelegate)), SetBreakOffDelegate)
                status = dlg(FtHandle)
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Задаёт число повторов (reset pipe retry count). По умолчанию 50.
        ''' </summary>
        ''' <param name="resetPipeRetryCount">Число повторов. В электрически шумных средах бОльшие значения лучше.</param>
        Public Sub SetResetPipeRetryCount(resetPipeRetryCount As Integer)
            Dim dlg As SetResetPipeRetryCountDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_SetResetPipeRetryCount), GetType(SetResetPipeRetryCountDelegate)), SetResetPipeRetryCountDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, CUInt(resetPipeRetryCount))
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Задаёт время простоя шины USB.
        ''' </summary>
        ''' <param name="deadmanTimeout">Время простоя, мс.</param>
        Public Sub SetDeadmanTimeout(Optional deadmanTimeout As UInteger = FT_DEFAULT_DEADMAN_TIMEOUT)
            Dim dlg As SetDeadmanTimeoutDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_SetDeadmanTimeout), GetType(SetDeadmanTimeoutDelegate)), SetDeadmanTimeoutDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, deadmanTimeout)
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Задаёт задержку в шине USB.
        ''' </summary>
        ''' <param name="latency">Задержка, мс.
        ''' Значения 2...255 мс для FT232BM, FT245BM и FT2232.
        ''' Значения 0...255 мс для прочих устройств.</param>
        Public Sub SetLatency(Optional latency As Byte = FT_DEFAULT_LATENCY)
            If (latency < 2) Then
                If ((DeviceType = FT_DEVICE.FT_DEVICE_BM) OrElse (DeviceType = FT_DEVICE.FT_DEVICE_2232)) Then
                    latency = 2
                End If
            End If
            Dim dlg As SetLatencyTimerDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_SetLatencyTimer), GetType(SetLatencyTimerDelegate)), SetLatencyTimerDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, latency)
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Получает значение задержки, мс.
        ''' </summary>
        Public Function GetLatency() As Byte
            Dim latency As Byte
            Dim dlg As GetLatencyTimerDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_GetLatencyTimer), GetType(GetLatencyTimerDelegate)), GetLatencyTimerDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, latency)
            CheckErrors(status)
            Return latency
        End Function

        ''' <summary>
        ''' Задаёт размеры передающего и принимающего буферов USB.
        ''' </summary>
        ''' <param name="inSize">Резмер в байтах входящего буфера USB. Кратно 64. Значение по умолчанию 4 кб.</param>
        ''' <param name="outSize">Резмер в байтах исходящего буфера USB. Кратно 64. Значение по умолчанию 4 кб.</param>
        Public Sub SetTransferSize(Optional inSize As Integer = FT_DEFAULT_IN_TRANSFER_SIZE, Optional outSize As Integer = FT_DEFAULT_OUT_TRANSFER_SIZE)
            Dim dlg As SetUsbParametersDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_SetUSBParameters), GetType(SetUsbParametersDelegate)), SetUsbParametersDelegate)
            Dim status As FT_STATUS = dlg.Invoke(FtHandle, CUInt(inSize), CUInt(outSize))
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Задаёт символы для событий "получение данных" и "возникновение ошибки".
        ''' </summary>
        ''' <param name="eventChar">Символ, который будет вызывать передачу в данных хост при его получении.</param>
        ''' <param name="eventCharEnable">Включает (True) или выключает (False) символ <paramref name="eventChar"/>.</param>
        ''' <param name="errorChar">Символ, который будет вставлен в поток данных, чтобы показать возникновение ошибки.</param>
        ''' <param name="errorCharEnable">Включает (True) или выключает (False) символ <paramref name="errorChar"/>.</param>
        Public Sub SetCharacters(eventChar As Byte, eventCharEnable As Boolean, errorChar As Byte, errorCharEnable As Boolean)
            Dim dlg As SetCharsDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_SetChars), GetType(SetCharsDelegate)), SetCharsDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, eventChar, Convert.ToByte(eventCharEnable), errorChar, Convert.ToByte(errorCharEnable))
            CheckErrors(status)
        End Sub

#End Region '/ПАРАМЕТРЫ

#Region "ЧТЕНИЕ, ЗАПИСЬ"

        ''' <summary>
        ''' Читает данные из открытого устройства.
        ''' </summary>
        ''' <param name="numBytesToRead">Число байтов, которые нужно прочитать.</param>
        Public Function Read(numBytesToRead As Integer) As Byte()
            Dim data(numBytesToRead - 1) As Byte
            If (data.Length < numBytesToRead) Then
                numBytesToRead = data.Length
            End If
            Dim numBytesRed As UInteger
            Dim dlg As ReadDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_Read), GetType(ReadDelegate)), ReadDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, data, CUInt(numBytesToRead), numBytesRed)
            CheckErrors(status)
            Return data
        End Function

        ''' <summary>
        ''' Записывает данные в открытое устройство и возвращает число реально переданных данных.
        ''' </summary>
        ''' <param name="data">Массив данных для записи.</param>
        ''' <param name="numBytesToWrite">Число байтов для записи.</param>
        Public Function Write(data As Byte(), numBytesToWrite As Integer) As Integer
            Dim numBytesWritten As UInteger = 0UI
            Dim dlg As WriteDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_Write), GetType(WriteDelegate)), WriteDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, data, CUInt(numBytesToWrite), numBytesWritten)
            CheckErrors(status)
            Return CInt(numBytesWritten)
        End Function

#End Region '/ЧТЕНИЕ, ЗАПИСЬ

#Region "УПРАВЛЕНИЕ"

        ''' <summary>
        ''' Сбрасывает открытое устройство.
        ''' </summary>
        Public Sub ResetDevice()
            Dim dlg As ResetDeviceDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_ResetDevice), GetType(ResetDeviceDelegate)), ResetDeviceDelegate)
            Dim status As FT_STATUS = dlg(FtHandle)
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Очищает заданные буферы.
        ''' Purge buffer constant definitions.
        ''' </summary>
        Public Sub Purge(purgeMask As FT_PURGE)
            Dim dlg As PurgeDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_Purge), GetType(PurgeDelegate)), PurgeDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, purgeMask)
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Регистрирует уведомление о событиях.
        ''' </summary>
        ''' <remarks>
        ''' После регистрации уведомления, событие может быть перехвачено методами ожидания <see cref="EventWaitHandle.WaitOne()"/>.
        ''' Если мониторятся несколько типов событий, вызывающее событие можно определить вызовом метода <see cref="GetEventType()"/>.
        ''' </remarks>
        ''' <param name="eventMask">Тип сигнального события.</param>
        ''' <param name="eventHandle">Указатель обработчика события.</param>
        Public Sub SetEventNotification(eventMask As FT_EVENTS, eventHandle As EventWaitHandle)
            Dim dlg As SetEventNotificationDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_SetEventNotification), GetType(SetEventNotificationDelegate)), SetEventNotificationDelegate)
            Dim status As FT_STATUS = dlg.Invoke(FtHandle, eventMask, eventHandle.SafeWaitHandle)
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Останавливает работу планировщика шины USB (опрос заданий).
        ''' </summary>
        Public Sub StopInTask()
            Dim dlg As StopInTaskDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_StopInTask), GetType(StopInTaskDelegate)), StopInTaskDelegate)
            Dim status As FT_STATUS = dlg(FtHandle)
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Возобновляет работу планировщика заданий шины USB (опрос заданий).
        ''' </summary>
        Public Sub RestartInTask()
            Dim dlg As RestartInTaskDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_RestartInTask), GetType(RestartInTaskDelegate)), RestartInTaskDelegate)
            Dim status As FT_STATUS = dlg(FtHandle)
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Перезагружает порт устройства.
        ''' </summary>
        Public Sub ResetPort()
            Dim dlg As ResetPortDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_ResetPort), GetType(ResetPortDelegate)), ResetPortDelegate)
            Dim status As FT_STATUS = dlg(FtHandle)
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Вызывает перенумерацию устройств на шине USB. Эквивалентно извлечению и вставке устройства USB.
        ''' </summary>
        Public Sub CyclePort()
            Dim dlg As CyclePortDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_CyclePort), GetType(CyclePortDelegate)), CyclePortDelegate)
            Dim status As FT_STATUS = dlg.Invoke(FtHandle)
            CheckErrors(status)

            Dim cDlg As CloseDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_Close), GetType(CloseDelegate)), CloseDelegate)
            status = cDlg.Invoke(FtHandle)
            If (status = FT_STATUS.FT_OK) Then
                FtHandle = CLOSED_HANDLE
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Переводит устройство в заданный режим.
        ''' </summary>
        ''' <param name="mask">Задаёт направления выводов. 0 соответствует ВХОДУ, 1 - ВЫХОДУ.
        ''' В режиме CBUS Bit Bang, верхняя тетрада байта задаёт направление, а нижняя - уровень (0 - низкий, 1 - высокий).
        ''' </param>
        ''' <param name="bitMode">Режим работы.
        ''' <list type="bullet">
        ''' <item>Для FT232H валидны <see cref="FT_BIT_MODES.FT_BIT_MODE_RESET"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_ASYNC_BITBANG"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_MPSSE"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_SYNC_BITBANG"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_CBUS_BITBANG"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_MCU_HOST"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_FAST_SERIAL"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_SYNC_FIFO"/>.</item>
        ''' <item>Для FT2232H валидны <see cref="FT_BIT_MODES.FT_BIT_MODE_RESET"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_ASYNC_BITBANG"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_MPSSE"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_SYNC_BITBANG"/>, FT_BIT_MODE_MCU_HOST, FT_BIT_MODE_FAST_SERIAL, FT_BIT_MODE_SYNC_FIFO.</item>
        ''' <item>Для FT4232H валидны <see cref="FT_BIT_MODES.FT_BIT_MODE_RESET"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_ASYNC_BITBANG"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_MPSSE"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_SYNC_BITBANG"/>.</item>
        ''' <item>Для FT232R валидны <see cref="FT_BIT_MODES.FT_BIT_MODE_RESET"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_ASYNC_BITBANG"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_SYNC_BITBANG"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_CBUS_BITBANG"/>.</item>
        ''' <item>Для FT245R валидны <see cref="FT_BIT_MODES.FT_BIT_MODE_RESET"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_ASYNC_BITBANG"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_SYNC_BITBANG"/>.</item>
        ''' <item>Для FT2232 валидны <see cref="FT_BIT_MODES.FT_BIT_MODE_RESET"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_ASYNC_BITBANG"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_MPSSE"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_SYNC_BITBANG"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_MCU_HOST"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_FAST_SERIAL"/>.</item>
        ''' <item>Для FT232B и FT245B валидны <see cref="FT_BIT_MODES.FT_BIT_MODE_RESET"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_ASYNC_BITBANG"/>.</item>
        ''' </list>
        ''' </param>
        Public Sub SetBitMode(mask As Byte, bitMode As FT_BIT_MODES)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            Select Case DeviceType
                Case FT_DEVICE.FT_DEVICE_AM
                    CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)

                Case FT_DEVICE.FT_DEVICE_100AX
                    CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)

                Case FT_DEVICE.FT_DEVICE_BM
                    If (bitMode <> FT_BIT_MODES.FT_BIT_MODE_RESET) Then
                        If ((bitMode And 1) = 0) Then
                            CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                        End If
                    End If

                Case FT_DEVICE.FT_DEVICE_2232
                    If (bitMode <> FT_BIT_MODES.FT_BIT_MODE_RESET) Then
                        If ((bitMode And &H1F) = 0) Then
                            CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                        End If
                        If (bitMode = FT_BIT_MODES.FT_BIT_MODE_MPSSE) AndAlso (Me.InterfaceIdentifier <> "A") Then
                            CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                        End If
                    End If

                Case FT_DEVICE.FT_DEVICE_232R
                    If (bitMode <> FT_BIT_MODES.FT_BIT_MODE_RESET) Then
                        If ((bitMode And 37) = 0) Then
                            CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                        End If
                    End If

                Case FT_DEVICE.FT_DEVICE_2232H
                    If (bitMode <> FT_BIT_MODES.FT_BIT_MODE_RESET) Then
                        If ((bitMode And 95) = 0) Then
                            CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                        End If
                        If ((bitMode = FT_BIT_MODES.FT_BIT_MODE_MCU_HOST) Or (bitMode = FT_BIT_MODES.FT_BIT_MODE_SYNC_FIFO)) AndAlso (Me.InterfaceIdentifier <> "A") Then
                            CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                        End If
                    End If

                Case FT_DEVICE.FT_DEVICE_4232H
                    If (bitMode <> FT_BIT_MODES.FT_BIT_MODE_RESET) Then
                        If ((bitMode And 7) = 0) Then
                            CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                        End If
                        If (bitMode = FT_BIT_MODES.FT_BIT_MODE_MPSSE) AndAlso (Me.InterfaceIdentifier <> "A") AndAlso (Me.InterfaceIdentifier <> "B") Then
                            CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                        End If
                    End If

                Case FT_DEVICE.FT_DEVICE_232H
                    If (bitMode > FT_BIT_MODES.FT_BIT_MODE_SYNC_FIFO) Then
                        CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                    End If
            End Select

            If (DeviceType = FT_DEVICE.FT_DEVICE_AM) Then
                CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)

            ElseIf (DeviceType = FT_DEVICE.FT_DEVICE_100AX) Then
                CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)

            ElseIf (DeviceType = FT_DEVICE.FT_DEVICE_BM) AndAlso (bitMode <> FT_BIT_MODES.FT_BIT_MODE_RESET) Then
                If ((bitMode And 1) = 0) Then
                    CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                End If

            ElseIf (DeviceType = FT_DEVICE.FT_DEVICE_2232) AndAlso (bitMode <> FT_BIT_MODES.FT_BIT_MODE_RESET) Then
                If ((bitMode And &H1F) = 0) Then
                    CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                End If
                If (bitMode = FT_BIT_MODES.FT_BIT_MODE_MPSSE) AndAlso (Me.InterfaceIdentifier <> "A") Then
                    CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                End If

            ElseIf (DeviceType = FT_DEVICE.FT_DEVICE_232R) AndAlso (bitMode <> FT_BIT_MODES.FT_BIT_MODE_RESET) Then
                If ((bitMode And 37) = 0) Then
                    CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                End If

            ElseIf (DeviceType = FT_DEVICE.FT_DEVICE_2232H) AndAlso (bitMode <> FT_BIT_MODES.FT_BIT_MODE_RESET) Then
                If ((bitMode And 95) = 0) Then
                    CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                End If
                If ((bitMode = FT_BIT_MODES.FT_BIT_MODE_MCU_HOST) Or (bitMode = FT_BIT_MODES.FT_BIT_MODE_SYNC_FIFO)) AndAlso (Me.InterfaceIdentifier <> "A") Then
                    CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                End If

            ElseIf (DeviceType = FT_DEVICE.FT_DEVICE_4232H) AndAlso (bitMode <> FT_BIT_MODES.FT_BIT_MODE_RESET) Then
                If ((bitMode And 7) = 0) Then
                    CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                End If
                If (bitMode = FT_BIT_MODES.FT_BIT_MODE_MPSSE) AndAlso (Me.InterfaceIdentifier <> "A") AndAlso (Me.InterfaceIdentifier <> "B") Then
                    CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                End If

            ElseIf (DeviceType = FT_DEVICE.FT_DEVICE_232H) AndAlso (bitMode <> FT_BIT_MODES.FT_BIT_MODE_RESET) AndAlso (bitMode > FT_BIT_MODES.FT_BIT_MODE_SYNC_FIFO) Then
                CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)

            End If

            Dim dlg As SetBitModeDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_SetBitMode), GetType(SetBitModeDelegate)), SetBitModeDelegate)
            status = dlg.Invoke(FtHandle, mask, bitMode)
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Возвращает состояние бита по индексу <paramref name="index"/>.
        ''' </summary>
        Public Function GetGpio(index As Integer) As Boolean
            Dim b As Byte = GetPinStates()
            Return CBool((b >> index) And 1)
        End Function

        ''' <summary>
        ''' Устанавливает для выбранного GPIO значение <paramref name="value"/>.
        ''' </summary>
        ''' <remarks>
        ''' 1 - ВЫХОД, 0 - ВХОД.
        ''' SetBitMode(0000_1111, Mode) 'all gpio high 
        ''' SetBitMode(0001_1111, Mode) 'gpio0 low
        ''' SetBitMode(0010_1111, Mode) 'gpio1 low
        ''' SetBitMode(0100_1111, Mode) 'gpio2 low
        ''' SetBitMode(1000_1111, Mode) 'gpio3 low
        ''' </remarks>
        Public Sub SetGpio(index As Integer, value As Boolean)
            If (index < 0) OrElse (index > 3) Then
                Throw New ArgumentException("Номер GPIO должен лежать в диапазоне от 0 до 3.")
            End If
            Dim b As Byte = GetPinStates()
            Dim bv As New Specialized.BitVector32(b)
            Dim shift As Integer = 4 + index
            bv(1 << shift) = value
            Dim mask As Byte = CByte(bv.Data Xor &HFF)
            SetBitMode(mask, FT_BIT_MODES.FT_BIT_MODE_ASYNC_BITBANG) 
        End Sub

        ''' <summary>
        ''' Получает данные от FT4222 используя командный интерфейс вендора.
        ''' </summary>
        Public Sub VendorCmdGet(request As UShort, buf As Byte(), len As UShort)
            Dim dlg As VendorCmdGetDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_VendorCmdGet), GetType(VendorCmdGetDelegate)), VendorCmdGetDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, request, buf, len)
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Выставляет данные FT4222 используя командный интерфейс вендора.
        ''' </summary>
        Public Sub VendorCmdSet(request As UShort, buf As Byte(), len As UShort)
            Dim dlg As VendorCmdSetDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_VendorCmdSet), GetType(VendorCmdSetDelegate)), VendorCmdSetDelegate)
            Dim status As FT_STATUS = dlg(FtHandle, request, buf, len)
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Проверяет изменения состава аппаратуры на шине USB.
        ''' </summary>
        ''' <remarks>
        ''' Эквивалентно нажатию кнопки "Обновить конфигурацию оборудования" в менеджере устройств.
        ''' </remarks>
        Public Shared Sub Rescan()
            Dim dlg As RescanDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_Rescan), GetType(RescanDelegate)), RescanDelegate)
            Dim status As FT_STATUS = dlg.Invoke()
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Принудительно перезагружает драйверы для устройств с заданными VID и PID.
        ''' </summary>
        ''' <remarks>Если VID и PID равны 0, будет перезагружен драйвер USB хаба, что вызовет перезагрузку всех подключённых к шине USB устройств.</remarks>
        Public Shared Sub Reload(vendorID As UShort, productID As UShort)
            Dim dlg As ReloadDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(Pft_Reload), GetType(ReloadDelegate)), ReloadDelegate)
            Dim status As FT_STATUS = dlg(vendorID, productID)
            CheckErrors(status)
        End Sub


#End Region '/УПРАВЛЕНИЕ

#Region "HELPERS"

        ''' <summary>
        ''' Проверяет результат выполнения метода и выбрасывает исключение, если статус с ошибкой.
        ''' </summary>
        Public Shared Sub CheckErrors(status As FT_STATUS, Optional additionalInfo As FT_ERROR = FT_ERROR.FT_NO_ERROR)
            Select Case status
                Case FT_STATUS.FT_OK
                    'ничего не делаем
                Case FT_STATUS.FT_OTHER_ERROR
                    Throw New FT_EXCEPTION("Ошибка при попытке соединения с устройством FTDI.")
                Case FT_STATUS.FT_INVALID_HANDLE
                    Throw New FT_EXCEPTION("Неверный дескриптор устройства FTDI.")
                Case FT_STATUS.FT_DEVICE_NOT_FOUND
                    Throw New FT_EXCEPTION("Устройство FTDI не найдено.")
                Case FT_STATUS.FT_DEVICE_NOT_OPENED
                    Throw New FT_EXCEPTION("Устройство FTDI не открыто.")
                Case FT_STATUS.FT_IO_ERROR
                    Throw New FT_EXCEPTION("Ошибка ввода-вывода устройства FTDI.")
                Case FT_STATUS.FT_INSUFFICIENT_RESOURCES
                    Throw New FT_EXCEPTION("Недостаточно ресурсов.")
                Case FT_STATUS.FT_INVALID_PARAMETER
                    Throw New FT_EXCEPTION("Неверный параметр вызываемой функции FTD2XX.")
                Case FT_STATUS.FT_INVALID_BAUD_RATE
                    Throw New FT_EXCEPTION("Неверный битрейт устройства FTDI.")
                Case FT_STATUS.FT_DEVICE_NOT_OPENED_FOR_ERASE
                    Throw New FT_EXCEPTION("Устройство FTDI не открыто для стирания ЭСППЗУ.")
                Case FT_STATUS.FT_DEVICE_NOT_OPENED_FOR_WRITE
                    Throw New FT_EXCEPTION("Устройство FTDI не открыто для записи.")
                Case FT_STATUS.FT_FAILED_TO_WRITE_DEVICE
                    Throw New FT_EXCEPTION("Сбой при попытке записи в устройство FTDI.")
                Case FT_STATUS.FT_EEPROM_READ_FAILED
                    Throw New FT_EXCEPTION("Сбой чтения EEPROM устройства FTDI.")
                Case FT_STATUS.FT_EEPROM_WRITE_FAILED
                    Throw New FT_EXCEPTION("Сбой записи EEPROM устройства FTDI.")
                Case FT_STATUS.FT_EEPROM_ERASE_FAILED
                    Throw New FT_EXCEPTION("Сбой при стирании EEPROM устройства FTDI.")
                Case FT_STATUS.FT_EEPROM_NOT_PRESENT
                    Throw New FT_EXCEPTION("EEPROM не подходит для устройства FTDI.")
                Case FT_STATUS.FT_EEPROM_NOT_PROGRAMMED
                    Throw New FT_EXCEPTION("EEPROM устройства FTDI не запрограммировано.")
                Case FT_STATUS.FT_INVALID_ARGS
                    Throw New FT_EXCEPTION("Неверные аргументы при вызове функции FTD2XX.")
            End Select
            Select Case additionalInfo
                Case FT_ERROR.FT_NO_ERROR
                    Return
                Case FT_ERROR.FT_INCORRECT_DEVICE
                    Throw New FT_EXCEPTION("Тип текущего устройства не совпадает со структурой памяти EEPROM.")
                Case FT_ERROR.FT_INVALID_BITMODE
                    Throw New FT_EXCEPTION("Указанный битовый режим не является допустимым для текущего устройства.")
                Case FT_ERROR.FT_BUFFER_SIZE
                    Throw New FT_EXCEPTION("Предоставленный буфер имеет недостаточный размер.")
                Case Else
                    Debug.WriteLine("Значение не содержится в перечислении FT_ERROR.")
            End Select
        End Sub

#End Region '/HELPERS

#Region "NATIVE"

        ''' <summary>
        ''' Задаёт произвольный путь к библиотеке ftd2xx.
        ''' </summary>
        Public Shared Sub SetLibraryPath(libPath As String)
            UnloadLibrary()
            _LibPath = libPath
        End Sub

        <DllImport("kernel32.dll", SetLastError:=True)>
        Private Shared Function LoadLibrary(dllToLoad As String) As Integer
        End Function

        <DllImport("kernel32.dll", SetLastError:=True)>
        Private Shared Function FreeLibrary(hModule As Integer) As Boolean
        End Function

        <DllImport("kernel32.dll", SetLastError:=True)>
        Private Shared Function GetProcAddress(hModule As Integer, procedureName As String) As Integer
        End Function

        ''' <summary>
        ''' Дескриптор библиотеки ftd2xx.
        ''' </summary>
        Private Shared ReadOnly Property Ftd2xxDllHandle As Integer
            Get
                If (_Ftd2xxDllHandle = CLOSED_HANDLE) Then
                    _Ftd2xxDllHandle = LoadLibrary(LibPath)
                    If (Not IsLibraryLoaded) Then
                        Throw New FT_EXCEPTION("Ошибка загрузки библиотеки ftd2xx.dll.")
                    End If
                    FindFunctionPointers()
                End If
                Return _Ftd2xxDllHandle
            End Get
        End Property
        Private Shared _Ftd2xxDllHandle As Integer = CLOSED_HANDLE

        Private Shared ReadOnly Property Pft_CreateDeviceInfoList As Integer
            Get
                If (_Pft_CreateDeviceInfoList = NOT_INIT) Then
                    _Pft_CreateDeviceInfoList = GetProcAddress(Ftd2xxDllHandle, "FT_CreateDeviceInfoList")
                End If
                Return _Pft_CreateDeviceInfoList
            End Get
        End Property
        Private Shared _Pft_CreateDeviceInfoList As Integer = NOT_INIT

        'Статические поля:
        Private Shared Pft_GetDeviceInfoDetail As Integer = NOT_INIT
        Private Shared Pft_Open As Integer = NOT_INIT
        Private Shared Pft_OpenEx As Integer = NOT_INIT
        Private Shared Pft_Close As Integer = NOT_INIT
        Private Shared Pft_Read As Integer = NOT_INIT
        Private Shared Pft_Write As Integer = NOT_INIT
        Private Shared Pft_GetQueueStatus As Integer = NOT_INIT
        Private Shared Pft_GetModemStatus As Integer = NOT_INIT
        Private Shared Pft_GetStatus As Integer = NOT_INIT
        Private Shared Pft_SetBaudRate As Integer = NOT_INIT
        Private Shared Pft_SetDataCharacteristics As Integer = NOT_INIT
        Private Shared Pft_SetFlowControl As Integer = NOT_INIT
        Private Shared Pft_SetDtr As Integer = NOT_INIT
        Private Shared Pft_ClrDtr As Integer = NOT_INIT
        Private Shared Pft_SetRts As Integer = NOT_INIT
        Private Shared Pft_ClrRts As Integer = NOT_INIT
        Private Shared Pft_ResetDevice As Integer = NOT_INIT
        Private Shared Pft_ResetPort As Integer = NOT_INIT
        Private Shared Pft_CyclePort As Integer = NOT_INIT
        Private Shared Pft_Rescan As Integer = NOT_INIT
        Private Shared Pft_Reload As Integer = NOT_INIT
        Private Shared Pft_Purge As Integer = NOT_INIT
        Private Shared Pft_SetTimeouts As Integer = NOT_INIT
        Private Shared Pft_SetBreakOn As Integer = NOT_INIT
        Private Shared Pft_SetBreakOff As Integer = NOT_INIT
        Private Shared Pft_GetDeviceInfo As Integer = NOT_INIT
        Private Shared Pft_SetResetPipeRetryCount As Integer = NOT_INIT
        Private Shared Pft_StopInTask As Integer = NOT_INIT
        Private Shared Pft_RestartInTask As Integer = NOT_INIT
        Private Shared Pft_GetDriverVersion As Integer = NOT_INIT
        Private Shared Pft_SetDeadmanTimeout As Integer = NOT_INIT
        Private Shared Pft_SetChars As Integer = NOT_INIT
        Private Shared Pft_SetEventNotification As Integer = NOT_INIT
        Private Shared Pft_GetComPortNumber As Integer = NOT_INIT
        Private Shared Pft_SetLatencyTimer As Integer = NOT_INIT
        Private Shared Pft_GetLatencyTimer As Integer = NOT_INIT
        Private Shared Pft_SetBitMode As Integer = NOT_INIT
        Private Shared Pft_GetBitMode As Integer = NOT_INIT
        Private Shared Pft_SetUSBParameters As Integer = NOT_INIT
        Private Shared Pft_VendorCmdGet As Integer = NOT_INIT
        Private Shared Pft_VendorCmdSet As Integer = NOT_INIT

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function CreateDeviceInfoListDelegate(ByRef numdevs As UInteger) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function GetDeviceInfoDetailDelegate(index As UInteger, ByRef flags As UInteger, ByRef chiptype As FT_DEVICE, ByRef id As UInteger, ByRef locid As UInteger, serialnumber As Byte(), description As Byte(), ByRef ftHandle As Integer) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function OpenDelegate(index As UInteger, ByRef ftHandle As Integer) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function OpenExDelegate(devstring As String, dwFlags As UInteger, ByRef ftHandle As Integer) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function OpenExLocDelegate(devloc As UInteger, dwFlags As UInteger, ByRef ftHandle As Integer) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function CloseDelegate(ftHandle As Integer) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function ReadDelegate(ftHandle As Integer, lpBuffer As Byte(), dwBytesToRead As UInteger, ByRef lpdwBytesReturned As UInteger) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function WriteDelegate(ftHandle As Integer, lpBuffer As Byte(), dwBytesToWrite As UInteger, ByRef lpdwBytesWritten As UInteger) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function GetQueueStatusDelegate(ftHandle As Integer, ByRef lpdwAmountInRxQueue As UInteger) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function GetModemStatusDelegate(ftHandle As Integer, ByRef lpdwModemStatus As UInteger) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function GetStatusDelegate(ftHandle As Integer, ByRef lpdwAmountInRxQueue As UInteger, ByRef lpdwAmountInTxQueue As UInteger, ByRef lpdwEventStatus As UInteger) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function SetBaudRateDelegate(ftHandle As Integer, dwBaudRate As UInteger) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function SetDataCharacteristicsDelegate(ftHandle As Integer, uWordLength As Byte, uStopBits As Byte, uParity As Byte) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function SetFlowControlDelegate(ftHandle As Integer, usFlowControl As UShort, uXon As Byte, uXoff As Byte) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function SetDtrDelegate(ftHandle As Integer) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function ClrDtrDelegate(ftHandle As Integer) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function SetRtsDelegate(ftHandle As Integer) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function ClrRtsDelegate(ftHandle As Integer) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function ResetDeviceDelegate(ftHandle As Integer) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function ResetPortDelegate(ftHandle As Integer) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function CyclePortDelegate(ftHandle As Integer) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function RescanDelegate() As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function ReloadDelegate(wVID As UShort, wPID As UShort) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function PurgeDelegate(ftHandle As Integer, dwMask As UInteger) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function SetTimeoutsDelegate(ftHandle As Integer, dwReadTimeout As UInteger, dwWriteTimeout As UInteger) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function SetBreakOnDelegate(ftHandle As Integer) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function SetBreakOffDelegate(ftHandle As Integer) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function GetDeviceInfoDelegate(ftHandle As Integer, ByRef pftType As FT_DEVICE, ByRef lpdwID As UInteger, pcSerialNumber As Byte(), pcDescription As Byte(), pvDummy As Integer) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function SetResetPipeRetryCountDelegate(ftHandle As Integer, dwCount As UInteger) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function StopInTaskDelegate(ftHandle As Integer) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function RestartInTaskDelegate(ftHandle As Integer) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function GetDriverVersionDelegate(ftHandle As Integer, ByRef lpdwDriverVersion As UInteger) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function GetLibraryVersionDelegate(ByRef lpdwLibraryVersion As UInteger) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function SetDeadmanTimeoutDelegate(ftHandle As Integer, dwDeadmanTimeout As UInteger) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function SetCharsDelegate(ftHandle As Integer, uEventCh As Byte, uEventChEn As Byte, uErrorCh As Byte, uErrorChEn As Byte) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function SetEventNotificationDelegate(ftHandle As Integer, dwEventMask As UInteger, hEvent As SafeHandle) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function GetComPortNumberDelegate(ftHandle As Integer, ByRef dwComPortNumber As Integer) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function SetLatencyTimerDelegate(ftHandle As Integer, ucLatency As Byte) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function GetLatencyTimerDelegate(ftHandle As Integer, ByRef ucLatency As Byte) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function SetBitModeDelegate(ftHandle As Integer, ucMask As Byte, ucMode As Byte) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function GetBitModeDelegate(ftHandle As Integer, ByRef ucMode As Byte) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function SetUsbParametersDelegate(ftHandle As Integer, dwInTransferSize As UInteger, dwOutTransferSize As UInteger) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function VendorCmdGetDelegate(ftHandle As Integer, request As UShort, buf As Byte(), len As UShort) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function VendorCmdSetDelegate(ftHandle As Integer, request As UShort, buf As Byte(), len As UShort) As FT_STATUS

        ''' <summary>
        ''' Ищет указатели на нативные функции в библиотеке ftd2xx.
        ''' </summary>
        Private Shared Sub FindFunctionPointers()
            If (Pft_GetDeviceInfoDetail = NOT_INIT) Then
                Pft_GetDeviceInfoDetail = GetProcAddress(Ftd2xxDllHandle, "FT_GetDeviceInfoDetail")
            End If
            If (Pft_Open = NOT_INIT) Then
                Pft_Open = GetProcAddress(Ftd2xxDllHandle, "FT_Open")
            End If
            If (Pft_OpenEx = NOT_INIT) Then
                Pft_OpenEx = GetProcAddress(Ftd2xxDllHandle, "FT_OpenEx")
            End If
            If (Pft_Close = NOT_INIT) Then
                Pft_Close = GetProcAddress(Ftd2xxDllHandle, "FT_Close")
            End If
            If (Pft_Read = NOT_INIT) Then
                Pft_Read = GetProcAddress(Ftd2xxDllHandle, "FT_Read")
            End If
            If (Pft_Write = NOT_INIT) Then
                Pft_Write = GetProcAddress(Ftd2xxDllHandle, "FT_Write")
            End If
            If (Pft_GetQueueStatus = NOT_INIT) Then
                Pft_GetQueueStatus = GetProcAddress(Ftd2xxDllHandle, "FT_GetQueueStatus")
            End If
            If (Pft_GetModemStatus = NOT_INIT) Then
                Pft_GetModemStatus = GetProcAddress(Ftd2xxDllHandle, "FT_GetModemStatus")
            End If
            If (Pft_GetStatus = NOT_INIT) Then
                Pft_GetStatus = GetProcAddress(Ftd2xxDllHandle, "FT_GetStatus")
            End If
            If (Pft_SetBaudRate = NOT_INIT) Then
                Pft_SetBaudRate = GetProcAddress(Ftd2xxDllHandle, "FT_SetBaudRate")
            End If
            If (Pft_SetDataCharacteristics = NOT_INIT) Then
                Pft_SetDataCharacteristics = GetProcAddress(Ftd2xxDllHandle, "FT_SetDataCharacteristics")
            End If
            If (Pft_SetFlowControl = NOT_INIT) Then
                Pft_SetFlowControl = GetProcAddress(Ftd2xxDllHandle, "FT_SetFlowControl")
            End If
            If (Pft_SetDtr = NOT_INIT) Then
                Pft_SetDtr = GetProcAddress(Ftd2xxDllHandle, "FT_SetDtr")
            End If
            If (Pft_ClrDtr = NOT_INIT) Then
                Pft_ClrDtr = GetProcAddress(Ftd2xxDllHandle, "FT_ClrDtr")
            End If
            If (Pft_SetRts = NOT_INIT) Then
                Pft_SetRts = GetProcAddress(Ftd2xxDllHandle, "FT_SetRts")
            End If
            If (Pft_ClrRts = NOT_INIT) Then
                Pft_ClrRts = GetProcAddress(Ftd2xxDllHandle, "FT_ClrRts")
            End If
            If (Pft_ResetDevice = NOT_INIT) Then
                Pft_ResetDevice = GetProcAddress(Ftd2xxDllHandle, "FT_ResetDevice")
            End If
            If (Pft_ResetPort = NOT_INIT) Then
                Pft_ResetPort = GetProcAddress(Ftd2xxDllHandle, "FT_ResetPort")
            End If
            If (Pft_CyclePort = NOT_INIT) Then
                Pft_CyclePort = GetProcAddress(Ftd2xxDllHandle, "FT_CyclePort")
            End If
            If (Pft_Rescan = NOT_INIT) Then
                Pft_Rescan = GetProcAddress(Ftd2xxDllHandle, "FT_Rescan")
            End If
            If (Pft_Reload = NOT_INIT) Then
                Pft_Reload = GetProcAddress(Ftd2xxDllHandle, "FT_Reload")
            End If
            If (Pft_Purge = NOT_INIT) Then
                Pft_Purge = GetProcAddress(Ftd2xxDllHandle, "FT_Purge")
            End If
            If (Pft_SetTimeouts = NOT_INIT) Then
                Pft_SetTimeouts = GetProcAddress(Ftd2xxDllHandle, "FT_SetTimeouts")
            End If
            If (Pft_SetBreakOn = NOT_INIT) Then
                Pft_SetBreakOn = GetProcAddress(Ftd2xxDllHandle, "FT_SetBreakOn")
            End If
            If (Pft_SetBreakOff = NOT_INIT) Then
                Pft_SetBreakOff = GetProcAddress(Ftd2xxDllHandle, "FT_SetBreakOff")
            End If
            If (Pft_GetDeviceInfo = NOT_INIT) Then
                Pft_GetDeviceInfo = GetProcAddress(Ftd2xxDllHandle, "FT_GetDeviceInfo")
            End If
            If (Pft_SetResetPipeRetryCount = NOT_INIT) Then
                Pft_SetResetPipeRetryCount = GetProcAddress(Ftd2xxDllHandle, "FT_SetResetPipeRetryCount")
            End If
            If (Pft_StopInTask = NOT_INIT) Then
                Pft_StopInTask = GetProcAddress(Ftd2xxDllHandle, "FT_StopInTask")
            End If
            If (Pft_RestartInTask = NOT_INIT) Then
                Pft_RestartInTask = GetProcAddress(Ftd2xxDllHandle, "FT_RestartInTask")
            End If
            If (Pft_GetDriverVersion = NOT_INIT) Then
                Pft_GetDriverVersion = GetProcAddress(Ftd2xxDllHandle, "FT_GetDriverVersion")
            End If
            If (Pft_SetDeadmanTimeout = NOT_INIT) Then
                Pft_SetDeadmanTimeout = GetProcAddress(Ftd2xxDllHandle, "FT_SetDeadmanTimeout")
            End If
            If (Pft_SetChars = NOT_INIT) Then
                Pft_SetChars = GetProcAddress(Ftd2xxDllHandle, "FT_SetChars")
            End If
            If (Pft_SetEventNotification = NOT_INIT) Then
                Pft_SetEventNotification = GetProcAddress(Ftd2xxDllHandle, "FT_SetEventNotification")
            End If
            If (Pft_GetComPortNumber = NOT_INIT) Then
                Pft_GetComPortNumber = GetProcAddress(Ftd2xxDllHandle, "FT_GetComPortNumber")
            End If
            If (Pft_SetLatencyTimer = NOT_INIT) Then
                Pft_SetLatencyTimer = GetProcAddress(Ftd2xxDllHandle, "FT_SetLatencyTimer")
            End If
            If (Pft_GetLatencyTimer = NOT_INIT) Then
                Pft_GetLatencyTimer = GetProcAddress(Ftd2xxDllHandle, "FT_GetLatencyTimer")
            End If
            If (Pft_SetBitMode = NOT_INIT) Then
                Pft_SetBitMode = GetProcAddress(Ftd2xxDllHandle, "FT_SetBitMode")
            End If
            If (Pft_GetBitMode = NOT_INIT) Then
                Pft_GetBitMode = GetProcAddress(Ftd2xxDllHandle, "FT_GetBitMode")
            End If
            If (Pft_SetUSBParameters = NOT_INIT) Then
                Pft_SetUSBParameters = GetProcAddress(Ftd2xxDllHandle, "FT_SetUSBParameters")
            End If
            If (Pft_VendorCmdGet = NOT_INIT) Then
                Pft_VendorCmdGet = GetProcAddress(Ftd2xxDllHandle, "FT_VendorCmdGet")
            End If
            If (Pft_VendorCmdSet = NOT_INIT) Then
                Pft_VendorCmdSet = GetProcAddress(Ftd2xxDllHandle, "FT_VendorCmdSet")
            End If
        End Sub

#End Region '/NATIVE

#Region "NESTED TYPES"

        Public Enum FT_STATUS
            FT_OK
            FT_INVALID_HANDLE
            FT_DEVICE_NOT_FOUND
            FT_DEVICE_NOT_OPENED
            FT_IO_ERROR
            FT_INSUFFICIENT_RESOURCES
            FT_INVALID_PARAMETER
            FT_INVALID_BAUD_RATE
            FT_DEVICE_NOT_OPENED_FOR_ERASE
            FT_DEVICE_NOT_OPENED_FOR_WRITE
            FT_FAILED_TO_WRITE_DEVICE
            FT_EEPROM_READ_FAILED
            FT_EEPROM_WRITE_FAILED
            FT_EEPROM_ERASE_FAILED
            FT_EEPROM_NOT_PRESENT
            FT_EEPROM_NOT_PROGRAMMED
            FT_INVALID_ARGS
            FT_OTHER_ERROR
        End Enum

        Public Enum FT_ERROR
            FT_NO_ERROR
            FT_INCORRECT_DEVICE
            FT_INVALID_BITMODE
            FT_BUFFER_SIZE
        End Enum

        Public Enum FT_DATA_BITS As Byte
            FT_BITS_7 = 7
            FT_BITS_8 = 8
        End Enum

        Public Enum FT_STOP_BITS As Byte
            FT_STOP_BITS_1 = 0
            FT_STOP_BITS_2 = 2
        End Enum

        Public Enum FT_PARITY As Byte
            FT_PARITY_NONE
            FT_PARITY_ODD
            FT_PARITY_EVEN
            FT_PARITY_MARK
            FT_PARITY_SPACE
        End Enum

        Public Enum FT_FLOW_CONTROL As UShort
            FT_FLOW_NONE = 0US
            FT_FLOW_RTS_CTS = 256US
            FT_FLOW_DTR_DSR = 512US
            FT_FLOW_XON_XOFF = 1024US
        End Enum

        <Flags()>
        Public Enum FT_PURGE As Byte
            FT_PURGE_RX = 1
            FT_PURGE_TX = 2
        End Enum

        Public Enum FT_MODEM_STATUS As Byte
            FT_CTS = &H10
            FT_DSR = &H20
            FT_RI = &H40
            FT_DCD = &H80
        End Enum

        <Flags()>
        Public Enum FT_LINE_STATUS As Byte
            FT_OE = 2
            FT_PE = 4
            FT_FE = 8
            FT_BI = 16
        End Enum

        <Flags()>
        Public Enum FT_EVENTS As UInteger
            FT_EVENT_RXCHAR = 1UI
            FT_EVENT_MODEM_STATUS = 2UI
            FT_EVENT_LINE_STATUS = 4UI
        End Enum

        <Flags()>
        Public Enum FT_BIT_MODES As Byte
            FT_BIT_MODE_RESET = 0
            FT_BIT_MODE_ASYNC_BITBANG = 1
            FT_BIT_MODE_MPSSE = 2
            FT_BIT_MODE_SYNC_BITBANG = 4
            FT_BIT_MODE_MCU_HOST = 8
            FT_BIT_MODE_FAST_SERIAL = 16
            FT_BIT_MODE_CBUS_BITBANG = 32
            FT_BIT_MODE_SYNC_FIFO = 64
        End Enum

        Public Enum FT_CBUS_OPTIONS As Byte
            FT_CBUS_TXDEN
            FT_CBUS_PWRON
            FT_CBUS_RXLED
            FT_CBUS_TXLED
            FT_CBUS_TXRXLED
            FT_CBUS_SLEEP
            FT_CBUS_CLK48
            FT_CBUS_CLK24
            FT_CBUS_CLK12
            FT_CBUS_CLK6
            FT_CBUS_IOMODE
            FT_CBUS_BITBANG_WR
            FT_CBUS_BITBANG_RD
        End Enum

        Public Enum FT_232H_CBUS_OPTIONS As Byte
            FT_CBUS_TRISTATE
            FT_CBUS_RXLED
            FT_CBUS_TXLED
            FT_CBUS_TXRXLED
            FT_CBUS_PWREN
            FT_CBUS_SLEEP
            FT_CBUS_DRIVE_0
            FT_CBUS_DRIVE_1
            FT_CBUS_IOMODE
            FT_CBUS_TXDEN
            FT_CBUS_CLK30
            FT_CBUS_CLK15
            FT_CBUS_CLK7_5
        End Enum

        Public Enum FT_XSERIES_CBUS_OPTIONS As Byte
            FT_CBUS_TRISTATE
            FT_CBUS_RXLED
            FT_CBUS_TXLED
            FT_CBUS_TXRXLED
            FT_CBUS_PWREN
            FT_CBUS_SLEEP
            FT_CBUS_Drive_0
            FT_CBUS_Drive_1
            FT_CBUS_GPIO
            FT_CBUS_TXDEN
            FT_CBUS_CLK24MHz
            FT_CBUS_CLK12MHz
            FT_CBUS_CLK6MHz
            FT_CBUS_BCD_Charger
            FT_CBUS_BCD_Charger_N
            FT_CBUS_I2C_TXE
            FT_CBUS_I2C_RXF
            FT_CBUS_VBUS_Sense
            FT_CBUS_BitBang_WR
            FT_CBUS_BitBang_RD
            FT_CBUS_Time_Stamp
            FT_CBUS_Keep_Awake
        End Enum

        Public Enum FT_FLAGS As UInteger
            FT_FLAGS_OPENED = 1UI
            FT_FLAGS_HISPEED = 2UI
        End Enum

        Public Enum FT_DRIVE_CURRENT As Byte
            FT_DRIVE_CURRENT_4MA = 4
            FT_DRIVE_CURRENT_8MA = 8
            FT_DRIVE_CURRENT_12MA = 12
            FT_DRIVE_CURRENT_16MA = 16
        End Enum

        Public Enum FT_DEVICE
            FT_DEVICE_BM = 0
            FT_DEVICE_AM
            FT_DEVICE_100AX
            FT_DEVICE_UNKNOWN
            FT_DEVICE_2232
            FT_DEVICE_232R
            FT_DEVICE_2232H
            FT_DEVICE_4232H
            FT_DEVICE_232H
            FT_DEVICE_X_SERIES
            FT_DEVICE_4222H_0
            FT_DEVICE_4222H_1_2
            FT_DEVICE_4222H_3
            FT_DEVICE_4222_PROG
        End Enum

        Public Structure FT_DEVICE_INFO_NODE
            Public Flags As UInteger
            Public Type As FT_DEVICE
            Public ID As UInteger
            Public LocId As UInteger
            Public SerialNumber As String
            Public Description As String
            Public FtHandle As Integer
        End Structure

        <Serializable()>
        Public Class FT_EXCEPTION
            Inherits Exception

            Public Sub New()
            End Sub

            Public Sub New(message As String)
                MyBase.New(message)
            End Sub

            Public Sub New(message As String, inner As Exception)
                MyBase.New(message, inner)
            End Sub

            Protected Sub New(info As SerializationInfo, context As StreamingContext)
                MyBase.New(info, context)
            End Sub

        End Class '/FT_EXCEPTION

#End Region '/NESTED TYPES

#Region "DISPOSABLE"

        Private DisposedValue As Boolean

        Protected Overridable Sub Dispose(disposing As Boolean)
            If (Not DisposedValue) Then
                If disposing Then
                    Close()
                End If
                DisposedValue = True
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

#End Region '/DISPOSABLE

    End Class '/Ftdi

End Namespace
