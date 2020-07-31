Imports System.Runtime.InteropServices
Imports System.Runtime.Serialization
Imports System.Threading
Imports System.Text

Namespace FTD2XX_NET

    ''' <summary>
    ''' Класс-обёртка для работы с устройствами FTDI.
    ''' </summary>
    Partial Public Class Ftdi

#Region "NATIVE"

        ''' <summary>
        ''' Built-in Windows API functions to allow us to dynamically load our own DLL.
        ''' Will allow us to use old versions of the DLL that do not have all of these functions available.
        ''' </summary>
        ''' <param name="dllToLoad"></param>
        <DllImport("kernel32.dll", SetLastError:=True)>
        Private Shared Function LoadLibrary(dllToLoad As String) As Integer
        End Function

        <DllImport("kernel32.dll", SetLastError:=True)>
        Private Shared Function GetProcAddress(hModule As Integer, procedureName As String) As Integer
        End Function

        <DllImport("kernel32.dll", SetLastError:=True)>
        Private Shared Function FreeLibrary(hModule As Integer) As Boolean
        End Function

#End Region '/NATIVE

#Region "CTOR"

        ''' <summary>
        ''' Constructor for the <see cref="Ftdi"/> class.
        ''' </summary>
        Public Sub New()
            Me.New("ftd2xx.dll")
        End Sub

        ''' <summary>
        ''' Non default constructor allowing passing of string for dll handle.
        ''' </summary>
        ''' <param name="path">Путь к библиотеке ftd2xx.</param>
        Public Sub New(path As String)
            ftd2xxDllHandle = LoadLibrary(path)
            If LibraryLoaded Then
                FindFunctionPointers()
                FindEeFunctionPointers() 'если нужна работа с ЭСППЗУ
            Else
                Throw New Exception("Ошибка загрузки библиотеки ftd2xx.")
            End If
        End Sub

        ''' <summary>
        ''' Ищет указатели на нативные функции по их названию и описателю библиотеки ftd2xx.
        ''' </summary>
        Private Sub FindFunctionPointers()
            pFT_CreateDeviceInfoList = GetProcAddress(ftd2xxDllHandle, "FT_CreateDeviceInfoList")
            pFT_GetDeviceInfoDetail = GetProcAddress(ftd2xxDllHandle, "FT_GetDeviceInfoDetail")
            pFT_Open = GetProcAddress(ftd2xxDllHandle, "FT_Open")
            pFT_OpenEx = GetProcAddress(ftd2xxDllHandle, "FT_OpenEx")
            pFT_Close = GetProcAddress(ftd2xxDllHandle, "FT_Close")
            pFT_Read = GetProcAddress(ftd2xxDllHandle, "FT_Read")
            pFT_Write = GetProcAddress(ftd2xxDllHandle, "FT_Write")
            pFT_GetQueueStatus = GetProcAddress(ftd2xxDllHandle, "FT_GetQueueStatus")
            pFT_GetModemStatus = GetProcAddress(ftd2xxDllHandle, "FT_GetModemStatus")
            pFT_GetStatus = GetProcAddress(ftd2xxDllHandle, "FT_GetStatus")
            pFT_SetBaudRate = GetProcAddress(ftd2xxDllHandle, "FT_SetBaudRate")
            pFT_SetDataCharacteristics = GetProcAddress(ftd2xxDllHandle, "FT_SetDataCharacteristics")
            pFT_SetFlowControl = GetProcAddress(ftd2xxDllHandle, "FT_SetFlowControl")
            pFT_SetDtr = GetProcAddress(ftd2xxDllHandle, "FT_SetDtr")
            pFT_ClrDtr = GetProcAddress(ftd2xxDllHandle, "FT_ClrDtr")
            pFT_SetRts = GetProcAddress(ftd2xxDllHandle, "FT_SetRts")
            pFT_ClrRts = GetProcAddress(ftd2xxDllHandle, "FT_ClrRts")
            pFT_ResetDevice = GetProcAddress(ftd2xxDllHandle, "FT_ResetDevice")
            pFT_ResetPort = GetProcAddress(ftd2xxDllHandle, "FT_ResetPort")
            pFT_CyclePort = GetProcAddress(ftd2xxDllHandle, "FT_CyclePort")
            pFT_Rescan = GetProcAddress(ftd2xxDllHandle, "FT_Rescan")
            pFT_Reload = GetProcAddress(ftd2xxDllHandle, "FT_Reload")
            pFT_Purge = GetProcAddress(ftd2xxDllHandle, "FT_Purge")
            pFT_SetTimeouts = GetProcAddress(ftd2xxDllHandle, "FT_SetTimeouts")
            pFT_SetBreakOn = GetProcAddress(ftd2xxDllHandle, "FT_SetBreakOn")
            pFT_SetBreakOff = GetProcAddress(ftd2xxDllHandle, "FT_SetBreakOff")
            pFT_GetDeviceInfo = GetProcAddress(ftd2xxDllHandle, "FT_GetDeviceInfo")
            pFT_SetResetPipeRetryCount = GetProcAddress(ftd2xxDllHandle, "FT_SetResetPipeRetryCount")
            pFT_StopInTask = GetProcAddress(ftd2xxDllHandle, "FT_StopInTask")
            pFT_RestartInTask = GetProcAddress(ftd2xxDllHandle, "FT_RestartInTask")
            pFT_GetDriverVersion = GetProcAddress(ftd2xxDllHandle, "FT_GetDriverVersion")
            pFT_GetLibraryVersion = GetProcAddress(ftd2xxDllHandle, "FT_GetLibraryVersion")
            pFT_SetDeadmanTimeout = GetProcAddress(ftd2xxDllHandle, "FT_SetDeadmanTimeout")
            pFT_SetChars = GetProcAddress(ftd2xxDllHandle, "FT_SetChars")
            pFT_SetEventNotification = GetProcAddress(ftd2xxDllHandle, "FT_SetEventNotification")
            pFT_GetComPortNumber = GetProcAddress(ftd2xxDllHandle, "FT_GetComPortNumber")
            pFT_SetLatencyTimer = GetProcAddress(ftd2xxDllHandle, "FT_SetLatencyTimer")
            pFT_GetLatencyTimer = GetProcAddress(ftd2xxDllHandle, "FT_GetLatencyTimer")
            pFT_SetBitMode = GetProcAddress(ftd2xxDllHandle, "FT_SetBitMode")
            pFT_GetBitMode = GetProcAddress(ftd2xxDllHandle, "FT_GetBitMode")
            pFT_SetUSBParameters = GetProcAddress(ftd2xxDllHandle, "FT_SetUSBParameters")
            pFT_VendorCmdGet = GetProcAddress(ftd2xxDllHandle, "FT_VendorCmdGet")
            pFT_VendorCmdSet = GetProcAddress(ftd2xxDllHandle, "FT_VendorCmdSet")
        End Sub

        ''' <summary>
        ''' Destructor for the <see cref="Ftdi"/> class.
        ''' </summary>
        Protected Overrides Sub Finalize()
            FreeLibrary(ftd2xxDllHandle)
            ftd2xxDllHandle = 0
        End Sub

#End Region '/CTOR

#Region "CLOSED FIELDS"

        Private ftd2xxDllHandle As Integer
        Private ftHandle As Integer
        Private pFT_CreateDeviceInfoList As Integer
        Private pFT_GetDeviceInfoDetail As Integer
        Private pFT_Open As Integer
        Private pFT_OpenEx As Integer
        Private pFT_Close As Integer
        Private pFT_Read As Integer
        Private pFT_Write As Integer
        Private pFT_GetQueueStatus As Integer
        Private pFT_GetModemStatus As Integer
        Private pFT_GetStatus As Integer
        Private pFT_SetBaudRate As Integer
        Private pFT_SetDataCharacteristics As Integer
        Private pFT_SetFlowControl As Integer
        Private pFT_SetDtr As Integer
        Private pFT_ClrDtr As Integer
        Private pFT_SetRts As Integer
        Private pFT_ClrRts As Integer
        Private pFT_ResetDevice As Integer
        Private pFT_ResetPort As Integer
        Private pFT_CyclePort As Integer
        Private pFT_Rescan As Integer
        Private pFT_Reload As Integer
        Private pFT_Purge As Integer
        Private pFT_SetTimeouts As Integer
        Private pFT_SetBreakOn As Integer
        Private pFT_SetBreakOff As Integer
        Private pFT_GetDeviceInfo As Integer
        Private pFT_SetResetPipeRetryCount As Integer
        Private pFT_StopInTask As Integer
        Private pFT_RestartInTask As Integer
        Private pFT_GetDriverVersion As Integer
        Private pFT_GetLibraryVersion As Integer
        Private pFT_SetDeadmanTimeout As Integer
        Private pFT_SetChars As Integer
        Private pFT_SetEventNotification As Integer
        Private pFT_GetComPortNumber As Integer
        Private pFT_SetLatencyTimer As Integer
        Private pFT_GetLatencyTimer As Integer
        Private pFT_SetBitMode As Integer
        Private pFT_GetBitMode As Integer
        Private pFT_SetUSBParameters As Integer
        Private pFT_VendorCmdGet As Integer
        Private pFT_VendorCmdSet As Integer

        'Константы:
        Private Const FT_OPEN_BY_SERIAL_NUMBER As UInteger = 1UI
        Private Const FT_OPEN_BY_DESCRIPTION As UInteger = 2UI
        Private Const FT_OPEN_BY_LOCATION As UInteger = 4UI
        Private Const FT_DEFAULT_BAUD_RATE As UInteger = 9600UI
        Private Const FT_COM_PORT_NOT_ASSIGNED As Integer = -1
        Private Const FT_DEFAULT_DEADMAN_TIMEOUT As UInteger = 5000UI
        Private Const FT_DEFAULT_LATENCY As Byte = 16
        Private Const FT_DEFAULT_IN_TRANSFER_SIZE As UInteger = 4096UI
        Private Const FT_DEFAULT_OUT_TRANSFER_SIZE As UInteger = 4096UI

#End Region '/CLOSED FIELDS

#Region "READ-ONLY PROPS"

        ''' <summary>
        ''' Загружена ли библиотека.
        ''' </summary>
        Public ReadOnly Property LibraryLoaded As Boolean
            Get
                Return (ftd2xxDllHandle <> 0)
            End Get
        End Property

        ''' <summary>
        ''' Открыто ли устройство.
        ''' </summary>
        Public ReadOnly Property IsDeviceOpen As Boolean
            Get
                Return (ftHandle <> 0)
            End Get
        End Property

        ''' <summary>
        ''' Признак канала (A, B и т.д.).
        ''' </summary>
        Private ReadOnly Property InterfaceIdentifier As String
            Get
                If IsDeviceOpen Then
                    Dim ft_DEVICE As FT_DEVICE = GetDeviceType()
                    If (ft_DEVICE = FT_DEVICE.FT_DEVICE_2232) OrElse (ft_DEVICE = FT_DEVICE.FT_DEVICE_2232H) OrElse (ft_DEVICE = FT_DEVICE.FT_DEVICE_4232H) Then
                        Dim text As String = GetDescription()
                        Return text.Substring(text.Length - 1)
                    End If
                End If
                Return String.Empty
            End Get
        End Property

#End Region '/READ-ONLY PROPS

#Region "OPEN, CLOSE"

        ''' <summary>
        ''' Opens the FTDI device with the specified index.  
        ''' </summary>
        ''' <param name="index">Index of the device to open.
        ''' Note that this cannot be guaranteed to open a specific device.</param>
        ''' <remarks>Initialises the device to 8 data bits, 1 stop bit, no parity, no flow control and 9600 Baud.</remarks>
        Public Sub OpenByIndex(index As UInteger)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded Then
                Dim tFT_Open As OpenDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_Open), GetType(OpenDelegate)), OpenDelegate)
                status = tFT_Open(index, ftHandle)
                If (status <> FT_STATUS.FT_OK) Then
                    ftHandle = 0
                End If
                If IsDeviceOpen Then
                    Dim uWordLength As Byte = 8
                    Dim uStopBits As Byte = 0
                    Dim uParity As Byte = 0
                    Dim setDCDeleg As SetDataCharacteristicsDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_SetDataCharacteristics), GetType(SetDataCharacteristicsDelegate)), SetDataCharacteristicsDelegate)
                    status = setDCDeleg.Invoke(ftHandle, uWordLength, uStopBits, uParity)

                    Dim usFlowControl As UShort = 0US
                    Dim uXon As Byte = 17
                    Dim uXoff As Byte = 19
                    Dim setFCDeleg As SetFlowControlDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_SetFlowControl), GetType(SetFlowControlDelegate)), SetFlowControlDelegate)
                    status = status Or setFCDeleg(ftHandle, usFlowControl, uXon, uXoff)

                    Dim dwBaudRate As UInteger = FT_DEFAULT_BAUD_RATE
                    Dim setBRDeleg As SetBaudRateDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_SetBaudRate), GetType(SetBaudRateDelegate)), SetBaudRateDelegate)
                    status = status Or setBRDeleg(ftHandle, dwBaudRate)
                End If
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Opens the FTDI device with the specified serial number.  
        ''' </summary>
        ''' <param name="serialnumber">Serial number of the device to open.</param>
        ''' <remarks>Initialises the device to 8 data bits, 1 stop bit, no parity, no flow control and 9600 Baud.</remarks>
        Public Sub OpenBySerialNumber(serialnumber As String)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded Then
                Dim tFT_OpenEx As OpenExDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_OpenEx, IntPtr), GetType(OpenExDelegate)), OpenExDelegate)
                status = tFT_OpenEx(serialnumber, FT_OPEN_BY_SERIAL_NUMBER, ftHandle)
                If (status <> FT_STATUS.FT_OK) Then
                    ftHandle = 0
                End If
                If IsDeviceOpen Then
                    Dim uWordLength As Byte = 8
                    Dim uStopBits As Byte = 0
                    Dim uParity As Byte = 0
                    Dim setDCDeleg As SetDataCharacteristicsDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetDataCharacteristics, IntPtr), GetType(SetDataCharacteristicsDelegate)), SetDataCharacteristicsDelegate)
                    status = setDCDeleg(ftHandle, uWordLength, uStopBits, uParity)

                    Dim usFlowControl As UShort = 0US
                    Dim uXon As Byte = 17
                    Dim uXoff As Byte = 19
                    Dim setFCDeleg As SetFlowControlDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetFlowControl, IntPtr), GetType(SetFlowControlDelegate)), SetFlowControlDelegate)
                    status = status Or setFCDeleg(ftHandle, usFlowControl, uXon, uXoff)

                    Dim dwBaudRate As UInteger = FT_DEFAULT_BAUD_RATE
                    Dim setBRDeleg As SetBaudRateDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetBaudRate, IntPtr), GetType(SetBaudRateDelegate)), SetBaudRateDelegate)
                    status = status Or setBRDeleg(ftHandle, dwBaudRate)
                End If
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Opens the FTDI device with the specified description.  
        ''' </summary>
        ''' <param name="description">Description of the device to open.</param>
        ''' <remarks>Initialises the device to 8 data bits, 1 stop bit, no parity, no flow control and 9600 Baud.</remarks>
        Public Sub OpenByDescription(description As String)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded Then
                Dim tFT_OpenEx As OpenExDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_OpenEx, IntPtr), GetType(OpenExDelegate)), OpenExDelegate)
                status = tFT_OpenEx(description, FT_OPEN_BY_DESCRIPTION, ftHandle)
                If (status <> FT_STATUS.FT_OK) Then
                    ftHandle = 0
                End If
                If IsDeviceOpen Then
                    Dim uWordLength As Byte = 8
                    Dim uStopBits As Byte = 0
                    Dim uParity As Byte = 0
                    Dim setDCDeleg As SetDataCharacteristicsDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetDataCharacteristics, IntPtr), GetType(SetDataCharacteristicsDelegate)), SetDataCharacteristicsDelegate)
                    status = setDCDeleg(ftHandle, uWordLength, uStopBits, uParity)

                    Dim usFlowControl As UShort = 0US
                    Dim uXon As Byte = 17
                    Dim uXoff As Byte = 19
                    Dim setFCDeleg As SetFlowControlDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetFlowControl, IntPtr), GetType(SetFlowControlDelegate)), SetFlowControlDelegate)
                    status = status Or setFCDeleg(ftHandle, usFlowControl, uXon, uXoff)

                    Dim dwBaudRate As UInteger = FT_DEFAULT_BAUD_RATE
                    Dim setBRDeleg As SetBaudRateDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetBaudRate, IntPtr), GetType(SetBaudRateDelegate)), SetBaudRateDelegate)
                    status = status Or setBRDeleg(ftHandle, dwBaudRate)
                End If
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Opens the FTDI device at the specified physical location.  
        ''' </summary>
        ''' <param name="location">Location of the device to open.</param>
        ''' <remarks>Initialises the device to 8 data bits, 1 stop bit, no parity, no flow control and 9600 Baud.</remarks>
        Public Sub OpenByLocation(location As UInteger)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded Then
                Dim tFT_OpenExLoc As OpenExLocDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_OpenEx, IntPtr), GetType(OpenExLocDelegate)), OpenExLocDelegate)
                status = tFT_OpenExLoc(location, FT_OPEN_BY_LOCATION, ftHandle)
                If (status <> FT_STATUS.FT_OK) Then
                    ftHandle = 0
                End If
                If IsDeviceOpen Then
                    Dim uWordLength As Byte = 8
                    Dim uStopBits As Byte = 0
                    Dim uParity As Byte = 0
                    Dim setDCDeleg As SetDataCharacteristicsDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetDataCharacteristics, IntPtr), GetType(SetDataCharacteristicsDelegate)), SetDataCharacteristicsDelegate)
                    status = setDCDeleg(ftHandle, uWordLength, uStopBits, uParity)

                    Dim usFlowControl As UShort = 0US
                    Dim uXon As Byte = 17
                    Dim uXoff As Byte = 19
                    Dim setFCDeleg As SetFlowControlDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetFlowControl, IntPtr), GetType(SetFlowControlDelegate)), SetFlowControlDelegate)
                    status = status Or setFCDeleg(ftHandle, usFlowControl, uXon, uXoff)

                    Dim dwBaudRate As UInteger = FT_DEFAULT_BAUD_RATE
                    Dim setBRDeleg As SetBaudRateDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetBaudRate, IntPtr), GetType(SetBaudRateDelegate)), SetBaudRateDelegate)
                    status = status Or setBRDeleg(ftHandle, dwBaudRate)
                End If
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Closes the handle to an open FTDI device.  
        ''' </summary>
        Public Sub Close()
            Dim ft_STATUS As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded Then
                Dim closeDeleg As CloseDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_Close), GetType(CloseDelegate)), CloseDelegate)
                ft_STATUS = closeDeleg.Invoke(ftHandle)
                If (ft_STATUS = FT_STATUS.FT_OK) Then
                    ftHandle = 0
                End If
            End If
            CheckErrors(ft_STATUS)
        End Sub

#End Region '/OPEN, CLOSE

#Region "INFO"

        ''' <summary>
        ''' Gets the number of FTDI devices available.  
        ''' </summary>
        ''' <returns>The number of FTDI devices available.</returns>
        Public Function GetNumberOfDevices() As Integer
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            Dim devcount As UInteger = 0
            If LibraryLoaded Then
                Dim cdilDeleg As CreateDeviceInfoListDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_CreateDeviceInfoList), GetType(CreateDeviceInfoListDelegate)), CreateDeviceInfoListDelegate)
                status = cdilDeleg.Invoke(devcount)
            End If
            CheckErrors(status)
            Return CInt(devcount)
        End Function

        ''' <summary>
        ''' Gets information on all of the FTDI devices available.  
        ''' </summary>
        ''' <returns>An array of type <see cref="FT_DEVICE_INFO_NODE"/> to contain the device information for all available devices.</returns>
        ''' <exception cref="FTD2XX_NET.FTDI.FT_EXCEPTION">Thrown when the supplied buffer is not large enough to contain the device info list.</exception>
        Public Function GetDeviceList() As FT_DEVICE_INFO_NODE()
            Dim ft_STATUS As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded Then
                Dim numDevices As UInteger = 0UI
                Dim tFT_CreateDeviceInfoList As CreateDeviceInfoListDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_CreateDeviceInfoList), GetType(CreateDeviceInfoListDelegate)), CreateDeviceInfoListDelegate)
                ft_STATUS = tFT_CreateDeviceInfoList(numDevices)
                CheckErrors(ft_STATUS)
                If (numDevices > 0) Then
                    Dim devicelist(CInt(numDevices - 1)) As FT_DEVICE_INFO_NODE
                    Dim gdiDeleg As GetDeviceInfoDetailDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_GetDeviceInfoDetail), GetType(GetDeviceInfoDetailDelegate)), GetDeviceInfoDetailDelegate)
                    For i As Integer = 0 To CInt(numDevices - 1)
                        devicelist(i) = New FT_DEVICE_INFO_NODE()
                        Dim sn As Byte() = New Byte(15) {}
                        Dim descr As Byte() = New Byte(63) {}
                        ft_STATUS = gdiDeleg(CUInt(i), devicelist(i).Flags, devicelist(i).Type, devicelist(i).ID, devicelist(i).LocId, sn, descr, devicelist(i).ftHandle)
                        CheckErrors(ft_STATUS)
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
            Else
                CheckErrors(ft_STATUS)
            End If
            Return New FT_DEVICE_INFO_NODE() {}
        End Function

        ''' <summary>
        ''' Gets the chip type of the current device.
        ''' </summary>
        ''' <returns>The FTDI chip type of the current device.</returns>
        Public Function GetDeviceType() As FT_DEVICE
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded Then
                If (pFT_GetDeviceInfo <> 0) Then
                    If IsDeviceOpen Then
                        Dim deviceType As FT_DEVICE = FT_DEVICE.FT_DEVICE_UNKNOWN
                        Dim num As UInteger = 0
                        Dim sn As Byte() = New Byte(15) {}
                        Dim descr As Byte() = New Byte(63) {}
                        Dim gdiDeleg As GetDeviceInfoDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_GetDeviceInfo, IntPtr), GetType(GetDeviceInfoDelegate)), GetDeviceInfoDelegate)
                        status = gdiDeleg(ftHandle, deviceType, num, sn, descr, 0)
                        CheckErrors(status)
                        Return deviceType
                    End If
                End If
            End If
            CheckErrors(status)
            Return FT_DEVICE.FT_DEVICE_UNKNOWN
        End Function

        ''' <summary>
        ''' Gets the Vendor ID and Product ID of the current device.
        ''' </summary>
        ''' <returns>The device ID (Vendor ID and Product ID) of the current device.</returns>
        Public Function GetDeviceID() As UInteger
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            Dim deviceID As UInteger
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim dev As FT_DEVICE = FT_DEVICE.FT_DEVICE_UNKNOWN
                Dim sn As Byte() = New Byte(15) {}
                Dim descr As Byte() = New Byte(63) {}
                Dim gdiDeleg As GetDeviceInfoDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_GetDeviceInfo, IntPtr), GetType(GetDeviceInfoDelegate)), GetDeviceInfoDelegate)
                status = gdiDeleg(ftHandle, dev, deviceID, sn, descr, 0)
            End If
            CheckErrors(status)
            Return deviceID
        End Function

        ''' <summary>
        ''' Gets the description of the current device.
        ''' </summary>
        ''' <returns>The description of the current device.</returns>
        Public Function GetDescription() As String
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            Dim description As String = String.Empty
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim sn As Byte() = New Byte(15) {}
                Dim descr As Byte() = New Byte(63) {}
                Dim num As UInteger = 0UI
                Dim dev As FT_DEVICE = FT_DEVICE.FT_DEVICE_UNKNOWN
                Dim gdiDeleg As GetDeviceInfoDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_GetDeviceInfo, IntPtr), GetType(GetDeviceInfoDelegate)), GetDeviceInfoDelegate)
                status = gdiDeleg(ftHandle, dev, num, sn, descr, 0)
                description = Encoding.ASCII.GetString(descr)
                description = description.Substring(0, description.IndexOf(vbNullChar))
            End If
            CheckErrors(status)
            Return description
        End Function

        ''' <summary>
        ''' Gets the serial number of the current device.
        ''' </summary>
        Public Function GetSerialNumber() As String
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            Dim serialNumber As String = String.Empty
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim sn As Byte() = New Byte(15) {}
                Dim descr As Byte() = New Byte(63) {}
                Dim num As UInteger = 0UI
                Dim dev As FT_DEVICE = FT_DEVICE.FT_DEVICE_UNKNOWN
                Dim gdiDeleg As GetDeviceInfoDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_GetDeviceInfo, IntPtr), GetType(GetDeviceInfoDelegate)), GetDeviceInfoDelegate)
                status = gdiDeleg(ftHandle, dev, num, sn, descr, 0)
                serialNumber = Encoding.ASCII.GetString(sn)
                serialNumber = serialNumber.Substring(0, serialNumber.IndexOf(vbNullChar))
            End If
            CheckErrors(status)
            Return serialNumber
        End Function

        ''' <summary>
        ''' Gets the number of bytes available in the receive buffer.
        ''' </summary>
        Public Function GetRxBytesAvailable() As Integer
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            Dim rxQueue As UInteger
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim gqsDeleg As GetQueueStatusDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_GetQueueStatus, IntPtr), GetType(GetQueueStatusDelegate)), GetQueueStatusDelegate)
                status = gqsDeleg(ftHandle, rxQueue)
            End If
            CheckErrors(status)
            Return CInt(rxQueue)
        End Function

        ''' <summary>
        ''' Gets the number of bytes waiting in the transmit buffer.
        ''' </summary>
        Public Function GetTxBytesWaiting() As Integer
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            Dim txQueue As UInteger
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim inTxLen As UInteger = 0UI
                Dim eventType As UInteger = 0UI
                Dim gsDeleg As GetStatusDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_GetStatus, IntPtr), GetType(GetStatusDelegate)), GetStatusDelegate)
                status = gsDeleg(ftHandle, inTxLen, txQueue, eventType)
            End If
            CheckErrors(status)
            Return CInt(txQueue)
        End Function

        ''' <summary>
        ''' Gets the event type after an event has fired. Can be used to distinguish which event has been triggered when waiting on multiple event types.
        ''' </summary>
        Public Function GetEventType() As UInteger
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            Dim eventType As UInteger
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim inTxLen As UInteger = 0UI
                Dim txQueue As UInteger = 0UI
                Dim gsDeleg As GetStatusDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_GetStatus, IntPtr), GetType(GetStatusDelegate)), GetStatusDelegate)
                status = gsDeleg(ftHandle, inTxLen, txQueue, eventType)
            End If
            CheckErrors(status)
            Return eventType
        End Function

        ''' <summary>
        ''' Gets the current modem status - a bit map representaion of the current modem status.
        ''' </summary>
        ''' <returns>A bit map representaion of the current modem status.</returns>
        Public Function GetModemStatus() As Byte
            Dim result As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            Dim modemStatus As Byte
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim status As UInteger = 0UI
                Dim gmsDeleg As GetModemStatusDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_GetModemStatus, IntPtr), GetType(GetModemStatusDelegate)), GetModemStatusDelegate)
                result = gmsDeleg(ftHandle, status)
                modemStatus = Convert.ToByte(status And &HFF)
            End If
            CheckErrors(result)
            Return modemStatus
        End Function

        ''' <summary>
        ''' Gets the current line status - a bit map representaion of the current line status.
        ''' </summary>
        Public Function GetLineStatus() As Byte
            Dim result As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            Dim lineStatus As Byte
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim status As UInteger = 0UI
                Dim gmsDeleg As GetModemStatusDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_GetModemStatus, IntPtr), GetType(GetModemStatusDelegate)), GetModemStatusDelegate)
                result = gmsDeleg(ftHandle, status)
                lineStatus = Convert.ToByte((status >> 8) And &HFF)
            End If
            CheckErrors(result)
            Return lineStatus
        End Function

        ''' <summary>
        ''' Gets the corresponding COM port number for the current device. If no COM port is exposed, an empty string is returned.
        ''' </summary>
        Public Function GetComPort() As String
            Dim result As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            Dim comPortName As String = String.Empty
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim portNum As Integer = FT_COM_PORT_NOT_ASSIGNED
                Dim gpnDeleg As GetComPortNumberDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_GetComPortNumber, IntPtr), GetType(GetComPortNumberDelegate)), GetComPortNumberDelegate)
                result = gpnDeleg(ftHandle, portNum)
                If (portNum <> FT_COM_PORT_NOT_ASSIGNED) Then
                    comPortName = "COM" & portNum.ToString()
                End If
            End If
            CheckErrors(result)
            Return comPortName
        End Function

        ''' <summary>
        ''' Gets the instantaneous state of the device IO pins - a bitmap value containing the instantaneous state of the device IO pins.
        ''' </summary>
        Public Function GetPinStates() As Byte
            Dim result As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            Dim bitMode As Byte
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim gbmDeleg As GetBitModeDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_GetBitMode, IntPtr), GetType(GetBitModeDelegate)), GetBitModeDelegate)
                result = gbmDeleg(ftHandle, bitMode)
            End If
            CheckErrors(result)
            Return bitMode
        End Function

#End Region '/INFO

#Region "SETTINGS"

        ''' <summary>
        ''' Sets the current Baud rate.
        ''' </summary>
        ''' <param name="baudRate">The desired Baud rate for the device.</param>
        Public Sub SetBaudRate(baudRate As Integer)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim brDeleg As SetBaudRateDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetBaudRate, IntPtr), GetType(SetBaudRateDelegate)), SetBaudRateDelegate)
                status = brDeleg(ftHandle, CUInt(baudRate))
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Sets the data bits, stop bits and parity for the device.
        ''' </summary>
        ''' <param name="dataBits">The number of data bits for UART data.</param>
        ''' <param name="stopBits">The number of stop bits for UART data.</param>
        ''' <param name="parity">The parity of the UART data.</param>
        Public Sub SetDataCharacteristics(dataBits As FT_DATA_BITS, stopBits As FT_STOP_BITS, parity As FT_PARITY)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim dcDeleg As SetDataCharacteristicsDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetDataCharacteristics, IntPtr), GetType(SetDataCharacteristicsDelegate)), SetDataCharacteristicsDelegate)
                status = dcDeleg(ftHandle, dataBits, stopBits, parity)
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Sets the flow control type.
        ''' </summary>
        ''' <param name="FlowControl">The type of flow control for the UART. </param>
        ''' <param name="Xon">The Xon character for Xon/Xoff flow control. Ignored if not using Xon/XOff flow control.</param>
        ''' <param name="Xoff">The Xoff character for Xon/Xoff flow control. Ignored if not using Xon/XOff flow control.</param>
        Public Sub SetFlowControl(flowControl As FT_FLOW_CONTROL, xon As Byte, xoff As Byte)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim fcDeleg As SetFlowControlDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetFlowControl, IntPtr), GetType(SetFlowControlDelegate)), SetFlowControlDelegate)
                status = fcDeleg(ftHandle, flowControl, xon, xoff)
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Asserts or de-asserts the Request To Send (RTS) line.
        ''' </summary>
        ''' <param name="enable">If true, asserts RTS. If false, de-asserts RTS.</param>
        Public Sub SetRts(enable As Boolean)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                If enable Then
                    Dim srtDlg As SetRtsDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetRts, IntPtr), GetType(SetRtsDelegate)), SetRtsDelegate)
                    status = srtDlg(ftHandle)
                Else
                    Dim crtDlg As ClrRtsDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_ClrRts, IntPtr), GetType(ClrRtsDelegate)), ClrRtsDelegate)
                    status = crtDlg(ftHandle)
                End If
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Asserts or de-asserts the Data Terminal Ready (DTR) line.
        ''' </summary>
        ''' <param name="enable">If true, asserts DTR. If false, de-asserts DTR.</param>
        Public Sub SetDtr(enable As Boolean)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                If enable Then
                    Dim sdrDlg As SetDtrDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetDtr, IntPtr), GetType(SetDtrDelegate)), SetDtrDelegate)
                    status = sdrDlg(ftHandle)
                Else
                    Dim cdrDlg As ClrDtrDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_ClrDtr, IntPtr), GetType(ClrDtrDelegate)), ClrDtrDelegate)
                    status = cdrDlg(ftHandle)
                End If
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Sets the read and write timeout values.
        ''' </summary>
        ''' <param name="readTimeout">Read timeout value in ms. A value of 0 indicates an infinite timeout.</param>
        ''' <param name="writeTimeout">Write timeout value in ms. A value of 0 indicates an infinite timeout.</param>
        Public Sub SetTimeouts(readTimeout As UInteger, writeTimeout As UInteger)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim stDlg As SetTimeoutsDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetTimeouts, IntPtr), GetType(SetTimeoutsDelegate)), SetTimeoutsDelegate)
                status = stDlg(ftHandle, readTimeout, writeTimeout)
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Sets or clears the break state.
        ''' </summary>
        ''' <param name="enable">If true, sets break on. If false, sets break off.</param>
        Public Sub SetBreak(enable As Boolean)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                If enable Then
                    Dim sbDlg As SetBreakOnDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_SetBreakOn), GetType(SetBreakOnDelegate)), SetBreakOnDelegate)
                    status = sbDlg(ftHandle)
                Else
                    Dim sboDlg As SetBreakOffDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_SetBreakOff), GetType(SetBreakOffDelegate)), SetBreakOffDelegate)
                    status = sboDlg(ftHandle)
                End If
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Gets or sets the reset pipe retry count. Default value is 50.
        ''' </summary>
        ''' <param name="resetPipeRetryCount">The reset pipe retry count. Electrically noisy environments may benefit from a larger value.</param>
        Public Sub SetResetPipeRetryCount(resetPipeRetryCount As Integer)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim rpDlg As SetResetPipeRetryCountDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_SetResetPipeRetryCount), GetType(SetResetPipeRetryCountDelegate)), SetResetPipeRetryCountDelegate)
                status = rpDlg(ftHandle, CUInt(resetPipeRetryCount))
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Gets the current FTDIBUS.SYS driver version number.
        ''' </summary>
        ''' <returns>The current driver version number.</returns>
        Public Function GetDriverVersion() As UInteger
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            Dim driverVersion As UInteger
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim dvDlg As GetDriverVersionDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_GetDriverVersion), GetType(GetDriverVersionDelegate)), GetDriverVersionDelegate)
                status = dvDlg(ftHandle, driverVersion)
            End If
            CheckErrors(status)
            Return driverVersion
        End Function

        ''' <summary>
        ''' Gets the current FTD2XX.DLL driver version number.
        ''' </summary>
        Public Function GetLibraryVersion() As UInteger
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            Dim libraryVersion As UInteger
            If LibraryLoaded Then
                Dim dvDlg As GetLibraryVersionDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_GetLibraryVersion), GetType(GetLibraryVersionDelegate)), GetLibraryVersionDelegate)
                status = dvDlg.Invoke(libraryVersion)
            End If
            CheckErrors(status)
            Return libraryVersion
        End Function

        ''' <summary>
        ''' Sets the USB deadman timeout value.
        ''' </summary>
        ''' <param name="deadmanTimeout">The deadman timeout value in ms.</param>
        Public Sub SetDeadmanTimeout(Optional deadmanTimeout As UInteger = FT_DEFAULT_DEADMAN_TIMEOUT)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim dtDlg As SetDeadmanTimeoutDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_SetDeadmanTimeout), GetType(SetDeadmanTimeoutDelegate)), SetDeadmanTimeoutDelegate)
                status = dtDlg(ftHandle, deadmanTimeout)
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Sets the value of the latency timer.
        ''' </summary>
        ''' <param name="Latency">The latency timer value in ms.
        ''' Valid values are 2...255 ms for FT232BM, FT245BM and FT2232 devices.
        ''' Valid values are 0...255 ms for other devices.</param>
        Public Sub SetLatency(Optional latency As Byte = FT_DEFAULT_LATENCY)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim dev As FT_DEVICE = GetDeviceType()
                If ((dev = FT_DEVICE.FT_DEVICE_BM) OrElse (dev = FT_DEVICE.FT_DEVICE_2232)) AndAlso (latency < 2) Then
                    latency = 2
                End If
                Dim slDlg As SetLatencyTimerDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_SetLatencyTimer), GetType(SetLatencyTimerDelegate)), SetLatencyTimerDelegate)
                status = slDlg(ftHandle, latency)
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Gets the value of the latency timer. Default value is <see cref="FT_DEFAULT_LATENCY "/> ms.
        ''' </summary>
        ''' <returns>The latency timer value in ms.</returns>
        Public Function GetLatency() As Byte
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            Dim latency As Byte
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim ltDlg As GetLatencyTimerDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_GetLatencyTimer), GetType(GetLatencyTimerDelegate)), GetLatencyTimerDelegate)
                status = ltDlg(ftHandle, latency)
            End If
            CheckErrors(status)
            Return latency
        End Function

        ''' <summary>
        ''' Sets the USB IN and OUT transfer sizes.
        ''' </summary>
        ''' <param name="inTs">The USB IN transfer size in bytes.</param>
        ''' <param name="outTs">The USB OUT transfer size in bytes.</param>
        Public Sub InTransferSize(Optional inTs As UInteger = FT_DEFAULT_IN_TRANSFER_SIZE, Optional outTs As UInteger = FT_DEFAULT_OUT_TRANSFER_SIZE)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim deleg As SetUsbParametersDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_SetUSBParameters), GetType(SetUsbParametersDelegate)), SetUsbParametersDelegate)
                status = deleg.Invoke(ftHandle, inTs, outTs)
            End If
            CheckErrors(status)
        End Sub

        '''<summary>
        '''Sets an event character, an error character and enables or disables them.
        '''</summary>
        '''<param name="eventChar">A character that will be tigger an IN to the host when this character is received.</param>
        '''<param name="eventCharEnable">Determines if the <paramref name="eventChar"/> is enabled or disabled.</param>
        '''<param name="errorChar">A character that will be inserted into the data stream to indicate that an error has occurred.</param>
        '''<param name="errorCharEnable">Determines if the <paramref name="errorChar"/> is enabled or disabled.</param>
        Public Sub SetCharacters(eventChar As Byte, eventCharEnable As Boolean, errorChar As Byte, errorCharEnable As Boolean)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim scDlg As SetCharsDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_SetChars), GetType(SetCharsDelegate)), SetCharsDelegate)
                status = scDlg(ftHandle, eventChar, Convert.ToByte(eventCharEnable), errorChar, Convert.ToByte(errorCharEnable))
            End If
            CheckErrors(status)
        End Sub

#End Region '/SETTINGS

#Region "READ, WRITE"

        ''' <summary>
        ''' Read data from an open FTDI device.
        ''' </summary>
        ''' <param name="numBytesToRead">The number of bytes requested from the device.</param>
        Public Function Read(numBytesToRead As Integer) As Byte()
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            Dim dataBuffer(numBytesToRead - 1) As Byte
            If LibraryLoaded AndAlso IsDeviceOpen Then
                If (dataBuffer.Length < numBytesToRead) Then
                    numBytesToRead = dataBuffer.Length
                End If
                Dim numBytesWereRead As UInteger
                Dim rdDlg As ReadDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_Read), GetType(ReadDelegate)), ReadDelegate)
                status = rdDlg(ftHandle, dataBuffer, CUInt(numBytesToRead), numBytesWereRead)
            End If
            CheckErrors(status)
            Return dataBuffer
        End Function

        ''' <summary>
        ''' Write data to an open FTDI device and returns number of bytes actually written to the device.
        ''' </summary>
        ''' <param name="dataBuffer">An array of bytes which contains the data to be written to the device.</param>
        ''' <param name="numBytesToWrite">The number of bytes to be written to the device.</param>
        ''' <returns>The number of bytes actually written to the device.</returns>
        Public Function Write(dataBuffer As Byte(), numBytesToWrite As Integer) As Integer
            Dim numBytesWritten As UInteger = 0UI
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim wDlg As WriteDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_Write), GetType(WriteDelegate)), WriteDelegate)
                status = wDlg(ftHandle, dataBuffer, CUInt(numBytesToWrite), numBytesWritten)
            End If
            CheckErrors(status)
            Return CInt(numBytesWritten)
        End Function

#End Region '/READ, WRITE

#Region "MANAGE"

        ''' <summary>
        ''' Reset an open FTDI device.
        ''' </summary>
        Public Sub ResetDevice()
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim rstDlg As ResetDeviceDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_ResetDevice), GetType(ResetDeviceDelegate)), ResetDeviceDelegate)
                status = rstDlg(ftHandle)
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Purge buffer constant definitions
        ''' </summary>
        Public Sub Purge(purgemask As UInteger)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim pDlg As PurgeDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_Purge), GetType(PurgeDelegate)), PurgeDelegate)
                status = pDlg(ftHandle, purgemask)
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Register for event notification.
        ''' </summary>
        ''' <remarks>After setting event notification, the event can be caught by executing the <see cref="EventWaitHandle.WaitOne()"/> method of the <see cref="EventWaitHandle"/>. 
        ''' If multiple event types are being monitored, the event that fired can be determined from the <see cref="GetEventType()"/> method.</remarks>
        ''' <param name="eventmask">The type of events to signal. 
        ''' Can be any combination of the following: <see cref="FT_EVENTS.FT_EVENT_RXCHAR"/>, <see cref="FT_EVENTS.FT_EVENT_MODEM_STATUS" />, <see cref="FT_EVENTS.FT_EVENT_LINE_STATUS"/>.</param>
        ''' <param name="eventhandle">Handle to the event that will receive the notification</param>
        Public Sub SetEventNotification(eventmask As UInteger, eventhandle As EventWaitHandle)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim deleg As SetEventNotificationDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_SetEventNotification), GetType(SetEventNotificationDelegate)), SetEventNotificationDelegate)
                status = deleg(ftHandle, eventmask, eventhandle.SafeWaitHandle)
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Stops the driver issuing USB in requests.
        ''' </summary>
        Public Sub StopInTask()
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim stplg As StopInTaskDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_StopInTask), GetType(StopInTaskDelegate)), StopInTaskDelegate)
                status = stplg(ftHandle)
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Resumes the driver issuing USB in requests.
        ''' </summary>
        Public Sub RestartInTask()
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim rstDlg As RestartInTaskDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_RestartInTask), GetType(RestartInTaskDelegate)), RestartInTaskDelegate)
                status = rstDlg(ftHandle)
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Resets the device port.
        ''' </summary>
        Public Sub ResetPort()
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim dlg As ResetPortDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_ResetPort), GetType(ResetPortDelegate)), ResetPortDelegate)
                status = dlg(ftHandle)
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Causes the device to be re-enumerated on the USB bus.
        ''' This is equivalent to unplugging and replugging the device.
        ''' Also calls <see cref="Close()"/> is successful, so no need to call this separately in the application.
        ''' </summary>
        Public Sub CyclePort()
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim cyDlg As CyclePortDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_CyclePort), GetType(CyclePortDelegate)), CyclePortDelegate)
                status = cyDlg(ftHandle)
                CheckErrors(status)

                Dim cDlg As CloseDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_Close), GetType(CloseDelegate)), CloseDelegate)
                status = cDlg(ftHandle)
                If (status = FT_STATUS.FT_OK) Then
                    ftHandle = 0
                End If
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Causes the system to check for USB hardware changes. 
        ''' This is equivalent to clicking on the "Scan for hardware changes" button in the Device Manager.
        ''' </summary>
        Public Sub Rescan()
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded Then
                Dim d As RescanDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_Rescan), GetType(RescanDelegate)), RescanDelegate)
                status = d.Invoke()
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Forces a reload of the driver for devices with a specific VID and PID combination.
        ''' </summary>
        ''' <remarks>If the VID and PID parameters are 0, the drivers for USB root hubs will be reloaded, causing all USB devices connected to reload their drivers</remarks>
        ''' <param name="VendorID">Vendor ID of the devices to have the driver reloaded</param>
        ''' <param name="ProductID">Product ID of the devices to have the driver reloaded</param>
        Public Sub Reload(vendorID As UShort, productID As UShort)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded Then
                Dim dlg As ReloadDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_Reload), GetType(ReloadDelegate)), ReloadDelegate)
                status = dlg(vendorID, productID)
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Puts the device in a mode other than the default UART or FIFO mode.
        ''' </summary>
        ''' <param name="mask">Sets up which bits are inputs And which are outputs. 
        ''' A bit value of 0 sets the corresponding pin to an input, a bit value of 1 sets the corresponding pin to an output.
        ''' In the case of CBUS Bit Bang, the upper nibble of this value controls which pins are inputs and outputs, while the lower nibble controls which of the outputs are high and low.</param>
        ''' <param name="bitMode">
        ''' <list type="bullet">
        ''' <item>For FT232H devices, valid values are <see cref="FT_BIT_MODES.FT_BIT_MODE_RESET"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_ASYNC_BITBANG"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_MPSSE"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_SYNC_BITBANG"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_CBUS_BITBANG"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_MCU_HOST"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_FAST_SERIAL"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_SYNC_FIFO"/>.</item>
        ''' <item>For FT2232H devices, valid values are <see cref="FT_BIT_MODES.FT_BIT_MODE_RESET"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_ASYNC_BITBANG"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_MPSSE"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_SYNC_BITBANG"/>, FT_BIT_MODE_MCU_HOST, FT_BIT_MODE_FAST_SERIAL, FT_BIT_MODE_SYNC_FIFO.</item>
        ''' <item>For FT4232H devices, valid values are <see cref="FT_BIT_MODES.FT_BIT_MODE_RESET"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_ASYNC_BITBANG"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_MPSSE"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_SYNC_BITBANG"/>.</item>
        ''' <item>For FT232R devices, valid values are <see cref="FT_BIT_MODES.FT_BIT_MODE_RESET"/>,  <see cref="FT_BIT_MODES.FT_BIT_MODE_ASYNC_BITBANG"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_SYNC_BITBANG"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_CBUS_BITBANG"/>.</item>
        ''' <item>For FT245R devices, valid values are <see cref="FT_BIT_MODES.FT_BIT_MODE_RESET"/>,  <see cref="FT_BIT_MODES.FT_BIT_MODE_ASYNC_BITBANG"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_SYNC_BITBANG"/>.</item>
        ''' <item>For FT2232 devices, valid values are <see cref="FT_BIT_MODES.FT_BIT_MODE_RESET"/>,  <see cref="FT_BIT_MODES.FT_BIT_MODE_ASYNC_BITBANG"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_MPSSE"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_SYNC_BITBANG"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_MCU_HOST"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_FAST_SERIAL"/>.</item>
        ''' <item>For FT232B And FT245B devices, valid values are <see cref="FT_BIT_MODES.FT_BIT_MODE_RESET"/>, <see cref="FT_BIT_MODES.FT_BIT_MODE_ASYNC_BITBANG"/>.</item>
        ''' </list>
        ''' </param>
        ''' <exception cref="FTD2XX_NET.Ftdi.FT_EXCEPTION">Thrown when the current device does not support the requested bit mode.</exception>
        Public Sub SetBitMode(mask As Byte, bitMode As FT_BIT_MODES)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                Dim dev As FT_DEVICE = GetDeviceType()
                Select Case dev
                    Case FT_DEVICE.FT_DEVICE_AM
                        CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                    Case FT_DEVICE.FT_DEVICE_100AX
                        CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                    Case FT_DEVICE.FT_DEVICE_BM
                        If (bitMode <> 0) Then
                            If ((bitMode And 1) = 0) Then
                                CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                            End If
                        End If
                    Case FT_DEVICE.FT_DEVICE_2232
                        If (bitMode <> 0) Then
                            If ((bitMode And 31) = 0) Then
                                CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                            End If
                            If (bitMode = 2) AndAlso (Me.InterfaceIdentifier <> "A") Then
                                CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                            End If
                        End If
                    Case FT_DEVICE.FT_DEVICE_232R
                        If (bitMode <> 0) Then
                            If ((bitMode And 37) = 0) Then
                                CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                            End If
                        End If
                    Case FT_DEVICE.FT_DEVICE_2232H
                        If (bitMode <> 0) Then
                            If ((bitMode And 95) = 0) Then
                                CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                            End If
                            If (bitMode = 8 Or bitMode = 64) AndAlso (Me.InterfaceIdentifier <> "A") Then
                                CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                            End If
                        End If
                    Case FT_DEVICE.FT_DEVICE_4232H
                        If (bitMode <> 0) Then
                            If ((bitMode And 7) = 0) Then
                                CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                            End If
                            If (bitMode = 2) AndAlso (Me.InterfaceIdentifier <> "A" AndAlso Me.InterfaceIdentifier <> "B") Then
                                CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                            End If
                        End If
                    Case FT_DEVICE.FT_DEVICE_232H
                        If (bitMode > 64) Then
                            CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                        End If
                End Select

                If (dev = FT_DEVICE.FT_DEVICE_AM) Then
                    CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                ElseIf (dev = FT_DEVICE.FT_DEVICE_100AX) Then
                    CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                ElseIf (dev = FT_DEVICE.FT_DEVICE_BM) AndAlso (bitMode <> 0) Then
                    If ((bitMode And 1) = 0) Then
                        CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                    End If
                ElseIf (dev = FT_DEVICE.FT_DEVICE_2232) AndAlso (bitMode <> 0) Then
                    If ((bitMode And 31) = 0) Then
                        CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                    End If
                    If (bitMode = 2) AndAlso (Me.InterfaceIdentifier <> "A") Then
                        CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                    End If
                ElseIf (dev = FT_DEVICE.FT_DEVICE_232R) AndAlso (bitMode <> 0) Then
                    If ((bitMode And 37) = 0) Then
                        CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                    End If

                ElseIf (dev = FT_DEVICE.FT_DEVICE_2232H) AndAlso (bitMode <> 0) Then
                    If ((bitMode And 95) = 0) Then
                        CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                    End If
                    If (bitMode = 8 Or bitMode = 64) AndAlso (Me.InterfaceIdentifier <> "A") Then
                        CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                    End If
                ElseIf (dev = FT_DEVICE.FT_DEVICE_4232H) AndAlso (bitMode <> 0) Then
                    If ((bitMode And 7) = 0) Then
                        CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                    End If
                    If (bitMode = 2) AndAlso (Me.InterfaceIdentifier <> "A" AndAlso Me.InterfaceIdentifier <> "B") Then
                        CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                    End If
                ElseIf (dev = FT_DEVICE.FT_DEVICE_232H) AndAlso (bitMode <> 0) AndAlso (bitMode > 64) Then
                    CheckErrors(status, FT_ERROR.FT_INVALID_BITMODE)
                End If
                Dim tFT_SetBitMode As SetBitModeDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_SetBitMode), GetType(SetBitModeDelegate)), SetBitModeDelegate)
                status = tFT_SetBitMode(ftHandle, mask, bitMode)
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Get data from the FT4222 using the vendor command interface.
        ''' </summary>
        Public Sub VendorCmdGet(request As UShort, buf As Byte(), len As UShort)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                If (pFT_VendorCmdGet <> 0) Then
                    Dim tFT_VendorCmdGet As VendorCmdGetDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_VendorCmdGet), GetType(VendorCmdGetDelegate)), VendorCmdGetDelegate)
                    status = tFT_VendorCmdGet(ftHandle, request, buf, len)
                End If
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Set data from the FT4222 using the vendor command interface.
        ''' </summary>
        Public Sub VendorCmdSet(request As UShort, buf As Byte(), len As UShort)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                If (pFT_VendorCmdSet <> 0) Then
                    Dim tFT_VendorCmdSet As VendorCmdSetDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_VendorCmdSet), GetType(VendorCmdSetDelegate)), VendorCmdSetDelegate)
                    status = tFT_VendorCmdSet(ftHandle, request, buf, len)
                End If
            End If
            CheckErrors(status)
        End Sub

#End Region '/MANAGE

#Region "HELPERS"

        ''' <summary>
        ''' Проверяет результат выполнения метода и выбрасывает исключение, если статус с ошибкой.
        ''' </summary>
        Private Sub CheckErrors(status As FT_STATUS, Optional additionalInfo As FT_ERROR = FT_ERROR.FT_NO_ERROR)
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

#Region "NESTED TYPES"

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

        Private Enum FT_ERROR
            FT_NO_ERROR
            FT_INCORRECT_DEVICE
            FT_INVALID_BITMODE
            FT_BUFFER_SIZE
        End Enum

        Public Enum FT_DATA_BITS As Byte
            FT_BITS_8 = 8
            FT_BITS_7 = 7
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

        Public Enum FT_PURGE As Byte
            FT_PURGE_RX = 1
            FT_PURGE_TX = 2
        End Enum

        <Flags()>
        Public Enum FT_MODEM_STATUS As Byte
            FT_CTS = 16
            FT_DSR = 32
            FT_RI = 64
            FT_DCD = 128
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
            FT_DEVICE_BM
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
            Public ftHandle As Integer
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

    End Class '/Ftdi

End Namespace
