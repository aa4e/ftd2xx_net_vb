Imports System.Runtime.InteropServices
Imports System.Runtime.Serialization
Imports System.Threading
Imports System.Text

Namespace FTD2XX_NET

    Public Class Ftdi

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
            Else
                Throw New Exception("Ошибка загрузки библиотеки ftd2xx.")
            End If
        End Sub

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
            pFT_ReadEE = GetProcAddress(ftd2xxDllHandle, "FT_ReadEE")
            pFT_WriteEE = GetProcAddress(ftd2xxDllHandle, "FT_WriteEE")
            pFT_EraseEE = GetProcAddress(ftd2xxDllHandle, "FT_EraseEE")
            pFT_EE_UASize = GetProcAddress(ftd2xxDllHandle, "FT_EE_UASize")
            pFT_EE_UARead = GetProcAddress(ftd2xxDllHandle, "FT_EE_UARead")
            pFT_EE_UAWrite = GetProcAddress(ftd2xxDllHandle, "FT_EE_UAWrite")
            pFT_EE_Read = GetProcAddress(ftd2xxDllHandle, "FT_EE_Read")
            pFT_EE_Program = GetProcAddress(ftd2xxDllHandle, "FT_EE_Program")
            pFT_EEPROM_Read = GetProcAddress(ftd2xxDllHandle, "FT_EEPROM_Read")
            pFT_EEPROM_Program = GetProcAddress(ftd2xxDllHandle, "FT_EEPROM_Program")
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
        Private pFT_ReadEE As Integer
        Private pFT_WriteEE As Integer
        Private pFT_EraseEE As Integer
        Private pFT_EE_UASize As Integer
        Private pFT_EE_UARead As Integer
        Private pFT_EE_UAWrite As Integer
        Private pFT_EE_Read As Integer
        Private pFT_EE_Program As Integer
        Private pFT_EEPROM_Read As Integer
        Private pFT_EEPROM_Program As Integer
        Private pFT_VendorCmdGet As Integer
        Private pFT_VendorCmdSet As Integer

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
            Dim ft_STATUS As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded Then
                If (pFT_Open <> 0) AndAlso (pFT_SetDataCharacteristics <> 0) AndAlso (pFT_SetFlowControl <> 0) AndAlso (pFT_SetBaudRate <> 0) Then
                    Dim tFT_Open As OpenDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_Open), GetType(OpenDelegate)), OpenDelegate)
                    ft_STATUS = tFT_Open(index, ftHandle)
                    If (ft_STATUS <> FT_STATUS.FT_OK) Then
                        ftHandle = 0
                    End If
                    If IsDeviceOpen Then
                        Dim uWordLength As Byte = 8
                        Dim uStopBits As Byte = 0
                        Dim uParity As Byte = 0
                        Dim setDCDeleg As SetDataCharacteristicsDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_SetDataCharacteristics), GetType(SetDataCharacteristicsDelegate)), SetDataCharacteristicsDelegate)
                        ft_STATUS = ft_STATUS Or setDCDeleg.Invoke(ftHandle, uWordLength, uStopBits, uParity)

                        Dim usFlowControl As UShort = 0US
                        Dim uXon As Byte = 17
                        Dim uXoff As Byte = 19
                        Dim setFCDeleg As SetFlowControlDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_SetFlowControl), GetType(SetFlowControlDelegate)), SetFlowControlDelegate)
                        ft_STATUS = ft_STATUS Or setFCDeleg(ftHandle, usFlowControl, uXon, uXoff)

                        Dim dwBaudRate As UInteger = FT_DEFAULT_BAUD_RATE
                        Dim setBRDeleg As SetBaudRateDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_SetBaudRate), GetType(SetBaudRateDelegate)), SetBaudRateDelegate)
                        ft_STATUS = ft_STATUS Or setBRDeleg(ftHandle, dwBaudRate)
                    End If
                Else
                    pFT_Open = 0
                    pFT_SetDataCharacteristics = 0
                    pFT_SetFlowControl = 0
                    pFT_SetBaudRate = 0
                End If
            End If
            CheckErrors(ft_STATUS)
        End Sub

        ''' <summary>
        ''' Opens the FTDI device with the specified serial number.  
        ''' </summary>
        ''' <param name="serialnumber">Serial number of the device to open.</param>
        ''' <remarks>Initialises the device to 8 data bits, 1 stop bit, no parity, no flow control and 9600 Baud.</remarks>
        Public Sub OpenBySerialNumber(serialnumber As String)
            Dim ft_STATUS As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded Then
                If (pFT_OpenEx <> 0) AndAlso (pFT_SetDataCharacteristics <> 0) AndAlso (pFT_SetFlowControl <> 0) AndAlso (pFT_SetBaudRate <> 0) Then
                    Dim tFT_OpenEx As OpenExDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_OpenEx, IntPtr), GetType(OpenExDelegate)), OpenExDelegate)
                    ft_STATUS = tFT_OpenEx(serialnumber, FT_OPEN_BY_SERIAL_NUMBER, ftHandle)
                    If (ft_STATUS <> FT_STATUS.FT_OK) Then
                        ftHandle = 0
                    End If
                    If IsDeviceOpen Then
                        Dim uWordLength As Byte = 8
                        Dim uStopBits As Byte = 0
                        Dim uParity As Byte = 0
                        Dim setDCDeleg As SetDataCharacteristicsDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetDataCharacteristics, IntPtr), GetType(SetDataCharacteristicsDelegate)), SetDataCharacteristicsDelegate)
                        ft_STATUS = ft_STATUS Or setDCDeleg(ftHandle, uWordLength, uStopBits, uParity)

                        Dim usFlowControl As UShort = 0US
                        Dim uXon As Byte = 17
                        Dim uXoff As Byte = 19
                        Dim setFCDeleg As SetFlowControlDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetFlowControl, IntPtr), GetType(SetFlowControlDelegate)), SetFlowControlDelegate)
                        ft_STATUS = ft_STATUS Or setFCDeleg(ftHandle, usFlowControl, uXon, uXoff)

                        Dim dwBaudRate As UInteger = FT_DEFAULT_BAUD_RATE
                        Dim setBRDeleg As SetBaudRateDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetBaudRate, IntPtr), GetType(SetBaudRateDelegate)), SetBaudRateDelegate)
                        ft_STATUS = ft_STATUS Or setBRDeleg(ftHandle, dwBaudRate)
                    End If
                Else
                    pFT_OpenEx = 0
                    pFT_SetDataCharacteristics = 0
                    pFT_SetFlowControl = 0
                    pFT_SetBaudRate = 0
                End If
            End If
            CheckErrors(ft_STATUS)
        End Sub

        ''' <summary>
        ''' Opens the FTDI device with the specified description.  
        ''' </summary>
        ''' <param name="description">Description of the device to open.</param>
        ''' <remarks>Initialises the device to 8 data bits, 1 stop bit, no parity, no flow control and 9600 Baud.</remarks>
        Public Sub OpenByDescription(description As String)
            Dim ft_STATUS As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded Then
                If (pFT_OpenEx <> 0) AndAlso (pFT_SetDataCharacteristics <> 0) AndAlso (pFT_SetFlowControl <> 0) AndAlso (pFT_SetBaudRate <> 0) Then
                    Dim tFT_OpenEx As OpenExDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_OpenEx, IntPtr), GetType(OpenExDelegate)), OpenExDelegate)
                    ft_STATUS = tFT_OpenEx(description, FT_OPEN_BY_DESCRIPTION, ftHandle)
                    If (ft_STATUS <> FT_STATUS.FT_OK) Then
                        ftHandle = 0
                    End If
                    If IsDeviceOpen Then
                        Dim uWordLength As Byte = 8
                        Dim uStopBits As Byte = 0
                        Dim uParity As Byte = 0
                        Dim setDCDeleg As SetDataCharacteristicsDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetDataCharacteristics, IntPtr), GetType(SetDataCharacteristicsDelegate)), SetDataCharacteristicsDelegate)
                        ft_STATUS = ft_STATUS Or setDCDeleg(ftHandle, uWordLength, uStopBits, uParity)

                        Dim usFlowControl As UShort = 0US
                        Dim uXon As Byte = 17
                        Dim uXoff As Byte = 19
                        Dim setFCDeleg As SetFlowControlDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetFlowControl, IntPtr), GetType(SetFlowControlDelegate)), SetFlowControlDelegate)
                        ft_STATUS = ft_STATUS Or setFCDeleg(ftHandle, usFlowControl, uXon, uXoff)

                        Dim dwBaudRate As UInteger = FT_DEFAULT_BAUD_RATE
                        Dim setBRDeleg As SetBaudRateDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetBaudRate, IntPtr), GetType(SetBaudRateDelegate)), SetBaudRateDelegate)
                        ft_STATUS = ft_STATUS Or setBRDeleg(ftHandle, dwBaudRate)
                    End If
                Else
                    pFT_OpenEx = 0
                    pFT_SetDataCharacteristics = 0
                    pFT_SetFlowControl = 0
                    pFT_SetBaudRate = 0
                End If
            End If
            CheckErrors(ft_STATUS)
        End Sub

        ''' <summary>
        ''' Opens the FTDI device at the specified physical location.  
        ''' </summary>
        ''' <param name="location">Location of the device to open.</param>
        ''' <remarks>Initialises the device to 8 data bits, 1 stop bit, no parity, no flow control and 9600 Baud.</remarks>
        Public Sub OpenByLocation(location As UInteger)
            Dim ft_STATUS As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded Then
                If (pFT_OpenEx <> 0) AndAlso (pFT_SetDataCharacteristics <> 0) AndAlso (pFT_SetFlowControl <> 0) AndAlso (pFT_SetBaudRate <> 0) Then
                    Dim tFT_OpenExLoc As OpenExLocDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_OpenEx, IntPtr), GetType(OpenExLocDelegate)), OpenExLocDelegate)
                    ft_STATUS = tFT_OpenExLoc(location, FT_OPEN_BY_LOCATION, ftHandle)
                    If (ft_STATUS <> FT_STATUS.FT_OK) Then
                        ftHandle = 0
                    End If
                    If IsDeviceOpen Then
                        Dim uWordLength As Byte = 8
                        Dim uStopBits As Byte = 0
                        Dim uParity As Byte = 0
                        Dim setDCDeleg As SetDataCharacteristicsDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetDataCharacteristics, IntPtr), GetType(SetDataCharacteristicsDelegate)), SetDataCharacteristicsDelegate)
                        ft_STATUS = ft_STATUS Or setDCDeleg(ftHandle, uWordLength, uStopBits, uParity)

                        Dim usFlowControl As UShort = 0US
                        Dim uXon As Byte = 17
                        Dim uXoff As Byte = 19
                        Dim setFCDeleg As SetFlowControlDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetFlowControl, IntPtr), GetType(SetFlowControlDelegate)), SetFlowControlDelegate)
                        ft_STATUS = ft_STATUS Or setFCDeleg(ftHandle, usFlowControl, uXon, uXoff)

                        Dim dwBaudRate As UInteger = FT_DEFAULT_BAUD_RATE
                        Dim setBRDeleg As SetBaudRateDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetBaudRate, IntPtr), GetType(SetBaudRateDelegate)), SetBaudRateDelegate)
                        ft_STATUS = ft_STATUS Or setBRDeleg(ftHandle, dwBaudRate)
                    End If
                Else
                    pFT_OpenEx = 0
                    pFT_SetDataCharacteristics = 0
                    pFT_SetFlowControl = 0
                    pFT_SetBaudRate = 0
                End If
            End If
            CheckErrors(ft_STATUS)
        End Sub

        ''' <summary>
        ''' Closes the handle to an open FTDI device.  
        ''' </summary>
        Public Sub Close()
            Dim ft_STATUS As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded Then
                If (pFT_Close <> 0) Then
                    Dim closeDeleg As CloseDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_Close), GetType(CloseDelegate)), CloseDelegate)
                    ft_STATUS = closeDeleg.Invoke(ftHandle)
                    If (ft_STATUS = FT_STATUS.FT_OK) Then
                        ftHandle = 0
                    End If
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
            Dim devcount As UInteger = 0
            If LibraryLoaded Then
                If (pFT_CreateDeviceInfoList <> 0) Then
                    Dim cdilDeleg As CreateDeviceInfoListDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_CreateDeviceInfoList), GetType(CreateDeviceInfoListDelegate)), CreateDeviceInfoListDelegate)
                    CheckErrors(cdilDeleg.Invoke(devcount))
                End If
            Else
                CheckErrors(FT_STATUS.FT_DLL_NOT_LOADED)
            End If
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
                If (pFT_CreateDeviceInfoList <> 0) AndAlso (pFT_GetDeviceInfoDetail <> 0) Then
                    Dim numDevices As UInteger = 0UI
                    Dim tFT_CreateDeviceInfoList As CreateDeviceInfoListDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_CreateDeviceInfoList), GetType(CreateDeviceInfoListDelegate)), CreateDeviceInfoListDelegate)
                    ft_STATUS = tFT_CreateDeviceInfoList(numDevices)
                    CheckErrors(ft_STATUS)
                    If (numDevices > 0) Then
                        Dim devicelist(CInt(numDevices - 1)) As FT_DEVICE_INFO_NODE
                        Dim tFT_GetDeviceInfoDetail As GetDeviceInfoDetailDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_GetDeviceInfoDetail), GetType(GetDeviceInfoDetailDelegate)), GetDeviceInfoDetailDelegate)
                        For i As Integer = 0 To CInt(numDevices - 1)
                            devicelist(i) = New FT_DEVICE_INFO_NODE()
                            Dim sn As Byte() = New Byte(15) {}
                            Dim descr As Byte() = New Byte(63) {}
                            ft_STATUS = tFT_GetDeviceInfoDetail(CUInt(i), devicelist(i).Flags, devicelist(i).Type, devicelist(i).ID, devicelist(i).LocId, sn, descr, devicelist(i).ftHandle)
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
                    pFT_CreateDeviceInfoList = 0
                    pFT_GetDeviceInfoDetail = 0
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
                        Dim tFT_GetDeviceInfo As GetDeviceInfoDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_GetDeviceInfo, IntPtr), GetType(GetDeviceInfoDelegate)), GetDeviceInfoDelegate)
                        status = tFT_GetDeviceInfo(ftHandle, deviceType, num, sn, descr, 0)
                        CheckErrors(status)
                        Return deviceType
                    End If
                Else
                    pFT_GetDeviceInfo = 0
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
            If LibraryLoaded Then
                If (pFT_GetDeviceInfo <> 0) Then
                    If IsDeviceOpen Then
                        Dim ft_DEVICE As FT_DEVICE = FT_DEVICE.FT_DEVICE_UNKNOWN
                        Dim sn As Byte() = New Byte(15) {}
                        Dim descr As Byte() = New Byte(63) {}
                        Dim tFT_GetDeviceInfo As GetDeviceInfoDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_GetDeviceInfo, IntPtr), GetType(GetDeviceInfoDelegate)), GetDeviceInfoDelegate)
                        status = tFT_GetDeviceInfo(ftHandle, ft_DEVICE, deviceID, sn, descr, 0)
                    End If
                Else
                    pFT_GetDeviceInfo = 0
                End If
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
            If LibraryLoaded Then
                If (pFT_GetDeviceInfo <> 0) Then
                    If IsDeviceOpen Then
                        Dim sn As Byte() = New Byte(15) {}
                        Dim descr As Byte() = New Byte(63) {}
                        Dim num As UInteger = 0UI
                        Dim ft_DEVICE As FT_DEVICE = FT_DEVICE.FT_DEVICE_UNKNOWN
                        Dim tFT_GetDeviceInfo As GetDeviceInfoDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_GetDeviceInfo, IntPtr), GetType(GetDeviceInfoDelegate)), GetDeviceInfoDelegate)
                        status = tFT_GetDeviceInfo(ftHandle, ft_DEVICE, num, sn, descr, 0)
                        description = Encoding.ASCII.GetString(descr)
                        description = description.Substring(0, description.IndexOf(vbNullChar))
                    End If
                Else
                    pFT_GetDeviceInfo = 0
                End If
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
            If LibraryLoaded Then
                If (pFT_GetDeviceInfo <> 0) Then
                    If IsDeviceOpen Then
                        Dim sn As Byte() = New Byte(15) {}
                        Dim descr As Byte() = New Byte(63) {}
                        Dim num As UInteger = 0UI
                        Dim ft_DEVICE As FT_DEVICE = FT_DEVICE.FT_DEVICE_UNKNOWN
                        Dim tFT_GetDeviceInfo As GetDeviceInfoDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_GetDeviceInfo, IntPtr), GetType(GetDeviceInfoDelegate)), GetDeviceInfoDelegate)
                        status = tFT_GetDeviceInfo(ftHandle, ft_DEVICE, num, sn, descr, 0)
                        serialNumber = Encoding.ASCII.GetString(sn)
                        serialNumber = serialNumber.Substring(0, serialNumber.IndexOf(vbNullChar))
                    End If
                Else
                    pFT_GetDeviceInfo = 0
                End If
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
            If LibraryLoaded Then
                If (pFT_GetQueueStatus <> 0) Then
                    If IsDeviceOpen Then
                        Dim tFT_GetQueueStatus As GetQueueStatusDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_GetQueueStatus, IntPtr), GetType(GetQueueStatusDelegate)), GetQueueStatusDelegate)
                        status = tFT_GetQueueStatus(ftHandle, rxQueue)
                    End If
                Else
                    pFT_GetQueueStatus = 0
                End If
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
            If LibraryLoaded Then
                If (pFT_GetStatus <> 0) Then
                    If IsDeviceOpen Then
                        Dim inTxLen As UInteger = 0UI
                        Dim eventType As UInteger = 0UI
                        Dim tFT_GetStatus As GetStatusDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_GetStatus, IntPtr), GetType(GetStatusDelegate)), GetStatusDelegate)
                        status = tFT_GetStatus(ftHandle, inTxLen, txQueue, eventType)
                    End If
                Else
                    pFT_GetStatus = 0
                End If
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
            If LibraryLoaded Then
                If (pFT_GetStatus <> 0) Then
                    If IsDeviceOpen Then
                        Dim inTxLen As UInteger = 0UI
                        Dim txQueue As UInteger = 0UI
                        Dim tFT_GetStatus As GetStatusDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_GetStatus, IntPtr), GetType(GetStatusDelegate)), GetStatusDelegate)
                        status = tFT_GetStatus(ftHandle, inTxLen, txQueue, eventType)
                    End If
                Else
                    pFT_GetStatus = 0
                End If
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
            If LibraryLoaded Then
                If (pFT_GetModemStatus <> 0) Then
                    Dim status As UInteger = 0UI
                    If IsDeviceOpen Then
                        Dim tFT_GetModemStatus As GetModemStatusDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_GetModemStatus, IntPtr), GetType(GetModemStatusDelegate)), GetModemStatusDelegate)
                        result = tFT_GetModemStatus(ftHandle, status)
                    End If
                    modemStatus = Convert.ToByte(status And &HFF)
                Else
                    pFT_GetModemStatus = 0
                End If
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
            If LibraryLoaded Then
                If (pFT_GetModemStatus <> 0) Then
                    Dim status As UInteger = 0UI
                    If IsDeviceOpen Then
                        Dim tFT_GetModemStatus As GetModemStatusDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_GetModemStatus, IntPtr), GetType(GetModemStatusDelegate)), GetModemStatusDelegate)
                        result = tFT_GetModemStatus(ftHandle, status)
                    End If
                    lineStatus = Convert.ToByte((status >> 8) And &HFF)
                Else
                    pFT_GetModemStatus = 0
                End If
            End If
            CheckErrors(result)
            Return lineStatus
        End Function

        ''' <summary>
        ''' Gets the size of the EEPROM user area size in bytes.
        ''' </summary>
        Public Function EeUserAreaSize() As Integer
            Dim result As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            Dim uaSize As UInteger
            If LibraryLoaded Then
                If (pFT_EE_UASize <> 0) Then
                    If IsDeviceOpen Then
                        Dim tFT_EE_UASize As EeUaSizeDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_EE_UASize, IntPtr), GetType(EeUaSizeDelegate)), EeUaSizeDelegate)
                        result = tFT_EE_UASize(ftHandle, uaSize)
                    End If
                Else
                    pFT_EE_UASize = 0
                End If
            End If
            CheckErrors(result)
            Return CInt(uaSize)
        End Function

        ''' <summary>
        ''' Gets the corresponding COM port number for the current device. If no COM port is exposed, an empty string is returned.
        ''' </summary>
        Public Function GetComPort() As String
            Dim result As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            Dim comPortName As String = String.Empty
            If LibraryLoaded Then
                If (pFT_GetComPortNumber <> 0) Then
                    Dim portNum As Integer = FT_COM_PORT_NOT_ASSIGNED
                    If IsDeviceOpen Then
                        Dim tFT_GetComPortNumber As GetComPortNumberDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_GetComPortNumber, IntPtr), GetType(GetComPortNumberDelegate)), GetComPortNumberDelegate)
                        result = tFT_GetComPortNumber(ftHandle, portNum)
                    End If
                    If (portNum <> FT_COM_PORT_NOT_ASSIGNED) Then
                        comPortName = "COM" & portNum.ToString()
                    End If
                Else
                    pFT_GetComPortNumber = 0
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
            If LibraryLoaded Then
                If (pFT_GetBitMode <> 0) Then
                    If IsDeviceOpen Then
                        Dim tFT_GetBitMode As GetBitModeDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_GetBitMode, IntPtr), GetType(GetBitModeDelegate)), GetBitModeDelegate)
                        result = tFT_GetBitMode(ftHandle, bitMode)
                    End If
                Else
                    pFT_GetBitMode = 0
                End If
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
            If LibraryLoaded Then
                If (pFT_SetBaudRate <> 0) Then
                    If IsDeviceOpen Then
                        Dim tFT_SetBaudRate As SetBaudRateDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetBaudRate, IntPtr), GetType(SetBaudRateDelegate)), SetBaudRateDelegate)
                        status = tFT_SetBaudRate(ftHandle, CUInt(baudRate))
                    End If
                Else
                    pFT_SetBaudRate = 0
                End If
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
            If LibraryLoaded Then
                If (pFT_SetDataCharacteristics <> 0) Then
                    If IsDeviceOpen Then
                        Dim tFT_SetDataCharacteristics As SetDataCharacteristicsDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetDataCharacteristics, IntPtr), GetType(SetDataCharacteristicsDelegate)), SetDataCharacteristicsDelegate)
                        status = tFT_SetDataCharacteristics(ftHandle, dataBits, stopBits, parity)
                    End If
                Else
                    pFT_SetDataCharacteristics = 0
                End If
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
            If LibraryLoaded Then
                If (pFT_SetFlowControl <> 0) Then
                    If IsDeviceOpen Then
                        Dim tFT_SetFlowControl As SetFlowControlDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetFlowControl, IntPtr), GetType(SetFlowControlDelegate)), SetFlowControlDelegate)
                        status = tFT_SetFlowControl(ftHandle, flowControl, xon, xoff)
                    End If
                Else
                    pFT_SetFlowControl = 0
                End If
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Asserts or de-asserts the Request To Send (RTS) line.
        ''' </summary>
        ''' <param name="enable">If true, asserts RTS. If false, de-asserts RTS.</param>
        Public Sub SetRts(enable As Boolean)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded Then
                If (pFT_SetRts <> 0) AndAlso (pFT_ClrRts <> 0) Then
                    If IsDeviceOpen Then
                        If enable Then
                            Dim tFT_SetRts As SetRtsDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetRts, IntPtr), GetType(SetRtsDelegate)), SetRtsDelegate)
                            status = tFT_SetRts(ftHandle)
                        Else
                            Dim tFT_ClrRts As ClrRtsDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_ClrRts, IntPtr), GetType(ClrRtsDelegate)), ClrRtsDelegate)
                            status = tFT_ClrRts(ftHandle)
                        End If
                    End If
                Else
                    pFT_SetRts = 0
                    pFT_ClrRts = 0
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
            If LibraryLoaded Then
                If (pFT_SetDtr <> 0) AndAlso (pFT_ClrDtr <> 0) Then
                    If IsDeviceOpen Then
                        If enable Then
                            Dim tFT_SetDtr As SetDtrDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetDtr, IntPtr), GetType(SetDtrDelegate)), SetDtrDelegate)
                            status = tFT_SetDtr(ftHandle)
                        Else
                            Dim tFT_ClrDtr As ClrDtrDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_ClrDtr, IntPtr), GetType(ClrDtrDelegate)), ClrDtrDelegate)
                            status = tFT_ClrDtr(ftHandle)
                        End If
                    End If
                Else
                    pFT_SetDtr = 0
                    pFT_ClrDtr = 0
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
                If (pFT_SetTimeouts <> 0) Then
                    Dim tFT_SetTimeouts As SetTimeoutsDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_SetTimeouts, IntPtr), GetType(SetTimeoutsDelegate)), SetTimeoutsDelegate)
                    status = tFT_SetTimeouts(ftHandle, readTimeout, writeTimeout)
                Else
                    pFT_SetTimeouts = 0
                End If
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
                If (pFT_SetBreakOn <> 0) AndAlso (pFT_SetBreakOff <> 0) Then
                    If enable Then
                        Dim tFT_SetBreakOn As SetBreakOnDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_SetBreakOn), GetType(SetBreakOnDelegate)), SetBreakOnDelegate)
                        status = tFT_SetBreakOn(ftHandle)
                    Else
                        Dim tFT_SetBreakOff As SetBreakOffDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_SetBreakOff), GetType(SetBreakOffDelegate)), SetBreakOffDelegate)
                        status = tFT_SetBreakOff(ftHandle)
                    End If
                Else
                    pFT_SetBreakOn = 0
                    pFT_SetBreakOff = 0
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
                If (pFT_SetResetPipeRetryCount <> 0) Then
                    Dim tFT_SetResetPipeRetryCount As SetResetPipeRetryCountDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_SetResetPipeRetryCount), GetType(SetResetPipeRetryCountDelegate)), SetResetPipeRetryCountDelegate)
                    status = tFT_SetResetPipeRetryCount(ftHandle, CUInt(resetPipeRetryCount))
                Else
                    pFT_SetResetPipeRetryCount = 0
                End If
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
                If (pFT_GetDriverVersion <> 0) Then
                    Dim tFT_GetDriverVersion As GetDriverVersionDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_GetDriverVersion), GetType(GetDriverVersionDelegate)), GetDriverVersionDelegate)
                    status = tFT_GetDriverVersion(ftHandle, driverVersion)
                Else
                    pFT_GetDriverVersion = 0
                End If
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
                If (pFT_GetLibraryVersion <> 0) Then
                    status = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_GetLibraryVersion), GetType(GetLibraryVersionDelegate)), GetLibraryVersionDelegate)(libraryVersion)
                Else
                    pFT_GetLibraryVersion = 0
                End If
            End If
            CheckErrors(status)
            Return libraryVersion
        End Function

        ''' <summary>
        ''' Sets the USB deadman timeout value.
        ''' </summary>
        ''' <param name="DeadmanTimeout">The deadman timeout value in ms. Default is <see cref="FT_DEFAULT_DEADMAN_TIMEOUT"/> ms.</param>
        Public Sub SetDeadmanTimeout(deadmanTimeout As UInteger)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                If (pFT_SetDeadmanTimeout <> 0) Then
                    Dim tFT_SetDeadmanTimeout As SetDeadmanTimeoutDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_SetDeadmanTimeout), GetType(SetDeadmanTimeoutDelegate)), SetDeadmanTimeoutDelegate)
                    status = tFT_SetDeadmanTimeout(ftHandle, deadmanTimeout)
                Else
                    pFT_SetDeadmanTimeout = 0
                End If
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Sets the value of the latency timer.
        ''' </summary>
        ''' <param name="Latency">The latency timer value in ms. Default value is <see cref="FT_DEFAULT_LATENCY "/> ms.
        ''' Valid values are 2...255ms for FT232BM, FT245BM and FT2232 devices.
        ''' Valid values are 0...255ms for other devices.</param>
        Public Sub SetLatency(latency As Byte)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                If (pFT_SetLatencyTimer <> 0) Then
                    Dim ft_DEVICE As FT_DEVICE = GetDeviceType()
                    If (ft_DEVICE = FT_DEVICE.FT_DEVICE_BM OrElse ft_DEVICE = FT_DEVICE.FT_DEVICE_2232) AndAlso (latency < 2) Then
                        latency = 2
                    End If
                    Dim tFT_SetLatencyTimer As SetLatencyTimerDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_SetLatencyTimer), GetType(SetLatencyTimerDelegate)), SetLatencyTimerDelegate)
                    status = tFT_SetLatencyTimer(ftHandle, latency)
                End If
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
                If (pFT_GetLatencyTimer <> 0) Then
                    Dim tFT_GetLatencyTimer As GetLatencyTimerDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_GetLatencyTimer), GetType(GetLatencyTimerDelegate)), GetLatencyTimerDelegate)
                    status = tFT_GetLatencyTimer(ftHandle, latency)
                End If
            End If
            CheckErrors(status)
            Return latency
        End Function

        ''' <summary>
        ''' Sets the USB IN and OUT transfer sizes.
        ''' </summary>
        ''' <param name="ts">The USB IN transfer size in bytes. Default value <see cref="FT_DEFAULT_IN_TRANSFER_SIZE"/> bytes.</param>
        Public Sub InTransferSize(ts As Integer)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                If (pFT_SetUSBParameters <> 0) Then
                    Dim deleg As SetUsbParametersDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_SetUSBParameters), GetType(SetUsbParametersDelegate)), SetUsbParametersDelegate)
                    status = deleg.Invoke(ftHandle, CUInt(ts), CUInt(ts))
                End If
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
                If (pFT_SetChars <> 0) Then
                    Dim tFT_SetChars As SetCharsDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_SetChars), GetType(SetCharsDelegate)), SetCharsDelegate)
                    status = tFT_SetChars(ftHandle, eventChar, Convert.ToByte(eventCharEnable), errorChar, Convert.ToByte(errorCharEnable))
                Else
                    pFT_SetChars = 0
                End If
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
                If (pFT_Read <> 0) Then
                    If (dataBuffer.Length < numBytesToRead) Then
                        numBytesToRead = dataBuffer.Length
                    End If
                    Dim numBytesWereRead As UInteger
                    Dim tFT_Read As ReadDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_Read), GetType(ReadDelegate)), ReadDelegate)
                    status = tFT_Read(ftHandle, dataBuffer, CUInt(numBytesToRead), numBytesWereRead)
                Else
                    pFT_Read = 0
                End If
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
                If (pFT_Write <> 0) Then
                    Dim tFT_Write As WriteDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_Write), GetType(WriteDelegate)), WriteDelegate)
                    status = tFT_Write(ftHandle, dataBuffer, CUInt(numBytesToWrite), numBytesWritten)
                Else
                    pFT_Write = 0
                End If
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
                If (pFT_ResetDevice <> 0) Then
                    Dim tFT_ResetDevice As ResetDeviceDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_ResetDevice), GetType(ResetDeviceDelegate)), ResetDeviceDelegate)
                    status = tFT_ResetDevice(ftHandle)
                Else
                    pFT_ResetDevice = 0
                End If
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Purge buffer constant definitions
        ''' </summary>
        Public Sub Purge(purgemask As UInteger)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                If (pFT_Purge <> 0) Then
                    Dim tFT_Purge As PurgeDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_Purge), GetType(PurgeDelegate)), PurgeDelegate)
                    status = tFT_Purge(ftHandle, purgemask)
                Else
                    pFT_Purge = 0
                End If
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
                If (pFT_SetEventNotification <> 0) Then
                    Dim deleg As SetEventNotificationDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_SetEventNotification), GetType(SetEventNotificationDelegate)), SetEventNotificationDelegate)
                    status = deleg(ftHandle, eventmask, eventhandle.SafeWaitHandle)
                Else
                    pFT_SetEventNotification = 0
                End If
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Stops the driver issuing USB in requests.
        ''' </summary>
        Public Sub StopInTask()
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                If (pFT_StopInTask <> 0) Then
                    Dim tFT_StopInTask As StopInTaskDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_StopInTask), GetType(StopInTaskDelegate)), StopInTaskDelegate)
                    status = tFT_StopInTask(ftHandle)
                Else
                    pFT_StopInTask = 0
                End If
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Resumes the driver issuing USB in requests.
        ''' </summary>
        Public Sub RestartInTask()
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                If (pFT_RestartInTask <> 0) Then
                    Dim tFT_RestartInTask As RestartInTaskDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_RestartInTask), GetType(RestartInTaskDelegate)), RestartInTaskDelegate)
                    status = tFT_RestartInTask(ftHandle)
                Else
                    pFT_RestartInTask = 0
                End If
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Resets the device port.
        ''' </summary>
        Public Sub ResetPort()
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                If (pFT_ResetPort <> 0) Then
                    Dim tFT_ResetPort As ResetPortDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_ResetPort), GetType(ResetPortDelegate)), ResetPortDelegate)
                    status = tFT_ResetPort(ftHandle)
                Else
                    pFT_ResetPort = 0
                End If
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Causes the device to be re-enumerated on the USB bus.
        ''' This is equivalent to unplugging and replugging the device.
        ''' Also calls <see cref="Close()"/> is successful, so no need to call this separately in the application.
        ''' </summary>
        Public Sub CyclePort()
            Dim ft_STATUS As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                If (pFT_CyclePort <> 0) AndAlso (pFT_Close <> 0) Then
                    Dim tFT_CyclePort As CyclePortDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_CyclePort), GetType(CyclePortDelegate)), CyclePortDelegate)
                    ft_STATUS = tFT_CyclePort(ftHandle)
                    CheckErrors(ft_STATUS)

                    Dim tFT_Close As CloseDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_Close), GetType(CloseDelegate)), CloseDelegate)
                    ft_STATUS = tFT_Close(ftHandle)
                    If (ft_STATUS = FT_STATUS.FT_OK) Then
                        ftHandle = 0
                    End If
                Else
                    pFT_CyclePort = 0
                    pFT_Close = 0
                End If
            End If
            CheckErrors(ft_STATUS)
        End Sub

        ''' <summary>
        ''' Causes the system to check for USB hardware changes. 
        ''' This is equivalent to clicking on the "Scan for hardware changes" button in the Device Manager.
        ''' </summary>
        Public Sub Rescan()
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded Then
                If (pFT_Rescan <> 0) Then
                    status = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_Rescan), GetType(RescanDelegate)), RescanDelegate)()
                Else
                    pFT_Rescan = 0
                End If
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
                If (pFT_Reload <> 0) Then
                    status = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_Reload), GetType(ReloadDelegate)), ReloadDelegate)(vendorID, productID)
                Else
                    pFT_Reload = 0
                End If
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
            Dim ft_STATUS As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                If (pFT_SetBitMode <> 0) Then
                    Dim ft_DEVICE As FT_DEVICE = GetDeviceType()
                    If (ft_DEVICE = FT_DEVICE.FT_DEVICE_AM) Then
                        CheckErrors(ft_STATUS, FT_ERROR.FT_INVALID_BITMODE)
                    ElseIf (ft_DEVICE = FT_DEVICE.FT_DEVICE_100AX) Then
                        CheckErrors(ft_STATUS, FT_ERROR.FT_INVALID_BITMODE)
                    ElseIf (ft_DEVICE = FT_DEVICE.FT_DEVICE_BM) AndAlso (bitMode <> 0) Then
                        If ((bitMode And 1) = 0) Then
                            CheckErrors(ft_STATUS, FT_ERROR.FT_INVALID_BITMODE)
                        End If
                    ElseIf (ft_DEVICE = FT_DEVICE.FT_DEVICE_2232) AndAlso (bitMode <> 0) Then
                        If ((bitMode And 31) = 0) Then
                            CheckErrors(ft_STATUS, FT_ERROR.FT_INVALID_BITMODE)
                        End If
                        If (bitMode = 2) AndAlso (Me.InterfaceIdentifier <> "A") Then
                            CheckErrors(ft_STATUS, FT_ERROR.FT_INVALID_BITMODE)
                        End If
                    ElseIf (ft_DEVICE = FT_DEVICE.FT_DEVICE_232R) AndAlso (bitMode <> 0) Then
                        If ((bitMode And 37) = 0) Then
                            CheckErrors(ft_STATUS, FT_ERROR.FT_INVALID_BITMODE)
                        End If
                    ElseIf (ft_DEVICE = FT_DEVICE.FT_DEVICE_2232H) AndAlso (bitMode <> 0) Then
                        If ((bitMode And 95) = 0) Then
                            CheckErrors(ft_STATUS, FT_ERROR.FT_INVALID_BITMODE)
                        End If
                        If (bitMode = 8 Or bitMode = 64) AndAlso (Me.InterfaceIdentifier <> "A") Then
                            CheckErrors(ft_STATUS, FT_ERROR.FT_INVALID_BITMODE)
                        End If
                    ElseIf (ft_DEVICE = FT_DEVICE.FT_DEVICE_4232H) AndAlso (bitMode <> 0) Then
                        If ((bitMode And 7) = 0) Then
                            CheckErrors(ft_STATUS, FT_ERROR.FT_INVALID_BITMODE)
                        End If
                        If (bitMode = 2) AndAlso (Me.InterfaceIdentifier <> "A" AndAlso Me.InterfaceIdentifier <> "B") Then
                            CheckErrors(ft_STATUS, FT_ERROR.FT_INVALID_BITMODE)
                        End If
                    ElseIf (ft_DEVICE = FT_DEVICE.FT_DEVICE_232H) AndAlso (bitMode <> 0) AndAlso (bitMode > 64) Then
                        CheckErrors(ft_STATUS, FT_ERROR.FT_INVALID_BITMODE)
                    End If
                    Dim tFT_SetBitMode As SetBitModeDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_SetBitMode), GetType(SetBitModeDelegate)), SetBitModeDelegate)
                    ft_STATUS = tFT_SetBitMode(ftHandle, mask, bitMode)
                Else
                    pFT_SetBitMode = 0
                End If
            End If
            CheckErrors(ft_STATUS)
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
                Else
                    pFT_VendorCmdGet = 0
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
                Else
                    pFT_VendorCmdSet = 0
                End If
            End If
            CheckErrors(status)
        End Sub

#End Region '/MANAGE

#Region "EEPROM"

        ''' <summary>
        ''' Reads an individual word value from a specified location in the device's EEPROM.
        ''' </summary>
        ''' <param name="address">The EEPROM location to read data from.</param>
        ''' <returns>The WORD value read from the EEPROM location specified in the <paramref name="address"/> paramter.</returns>
        Public Function ReadEepromLocation(address As UInteger) As Integer
            Dim eEValue As UShort
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                If (pFT_ReadEE <> 0) Then
                    Dim tFT_ReadEE As ReadEeDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_ReadEE, IntPtr), GetType(ReadEeDelegate)), ReadEeDelegate)
                    status = tFT_ReadEE(ftHandle, address, eEValue)
                Else
                    pFT_ReadEE = 0
                End If
            End If
            CheckErrors(status)
            Return eEValue
        End Function

        ''' <summary>
        ''' Writes an individual word value to a specified location in the device's EEPROM.
        ''' </summary>
        ''' <param name="address">The EEPROM location to read data from.</param>
        ''' <param name="eEValue">The WORD value to write to the EEPROM location specified by the <paramref name="address"/> parameter.</param>
        Public Sub WriteEepromLocation(address As UInteger, eEValue As UShort)
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                If (pFT_WriteEE <> 0) Then
                    Dim tFT_WriteEE As WriteEeDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_WriteEE, IntPtr), GetType(WriteEeDelegate)), WriteEeDelegate)
                    status = tFT_WriteEE(ftHandle, address, eEValue)
                Else
                    pFT_WriteEE = 0
                End If
            End If
            CheckErrors(status)
        End Sub

        ''' <summary>
        ''' Erases the device EEPROM.
        ''' </summary>
        ''' <exception cref="FTD2XX_NET.Ftdi.FT_EXCEPTION">Thrown when attempting to erase the EEPROM of a device with an internal EEPROM such as an FT232R or FT245R.</exception>
        Public Sub EraseEeprom()
            Dim ft_STATUS As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded AndAlso IsDeviceOpen Then
                If (pFT_EraseEE <> 0) Then
                    Dim ft_DEVICE As FT_DEVICE = GetDeviceType()
                    If (ft_DEVICE = FT_DEVICE.FT_DEVICE_232R) Then
                        CheckErrors(ft_STATUS, FT_ERROR.FT_INCORRECT_DEVICE)
                    End If
                    Dim tFT_EraseEE As EraseEeDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_EraseEE, IntPtr), GetType(EraseEeDelegate)), EraseEeDelegate)
                    ft_STATUS = tFT_EraseEE(ftHandle)
                Else
                    pFT_EraseEE = 0
                End If
            End If
            CheckErrors(ft_STATUS)
        End Sub

        ''' <summary>
        ''' Reads the EEPROM contents of an FT232B or FT245B device.
        ''' </summary>
        ''' <returns>An <see cref="FT232B_EEPROM_STRUCTURE"/> which contains only the relevant information for an FT232B and FT245B device.</returns>
        ''' <exception cref="FTD2XX_NET.FTDI.FT_EXCEPTION">Thrown When the current device does not match the type required by this method.</exception>
        Public Function ReadFt232bEeprom() As FT232B_EEPROM_STRUCTURE
            Dim ft_STATUS As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            Dim ee232b As New FT232B_EEPROM_STRUCTURE()
            If LibraryLoaded AndAlso IsDeviceOpen Then
                If (pFT_EE_Read <> 0) Then
                    Dim ft_DEVICE As FT_DEVICE = GetDeviceType()
                    If (ft_DEVICE <> FT_DEVICE.FT_DEVICE_BM) Then
                        CheckErrors(ft_STATUS, FT_ERROR.FT_INCORRECT_DEVICE)
                    End If

                    Dim ft_PROGRAM_DATA As New FT_PROGRAM_DATA() With {
                            .Signature1 = 0UI,
                            .Signature2 = UInteger.MaxValue,
                            .Version = 2UI,
                            .Manufacturer = Marshal.AllocHGlobal(32),
                            .ManufacturerID = Marshal.AllocHGlobal(16),
                            .Description = Marshal.AllocHGlobal(64),
                            .SerialNumber = Marshal.AllocHGlobal(16)
                        }
                    Dim tFT_EE_Read As EeReadDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_EE_Read, IntPtr), GetType(EeReadDelegate)), EeReadDelegate)
                    ft_STATUS = tFT_EE_Read(ftHandle, ft_PROGRAM_DATA)
                    ee232b.Manufacturer = Marshal.PtrToStringAnsi(ft_PROGRAM_DATA.Manufacturer)
                    ee232b.ManufacturerID = Marshal.PtrToStringAnsi(ft_PROGRAM_DATA.ManufacturerID)
                    ee232b.Description = Marshal.PtrToStringAnsi(ft_PROGRAM_DATA.Description)
                    ee232b.SerialNumber = Marshal.PtrToStringAnsi(ft_PROGRAM_DATA.SerialNumber)
                    Marshal.FreeHGlobal(ft_PROGRAM_DATA.Manufacturer)
                    Marshal.FreeHGlobal(ft_PROGRAM_DATA.ManufacturerID)
                    Marshal.FreeHGlobal(ft_PROGRAM_DATA.Description)
                    Marshal.FreeHGlobal(ft_PROGRAM_DATA.SerialNumber)
                    ee232b.VendorID = ft_PROGRAM_DATA.VendorID
                    ee232b.ProductID = ft_PROGRAM_DATA.ProductID
                    ee232b.MaxPower = ft_PROGRAM_DATA.MaxPower
                    ee232b.SelfPowered = Convert.ToBoolean(ft_PROGRAM_DATA.SelfPowered)
                    ee232b.RemoteWakeup = Convert.ToBoolean(ft_PROGRAM_DATA.RemoteWakeup)
                    ee232b.PullDownEnable = Convert.ToBoolean(ft_PROGRAM_DATA.PullDownEnable)
                    ee232b.SerNumEnable = Convert.ToBoolean(ft_PROGRAM_DATA.SerNumEnable)
                    ee232b.USBVersionEnable = Convert.ToBoolean(ft_PROGRAM_DATA.USBVersionEnable)
                    ee232b.USBVersion = ft_PROGRAM_DATA.USBVersion
                Else
                    pFT_EE_Read = 0
                End If
            End If
            CheckErrors(ft_STATUS)
            Return ee232b
        End Function

        ''' <summary>
        ''' Reads the EEPROM contents of an FT2232 device.
        ''' </summary>
        ''' <returns>An FT2232_EEPROM_STRUCTURE which contains only the relevant information for an FT2232 device.</returns>
        ''' <exception cref="FTD2XX_NET.FTDI.FT_EXCEPTION">Thrown when the current device does not match the type required by this method.</exception>
        Public Function ReadFt2232Eeprom() As FT2232_EEPROM_STRUCTURE
            Dim ft_STATUS As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            Dim ee2232 As New FT2232_EEPROM_STRUCTURE()
            If LibraryLoaded Then
                If (pFT_EE_Read <> 0) Then
                    If IsDeviceOpen Then
                        Dim ft_DEVICE As FT_DEVICE = GetDeviceType()
                        If (ft_DEVICE <> FT_DEVICE.FT_DEVICE_2232) Then
                            CheckErrors(ft_STATUS, FT_ERROR.FT_INCORRECT_DEVICE)
                        End If

                        Dim ft_PROGRAM_DATA As New FT_PROGRAM_DATA() With {
                            .Signature1 = 0UI,
                            .Signature2 = UInteger.MaxValue,
                            .Version = 2UI,
                            .Manufacturer = Marshal.AllocHGlobal(32),
                            .ManufacturerID = Marshal.AllocHGlobal(16),
                            .Description = Marshal.AllocHGlobal(64),
                            .SerialNumber = Marshal.AllocHGlobal(16)
                        }
                        Dim tFT_EE_Read As EeReadDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_EE_Read, IntPtr), GetType(EeReadDelegate)), EeReadDelegate)
                        ft_STATUS = tFT_EE_Read(ftHandle, ft_PROGRAM_DATA)
                        ee2232.Manufacturer = Marshal.PtrToStringAnsi(ft_PROGRAM_DATA.Manufacturer)
                        ee2232.ManufacturerID = Marshal.PtrToStringAnsi(ft_PROGRAM_DATA.ManufacturerID)
                        ee2232.Description = Marshal.PtrToStringAnsi(ft_PROGRAM_DATA.Description)
                        ee2232.SerialNumber = Marshal.PtrToStringAnsi(ft_PROGRAM_DATA.SerialNumber)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.Manufacturer)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.ManufacturerID)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.Description)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.SerialNumber)
                        ee2232.VendorID = ft_PROGRAM_DATA.VendorID
                        ee2232.ProductID = ft_PROGRAM_DATA.ProductID
                        ee2232.MaxPower = ft_PROGRAM_DATA.MaxPower
                        ee2232.SelfPowered = Convert.ToBoolean(ft_PROGRAM_DATA.SelfPowered)
                        ee2232.RemoteWakeup = Convert.ToBoolean(ft_PROGRAM_DATA.RemoteWakeup)
                        ee2232.PullDownEnable = Convert.ToBoolean(ft_PROGRAM_DATA.PullDownEnable5)
                        ee2232.SerNumEnable = Convert.ToBoolean(ft_PROGRAM_DATA.SerNumEnable5)
                        ee2232.USBVersionEnable = Convert.ToBoolean(ft_PROGRAM_DATA.USBVersionEnable5)
                        ee2232.USBVersion = ft_PROGRAM_DATA.USBVersion5
                        ee2232.AIsHighCurrent = Convert.ToBoolean(ft_PROGRAM_DATA.AIsHighCurrent)
                        ee2232.BIsHighCurrent = Convert.ToBoolean(ft_PROGRAM_DATA.BIsHighCurrent)
                        ee2232.IFAIsFifo = Convert.ToBoolean(ft_PROGRAM_DATA.IFAIsFifo)
                        ee2232.IFAIsFifoTar = Convert.ToBoolean(ft_PROGRAM_DATA.IFAIsFifoTar)
                        ee2232.IFAIsFastSer = Convert.ToBoolean(ft_PROGRAM_DATA.IFAIsFastSer)
                        ee2232.AIsVCP = Convert.ToBoolean(ft_PROGRAM_DATA.AIsVCP)
                        ee2232.IFBIsFifo = Convert.ToBoolean(ft_PROGRAM_DATA.IFBIsFifo)
                        ee2232.IFBIsFifoTar = Convert.ToBoolean(ft_PROGRAM_DATA.IFBIsFifoTar)
                        ee2232.IFBIsFastSer = Convert.ToBoolean(ft_PROGRAM_DATA.IFBIsFastSer)
                        ee2232.BIsVCP = Convert.ToBoolean(ft_PROGRAM_DATA.BIsVCP)
                    End If
                Else
                    pFT_EE_Read = 0
                End If
            End If
            CheckErrors(ft_STATUS)
            Return ee2232
        End Function

        ''' <summary>
        ''' Reads the EEPROM contents of an FT232R or FT245R device.
        ''' </summary>
        ''' <returns>An <see cref="FT232R_EEPROM_STRUCTURE"/> which contains only the relevant information for an FT232R and FT245R device.</returns>
        ''' <exception cref="FTD2XX_NET.FTDI.FT_EXCEPTION">Thrown when the current device does not match the type required by this method.</exception>
        Public Function ReadFt232rEeprom() As FT232R_EEPROM_STRUCTURE
            Dim ft_STATUS As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            Dim ee232r As New FT232R_EEPROM_STRUCTURE()
            If LibraryLoaded Then
                If (pFT_EE_Read <> 0) Then
                    If IsDeviceOpen Then
                        Dim ft_DEVICE As FT_DEVICE = GetDeviceType()
                        If (ft_DEVICE <> FT_DEVICE.FT_DEVICE_232R) Then
                            CheckErrors(ft_STATUS, FT_ERROR.FT_INCORRECT_DEVICE)
                        End If

                        Dim ft_PROGRAM_DATA As New FT_PROGRAM_DATA() With {
                            .Signature1 = 0UI,
                            .Signature2 = UInteger.MaxValue,
                            .Version = 2UI,
                            .Manufacturer = Marshal.AllocHGlobal(32),
                            .ManufacturerID = Marshal.AllocHGlobal(16),
                            .Description = Marshal.AllocHGlobal(64),
                            .SerialNumber = Marshal.AllocHGlobal(16)
                        }
                        Dim tFT_EE_Read As EeReadDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_EE_Read, IntPtr), GetType(EeReadDelegate)), EeReadDelegate)
                        ft_STATUS = tFT_EE_Read(ftHandle, ft_PROGRAM_DATA)
                        ee232r.Manufacturer = Marshal.PtrToStringAnsi(ft_PROGRAM_DATA.Manufacturer)
                        ee232r.ManufacturerID = Marshal.PtrToStringAnsi(ft_PROGRAM_DATA.ManufacturerID)
                        ee232r.Description = Marshal.PtrToStringAnsi(ft_PROGRAM_DATA.Description)
                        ee232r.SerialNumber = Marshal.PtrToStringAnsi(ft_PROGRAM_DATA.SerialNumber)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.Manufacturer)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.ManufacturerID)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.Description)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.SerialNumber)
                        ee232r.VendorID = ft_PROGRAM_DATA.VendorID
                        ee232r.ProductID = ft_PROGRAM_DATA.ProductID
                        ee232r.MaxPower = ft_PROGRAM_DATA.MaxPower
                        ee232r.SelfPowered = Convert.ToBoolean(ft_PROGRAM_DATA.SelfPowered)
                        ee232r.RemoteWakeup = Convert.ToBoolean(ft_PROGRAM_DATA.RemoteWakeup)
                        ee232r.UseExtOsc = Convert.ToBoolean(ft_PROGRAM_DATA.UseExtOsc)
                        ee232r.HighDriveIOs = Convert.ToBoolean(ft_PROGRAM_DATA.HighDriveIOs)
                        ee232r.EndpointSize = ft_PROGRAM_DATA.EndpointSize
                        ee232r.PullDownEnable = Convert.ToBoolean(ft_PROGRAM_DATA.PullDownEnableR)
                        ee232r.SerNumEnable = Convert.ToBoolean(ft_PROGRAM_DATA.SerNumEnableR)
                        ee232r.InvertTXD = Convert.ToBoolean(ft_PROGRAM_DATA.InvertTXD)
                        ee232r.InvertRXD = Convert.ToBoolean(ft_PROGRAM_DATA.InvertRXD)
                        ee232r.InvertRTS = Convert.ToBoolean(ft_PROGRAM_DATA.InvertRTS)
                        ee232r.InvertCTS = Convert.ToBoolean(ft_PROGRAM_DATA.InvertCTS)
                        ee232r.InvertDTR = Convert.ToBoolean(ft_PROGRAM_DATA.InvertDTR)
                        ee232r.InvertDSR = Convert.ToBoolean(ft_PROGRAM_DATA.InvertDSR)
                        ee232r.InvertDCD = Convert.ToBoolean(ft_PROGRAM_DATA.InvertDCD)
                        ee232r.InvertRI = Convert.ToBoolean(ft_PROGRAM_DATA.InvertRI)
                        ee232r.Cbus0 = ft_PROGRAM_DATA.Cbus0
                        ee232r.Cbus1 = ft_PROGRAM_DATA.Cbus1
                        ee232r.Cbus2 = ft_PROGRAM_DATA.Cbus2
                        ee232r.Cbus3 = ft_PROGRAM_DATA.Cbus3
                        ee232r.Cbus4 = ft_PROGRAM_DATA.Cbus4
                        ee232r.RIsD2XX = Convert.ToBoolean(ft_PROGRAM_DATA.RIsD2XX)
                    End If
                Else
                    pFT_EE_Read = 0
                End If
            End If
            CheckErrors(ft_STATUS)
            Return ee232r
        End Function

        ''' <summary>
        ''' Reads the EEPROM contents of an FT2232H device.
        ''' </summary>
        ''' <returns>An FT2232H_EEPROM_STRUCTURE which contains only the relevant information for an FT2232H device.</returns>
        ''' <exception cref="FTD2XX_NET.FTDI.FT_EXCEPTION">Thrown when the current device does not match the type required by this method.</exception>
        Public Function ReadFt2232hEeprom() As FT2232H_EEPROM_STRUCTURE
            Dim ft_STATUS As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            Dim ee2232h As New FT2232H_EEPROM_STRUCTURE()
            If LibraryLoaded Then
                If (pFT_EE_Read <> 0) Then
                    If IsDeviceOpen Then
                        Dim ft_DEVICE As FT_DEVICE = GetDeviceType()
                        If (ft_DEVICE <> FT_DEVICE.FT_DEVICE_2232H) Then
                            CheckErrors(ft_STATUS, FT_ERROR.FT_INCORRECT_DEVICE)
                        End If

                        Dim ft_PROGRAM_DATA As New FT_PROGRAM_DATA() With {
                            .Signature1 = 0UI,
                            .Signature2 = UInteger.MaxValue,
                            .Version = 3UI,
                            .Manufacturer = Marshal.AllocHGlobal(32),
                            .ManufacturerID = Marshal.AllocHGlobal(16),
                            .Description = Marshal.AllocHGlobal(64),
                            .SerialNumber = Marshal.AllocHGlobal(16)
                        }
                        Dim tFT_EE_Read As EeReadDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_EE_Read, IntPtr), GetType(EeReadDelegate)), EeReadDelegate)
                        ft_STATUS = tFT_EE_Read(ftHandle, ft_PROGRAM_DATA)
                        ee2232h.Manufacturer = Marshal.PtrToStringAnsi(ft_PROGRAM_DATA.Manufacturer)
                        ee2232h.ManufacturerID = Marshal.PtrToStringAnsi(ft_PROGRAM_DATA.ManufacturerID)
                        ee2232h.Description = Marshal.PtrToStringAnsi(ft_PROGRAM_DATA.Description)
                        ee2232h.SerialNumber = Marshal.PtrToStringAnsi(ft_PROGRAM_DATA.SerialNumber)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.Manufacturer)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.ManufacturerID)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.Description)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.SerialNumber)
                        ee2232h.VendorID = ft_PROGRAM_DATA.VendorID
                        ee2232h.ProductID = ft_PROGRAM_DATA.ProductID
                        ee2232h.MaxPower = ft_PROGRAM_DATA.MaxPower
                        ee2232h.SelfPowered = Convert.ToBoolean(ft_PROGRAM_DATA.SelfPowered)
                        ee2232h.RemoteWakeup = Convert.ToBoolean(ft_PROGRAM_DATA.RemoteWakeup)
                        ee2232h.PullDownEnable = Convert.ToBoolean(ft_PROGRAM_DATA.PullDownEnable7)
                        ee2232h.SerNumEnable = Convert.ToBoolean(ft_PROGRAM_DATA.SerNumEnable7)
                        ee2232h.ALSlowSlew = Convert.ToBoolean(ft_PROGRAM_DATA.ALSlowSlew)
                        ee2232h.ALSchmittInput = Convert.ToBoolean(ft_PROGRAM_DATA.ALSchmittInput)
                        ee2232h.ALDriveCurrent = ft_PROGRAM_DATA.ALDriveCurrent
                        ee2232h.AHSlowSlew = Convert.ToBoolean(ft_PROGRAM_DATA.AHSlowSlew)
                        ee2232h.AHSchmittInput = Convert.ToBoolean(ft_PROGRAM_DATA.AHSchmittInput)
                        ee2232h.AHDriveCurrent = ft_PROGRAM_DATA.AHDriveCurrent
                        ee2232h.BLSlowSlew = Convert.ToBoolean(ft_PROGRAM_DATA.BLSlowSlew)
                        ee2232h.BLSchmittInput = Convert.ToBoolean(ft_PROGRAM_DATA.BLSchmittInput)
                        ee2232h.BLDriveCurrent = ft_PROGRAM_DATA.BLDriveCurrent
                        ee2232h.BHSlowSlew = Convert.ToBoolean(ft_PROGRAM_DATA.BHSlowSlew)
                        ee2232h.BHSchmittInput = Convert.ToBoolean(ft_PROGRAM_DATA.BHSchmittInput)
                        ee2232h.BHDriveCurrent = ft_PROGRAM_DATA.BHDriveCurrent
                        ee2232h.IFAIsFifo = Convert.ToBoolean(ft_PROGRAM_DATA.IFAIsFifo7)
                        ee2232h.IFAIsFifoTar = Convert.ToBoolean(ft_PROGRAM_DATA.IFAIsFifoTar7)
                        ee2232h.IFAIsFastSer = Convert.ToBoolean(ft_PROGRAM_DATA.IFAIsFastSer7)
                        ee2232h.AIsVCP = Convert.ToBoolean(ft_PROGRAM_DATA.AIsVCP7)
                        ee2232h.IFBIsFifo = Convert.ToBoolean(ft_PROGRAM_DATA.IFBIsFifo7)
                        ee2232h.IFBIsFifoTar = Convert.ToBoolean(ft_PROGRAM_DATA.IFBIsFifoTar7)
                        ee2232h.IFBIsFastSer = Convert.ToBoolean(ft_PROGRAM_DATA.IFBIsFastSer7)
                        ee2232h.BIsVCP = Convert.ToBoolean(ft_PROGRAM_DATA.BIsVCP7)
                        ee2232h.PowerSaveEnable = Convert.ToBoolean(ft_PROGRAM_DATA.PowerSaveEnable)
                    End If
                Else
                    pFT_EE_Read = 0
                End If
            End If
            CheckErrors(ft_STATUS)
            Return ee2232h
        End Function

        ''' <summary>
        ''' Reads the EEPROM contents of an FT4232H device.
        ''' </summary>
        ''' <returns>An FT4232H_EEPROM_STRUCTURE which contains only the relevant information for an FT4232H device.</returns>
        ''' <exception cref="FTD2XX_NET.FTDI.FT_EXCEPTION">Thrown when the current device does not match the type required by this method.</exception>
        Public Function ReadFt4232hEeprom() As FT4232H_EEPROM_STRUCTURE
            Dim ft_STATUS As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            Dim ee4232h As New FT4232H_EEPROM_STRUCTURE()
            If LibraryLoaded Then
                If (pFT_EE_Read <> 0) Then
                    If IsDeviceOpen Then
                        Dim ft_DEVICE As FT_DEVICE = GetDeviceType()
                        If (ft_DEVICE <> FT_DEVICE.FT_DEVICE_4232H) Then
                            CheckErrors(ft_STATUS, FT_ERROR.FT_INCORRECT_DEVICE)
                        End If

                        Dim ft_PROGRAM_DATA As New FT_PROGRAM_DATA() With {
                            .Signature1 = 0UI,
                            .Signature2 = UInteger.MaxValue,
                            .Version = 4UI,
                            .Manufacturer = Marshal.AllocHGlobal(32),
                            .ManufacturerID = Marshal.AllocHGlobal(16),
                            .Description = Marshal.AllocHGlobal(64),
                            .SerialNumber = Marshal.AllocHGlobal(16)
                        }
                        Dim tFT_EE_Read As EeReadDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_EE_Read, IntPtr), GetType(EeReadDelegate)), EeReadDelegate)
                        ft_STATUS = tFT_EE_Read(ftHandle, ft_PROGRAM_DATA)
                        ee4232h.Manufacturer = Marshal.PtrToStringAnsi(ft_PROGRAM_DATA.Manufacturer)
                        ee4232h.ManufacturerID = Marshal.PtrToStringAnsi(ft_PROGRAM_DATA.ManufacturerID)
                        ee4232h.Description = Marshal.PtrToStringAnsi(ft_PROGRAM_DATA.Description)
                        ee4232h.SerialNumber = Marshal.PtrToStringAnsi(ft_PROGRAM_DATA.SerialNumber)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.Manufacturer)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.ManufacturerID)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.Description)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.SerialNumber)
                        ee4232h.VendorID = ft_PROGRAM_DATA.VendorID
                        ee4232h.ProductID = ft_PROGRAM_DATA.ProductID
                        ee4232h.MaxPower = ft_PROGRAM_DATA.MaxPower
                        ee4232h.SelfPowered = Convert.ToBoolean(ft_PROGRAM_DATA.SelfPowered)
                        ee4232h.RemoteWakeup = Convert.ToBoolean(ft_PROGRAM_DATA.RemoteWakeup)
                        ee4232h.PullDownEnable = Convert.ToBoolean(ft_PROGRAM_DATA.PullDownEnable8)
                        ee4232h.SerNumEnable = Convert.ToBoolean(ft_PROGRAM_DATA.SerNumEnable8)
                        ee4232h.ASlowSlew = Convert.ToBoolean(ft_PROGRAM_DATA.ASlowSlew)
                        ee4232h.ASchmittInput = Convert.ToBoolean(ft_PROGRAM_DATA.ASchmittInput)
                        ee4232h.ADriveCurrent = ft_PROGRAM_DATA.ADriveCurrent
                        ee4232h.BSlowSlew = Convert.ToBoolean(ft_PROGRAM_DATA.BSlowSlew)
                        ee4232h.BSchmittInput = Convert.ToBoolean(ft_PROGRAM_DATA.BSchmittInput)
                        ee4232h.BDriveCurrent = ft_PROGRAM_DATA.BDriveCurrent
                        ee4232h.CSlowSlew = Convert.ToBoolean(ft_PROGRAM_DATA.CSlowSlew)
                        ee4232h.CSchmittInput = Convert.ToBoolean(ft_PROGRAM_DATA.CSchmittInput)
                        ee4232h.CDriveCurrent = ft_PROGRAM_DATA.CDriveCurrent
                        ee4232h.DSlowSlew = Convert.ToBoolean(ft_PROGRAM_DATA.DSlowSlew)
                        ee4232h.DSchmittInput = Convert.ToBoolean(ft_PROGRAM_DATA.DSchmittInput)
                        ee4232h.DDriveCurrent = ft_PROGRAM_DATA.DDriveCurrent
                        ee4232h.ARIIsTXDEN = Convert.ToBoolean(ft_PROGRAM_DATA.ARIIsTXDEN)
                        ee4232h.BRIIsTXDEN = Convert.ToBoolean(ft_PROGRAM_DATA.BRIIsTXDEN)
                        ee4232h.CRIIsTXDEN = Convert.ToBoolean(ft_PROGRAM_DATA.CRIIsTXDEN)
                        ee4232h.DRIIsTXDEN = Convert.ToBoolean(ft_PROGRAM_DATA.DRIIsTXDEN)
                        ee4232h.AIsVCP = Convert.ToBoolean(ft_PROGRAM_DATA.AIsVCP8)
                        ee4232h.BIsVCP = Convert.ToBoolean(ft_PROGRAM_DATA.BIsVCP8)
                        ee4232h.CIsVCP = Convert.ToBoolean(ft_PROGRAM_DATA.CIsVCP8)
                        ee4232h.DIsVCP = Convert.ToBoolean(ft_PROGRAM_DATA.DIsVCP8)
                    End If
                Else
                    pFT_EE_Read = 0
                End If
            End If
            CheckErrors(ft_STATUS)
            Return ee4232h
        End Function

        ''' <summary>
        ''' Reads the EEPROM contents of an FT232H device.
        ''' </summary>
        ''' <returns>An FT232H_EEPROM_STRUCTURE which contains only the relevant information for an FT232H device.</returns>
        ''' <exception cref="FTD2XX_NET.FTDI.FT_EXCEPTION">Thrown when the current device does not match the type required by this method.</exception>
        Public Function ReadFt232hEeprom() As FT232H_EEPROM_STRUCTURE
            Dim ft_STATUS As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            Dim ee232h As New FT232H_EEPROM_STRUCTURE()
            If LibraryLoaded Then
                If (pFT_EE_Read <> 0) Then
                    If IsDeviceOpen Then
                        Dim ft_DEVICE As FT_DEVICE = GetDeviceType()
                        If (ft_DEVICE <> FT_DEVICE.FT_DEVICE_232H) Then
                            CheckErrors(ft_STATUS, FT_ERROR.FT_INCORRECT_DEVICE)
                        End If

                        Dim ft_PROGRAM_DATA As New FT_PROGRAM_DATA() With {
                            .Signature1 = 0UI,
                            .Signature2 = UInteger.MaxValue,
                            .Version = 5UI,
                            .Manufacturer = Marshal.AllocHGlobal(32),
                            .ManufacturerID = Marshal.AllocHGlobal(16),
                            .Description = Marshal.AllocHGlobal(64),
                            .SerialNumber = Marshal.AllocHGlobal(16)
                        }
                        Dim tFT_EE_Read As EeReadDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_EE_Read), GetType(EeReadDelegate)), EeReadDelegate)
                        ft_STATUS = tFT_EE_Read(ftHandle, ft_PROGRAM_DATA)
                        ee232h.Manufacturer = Marshal.PtrToStringAnsi(ft_PROGRAM_DATA.Manufacturer)
                        ee232h.ManufacturerID = Marshal.PtrToStringAnsi(ft_PROGRAM_DATA.ManufacturerID)
                        ee232h.Description = Marshal.PtrToStringAnsi(ft_PROGRAM_DATA.Description)
                        ee232h.SerialNumber = Marshal.PtrToStringAnsi(ft_PROGRAM_DATA.SerialNumber)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.Manufacturer)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.ManufacturerID)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.Description)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.SerialNumber)
                        ee232h.VendorID = ft_PROGRAM_DATA.VendorID
                        ee232h.ProductID = ft_PROGRAM_DATA.ProductID
                        ee232h.MaxPower = ft_PROGRAM_DATA.MaxPower
                        ee232h.SelfPowered = Convert.ToBoolean(ft_PROGRAM_DATA.SelfPowered)
                        ee232h.RemoteWakeup = Convert.ToBoolean(ft_PROGRAM_DATA.RemoteWakeup)
                        ee232h.PullDownEnable = Convert.ToBoolean(ft_PROGRAM_DATA.PullDownEnableH)
                        ee232h.SerNumEnable = Convert.ToBoolean(ft_PROGRAM_DATA.SerNumEnableH)
                        ee232h.ACSlowSlew = Convert.ToBoolean(ft_PROGRAM_DATA.ACSlowSlewH)
                        ee232h.ACSchmittInput = Convert.ToBoolean(ft_PROGRAM_DATA.ACSchmittInputH)
                        ee232h.ACDriveCurrent = ft_PROGRAM_DATA.ACDriveCurrentH
                        ee232h.ADSlowSlew = Convert.ToBoolean(ft_PROGRAM_DATA.ADSlowSlewH)
                        ee232h.ADSchmittInput = Convert.ToBoolean(ft_PROGRAM_DATA.ADSchmittInputH)
                        ee232h.ADDriveCurrent = ft_PROGRAM_DATA.ADDriveCurrentH
                        ee232h.Cbus0 = ft_PROGRAM_DATA.Cbus0H
                        ee232h.Cbus1 = ft_PROGRAM_DATA.Cbus1H
                        ee232h.Cbus2 = ft_PROGRAM_DATA.Cbus2H
                        ee232h.Cbus3 = ft_PROGRAM_DATA.Cbus3H
                        ee232h.Cbus4 = ft_PROGRAM_DATA.Cbus4H
                        ee232h.Cbus5 = ft_PROGRAM_DATA.Cbus5H
                        ee232h.Cbus6 = ft_PROGRAM_DATA.Cbus6H
                        ee232h.Cbus7 = ft_PROGRAM_DATA.Cbus7H
                        ee232h.Cbus8 = ft_PROGRAM_DATA.Cbus8H
                        ee232h.Cbus9 = ft_PROGRAM_DATA.Cbus9H
                        ee232h.IsFifo = Convert.ToBoolean(ft_PROGRAM_DATA.IsFifoH)
                        ee232h.IsFifoTar = Convert.ToBoolean(ft_PROGRAM_DATA.IsFifoTarH)
                        ee232h.IsFastSer = Convert.ToBoolean(ft_PROGRAM_DATA.IsFastSerH)
                        ee232h.IsFT1248 = Convert.ToBoolean(ft_PROGRAM_DATA.IsFT1248H)
                        ee232h.FT1248Cpol = Convert.ToBoolean(ft_PROGRAM_DATA.FT1248CpolH)
                        ee232h.FT1248Lsb = Convert.ToBoolean(ft_PROGRAM_DATA.FT1248LsbH)
                        ee232h.FT1248FlowControl = Convert.ToBoolean(ft_PROGRAM_DATA.FT1248FlowControlH)
                        ee232h.IsVCP = Convert.ToBoolean(ft_PROGRAM_DATA.IsVCPH)
                        ee232h.PowerSaveEnable = Convert.ToBoolean(ft_PROGRAM_DATA.PowerSaveEnableH)
                    End If
                Else
                    pFT_EE_Read = 0
                End If
            End If
            CheckErrors(ft_STATUS)
            Return ee232h
        End Function

        ''' <summary>
        ''' Reads the EEPROM contents of an X-Series device.
        ''' </summary>
        ''' <returns>An FT_XSERIES_EEPROM_STRUCTURE which contains only the relevant information for an X-Series device.</returns>
        ''' <exception cref="FTD2XX_NET.FTDI.FT_EXCEPTION">Thrown when the current device does not match the type required by this method.</exception>
        Public Function ReadXSeriesEeprom() As FT_XSERIES_EEPROM_STRUCTURE
            Dim ft_STATUS As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            Dim eeX As New FT_XSERIES_EEPROM_STRUCTURE()
            If LibraryLoaded Then
                If (pFT_EEPROM_Read <> 0) Then
                    If IsDeviceOpen Then
                        Dim devType As FT_DEVICE = GetDeviceType()
                        If (devType <> FT_DEVICE.FT_DEVICE_X_SERIES) Then
                            CheckErrors(ft_STATUS, FT_ERROR.FT_INCORRECT_DEVICE)
                        End If
                        Dim header As New FT_EEPROM_HEADER() With {
                            .deviceType = 9UI
                        }
                        Dim ft_XSERIES_DATA As New FT_XSERIES_DATA() With {
                            .common = header
                        }
                        Dim strSize As Integer = Marshal.SizeOf(ft_XSERIES_DATA)
                        Dim intPtr As IntPtr = Marshal.AllocHGlobal(strSize)
                        Marshal.StructureToPtr(ft_XSERIES_DATA, intPtr, False)
                        Dim man As Byte() = New Byte(31) {}
                        Dim id As Byte() = New Byte(15) {}
                        Dim descr As Byte() = New Byte(63) {}
                        Dim sn As Byte() = New Byte(15) {}
                        Dim tFT_EEPROM_Read As EepromReadDelegate = CType(Marshal.GetDelegateForFunctionPointer(New IntPtr(pFT_EEPROM_Read), GetType(EepromReadDelegate)), EepromReadDelegate)
                        ft_STATUS = tFT_EEPROM_Read(ftHandle, intPtr, CUInt(strSize), man, id, descr, sn)
                        If (ft_STATUS = FT_STATUS.FT_OK) Then
                            ft_XSERIES_DATA = CType(Marshal.PtrToStructure(intPtr, GetType(FT_XSERIES_DATA)), FT_XSERIES_DATA)
                            Dim utf8Encoding As New UTF8Encoding()
                            eeX.Manufacturer = utf8Encoding.GetString(man)
                            eeX.ManufacturerID = utf8Encoding.GetString(id)
                            eeX.Description = utf8Encoding.GetString(descr)
                            eeX.SerialNumber = utf8Encoding.GetString(sn)
                            eeX.VendorID = ft_XSERIES_DATA.common.VendorId
                            eeX.ProductID = ft_XSERIES_DATA.common.ProductId
                            eeX.MaxPower = ft_XSERIES_DATA.common.MaxPower
                            eeX.SelfPowered = Convert.ToBoolean(ft_XSERIES_DATA.common.SelfPowered)
                            eeX.RemoteWakeup = Convert.ToBoolean(ft_XSERIES_DATA.common.RemoteWakeup)
                            eeX.SerNumEnable = Convert.ToBoolean(ft_XSERIES_DATA.common.SerNumEnable)
                            eeX.PullDownEnable = Convert.ToBoolean(ft_XSERIES_DATA.common.PullDownEnable)
                            eeX.Cbus0 = ft_XSERIES_DATA.Cbus0
                            eeX.Cbus1 = ft_XSERIES_DATA.Cbus1
                            eeX.Cbus2 = ft_XSERIES_DATA.Cbus2
                            eeX.Cbus3 = ft_XSERIES_DATA.Cbus3
                            eeX.Cbus4 = ft_XSERIES_DATA.Cbus4
                            eeX.Cbus5 = ft_XSERIES_DATA.Cbus5
                            eeX.Cbus6 = ft_XSERIES_DATA.Cbus6
                            eeX.ACDriveCurrent = ft_XSERIES_DATA.ACDriveCurrent
                            eeX.ACSchmittInput = ft_XSERIES_DATA.ACSchmittInput
                            eeX.ACSlowSlew = ft_XSERIES_DATA.ACSlowSlew
                            eeX.ADDriveCurrent = ft_XSERIES_DATA.ADDriveCurrent
                            eeX.ADSchmittInput = ft_XSERIES_DATA.ADSchmittInput
                            eeX.ADSlowSlew = ft_XSERIES_DATA.ADSlowSlew
                            eeX.BCDDisableSleep = ft_XSERIES_DATA.BCDDisableSleep
                            eeX.BCDEnable = ft_XSERIES_DATA.BCDEnable
                            eeX.BCDForceCbusPWREN = ft_XSERIES_DATA.BCDForceCbusPWREN
                            eeX.FT1248Cpol = ft_XSERIES_DATA.FT1248Cpol
                            eeX.FT1248FlowControl = ft_XSERIES_DATA.FT1248FlowControl
                            eeX.FT1248Lsb = ft_XSERIES_DATA.FT1248Lsb
                            eeX.I2CDeviceId = ft_XSERIES_DATA.I2CDeviceId
                            eeX.I2CDisableSchmitt = ft_XSERIES_DATA.I2CDisableSchmitt
                            eeX.I2CSlaveAddress = ft_XSERIES_DATA.I2CSlaveAddress
                            eeX.InvertCTS = ft_XSERIES_DATA.InvertCTS
                            eeX.InvertDCD = ft_XSERIES_DATA.InvertDCD
                            eeX.InvertDSR = ft_XSERIES_DATA.InvertDSR
                            eeX.InvertDTR = ft_XSERIES_DATA.InvertDTR
                            eeX.InvertRI = ft_XSERIES_DATA.InvertRI
                            eeX.InvertRTS = ft_XSERIES_DATA.InvertRTS
                            eeX.InvertRXD = ft_XSERIES_DATA.InvertRXD
                            eeX.InvertTXD = ft_XSERIES_DATA.InvertTXD
                            eeX.PowerSaveEnable = ft_XSERIES_DATA.PowerSaveEnable
                            eeX.RS485EchoSuppress = ft_XSERIES_DATA.RS485EchoSuppress
                            eeX.IsVCP = ft_XSERIES_DATA.DriverType
                            Marshal.DestroyStructure(intPtr, GetType(FT_XSERIES_DATA))
                            Marshal.FreeHGlobal(intPtr)
                        End If
                    End If
                Else
                    pFT_EE_Read = 0
                End If
            End If
            CheckErrors(ft_STATUS)
            Return eeX
        End Function

        ''' <summary>
        ''' Writes the specified values to the EEPROM of an FT232B or FT245B device.
        ''' </summary>
        ''' <param name="ee232b">The EEPROM settings to be written to the device</param>
        ''' <remarks>If the strings are too long, they will be truncated to their maximum permitted lengths</remarks>
        ''' <exception cref="FTD2XX_NET.FTDI.FT_EXCEPTION">Thrown when the current device does not match the type required by this method.</exception>
        Public Sub WriteFt232bEeprom(ee232b As FT232B_EEPROM_STRUCTURE)
            Dim ft_STATUS As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded Then
                If (pFT_EE_Program <> 0) Then
                    If IsDeviceOpen Then
                        Dim ft_DEVICE As FT_DEVICE = GetDeviceType()
                        If (ft_DEVICE <> FT_DEVICE.FT_DEVICE_BM) Then
                            CheckErrors(ft_STATUS, FT_ERROR.FT_INCORRECT_DEVICE)
                        End If
                        If (ee232b.VendorID = 0US) OrElse (ee232b.ProductID = 0US) Then
                            CheckErrors(FT_STATUS.FT_INVALID_PARAMETER)
                        End If

                        Dim ft_PROGRAM_DATA As New FT_PROGRAM_DATA() With {
                            .Signature1 = 0UI,
                            .Signature2 = UInteger.MaxValue,
                            .Version = 2UI,
                            .Manufacturer = Marshal.AllocHGlobal(32),
                            .ManufacturerID = Marshal.AllocHGlobal(16),
                            .Description = Marshal.AllocHGlobal(64),
                            .SerialNumber = Marshal.AllocHGlobal(16)
                        }
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
                        ft_PROGRAM_DATA.Manufacturer = Marshal.StringToHGlobalAnsi(ee232b.Manufacturer)
                        ft_PROGRAM_DATA.ManufacturerID = Marshal.StringToHGlobalAnsi(ee232b.ManufacturerID)
                        ft_PROGRAM_DATA.Description = Marshal.StringToHGlobalAnsi(ee232b.Description)
                        ft_PROGRAM_DATA.SerialNumber = Marshal.StringToHGlobalAnsi(ee232b.SerialNumber)
                        ft_PROGRAM_DATA.VendorID = ee232b.VendorID
                        ft_PROGRAM_DATA.ProductID = ee232b.ProductID
                        ft_PROGRAM_DATA.MaxPower = ee232b.MaxPower
                        ft_PROGRAM_DATA.SelfPowered = Convert.ToUInt16(ee232b.SelfPowered)
                        ft_PROGRAM_DATA.RemoteWakeup = Convert.ToUInt16(ee232b.RemoteWakeup)
                        ft_PROGRAM_DATA.Rev4 = Convert.ToByte(True)
                        ft_PROGRAM_DATA.PullDownEnable = Convert.ToByte(ee232b.PullDownEnable)
                        ft_PROGRAM_DATA.SerNumEnable = Convert.ToByte(ee232b.SerNumEnable)
                        ft_PROGRAM_DATA.USBVersionEnable = Convert.ToByte(ee232b.USBVersionEnable)
                        ft_PROGRAM_DATA.USBVersion = ee232b.USBVersion
                        Dim tFT_EE_Program As EeProgramDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_EE_Program, IntPtr), GetType(EeProgramDelegate)), EeProgramDelegate)
                        ft_STATUS = ft_STATUS Or tFT_EE_Program(ftHandle, ft_PROGRAM_DATA)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.Manufacturer)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.ManufacturerID)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.Description)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.SerialNumber)
                    End If
                Else
                    pFT_EE_Program = 0
                End If
            End If
            CheckErrors(ft_STATUS)
        End Sub

        ''' <summary>
        ''' Writes the specified values to the EEPROM of an FT2232 device.
        ''' </summary>
        ''' <param name="ee2232">The EEPROM settings to be written to the device</param>
        ''' <remarks>If the strings are too long, they will be truncated to their maximum permitted lengths</remarks>
        ''' <exception cref="FTD2XX_NET.FTDI.FT_EXCEPTION">Thrown when the current device does not match the type required by this method.</exception>
        Public Sub WriteFt2232Eeprom(ee2232 As FT2232_EEPROM_STRUCTURE)
            Dim ft_STATUS As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded Then
                If (pFT_EE_Program <> 0) Then
                    If IsDeviceOpen Then
                        Dim ft_DEVICE As FT_DEVICE = GetDeviceType()
                        If (ft_DEVICE <> FT_DEVICE.FT_DEVICE_2232) Then
                            CheckErrors(ft_STATUS, FT_ERROR.FT_INCORRECT_DEVICE)
                        End If
                        If (ee2232.VendorID = 0US) OrElse (ee2232.ProductID = 0US) Then
                            CheckErrors(FT_STATUS.FT_INVALID_PARAMETER)
                        End If

                        Dim ft_PROGRAM_DATA As New FT_PROGRAM_DATA() With {
                            .Signature1 = 0UI,
                            .Signature2 = UInteger.MaxValue,
                            .Version = 2UI,
                            .Manufacturer = Marshal.AllocHGlobal(32),
                            .ManufacturerID = Marshal.AllocHGlobal(16),
                            .Description = Marshal.AllocHGlobal(64),
                            .SerialNumber = Marshal.AllocHGlobal(16)
                        }
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
                        ft_PROGRAM_DATA.Manufacturer = Marshal.StringToHGlobalAnsi(ee2232.Manufacturer)
                        ft_PROGRAM_DATA.ManufacturerID = Marshal.StringToHGlobalAnsi(ee2232.ManufacturerID)
                        ft_PROGRAM_DATA.Description = Marshal.StringToHGlobalAnsi(ee2232.Description)
                        ft_PROGRAM_DATA.SerialNumber = Marshal.StringToHGlobalAnsi(ee2232.SerialNumber)
                        ft_PROGRAM_DATA.VendorID = ee2232.VendorID
                        ft_PROGRAM_DATA.ProductID = ee2232.ProductID
                        ft_PROGRAM_DATA.MaxPower = ee2232.MaxPower
                        ft_PROGRAM_DATA.SelfPowered = Convert.ToUInt16(ee2232.SelfPowered)
                        ft_PROGRAM_DATA.RemoteWakeup = Convert.ToUInt16(ee2232.RemoteWakeup)
                        ft_PROGRAM_DATA.Rev5 = Convert.ToByte(True)
                        ft_PROGRAM_DATA.PullDownEnable5 = Convert.ToByte(ee2232.PullDownEnable)
                        ft_PROGRAM_DATA.SerNumEnable5 = Convert.ToByte(ee2232.SerNumEnable)
                        ft_PROGRAM_DATA.USBVersionEnable5 = Convert.ToByte(ee2232.USBVersionEnable)
                        ft_PROGRAM_DATA.USBVersion5 = ee2232.USBVersion
                        ft_PROGRAM_DATA.AIsHighCurrent = Convert.ToByte(ee2232.AIsHighCurrent)
                        ft_PROGRAM_DATA.BIsHighCurrent = Convert.ToByte(ee2232.BIsHighCurrent)
                        ft_PROGRAM_DATA.IFAIsFifo = Convert.ToByte(ee2232.IFAIsFifo)
                        ft_PROGRAM_DATA.IFAIsFifoTar = Convert.ToByte(ee2232.IFAIsFifoTar)
                        ft_PROGRAM_DATA.IFAIsFastSer = Convert.ToByte(ee2232.IFAIsFastSer)
                        ft_PROGRAM_DATA.AIsVCP = Convert.ToByte(ee2232.AIsVCP)
                        ft_PROGRAM_DATA.IFBIsFifo = Convert.ToByte(ee2232.IFBIsFifo)
                        ft_PROGRAM_DATA.IFBIsFifoTar = Convert.ToByte(ee2232.IFBIsFifoTar)
                        ft_PROGRAM_DATA.IFBIsFastSer = Convert.ToByte(ee2232.IFBIsFastSer)
                        ft_PROGRAM_DATA.BIsVCP = Convert.ToByte(ee2232.BIsVCP)
                        Dim tFT_EE_Program As EeProgramDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_EE_Program, IntPtr), GetType(EeProgramDelegate)), EeProgramDelegate)
                        ft_STATUS = ft_STATUS Or tFT_EE_Program(ftHandle, ft_PROGRAM_DATA)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.Manufacturer)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.ManufacturerID)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.Description)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.SerialNumber)
                    End If
                Else
                    pFT_EE_Program = 0
                End If
            End If
            CheckErrors(ft_STATUS)
        End Sub

        ''' <summary>
        ''' Writes the specified values to the EEPROM of an FT232R or FT245R device.
        ''' </summary>
        ''' <param name="ee232r">The EEPROM settings to be written to the device</param>
        ''' <remarks>If the strings are too long, they will be truncated to their maximum permitted lengths</remarks>
        ''' <exception cref="FTD2XX_NET.FTDI.FT_EXCEPTION">Thrown when the current device does not match the type required by this method.</exception>
        Public Sub WriteFt232rEeprom(ee232r As FT232R_EEPROM_STRUCTURE)
            Dim ft_STATUS As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded Then
                If (pFT_EE_Program <> 0) Then
                    If IsDeviceOpen Then
                        Dim ft_DEVICE As FT_DEVICE = GetDeviceType()
                        If (ft_DEVICE <> FT_DEVICE.FT_DEVICE_232R) Then
                            CheckErrors(ft_STATUS, FT_ERROR.FT_INCORRECT_DEVICE)
                        End If
                        If (ee232r.VendorID = 0US) OrElse (ee232r.ProductID = 0US) Then
                            CheckErrors(FT_STATUS.FT_INVALID_PARAMETER)
                        End If

                        Dim ft_PROGRAM_DATA As New FT_PROGRAM_DATA() With {
                            .Signature1 = 0UI,
                            .Signature2 = UInteger.MaxValue,
                            .Version = 2UI,
                            .Manufacturer = Marshal.AllocHGlobal(32),
                            .ManufacturerID = Marshal.AllocHGlobal(16),
                            .Description = Marshal.AllocHGlobal(64),
                            .SerialNumber = Marshal.AllocHGlobal(16)
                        }
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
                        ft_PROGRAM_DATA.Manufacturer = Marshal.StringToHGlobalAnsi(ee232r.Manufacturer)
                        ft_PROGRAM_DATA.ManufacturerID = Marshal.StringToHGlobalAnsi(ee232r.ManufacturerID)
                        ft_PROGRAM_DATA.Description = Marshal.StringToHGlobalAnsi(ee232r.Description)
                        ft_PROGRAM_DATA.SerialNumber = Marshal.StringToHGlobalAnsi(ee232r.SerialNumber)
                        ft_PROGRAM_DATA.VendorID = ee232r.VendorID
                        ft_PROGRAM_DATA.ProductID = ee232r.ProductID
                        ft_PROGRAM_DATA.MaxPower = ee232r.MaxPower
                        ft_PROGRAM_DATA.SelfPowered = Convert.ToUInt16(ee232r.SelfPowered)
                        ft_PROGRAM_DATA.RemoteWakeup = Convert.ToUInt16(ee232r.RemoteWakeup)
                        ft_PROGRAM_DATA.PullDownEnableR = Convert.ToByte(ee232r.PullDownEnable)
                        ft_PROGRAM_DATA.SerNumEnableR = Convert.ToByte(ee232r.SerNumEnable)
                        ft_PROGRAM_DATA.UseExtOsc = Convert.ToByte(ee232r.UseExtOsc)
                        ft_PROGRAM_DATA.HighDriveIOs = Convert.ToByte(ee232r.HighDriveIOs)
                        ft_PROGRAM_DATA.EndpointSize = 64
                        ft_PROGRAM_DATA.PullDownEnableR = Convert.ToByte(ee232r.PullDownEnable)
                        ft_PROGRAM_DATA.SerNumEnableR = Convert.ToByte(ee232r.SerNumEnable)
                        ft_PROGRAM_DATA.InvertTXD = Convert.ToByte(ee232r.InvertTXD)
                        ft_PROGRAM_DATA.InvertRXD = Convert.ToByte(ee232r.InvertRXD)
                        ft_PROGRAM_DATA.InvertRTS = Convert.ToByte(ee232r.InvertRTS)
                        ft_PROGRAM_DATA.InvertCTS = Convert.ToByte(ee232r.InvertCTS)
                        ft_PROGRAM_DATA.InvertDTR = Convert.ToByte(ee232r.InvertDTR)
                        ft_PROGRAM_DATA.InvertDSR = Convert.ToByte(ee232r.InvertDSR)
                        ft_PROGRAM_DATA.InvertDCD = Convert.ToByte(ee232r.InvertDCD)
                        ft_PROGRAM_DATA.InvertRI = Convert.ToByte(ee232r.InvertRI)
                        ft_PROGRAM_DATA.Cbus0 = ee232r.Cbus0
                        ft_PROGRAM_DATA.Cbus1 = ee232r.Cbus1
                        ft_PROGRAM_DATA.Cbus2 = ee232r.Cbus2
                        ft_PROGRAM_DATA.Cbus3 = ee232r.Cbus3
                        ft_PROGRAM_DATA.Cbus4 = ee232r.Cbus4
                        ft_PROGRAM_DATA.RIsD2XX = Convert.ToByte(ee232r.RIsD2XX)
                        Dim tFT_EE_Program As EeProgramDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_EE_Program, IntPtr), GetType(EeProgramDelegate)), EeProgramDelegate)
                        ft_STATUS = ft_STATUS Or tFT_EE_Program(ftHandle, ft_PROGRAM_DATA)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.Manufacturer)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.ManufacturerID)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.Description)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.SerialNumber)
                    End If
                Else
                    pFT_EE_Program = 0
                End If
            End If
            CheckErrors(ft_STATUS)
        End Sub

        ''' <summary>
        ''' Writes the specified values to the EEPROM of an FT2232H device.
        ''' </summary>
        ''' <param name="ee2232h">The EEPROM settings to be written to the device</param>
        ''' <remarks>If the strings are too long, they will be truncated to their maximum permitted lengths</remarks>
        ''' <exception cref="FTD2XX_NET.FTDI.FT_EXCEPTION">Thrown when the current device does not match the type required by this method.</exception>
        Public Sub WriteFt2232hEeprom(ee2232h As FT2232H_EEPROM_STRUCTURE)
            Dim ft_STATUS As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded Then
                If (pFT_EE_Program <> 0) Then
                    If IsDeviceOpen Then
                        Dim ft_DEVICE As FT_DEVICE = GetDeviceType()
                        If (ft_DEVICE <> FT_DEVICE.FT_DEVICE_2232H) Then
                            CheckErrors(ft_STATUS, FT_ERROR.FT_INCORRECT_DEVICE)
                        End If
                        If (ee2232h.VendorID = 0US) OrElse (ee2232h.ProductID = 0US) Then
                            CheckErrors(FT_STATUS.FT_INVALID_PARAMETER)
                        End If

                        Dim ft_PROGRAM_DATA As New FT_PROGRAM_DATA() With {
                            .Signature1 = 0UI,
                            .Signature2 = UInteger.MaxValue,
                            .Version = 3UI,
                            .Manufacturer = Marshal.AllocHGlobal(32),
                            .ManufacturerID = Marshal.AllocHGlobal(16),
                            .Description = Marshal.AllocHGlobal(64),
                            .SerialNumber = Marshal.AllocHGlobal(16)
                        }
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
                        ft_PROGRAM_DATA.Manufacturer = Marshal.StringToHGlobalAnsi(ee2232h.Manufacturer)
                        ft_PROGRAM_DATA.ManufacturerID = Marshal.StringToHGlobalAnsi(ee2232h.ManufacturerID)
                        ft_PROGRAM_DATA.Description = Marshal.StringToHGlobalAnsi(ee2232h.Description)
                        ft_PROGRAM_DATA.SerialNumber = Marshal.StringToHGlobalAnsi(ee2232h.SerialNumber)
                        ft_PROGRAM_DATA.VendorID = ee2232h.VendorID
                        ft_PROGRAM_DATA.ProductID = ee2232h.ProductID
                        ft_PROGRAM_DATA.MaxPower = ee2232h.MaxPower
                        ft_PROGRAM_DATA.SelfPowered = Convert.ToUInt16(ee2232h.SelfPowered)
                        ft_PROGRAM_DATA.RemoteWakeup = Convert.ToUInt16(ee2232h.RemoteWakeup)
                        ft_PROGRAM_DATA.PullDownEnable7 = Convert.ToByte(ee2232h.PullDownEnable)
                        ft_PROGRAM_DATA.SerNumEnable7 = Convert.ToByte(ee2232h.SerNumEnable)
                        ft_PROGRAM_DATA.ALSlowSlew = Convert.ToByte(ee2232h.ALSlowSlew)
                        ft_PROGRAM_DATA.ALSchmittInput = Convert.ToByte(ee2232h.ALSchmittInput)
                        ft_PROGRAM_DATA.ALDriveCurrent = ee2232h.ALDriveCurrent
                        ft_PROGRAM_DATA.AHSlowSlew = Convert.ToByte(ee2232h.AHSlowSlew)
                        ft_PROGRAM_DATA.AHSchmittInput = Convert.ToByte(ee2232h.AHSchmittInput)
                        ft_PROGRAM_DATA.AHDriveCurrent = ee2232h.AHDriveCurrent
                        ft_PROGRAM_DATA.BLSlowSlew = Convert.ToByte(ee2232h.BLSlowSlew)
                        ft_PROGRAM_DATA.BLSchmittInput = Convert.ToByte(ee2232h.BLSchmittInput)
                        ft_PROGRAM_DATA.BLDriveCurrent = ee2232h.BLDriveCurrent
                        ft_PROGRAM_DATA.BHSlowSlew = Convert.ToByte(ee2232h.BHSlowSlew)
                        ft_PROGRAM_DATA.BHSchmittInput = Convert.ToByte(ee2232h.BHSchmittInput)
                        ft_PROGRAM_DATA.BHDriveCurrent = ee2232h.BHDriveCurrent
                        ft_PROGRAM_DATA.IFAIsFifo7 = Convert.ToByte(ee2232h.IFAIsFifo)
                        ft_PROGRAM_DATA.IFAIsFifoTar7 = Convert.ToByte(ee2232h.IFAIsFifoTar)
                        ft_PROGRAM_DATA.IFAIsFastSer7 = Convert.ToByte(ee2232h.IFAIsFastSer)
                        ft_PROGRAM_DATA.AIsVCP7 = Convert.ToByte(ee2232h.AIsVCP)
                        ft_PROGRAM_DATA.IFBIsFifo7 = Convert.ToByte(ee2232h.IFBIsFifo)
                        ft_PROGRAM_DATA.IFBIsFifoTar7 = Convert.ToByte(ee2232h.IFBIsFifoTar)
                        ft_PROGRAM_DATA.IFBIsFastSer7 = Convert.ToByte(ee2232h.IFBIsFastSer)
                        ft_PROGRAM_DATA.BIsVCP7 = Convert.ToByte(ee2232h.BIsVCP)
                        ft_PROGRAM_DATA.PowerSaveEnable = Convert.ToByte(ee2232h.PowerSaveEnable)
                        Dim tFT_EE_Program As EeProgramDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_EE_Program, IntPtr), GetType(EeProgramDelegate)), EeProgramDelegate)
                        ft_STATUS = ft_STATUS Or tFT_EE_Program(ftHandle, ft_PROGRAM_DATA)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.Manufacturer)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.ManufacturerID)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.Description)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.SerialNumber)
                    End If
                Else
                    pFT_EE_Program = 0
                End If
            End If
            CheckErrors(ft_STATUS)
        End Sub

        ''' <summary>
        ''' Writes the specified values to the EEPROM of an FT4232H device.
        ''' </summary>
        ''' <param name="ee4232h">The EEPROM settings to be written to the device</param>
        ''' <remarks>If the strings are too long, they will be truncated to their maximum permitted lengths</remarks>
        ''' <exception cref="FTD2XX_NET.FTDI.FT_EXCEPTION">Thrown when the current device does not match the type required by this method.</exception>
        Public Sub WriteFt4232hEeprom(ee4232h As FT4232H_EEPROM_STRUCTURE)
            Dim ft_STATUS As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded Then
                If (pFT_EE_Program <> 0) Then
                    If IsDeviceOpen Then
                        Dim ft_DEVICE As FT_DEVICE = GetDeviceType()
                        If (ft_DEVICE <> FT_DEVICE.FT_DEVICE_4232H) Then
                            CheckErrors(ft_STATUS, FT_ERROR.FT_INCORRECT_DEVICE)
                        End If
                        If (ee4232h.VendorID = 0US) OrElse (ee4232h.ProductID = 0US) Then
                            CheckErrors(FT_STATUS.FT_INVALID_PARAMETER)
                        End If

                        Dim ft_PROGRAM_DATA As New FT_PROGRAM_DATA() With {
                            .Signature1 = 0UI,
                            .Signature2 = UInteger.MaxValue,
                            .Version = 4UI,
                            .Manufacturer = Marshal.AllocHGlobal(32),
                            .ManufacturerID = Marshal.AllocHGlobal(16),
                            .Description = Marshal.AllocHGlobal(64),
                            .SerialNumber = Marshal.AllocHGlobal(16)
                        }
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
                        ft_PROGRAM_DATA.Manufacturer = Marshal.StringToHGlobalAnsi(ee4232h.Manufacturer)
                        ft_PROGRAM_DATA.ManufacturerID = Marshal.StringToHGlobalAnsi(ee4232h.ManufacturerID)
                        ft_PROGRAM_DATA.Description = Marshal.StringToHGlobalAnsi(ee4232h.Description)
                        ft_PROGRAM_DATA.SerialNumber = Marshal.StringToHGlobalAnsi(ee4232h.SerialNumber)
                        ft_PROGRAM_DATA.VendorID = ee4232h.VendorID
                        ft_PROGRAM_DATA.ProductID = ee4232h.ProductID
                        ft_PROGRAM_DATA.MaxPower = ee4232h.MaxPower
                        ft_PROGRAM_DATA.SelfPowered = Convert.ToUInt16(ee4232h.SelfPowered)
                        ft_PROGRAM_DATA.RemoteWakeup = Convert.ToUInt16(ee4232h.RemoteWakeup)
                        ft_PROGRAM_DATA.PullDownEnable8 = Convert.ToByte(ee4232h.PullDownEnable)
                        ft_PROGRAM_DATA.SerNumEnable8 = Convert.ToByte(ee4232h.SerNumEnable)
                        ft_PROGRAM_DATA.ASlowSlew = Convert.ToByte(ee4232h.ASlowSlew)
                        ft_PROGRAM_DATA.ASchmittInput = Convert.ToByte(ee4232h.ASchmittInput)
                        ft_PROGRAM_DATA.ADriveCurrent = ee4232h.ADriveCurrent
                        ft_PROGRAM_DATA.BSlowSlew = Convert.ToByte(ee4232h.BSlowSlew)
                        ft_PROGRAM_DATA.BSchmittInput = Convert.ToByte(ee4232h.BSchmittInput)
                        ft_PROGRAM_DATA.BDriveCurrent = ee4232h.BDriveCurrent
                        ft_PROGRAM_DATA.CSlowSlew = Convert.ToByte(ee4232h.CSlowSlew)
                        ft_PROGRAM_DATA.CSchmittInput = Convert.ToByte(ee4232h.CSchmittInput)
                        ft_PROGRAM_DATA.CDriveCurrent = ee4232h.CDriveCurrent
                        ft_PROGRAM_DATA.DSlowSlew = Convert.ToByte(ee4232h.DSlowSlew)
                        ft_PROGRAM_DATA.DSchmittInput = Convert.ToByte(ee4232h.DSchmittInput)
                        ft_PROGRAM_DATA.DDriveCurrent = ee4232h.DDriveCurrent
                        ft_PROGRAM_DATA.ARIIsTXDEN = Convert.ToByte(ee4232h.ARIIsTXDEN)
                        ft_PROGRAM_DATA.BRIIsTXDEN = Convert.ToByte(ee4232h.BRIIsTXDEN)
                        ft_PROGRAM_DATA.CRIIsTXDEN = Convert.ToByte(ee4232h.CRIIsTXDEN)
                        ft_PROGRAM_DATA.DRIIsTXDEN = Convert.ToByte(ee4232h.DRIIsTXDEN)
                        ft_PROGRAM_DATA.AIsVCP8 = Convert.ToByte(ee4232h.AIsVCP)
                        ft_PROGRAM_DATA.BIsVCP8 = Convert.ToByte(ee4232h.BIsVCP)
                        ft_PROGRAM_DATA.CIsVCP8 = Convert.ToByte(ee4232h.CIsVCP)
                        ft_PROGRAM_DATA.DIsVCP8 = Convert.ToByte(ee4232h.DIsVCP)
                        Dim tFT_EE_Program As EeProgramDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_EE_Program, IntPtr), GetType(EeProgramDelegate)), EeProgramDelegate)
                        ft_STATUS = ft_STATUS Or tFT_EE_Program(ftHandle, ft_PROGRAM_DATA)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.Manufacturer)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.ManufacturerID)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.Description)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.SerialNumber)
                    End If
                Else
                    pFT_EE_Program = 0
                End If
            End If
            CheckErrors(ft_STATUS)
        End Sub

        ''' <summary>
        ''' Writes the specified values to the EEPROM of an FT232H device.
        ''' </summary>
        ''' <param name="ee232h">The EEPROM settings to be written to the device</param>
        ''' <remarks>If the strings are too long, they will be truncated to their maximum permitted lengths</remarks>
        ''' <exception cref="FTD2XX_NET.FTDI.FT_EXCEPTION">Thrown when the current device does not match the type required by this method.</exception>
        Public Sub WriteFt232hEeprom(ee232h As FT232H_EEPROM_STRUCTURE)
            Dim ft_STATUS As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded Then
                If (pFT_EE_Program <> 0) Then
                    If IsDeviceOpen Then
                        Dim ft_DEVICE As FT_DEVICE = GetDeviceType()
                        If (ft_DEVICE <> FT_DEVICE.FT_DEVICE_232H) Then
                            CheckErrors(ft_STATUS, FT_ERROR.FT_INCORRECT_DEVICE)
                        End If
                        If (ee232h.VendorID = 0US) OrElse (ee232h.ProductID = 0US) Then
                            CheckErrors(FT_STATUS.FT_INVALID_PARAMETER)
                        End If
                        Dim ft_PROGRAM_DATA As New FT_PROGRAM_DATA() With {
                            .Signature1 = 0UI,
                            .Signature2 = UInteger.MaxValue,
                            .Version = 5UI,
                            .Manufacturer = Marshal.AllocHGlobal(32),
                            .ManufacturerID = Marshal.AllocHGlobal(16),
                            .Description = Marshal.AllocHGlobal(64),
                            .SerialNumber = Marshal.AllocHGlobal(16)
                        }
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
                        ft_PROGRAM_DATA.Manufacturer = Marshal.StringToHGlobalAnsi(ee232h.Manufacturer)
                        ft_PROGRAM_DATA.ManufacturerID = Marshal.StringToHGlobalAnsi(ee232h.ManufacturerID)
                        ft_PROGRAM_DATA.Description = Marshal.StringToHGlobalAnsi(ee232h.Description)
                        ft_PROGRAM_DATA.SerialNumber = Marshal.StringToHGlobalAnsi(ee232h.SerialNumber)
                        ft_PROGRAM_DATA.VendorID = ee232h.VendorID
                        ft_PROGRAM_DATA.ProductID = ee232h.ProductID
                        ft_PROGRAM_DATA.MaxPower = ee232h.MaxPower
                        ft_PROGRAM_DATA.SelfPowered = Convert.ToUInt16(ee232h.SelfPowered)
                        ft_PROGRAM_DATA.RemoteWakeup = Convert.ToUInt16(ee232h.RemoteWakeup)
                        ft_PROGRAM_DATA.PullDownEnableH = Convert.ToByte(ee232h.PullDownEnable)
                        ft_PROGRAM_DATA.SerNumEnableH = Convert.ToByte(ee232h.SerNumEnable)
                        ft_PROGRAM_DATA.ACSlowSlewH = Convert.ToByte(ee232h.ACSlowSlew)
                        ft_PROGRAM_DATA.ACSchmittInputH = Convert.ToByte(ee232h.ACSchmittInput)
                        ft_PROGRAM_DATA.ACDriveCurrentH = Convert.ToByte(ee232h.ACDriveCurrent)
                        ft_PROGRAM_DATA.ADSlowSlewH = Convert.ToByte(ee232h.ADSlowSlew)
                        ft_PROGRAM_DATA.ADSchmittInputH = Convert.ToByte(ee232h.ADSchmittInput)
                        ft_PROGRAM_DATA.ADDriveCurrentH = Convert.ToByte(ee232h.ADDriveCurrent)
                        ft_PROGRAM_DATA.Cbus0H = Convert.ToByte(ee232h.Cbus0)
                        ft_PROGRAM_DATA.Cbus1H = Convert.ToByte(ee232h.Cbus1)
                        ft_PROGRAM_DATA.Cbus2H = Convert.ToByte(ee232h.Cbus2)
                        ft_PROGRAM_DATA.Cbus3H = Convert.ToByte(ee232h.Cbus3)
                        ft_PROGRAM_DATA.Cbus4H = Convert.ToByte(ee232h.Cbus4)
                        ft_PROGRAM_DATA.Cbus5H = Convert.ToByte(ee232h.Cbus5)
                        ft_PROGRAM_DATA.Cbus6H = Convert.ToByte(ee232h.Cbus6)
                        ft_PROGRAM_DATA.Cbus7H = Convert.ToByte(ee232h.Cbus7)
                        ft_PROGRAM_DATA.Cbus8H = Convert.ToByte(ee232h.Cbus8)
                        ft_PROGRAM_DATA.Cbus9H = Convert.ToByte(ee232h.Cbus9)
                        ft_PROGRAM_DATA.IsFifoH = Convert.ToByte(ee232h.IsFifo)
                        ft_PROGRAM_DATA.IsFifoTarH = Convert.ToByte(ee232h.IsFifoTar)
                        ft_PROGRAM_DATA.IsFastSerH = Convert.ToByte(ee232h.IsFastSer)
                        ft_PROGRAM_DATA.IsFT1248H = Convert.ToByte(ee232h.IsFT1248)
                        ft_PROGRAM_DATA.FT1248CpolH = Convert.ToByte(ee232h.FT1248Cpol)
                        ft_PROGRAM_DATA.FT1248LsbH = Convert.ToByte(ee232h.FT1248Lsb)
                        ft_PROGRAM_DATA.FT1248FlowControlH = Convert.ToByte(ee232h.FT1248FlowControl)
                        ft_PROGRAM_DATA.IsVCPH = Convert.ToByte(ee232h.IsVCP)
                        ft_PROGRAM_DATA.PowerSaveEnableH = Convert.ToByte(ee232h.PowerSaveEnable)
                        Dim tFT_EE_Program As EeProgramDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_EE_Program, IntPtr), GetType(EeProgramDelegate)), EeProgramDelegate)
                        ft_STATUS = ft_STATUS Or tFT_EE_Program(ftHandle, ft_PROGRAM_DATA)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.Manufacturer)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.ManufacturerID)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.Description)
                        Marshal.FreeHGlobal(ft_PROGRAM_DATA.SerialNumber)
                    End If
                Else
                    pFT_EE_Program = 0
                End If
            End If
            CheckErrors(ft_STATUS)
        End Sub

        ''' <summary>
        ''' Writes the specified values to the EEPROM of an X-Series device.
        ''' </summary>
        ''' <remarks>If the strings are too long, they will be truncated to their maximum permitted lengths</remarks>
        ''' <exception cref="FTD2XX_NET.FTDI.FT_EXCEPTION">Thrown when the current device does not match the type required by this method.</exception>
        Public Sub WriteXSeriesEeprom(eeX As FT_XSERIES_EEPROM_STRUCTURE)
            Dim ft_STATUS As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded Then
                If (pFT_EEPROM_Program <> 0) Then
                    If IsDeviceOpen Then
                        Dim ft_DEVICE As FT_DEVICE = GetDeviceType()
                        If (ft_DEVICE <> FT_DEVICE.FT_DEVICE_X_SERIES) Then
                            CheckErrors(ft_STATUS, FT_ERROR.FT_INCORRECT_DEVICE)
                        End If
                        If (eeX.VendorID = 0US) OrElse (eeX.ProductID = 0US) Then
                            CheckErrors(FT_STATUS.FT_INVALID_PARAMETER)
                        End If
                        Dim ft_XSERIES_DATA As FT_XSERIES_DATA = Nothing
                        Dim unused As Byte() = New Byte(15) {}
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
                        Dim utf8Encoding As UTF8Encoding = New UTF8Encoding()
                        Dim manufacturer As Byte() = utf8Encoding.GetBytes(eeX.Manufacturer) 'len 32
                        Dim manufacturerID As Byte() = utf8Encoding.GetBytes(eeX.ManufacturerID) '16
                        Dim description As Byte() = utf8Encoding.GetBytes(eeX.Description) '64
                        Dim serialnumber As Byte() = utf8Encoding.GetBytes(eeX.SerialNumber) '16
                        ft_XSERIES_DATA.common.deviceType = 9UI
                        ft_XSERIES_DATA.common.VendorId = eeX.VendorID
                        ft_XSERIES_DATA.common.ProductId = eeX.ProductID
                        ft_XSERIES_DATA.common.MaxPower = eeX.MaxPower
                        ft_XSERIES_DATA.common.SelfPowered = Convert.ToByte(eeX.SelfPowered)
                        ft_XSERIES_DATA.common.RemoteWakeup = Convert.ToByte(eeX.RemoteWakeup)
                        ft_XSERIES_DATA.common.SerNumEnable = Convert.ToByte(eeX.SerNumEnable)
                        ft_XSERIES_DATA.common.PullDownEnable = Convert.ToByte(eeX.PullDownEnable)
                        ft_XSERIES_DATA.Cbus0 = eeX.Cbus0
                        ft_XSERIES_DATA.Cbus1 = eeX.Cbus1
                        ft_XSERIES_DATA.Cbus2 = eeX.Cbus2
                        ft_XSERIES_DATA.Cbus3 = eeX.Cbus3
                        ft_XSERIES_DATA.Cbus4 = eeX.Cbus4
                        ft_XSERIES_DATA.Cbus5 = eeX.Cbus5
                        ft_XSERIES_DATA.Cbus6 = eeX.Cbus6
                        ft_XSERIES_DATA.ACDriveCurrent = eeX.ACDriveCurrent
                        ft_XSERIES_DATA.ACSchmittInput = eeX.ACSchmittInput
                        ft_XSERIES_DATA.ACSlowSlew = eeX.ACSlowSlew
                        ft_XSERIES_DATA.ADDriveCurrent = eeX.ADDriveCurrent
                        ft_XSERIES_DATA.ADSchmittInput = eeX.ADSchmittInput
                        ft_XSERIES_DATA.ADSlowSlew = eeX.ADSlowSlew
                        ft_XSERIES_DATA.BCDDisableSleep = eeX.BCDDisableSleep
                        ft_XSERIES_DATA.BCDEnable = eeX.BCDEnable
                        ft_XSERIES_DATA.BCDForceCbusPWREN = eeX.BCDForceCbusPWREN
                        ft_XSERIES_DATA.FT1248Cpol = eeX.FT1248Cpol
                        ft_XSERIES_DATA.FT1248FlowControl = eeX.FT1248FlowControl
                        ft_XSERIES_DATA.FT1248Lsb = eeX.FT1248Lsb
                        ft_XSERIES_DATA.I2CDeviceId = eeX.I2CDeviceId
                        ft_XSERIES_DATA.I2CDisableSchmitt = eeX.I2CDisableSchmitt
                        ft_XSERIES_DATA.I2CSlaveAddress = eeX.I2CSlaveAddress
                        ft_XSERIES_DATA.InvertCTS = eeX.InvertCTS
                        ft_XSERIES_DATA.InvertDCD = eeX.InvertDCD
                        ft_XSERIES_DATA.InvertDSR = eeX.InvertDSR
                        ft_XSERIES_DATA.InvertDTR = eeX.InvertDTR
                        ft_XSERIES_DATA.InvertRI = eeX.InvertRI
                        ft_XSERIES_DATA.InvertRTS = eeX.InvertRTS
                        ft_XSERIES_DATA.InvertRXD = eeX.InvertRXD
                        ft_XSERIES_DATA.InvertTXD = eeX.InvertTXD
                        ft_XSERIES_DATA.PowerSaveEnable = eeX.PowerSaveEnable
                        ft_XSERIES_DATA.RS485EchoSuppress = eeX.RS485EchoSuppress
                        ft_XSERIES_DATA.DriverType = eeX.IsVCP
                        Dim num As Integer = Marshal.SizeOf(ft_XSERIES_DATA)
                        Dim intPtr As IntPtr = Marshal.AllocHGlobal(num)
                        Marshal.StructureToPtr(ft_XSERIES_DATA, intPtr, False)
                        Dim tFT_EEPROM_Program As EepromProgramDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(Me.pFT_EEPROM_Program, IntPtr), GetType(EepromProgramDelegate)), EepromProgramDelegate)
                        ft_STATUS = ft_STATUS Or tFT_EEPROM_Program(ftHandle, intPtr, CUInt(num), manufacturer, manufacturerID, description, serialnumber)
                    End If
                End If
            End If
            CheckErrors(ft_STATUS)
        End Sub

        ''' <summary>
        ''' Reads data from the user area of the device EEPROM.
        ''' </summary>
        ''' <returns>An array of bytes with the data read from the device EEPROM user area.</returns>
        Public Function EeReadUserArea() As Byte()
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            Dim userAreaDataBuffer As Byte() = New Byte() {}
            If LibraryLoaded Then
                If (pFT_EE_UASize <> 0) AndAlso (pFT_EE_UARead <> 0) Then
                    If IsDeviceOpen Then
                        Dim bufSize As UInteger = 0UI
                        Dim tFT_EE_UASize As EeUaSizeDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_EE_UASize, IntPtr), GetType(EeUaSizeDelegate)), EeUaSizeDelegate)
                        status = status Or tFT_EE_UASize(ftHandle, bufSize)
                        ReDim userAreaDataBuffer(CInt(bufSize - 1))
                        If (userAreaDataBuffer.Length >= bufSize) Then
                            Dim numBytesWereRead As UInteger = 0
                            Dim tFT_EE_UARead As EeUaReadDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_EE_UARead, IntPtr), GetType(EeUaReadDelegate)), EeUaReadDelegate)
                            status = status Or tFT_EE_UARead(ftHandle, userAreaDataBuffer, userAreaDataBuffer.Length, numBytesWereRead)
                        End If
                    End If
                Else
                    pFT_EE_UASize = 0
                    pFT_EE_UARead = 0
                End If
            End If
            CheckErrors(status)
            Return userAreaDataBuffer
        End Function

        ''' <summary>
        ''' Writes data to the user area of the device EEPROM.
        ''' </summary>
        ''' <param name="userAreaDataBuffer">An array of bytes which will be written to the device EEPROM user area.</param>
        Public Sub EeWriteUserArea(userAreaDataBuffer As Byte())
            Dim status As FT_STATUS = FT_STATUS.FT_OTHER_ERROR
            If LibraryLoaded Then
                If (pFT_EE_UASize <> 0) AndAlso (pFT_EE_UAWrite <> 0) Then
                    If IsDeviceOpen Then
                        Dim bufSize As UInteger = 0UI
                        Dim tFT_EE_UASize As EeUaSizeDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_EE_UASize, IntPtr), GetType(EeUaSizeDelegate)), EeUaSizeDelegate)
                        status = status Or tFT_EE_UASize(ftHandle, bufSize)
                        If (userAreaDataBuffer.Length <= bufSize) Then
                            Dim tFT_EE_UAWrite As EeUaWriteDelegate = CType(Marshal.GetDelegateForFunctionPointer(CType(pFT_EE_UAWrite, IntPtr), GetType(EeUaWriteDelegate)), EeUaWriteDelegate)
                            status = status Or tFT_EE_UAWrite(ftHandle, userAreaDataBuffer, userAreaDataBuffer.Length)
                        End If
                    End If
                Else
                    pFT_EE_UASize = 0
                    pFT_EE_UAWrite = 0
                End If
            End If
            CheckErrors(status)
        End Sub

#End Region '/EEPROM

#Region "HELPERS"

        ''' <summary>
        ''' Проверяет результат выполнения метода.
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
                    'Case Else
                    '    Return
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
        Private Delegate Function ReadEeDelegate(ftHandle As Integer, dwWordOffset As UInteger, ByRef lpwValue As UShort) As FT_STATUS

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
        Private Delegate Function EeReadDelegate(ftHandle As Integer, pData As FT_PROGRAM_DATA) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function EeProgramDelegate(ftHandle As Integer, pData As FT_PROGRAM_DATA) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function EepromReadDelegate(ftHandle As Integer, eepromData As IntPtr, eepromDataSize As UInteger, manufacturer As Byte(), manufacturerID As Byte(), description As Byte(), serialnumber As Byte()) As FT_STATUS

        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Function EepromProgramDelegate(ftHandle As Integer, eepromData As IntPtr, eepromDataSize As UInteger, manufacturer As Byte(), manufacturerID As Byte(), description As Byte(), serialnumber As Byte()) As FT_STATUS

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
            FT_DLL_NOT_LOADED
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
            Public ALDriveCurrent As Byte
            Public AHSlowSlew As Byte
            Public AHSchmittInput As Byte
            Public AHDriveCurrent As Byte
            Public BLSlowSlew As Byte
            Public BLSchmittInput As Byte
            Public BLDriveCurrent As Byte
            Public BHSlowSlew As Byte
            Public BHSchmittInput As Byte
            Public BHDriveCurrent As Byte
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
            Public ADriveCurrent As Byte
            Public BSlowSlew As Byte
            Public BSchmittInput As Byte
            Public BDriveCurrent As Byte
            Public CSlowSlew As Byte
            Public CSchmittInput As Byte
            Public CDriveCurrent As Byte
            Public DSlowSlew As Byte
            Public DSchmittInput As Byte
            Public DDriveCurrent As Byte
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
            Public ACDriveCurrentH As Byte
            Public ADSlowSlewH As Byte
            Public ADSchmittInputH As Byte
            Public ADDriveCurrentH As Byte
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
            Public ACDriveCurrent As Byte
            Public ADSlowSlew As Byte
            Public ADSchmittInput As Byte
            Public ADDriveCurrent As Byte
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
            Public VendorID As UShort = 1027US
            Public ProductID As UShort = 24577US
            Public Manufacturer As String = "FTDI"
            Public ManufacturerID As String = "FT"
            Public Description As String = "USB-Serial Converter"
            Public SerialNumber As String = ""
            Public MaxPower As UShort = 144US
            Public SelfPowered As Boolean
            Public RemoteWakeup As Boolean
        End Class

        Public Class FT232B_EEPROM_STRUCTURE
            Inherits FT_EEPROM_DATA
            Public PullDownEnable As Boolean
            Public SerNumEnable As Boolean = True
            Public USBVersionEnable As Boolean = True
            Public USBVersion As UShort = 512US
        End Class

        Public Class FT2232_EEPROM_STRUCTURE
            Inherits FT_EEPROM_DATA
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

        Public Class FT232R_EEPROM_STRUCTURE
            Inherits FT_EEPROM_DATA
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

        Public Class FT2232H_EEPROM_STRUCTURE
            Inherits FT_EEPROM_DATA
            Public PullDownEnable As Boolean
            Public SerNumEnable As Boolean = True
            Public ALSlowSlew As Boolean
            Public ALSchmittInput As Boolean
            Public ALDriveCurrent As Byte = 4
            Public AHSlowSlew As Boolean
            Public AHSchmittInput As Boolean
            Public AHDriveCurrent As Byte = 4
            Public BLSlowSlew As Boolean
            Public BLSchmittInput As Boolean
            Public BLDriveCurrent As Byte = 4
            Public BHSlowSlew As Boolean
            Public BHSchmittInput As Boolean
            Public BHDriveCurrent As Byte = 4
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

        Public Class FT4232H_EEPROM_STRUCTURE
            Inherits FT_EEPROM_DATA
            Public PullDownEnable As Boolean
            Public SerNumEnable As Boolean = True
            Public ASlowSlew As Boolean
            Public ASchmittInput As Boolean
            Public ADriveCurrent As Byte = 4
            Public BSlowSlew As Boolean
            Public BSchmittInput As Boolean
            Public BDriveCurrent As Byte = 4
            Public CSlowSlew As Boolean
            Public CSchmittInput As Boolean
            Public CDriveCurrent As Byte = 4
            Public DSlowSlew As Boolean
            Public DSchmittInput As Boolean
            Public DDriveCurrent As Byte = 4
            Public ARIIsTXDEN As Boolean
            Public BRIIsTXDEN As Boolean
            Public CRIIsTXDEN As Boolean
            Public DRIIsTXDEN As Boolean
            Public AIsVCP As Boolean = True
            Public BIsVCP As Boolean = True
            Public CIsVCP As Boolean = True
            Public DIsVCP As Boolean = True
        End Class

        Public Class FT232H_EEPROM_STRUCTURE
            Inherits FT_EEPROM_DATA
            Public PullDownEnable As Boolean
            Public SerNumEnable As Boolean = True
            Public ACSlowSlew As Boolean
            Public ACSchmittInput As Boolean
            Public ACDriveCurrent As Byte = 4
            Public ADSlowSlew As Boolean
            Public ADSchmittInput As Boolean
            Public ADDriveCurrent As Byte = 4
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

        Public Class FT_XSERIES_EEPROM_STRUCTURE
            Inherits FT_EEPROM_DATA
            Public PullDownEnable As Boolean
            Public SerNumEnable As Boolean = True
            Public USBVersionEnable As Boolean = True
            Public USBVersion As UShort = 512US
            Public ACSlowSlew As Byte
            Public ACSchmittInput As Byte
            Public ACDriveCurrent As Byte
            Public ADSlowSlew As Byte
            Public ADSchmittInput As Byte
            Public ADDriveCurrent As Byte
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

    End Class '/FTDI

End Namespace
