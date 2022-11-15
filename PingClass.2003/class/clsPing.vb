Option Explicit On 

Imports System
Imports System.Net
Imports System.Net.Sockets

'* Class PingClass
'* 
'*    Author : Paulo Santos
'*      Date : September 2002
'* Objective : Ping a host and return basic informations
'*
'* Class Properties
'*
'*   +------------+-------------+-------------------------------------------------+
'*   | Name       | Type        | Description                                     |
'*   +------------+-------------+-------------------------------------------------+
'*   | DataSize   | Integer     | Size of the data to be send to the host         |
'*   | Identifier | Integer     | Identifier of ping packet                       |
'*   | Sequence   | Integer     | Sequence of the packet                          |
'*   | LocalHost  | IPHostEntry | Info for Local Computer                         |
'*   | TimeOut    | Long        | Time (in millisecons) of the time-out value     |
'*   | Host       | Object      | Host Entry for the remote computer              |
'*   +------------+-------------+-------------------------------------------------+
'*
'* Class Methods
'*
'*   +-------------------+----------------------------------------------------+
'*   | Name              | Description                                        |
'*   +-------------------+----------------------------------------------------+
'*   | PingHost          | Pings the specified host                           |
'*   +-------------------+----------------------------------------------------+
'*
'* Parts of the code based on information from Visual Studio Magazine
'* more info : http://www.fawcette.com/vsm/2002_03/magazine/columns/qa/default.asp
'*
Public Class PingClass

    Private Const PING_ERROR As Integer = -1
    Private Const PING_ERROR_BASE As Long = &H8000

    Public Enum pingErrorCodes As Long
        Success = 0
        Unknown = (-1)
        HostNotFound = PING_ERROR_BASE + 1
        SocketDidntSend
        HostNotResponding
        TimeOut
        BufferTooSmall
        DestinationUnreahable
        HostNotReachable
        ProtocolNotReachable
        PortNotReachable
        NoResourcesAvailable
        BadOption
        HardwareError
        PacketTooBig
        ReqestedTimedOut
        BadRequest
        BadRoute
        TTLExprdInTransit
        TTLExprdReassemb
        ParameterProblem
        SourceQuench
        OptionTooBig
        BadDestination
        AddressDeleted
        SpecMTUChange
        MTUChange
        Unload
        GeneralFailure
    End Enum

    Private Const ICMP_ECHO As Integer = 8

    Private Const __Sequence As Byte = 0
    Private Const __Identifier As Byte = 0
    Private Const __BufferHeaderSize As Integer = 8

    Private __Error As pingErrorCodes
    Private __TimeOut As Long
    Private __DataSize As Byte
    Private __LocalHost As System.Net.IPHostEntry
    Private __RemoteHost As System.Net.IPHostEntry

    Public Structure PingAsyncResult
        Dim PingAsyncState As Object
        Dim PingError As pingErrorCodes
        Dim PingTime As Long
    End Structure

    Private Structure PingAsyncRequest
        Dim RemoteHost As String
        Dim PingCallback As PingCallback
        Dim PingAsyncResult As PingAsyncResult
    End Structure

    Public Delegate Sub PingCallback(ByVal par As PingAsyncResult)

    Public Property TimeOut() As Long
        Get
            Return __TimeOut
        End Get
        Set(ByVal Value As Long)
            __TimeOut = Value
        End Set
    End Property

    Public Property DataSize() As Byte
        Get
            Return __DataSize
        End Get
        Set(ByVal Value As Byte)
            __DataSize = Value
        End Set
    End Property

    Public ReadOnly Property LocalHost() As System.Net.IPHostEntry
        Get
            Return __LocalHost
        End Get
    End Property

    Public Property Host() As System.Net.IPHostEntry
        Get
            Return __RemoteHost
        End Get
        Set(ByVal Value As System.Net.IPHostEntry)
            __RemoteHost = Value
        End Set
    End Property

    Public Overloads Sub SetRemoteHost(ByVal fqdnHostName As String)

        '*
        '* Find the Host's IP address
        '*
        Try
            Host = System.Net.Dns.Resolve(fqdnHostName)
        Catch
            Host = Nothing
            __Error = pingErrorCodes.HostNotFound 'PING_ERROR_HOST_NOT_FOUND
        End Try

    End Sub

    Public Overloads Sub SetRemoteHost(ByVal ipHost As System.Net.IPAddress)
        SetRemoteHost(ipHost.ToString)
    End Sub

    '*
    '* Class Constructor
    '*
    Public Sub New()
        '*
        '* Initializes the parameters
        '*
        __DataSize = 32
        __TimeOut = 500
        __Error = pingErrorCodes.Success

        '*
        '* Get local IP and transform in EndPoint
        '*
        __LocalHost = System.Net.Dns.Resolve(System.Net.Dns.GetHostName())

        '*
        '* Defines Host
        '*
        __RemoteHost = Nothing

    End Sub

    '*
    '*   Function : PingHost
    '*     Author : Paulo Santos
    '*       Date : 05/09/2002
    '*  Objective : Pings a specified host
    '*
    '* Parameters : Host as String
    '*
    '*    Returns : Response time in milliseconds
    '*              (-1) if error
    '*
    Public Overloads Function Ping() As Long

        Dim intCount As Integer
        Dim aReplyBuffer(255) As Byte

        Dim intNBytes As Integer = 0

        Dim intEnd As Integer
        Dim intStart As Integer

        Dim epFrom As System.Net.EndPoint
        Dim epServer As System.Net.EndPoint
        Dim ipepServer As System.Net.IPEndPoint

        If (__RemoteHost Is Nothing) Then
            __Error = pingErrorCodes.Unknown
            Return (PING_ERROR)
        End If

        '*
        '* Transforms the IP address in EndPoint
        '*
        ipepServer = New System.Net.IPEndPoint(__RemoteHost.AddressList(0), 0)
        epServer = CType(ipepServer, System.Net.EndPoint)

        epFrom = New System.Net.IPEndPoint(__LocalHost.AddressList(0), 0)

        '*
        '* Builds the packet to send
        '*
        DataSize = Convert.ToByte(DataSize + __BufferHeaderSize)

        '*
        '* The packet must by an even number, so if the DataSize is and odd number adds one 
        '* 
        If (DataSize Mod 2 = 1) Then
            DataSize += Convert.ToByte(1)
        End If
        Dim aRequestBuffer(DataSize - 1) As Byte

        '*
        '* --- ICMP Echo Header Format ---
        '* (first 8 bytes of the data buffer)
        '*
        '* Buffer (0) ICMP Type Field
        '* Buffer (1) ICMP Code Field
        '*     (must be 0 for Echo and Echo Reply)
        '* Buffer (2) checksum hi
        '*     (must be 0 before checksum calc)
        '* Buffer (3) checksum lo
        '*     (must be 0 before checksum calc)
        '* Buffer (4) ID hi
        '* Buffer (5) ID lo
        '* Buffer (6) __Sequence hi
        '* Buffer (7) __Sequence lo
        '* Buffer (8)..(n)  Ping Data
        '*

        '*
        '* Set Type Field
        '*
        aRequestBuffer(0) = Convert.ToByte(ICMP_ECHO) ' ECHO Request

        '*
        '* Set ID field
        '*
        BitConverter.GetBytes(__Identifier).CopyTo(aRequestBuffer, 4)

        '*
        '* Set __Sequence
        '*
        BitConverter.GetBytes(__Sequence).CopyTo(aRequestBuffer, 6)

        '*
        '* Load some data into buffer
        '*
        Dim i As Integer
        For i = 8 To DataSize - 1
            aRequestBuffer(i) = Convert.ToByte(i Mod 8)
        Next i

        '*
        '* Calculate Checksum
        '*
        Call CreateChecksum(aRequestBuffer, DataSize, aRequestBuffer(2), aRequestBuffer(3))


        '*
        '* Try send the packet
        '*
        Try
            '*
            '* Create the socket
            '*
            Dim sckSocket As New System.Net.Sockets.Socket( _
                                            Net.Sockets.AddressFamily.InterNetwork, _
                                            Net.Sockets.SocketType.Raw, _
                                            Net.Sockets.ProtocolType.Icmp)
            sckSocket.Blocking = False

            '*
            '* Sends Package
            '*
            sckSocket.SendTo(aRequestBuffer, 0, DataSize, SocketFlags.None, ipepServer)

            '*
            '* Records the Start Time, after sending the Echo Request
            '*
            intStart = System.Environment.TickCount

            '*
            '* Waits for response
            '*
            Do
                System.Threading.Thread.Sleep(10)
                Application.DoEvents()
                Try
                    intNBytes = sckSocket.ReceiveFrom(aReplyBuffer, SocketFlags.None, epServer)
                Catch objErr As Exception
                End Try
            Loop Until (intNBytes > 0) Or ((System.Environment.TickCount - intStart) > TimeOut)

            '*
            '* Check to see if the TimeOut was hit
            '*
            If ((System.Environment.TickCount - intStart) > TimeOut) Then
                __Error = pingErrorCodes.TimeOut ' PING_ERROR_TIME_OUT
                Return (PING_ERROR)
            End If

            '*
            '* Records End Time
            '*
            intEnd = System.Environment.TickCount

            If (intNBytes > 0) Then
                '*
                '* Updates LastError with the state of the server
                '*
                Select Case aReplyBuffer(20)
                    Case 0 : __Error = pingErrorCodes.Success
                    Case 1 : __Error = pingErrorCodes.BufferTooSmall
                    Case 2 : __Error = pingErrorCodes.DestinationUnreahable
                    Case 3 : __Error = pingErrorCodes.HostNotReachable
                    Case 4 : __Error = pingErrorCodes.ProtocolNotReachable
                    Case 5 : __Error = pingErrorCodes.PortNotReachable
                    Case 6 : __Error = pingErrorCodes.NoResourcesAvailable
                    Case 7 : __Error = pingErrorCodes.BadOption
                    Case 8 : __Error = pingErrorCodes.HardwareError
                    Case 9 : __Error = pingErrorCodes.PacketTooBig
                    Case 10 : __Error = pingErrorCodes.ReqestedTimedOut
                    Case 11 : __Error = pingErrorCodes.BadRequest
                    Case 12 : __Error = pingErrorCodes.BadRoute
                    Case 13 : __Error = pingErrorCodes.TTLExprdInTransit
                    Case 14 : __Error = pingErrorCodes.TTLExprdReassemb
                    Case 15 : __Error = pingErrorCodes.ParameterProblem
                    Case 16 : __Error = pingErrorCodes.SourceQuench
                    Case 17 : __Error = pingErrorCodes.OptionTooBig
                    Case 18 : __Error = pingErrorCodes.BadDestination
                    Case 19 : __Error = pingErrorCodes.AddressDeleted
                    Case 20 : __Error = pingErrorCodes.SpecMTUChange
                    Case 21 : __Error = pingErrorCodes.MTUChange
                    Case 22 : __Error = pingErrorCodes.Unload
                    Case Else : __Error = pingErrorCodes.GeneralFailure
                End Select

            End If

            Return (intEnd - intStart)
        Catch oExcept As Exception
            '
        End Try

    End Function

    Public Overloads Shared Function Ping(ByVal fqdnHostName As String) As Long
        Dim p As New PingClass
        Call p.SetRemoteHost(fqdnHostName)
        Return p.Ping()
    End Function

    Public Overloads Shared Function Ping(ByVal ipAddress As System.Net.IPAddress) As Long
        Dim p As New PingClass
        Call p.SetRemoteHost(ipAddress)
        Return p.Ping()
    End Function

    Public ReadOnly Property LastError() As pingErrorCodes
        Get
            Return __Error
        End Get
    End Property

    Public Overloads Shared Function BeginPing(ByVal fqdnHostName As String, _
                                               ByVal requestCallback As PingCallback, _
                                               ByVal stateObject As Object) As PingAsyncResult

        Dim ipar As New PingAsyncResult
        Dim pingInfo As New PingAsyncRequest

        ipar.PingAsyncState = stateObject

        With pingInfo
            .RemoteHost = fqdnHostName
            .PingAsyncResult = ipar
            .PingCallback = requestCallback
        End With

        System.Threading.ThreadPool.QueueUserWorkItem(AddressOf DoAsyncPing, pingInfo)

        Return (ipar)

    End Function

    Public Overloads Shared Function BeginPing(ByVal ipAddress As System.Net.IPAddress, _
                                               ByVal requestCallback As PingCallback, _
                                               ByVal stateObject As Object) As PingAsyncResult

        Return BeginPing(ipAddress.ToString, requestCallback, stateObject)

    End Function

    Private Shared Sub DoAsyncPing(ByVal pingInfoObject As Object)

        Dim pingInfo As PingAsyncRequest = CType(pingInfoObject, PingAsyncRequest)

        Dim p As New PingClass
        p.SetRemoteHost(pingInfo.RemoteHost)
        pingInfo.PingAsyncResult.PingTime = p.Ping()
        pingInfo.PingAsyncResult.PingError = p.LastError

        Call pingInfo.PingCallback(pingInfo.PingAsyncResult)

    End Sub

    Public Shared Sub EndPing(ByVal par As PingAsyncResult)
        '*
        '* It would be good if we had something to do here
        '*
    End Sub

    '*
    '* ICMP requires a checksum that  is  the  one's 
    '* complement of the one's complement sum of the 
    '* 16-bit values in the data in the buffer.
    '*
    '* Use this procedure to load the Checksum field 
    '* of the buffer.
    '*
    '* The Checksum Field (hi and lo bytes) must be
    '* zero before calling this procedure.
    '*
    Private Sub CreateChecksum(ByRef data() As Byte, _
                               ByVal Size As Integer, _
                               ByRef HiByte As Byte, _
                               ByRef LoByte As Byte)

        Dim i As Integer
        Dim chk As Integer = 0

        For i = 0 To Size - 1 Step 2
            chk += Convert.ToInt32(data(i) * &H100 + data(i + 1))
        Next

        chk = Convert.ToInt32((chk And &HFFFF&) + Fix(chk / &H10000&))
        chk += Convert.ToInt32(Fix(chk / &H10000&))
        chk = Not (chk)

        HiByte = Convert.ToByte((chk And &HFF00) / &H100)
        LoByte = Convert.ToByte(chk And &HFF)

    End Sub

End Class
