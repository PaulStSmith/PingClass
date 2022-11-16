# PingClass
A simple class to ping a computer over a network. [archival]

This is an old project of mine, that I created back in 2002.<br/>
Imported into GitHub for archival puposes only.<br/>
There will be no updates.

<H2>Introduction</H2>
<P>This is an old class of mine. I created this class when I was developing a server monitoring system back in 2002, and I reasoned that the easiest way to know if the server was up is to ping it. </P>
<P>So, I started searching the web and I found a code developed for C# and after a few bumps I converted to VB .NET and the code works fine, up to the point of performance. </P>
<P>I use the <A href="http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cpref/html/frlrfSystemNetSocketsSocketClassTopic.asp" target="_new" shape="rect">Socket </A>class to create a packet of raw data to be sent to the host. The send part works fine, but the response (when I listen for the echo of the packet I just sent) takes ages to process. I've checked the docs on Microsoft about the <A href="http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cpref/html/frlrfSystemNetSocketsSocketClassTopic.asp" target="_new" shape="rect">ReceiveFrom</A> method to see if they have some insight about what's going on, and finally found a solution.</P>
<P>Since then I have published this code in some other places, always with good feedback from my fellow developers friends.</P>
<P>When I decided to publish it on my website I knew I needed to take another look at this class because so much time has passed.</P>
<P>One of the things that always bothered me with this class is that it was synchronous, <EM>i.e.</EM> when you call the method to ping the host it doesn't return until the host responds or a time-out is reached.</P>
<P>Hence I decided to go through it once more and make a second version of it with asynchronous methods too. This way is up to you to use the method that suits best your application.</P>
<P><STRONG>Note:</STRONG></P>
<P>In the .NET Framework 2.0 there is a namespace called <A href="http://msdn2.microsoft.com/en-us/library/3ew4thdx(en-US,VS.80).aspx" shape="rect">NetworkInformation </A> where a new class called <A href="http://msdn2.microsoft.com/en-us/library/system.net.networkinformation.ping.aspx" shape="rect">Ping</A> that supersedes this class.</P>
<H2>PingClass Members</H2>
<H3>Public Enumerators</H3>
<TABLE class="Members" id="table1" cellSpacing="0" cellPadding="2" rules="rows" border="1">
<TBODY>
<TR>
<TD class="Icon" vAlign="top" align="center">
<IMG height="16" alt="Public Constructor" src="http://pjondevelopment.50webs.com/images/class/enum.gif" width="16" /></TD>
<TD class="Definition" vAlign="top">
<CODE>pingErrorCodes</CODE><BR />
Enumerates the possible errors that might occur when pinging a remote host.
<PRE class="vb" xml:space="preserve">
Public Enum pingErrorCodes As Integer
    Success               = 0
    Unknown               = (-1)
    HostNotFound          = &amp;H8001
    SocketDidntSend       = &amp;H8002
    HostNotResponding     = &amp;H8003
    TimeOut               = &amp;H8004
    BufferTooSmall        = &amp;H8101
    DestinationUnreahable = &amp;H8102
    HostNotReachable      = &amp;H8103
    ProtocolNotReachable  = &amp;H8104
    PortNotReachable      = &amp;H8105
    NoResourcesAvailable  = &amp;H8106
    BadOption             = &amp;H8107
    HardwareError         = &amp;H8108
    PacketTooBig          = &amp;H8109
    ReqestedTimedOut      = &amp;H810A
    BadRequest            = &amp;H810B
    BadRoute              = &amp;H810C
    TTLExprdInTransit     = &amp;H810D
    TTLExprdReassemb      = &amp;H810E
    ParameterProblem      = &amp;H810F
    SourceQuench          = &amp;H8110
    OptionTooBig          = &amp;H8111
    BadDestination        = &amp;H8112
    AddressDeleted        = &amp;H8113
    SpecMTUChange         = &amp;H8114
    MTUChange             = &amp;H8115
    Unload                = &amp;H8116
    GeneralFailure        = &amp;H8117
End Enum</PRE></TD></TR></TBODY></TABLE>
<H3>Public Constructors</H3>
<TABLE class="Members" id="table2" cellSpacing="0" cellPadding="2" rules="rows" border="1">
<TBODY>
<TR>
<TD class="Icon" vAlign="top" align="center">
<IMG height="16" alt="Public Constructor" src="http://pjondevelopment.50webs.com/images/class/method.gif" width="16" /></TD>
<TD class="Definition" vAlign="top">
<CODE>PingClass Constructor</CODE><BR />
Initializes a new instance of the <B>PingClass</B> Class. 
<PRE class="vb" xml:space="preserve">Public Sub New()</PRE></TD></TR></TBODY></TABLE>
<H3>Public Properties</H3>
<TABLE class="Members" id="table3" cellSpacing="0" cellPadding="2" rules="rows" border="1">
<TBODY>
<TR>
<TD class="Icon" vAlign="top" align="center">
<IMG height="16" alt="Public Property" src="http://pjondevelopment.50webs.com/images/class/property.gif" width="16" /></TD>
<TD class="Definition" vAlign="top">
<CODE>DataSize</CODE><BR />
Gets or sets the size of the package to be send on the ping request. Default 32 bytes. 
<PRE class="vb" xml:space="preserve">Public Property DataSize() As Byte</PRE></TD></TR>
<TR>
<TD class="Icon" vAlign="top" align="center">
<IMG height="16" alt="Public Property" src="http://pjondevelopment.50webs.com/images/class/property.gif" width="16" /></TD>
<TD class="Definition" vAlign="top">
<CODE>LastError</CODE><BR />
<B>ReadOnly</B><BR />
The last error occurred when pinging a host.<BR />
<PRE class="vb" xml:space="preserve">Public ReadOnly Property LastError() _
                         As pingErrorCodes</PRE></TD></TR>
<TR>
<TD class="Icon" vAlign="top" align="center">
<IMG height="16" alt="Public Property" src="http://pjondevelopment.50webs.com/images/class/property.gif" width="16" /></TD>
<TD class="Definition" vAlign="top">
<CODE>LocalHost</CODE><BR />
<B>ReadOnly</B><BR />
A System.Net.IPHostEntry that represents the local computer.<BR />
<PRE class="vb" xml:space="preserve">Public ReadOnly Property LocalHost() _
                         As System.Net.IPHostEntry</PRE></TD></TR>
<TR>
<TD class="Icon" vAlign="top" align="center">
<IMG height="16" alt="Public Property" src="http://pjondevelopment.50webs.com/images/class/property.gif" width="16" /></TD>
<TD class="Definition" vAlign="top">
<CODE>RemoteHost</CODE><BR />
A System.Net.IPHostEntry that represents the remote host to be pinged.<BR />
<PRE class="vb" xml:space="preserve">Public Property RemoteHost() _
                         As System.Net.IPHostEntry</PRE></TD></TR>
<TR>
<TD class="Icon" vAlign="top" align="center">
<IMG height="16" alt="Public Property" src="http://pjondevelopment.50webs.com/images/class/property.gif" width="16" /></TD>
<TD class="Definition" vAlign="top">
<CODE>TimeOut</CODE><BR />
The amount of time in milliseconds in which the remote host must reply.<BR />
<PRE class="vb" xml:space="preserve">Public Property TimeOut() As Long</PRE></TD></TR></TBODY></TABLE>
<H3>Public Methods</H3>
<TABLE class="Members" id="table4" cellSpacing="0" cellPadding="2" rules="rows" border="1">
<TBODY>
<TR>
<TD class="Icon" vAlign="top" align="center">
<IMG height="16" alt="Public Method" src="http://pjondevelopment.50webs.com/images/class/method.gif" width="16" /></TD>
<TD class="Definition">
<CODE>BeginPing</CODE><BR />
Begins the asynchronous execution of a PING request to a specified host. 
<PRE class="vb" xml:space="preserve">Public Overloads Shared Function BeginPing( _
          ByVal fqdnHostName As String, _
          ByVal requestCallback As PingCallback, _
          ByVal stateObject As Object) _
                                 As PingAsyncResult

Public Overloads Shared Function BeginPing( _
          ByVal ipAddress As System.Net.IPAddress, _
          ByVal requestCallback As PingCallback, _
          ByVal stateObject As Object) _
                                 As PingAsyncResult</PRE></TD></TR>
<TR>
<TD class="Icon" vAlign="top" align="center">
<IMG height="16" alt="Public Method" src="http://pjondevelopment.50webs.com/images/class/method.gif" width="16" /></TD>
<TD class="Definition">
<CODE>EndPing</CODE><BR />
Ends an asynchronous PING request.
<PRE class="vb" xml:space="preserve">Public Shared Sub EndPing( _
                      ByVal par As PingAsyncResult)</PRE></TD></TR>
<TR>
<TD class="Icon" vAlign="top" align="center">
<IMG height="16" alt="Public Method" src="http://pjondevelopment.50webs.com/images/class/method.gif" width="16" /></TD>
<TD class="Definition">
<CODE>Ping</CODE><BR />
Pings the remote host and return the number of milliseconds that the host took to reply. 
<PRE class="vb" xml:space="preserve">Public Overloads Function Ping() As Long

Public Overloads _
          Shared Function Ping( _
              ByVal fqdnHostName As String) As Long

Public Overloads _
          Shared Function Ping( _
   ByVal ipAddress As System.Net.IPAddress) As Long</PRE></TD></TR>
<TR>
<TD class="Icon" vAlign="top" align="center">
<IMG height="16" alt="Public Method" src="http://pjondevelopment.50webs.com/images/class/method.gif" width="16" /></TD>
<TD class="Definition">
<CODE>SetRemoteHost</CODE><BR />
Sets the RemoteHost Property.
<PRE class="vb" xml:space="preserve">Public Overloads Sub SetRemoteHost( _
                      ByVal fqdnHostName As String)

Public Overloads Sub SetRemoteHost( _
           ByVal ipAddress As System.Net.IPAddress)</PRE></TD></TR></TBODY></TABLE>
<H2>Using the code</H2>
<P>For a quick and dirty sample just take a look at the follow snippet.</P>
<PRE class="vb" xml:space="preserve">MsgBox(&quot;yahoo.com took &quot; &amp; _
                  PingClass.Ping(&quot;yahoo.com&quot;) &amp; _
                                     &quot; milliseconds to reply&quot;)
</PRE>
<P>Yes, it is that easy.</P>
<H2>Background information</H2>
<P>Always a good source of information about the Internet is the &quot;Request For Comments&quot;. More specifically about ping and ping request are: </P>
<DL>
    <DT><A href="ftp://ftp.rfc-editor.org/in-notes/rfc792.txt" shape="rect"><LI class="download">&nbsp; RFC 792</LI></A></DT>
    <DD>Internet Control Message Protocol.</DD>
    <DT><A href="ftp://ftp.rfc-editor.org/in-notes/rfc1122.txt" shape="rect"><LI class="download">&nbsp; RFC 1122</LI></A></DT>
    <DD>Requirements for Internet Hosts -- Communication Layers.</DD>
    <DT><A href="ftp://ftp.rfc-editor.org/in-notes/rfc1812.txt" shape="rect"><LI class="download">&nbsp; RFC 1812</LI></A></DT>
    <DD>Requirements for IP Version 4 Routers.</DD>
</DL>
