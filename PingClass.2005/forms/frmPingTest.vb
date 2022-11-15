Public Class frmPingTest
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtHost As System.Windows.Forms.TextBox
    Friend WithEvents btnPing As System.Windows.Forms.Button
    Friend WithEvents txtLog As System.Windows.Forms.TextBox
    Friend WithEvents btnAsync As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPingTest))
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtHost = New System.Windows.Forms.TextBox
        Me.btnPing = New System.Windows.Forms.Button
        Me.txtLog = New System.Windows.Forms.TextBox
        Me.btnAsync = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "&Host to ping :"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtHost
        '
        Me.txtHost.Location = New System.Drawing.Point(84, 8)
        Me.txtHost.Name = "txtHost"
        Me.txtHost.Size = New System.Drawing.Size(154, 21)
        Me.txtHost.TabIndex = 1
        Me.txtHost.Text = ""
        '
        'btnPing
        '
        Me.btnPing.Location = New System.Drawing.Point(244, 8)
        Me.btnPing.Name = "btnPing"
        Me.btnPing.Size = New System.Drawing.Size(67, 21)
        Me.btnPing.TabIndex = 2
        Me.btnPing.Text = "&Sync"
        '
        'txtLog
        '
        Me.txtLog.Location = New System.Drawing.Point(6, 42)
        Me.txtLog.Multiline = True
        Me.txtLog.Name = "txtLog"
        Me.txtLog.ReadOnly = True
        Me.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtLog.Size = New System.Drawing.Size(378, 192)
        Me.txtLog.TabIndex = 3
        Me.txtLog.Text = ""
        '
        'btnAsync
        '
        Me.btnAsync.Location = New System.Drawing.Point(317, 8)
        Me.btnAsync.Name = "btnAsync"
        Me.btnAsync.Size = New System.Drawing.Size(67, 21)
        Me.btnAsync.TabIndex = 2
        Me.btnAsync.Text = "&Async"
        '
        'frmPingTest
        '
        Me.AcceptButton = Me.btnPing
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
        Me.ClientSize = New System.Drawing.Size(394, 242)
        Me.Controls.Add(Me.txtLog)
        Me.Controls.Add(Me.txtHost)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnPing)
        Me.Controls.Add(Me.btnAsync)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmPingTest"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Ping Test"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnPing_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPing.Click

        Dim objPingHost As New PingClass
        Dim lngPingReply As Long

        '*
        '* Is there any host in the TextBox?
        '*
        If (Trim(Me.txtHost.Text) = "") Then
            Call MsgBox("Please, type a host name or IP.", MsgBoxStyle.Exclamation, Application.ProductName)
            Exit Sub
        End If

        '*
        '* Prepares to Ping a Host
        '*
        objPingHost.TimeOut = 5000                      ' --> 5000 msec  Time Out
        objPingHost.DataSize = 32                       ' -->   32 bytes Package Size
        objPingHost.SetRemoteHost(Me.txtHost.Text.Trim) ' --> Host to Ping

        '*
        '* Writes in the Log
        '*
        Me.txtLog.Text &= "Trying to ping " & Trim(Me.txtHost.Text) & " ..." & vbCrLf

        '*
        '* Ping the Host and get the reply time
        '*
        lngPingReply = objPingHost.Ping()

        '*
        '* Check for error
        '*
        If (lngPingReply = (-1)) Then
            Me.txtLog.Text &= "Error : " & objPingHost.LastError.ToString & vbCrLf & vbCrLf
        Else
            Me.txtLog.Text &= "Ping Reply " & lngPingReply & " msec." & vbCrLf & vbCrLf
        End If

    End Sub

    Private Sub btnAsync_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAsync.Click

        Me.txtLog.Text &= "Starting async ping to " & Me.txtHost.Text & " ..." & vbCrLf

        Call PingClass.BeginPing(Me.txtHost.Text.Trim, AddressOf DoAsyncPing, Nothing)

        Me.txtLog.Text &= "Waiting for async to return " & Me.txtHost.Text & " ..." & vbCrLf

    End Sub

    Private Sub DoAsyncPing(ByVal par As PingClass.PingAsyncResult)

        Me.txtLog.Text &= "Async ping returned " & Me.txtHost.Text & vbCrLf

        Call PingClass.EndPing(par)
        If (par.PingError = PingClass.pingErrorCodes.Success) Then
            Me.txtLog.Text &= "Ping Reply " & par.PingTime & " msec." & vbCrLf & vbCrLf
        Else
            Me.txtLog.Text &= "Error : " & par.PingError.ToString & vbCrLf & vbCrLf
        End If

    End Sub

    Private Sub txtLog_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLog.TextChanged
        Me.txtHost.SelectionStart = Me.txtHost.Text.Length
    End Sub
End Class
