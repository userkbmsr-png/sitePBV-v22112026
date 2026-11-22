' ============================================================
'  FTP_Upload.vb  -  Visual Basic .NET (WinForms, .NET Framework 4.x)
'
'  Compilare (linie de comanda):
'  C:\Windows\Microsoft.NET\Framework64\v4.0.30319\vbc.exe /target:winexe /out:FTP_Upload.exe ^
'     /r:System.dll,System.Windows.Forms.dll,System.Drawing.dll,System.Security.dll FTP_Upload.vb
'
'  Fisiere create langa executabil:
'  Program instalat in Documente\Aplicatii; acolo se creeaza:
'     setari.ini   - setarile (parola codata cu DPAPI, nu in text clar)
'     incarcate.db - lista fisierelor deja incarcate
'  Folder local prestabilit: Documente\uploads ; teme/module: Documente\wp-content
' ============================================================
Option Strict On
Option Explicit On
Option Infer On

Imports System
Imports System.IO
Imports System.Net
Imports System.Reflection
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Security.Cryptography

Public Module Program
    <STAThread()>
    Public Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New MainForm())
    End Sub
End Module

' ------------------------------------------------------------
'  Setari (setari.ini)
' ------------------------------------------------------------
Public Class AppSettings
    ' Cai prestabilite (nu se mai citesc din setari.ini)
    Public Shared ReadOnly Property AppDir As String
        Get
            Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "roSite", "Aplicatii")
        End Get
    End Property

    Public ReadOnly Property LocalFolder As String
        Get
            Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "roSite", "uploads")
        End Get
    End Property

    Public Shared ReadOnly Property WpContentFolder As String
        Get
            Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "roSite", "wp-content")
        End Get
    End Property

    Public Server As String = ""
    Public User As String = ""
    Public Password As String = ""
    Public FtpFolder As String = "/"
    Public AutoAction As Boolean = False

    Private Shared ReadOnly Entropy As Byte() = Encoding.UTF8.GetBytes("FTP_Upload.v1")

    Public Shared ReadOnly Property IniPath As String
        Get
            Return Path.Combine(AppDir, "setari.ini")
        End Get
    End Property

    Public Shared Function EncodePassword(plain As String) As String
        If String.IsNullOrEmpty(plain) Then Return ""
        Dim b As Byte() = ProtectedData.Protect(Encoding.UTF8.GetBytes(plain), Entropy, DataProtectionScope.CurrentUser)
        Return Convert.ToBase64String(b)
    End Function

    Public Shared Function DecodePassword(coded As String) As String
        If String.IsNullOrEmpty(coded) Then Return ""
        Try
            Dim b As Byte() = ProtectedData.Unprotect(Convert.FromBase64String(coded), Entropy, DataProtectionScope.CurrentUser)
            Return Encoding.UTF8.GetString(b)
        Catch
            Return ""
        End Try
    End Function

    Public Shared Function Load() As AppSettings
        Dim s As New AppSettings()
        If Not File.Exists(IniPath) Then Return s
        For Each line As String In File.ReadAllLines(IniPath, Encoding.UTF8)
            Dim p As Integer = line.IndexOf("="c)
            If p <= 0 Then Continue For
            Dim k As String = line.Substring(0, p).Trim()
            Dim v As String = line.Substring(p + 1).Trim()
            Select Case k.ToLowerInvariant()
                Case "server" : s.Server = v
                Case "user" : s.User = v
                Case "parola" : s.Password = DecodePassword(v)
                Case "folderftp" : s.FtpFolder = If(v = "", "/", v)
                Case "actiuneautomata" : s.AutoAction = (v = "1")
            End Select
        Next
        Return s
    End Function

    Public Sub Save()
        Dim sb As New StringBuilder()
        sb.AppendLine("[Setari]")
        sb.AppendLine("Server=" & Server)
        sb.AppendLine("User=" & User)
        sb.AppendLine("Parola=" & EncodePassword(Password))
        sb.AppendLine("FolderFTP=" & FtpFolder)
        sb.AppendLine("ActiuneAutomata=" & If(AutoAction, "1", "0"))
        Directory.CreateDirectory(AppDir)
        File.WriteAllText(IniPath, sb.ToString(), Encoding.UTF8)
    End Sub

    Public Function IsValid() As Boolean
        Return Server <> "" AndAlso User <> ""
    End Function
End Class

' ------------------------------------------------------------
'  Functii FTP
' ------------------------------------------------------------
Public Class FtpHelper
    Public Shared Function BuildUrl(server As String, path As String) As String
        Dim s As String = server.Trim()
        If Not s.StartsWith("ftp://", StringComparison.OrdinalIgnoreCase) Then s = "ftp://" & s
        s = s.TrimEnd("/"c)
        Dim parts As String() = path.Split(New Char() {"/"c, "\"c}, StringSplitOptions.RemoveEmptyEntries)
        Dim sb As New StringBuilder(s)
        For Each seg As String In parts
            sb.Append("/"c).Append(Uri.EscapeDataString(seg))
        Next
        Return sb.ToString()
    End Function

    Public Shared Function CreateRequest(url As String, method As String, cfg As AppSettings) As FtpWebRequest
        Dim req As FtpWebRequest = CType(WebRequest.Create(url), FtpWebRequest)
        req.Method = method
        req.Credentials = New NetworkCredential(cfg.User, cfg.Password)
        req.UsePassive = True
        req.UseBinary = True
        req.KeepAlive = False
        req.Timeout = 30000
        req.ReadWriteTimeout = 60000
        Return req
    End Function

    ' Returneaza lista numelor de subfoldere din folderul dat
    Public Shared Function ListDirs(cfg As AppSettings, folder As String) As List(Of String)
        Dim res As New List(Of String)()
        Dim url As String = BuildUrl(cfg.Server, folder) & "/"
        Dim req As FtpWebRequest = CreateRequest(url, WebRequestMethods.Ftp.ListDirectoryDetails, cfg)
        Dim unixRx As New Regex("^d\S*\s+\d+\s+\S+\s+\S+\s+\d+\s+\S+\s+\S+\s+\S+\s+(.+)$")
        Dim dosRx As New Regex("^\d\d-\d\d-\d\d\s+\d\d:\d\d[AP]M\s+<DIR>\s+(.+)$", RegexOptions.IgnoreCase)
        Using resp As FtpWebResponse = CType(req.GetResponse(), FtpWebResponse)
            Using sr As New StreamReader(resp.GetResponseStream())
                Dim line As String = sr.ReadLine()
                Do While line IsNot Nothing
                    Dim m As Match = unixRx.Match(line)
                    If Not m.Success Then m = dosRx.Match(line)
                    If m.Success Then
                        Dim name As String = m.Groups(1).Value.Trim()
                        If name <> "." AndAlso name <> ".." Then res.Add(name)
                    End If
                    line = sr.ReadLine()
                Loop
            End Using
        End Using
        res.Sort(StringComparer.OrdinalIgnoreCase)
        Return res
    End Function

    ' Incarca un fisier (inlocuieste fisierul existent pe server)
    Public Shared Sub UploadFile(cfg As AppSettings, fullPath As String, remotePath As String,
                                 onBytes As Action(Of Integer), isCancelled As Func(Of Boolean))
        Dim req As FtpWebRequest = CreateRequest(BuildUrl(cfg.Server, remotePath), WebRequestMethods.Ftp.UploadFile, cfg)
        Using fs As New FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            req.ContentLength = fs.Length
            Using rs As Stream = req.GetRequestStream()
                Dim buf(32767) As Byte
                Dim n As Integer = fs.Read(buf, 0, buf.Length)
                Do While n > 0
                    If isCancelled() Then Throw New OperationCanceledException()
                    rs.Write(buf, 0, n)
                    onBytes(n)
                    n = fs.Read(buf, 0, buf.Length)
                Loop
            End Using
        End Using
        Using resp As FtpWebResponse = CType(req.GetResponse(), FtpWebResponse)
        End Using
    End Sub

    Public Shared Sub TestConnection(cfg As AppSettings)
        Dim url As String = BuildUrl(cfg.Server, "/") & "/"
        Dim req As FtpWebRequest = CreateRequest(url, WebRequestMethods.Ftp.ListDirectory, cfg)
        Using resp As FtpWebResponse = CType(req.GetResponse(), FtpWebResponse)
        End Using
    End Sub

    Public Shared Sub MakeDir(cfg As AppSettings, path As String)
        Try
            Dim req As FtpWebRequest = CreateRequest(BuildUrl(cfg.Server, path), WebRequestMethods.Ftp.MakeDirectory, cfg)
            Using resp As FtpWebResponse = CType(req.GetResponse(), FtpWebResponse)
            End Using
        Catch
            ' exista deja sau nu se poate crea - eroarea reala apare la upload
        End Try
    End Sub

    Public Shared Sub DeleteFile(cfg As AppSettings, path As String)
        Try
            Dim req As FtpWebRequest = CreateRequest(BuildUrl(cfg.Server, path), WebRequestMethods.Ftp.DeleteFile, cfg)
            Using resp As FtpWebResponse = CType(req.GetResponse(), FtpWebResponse)
            End Using
        Catch
        End Try
    End Sub
End Class

' ------------------------------------------------------------
'  Fisier de incarcat
' ------------------------------------------------------------
Public Class FileItem
    Public Rel As String      ' cale relativa, cu "/"
    Public Full As String
    Public Size As Long
    Public Ticks As Long
    Public Function Signature() As String
        Return Size.ToString() & "|" & Ticks.ToString()
    End Function
End Class

' ------------------------------------------------------------
'  Fereastra principala
' ------------------------------------------------------------
Public Class MainForm
    Inherits Form

    Private lblConn As New Label()
    Private lblStatus As New Label()
    Private pb As New ProgressBar()
    Private WithEvents btnSettings As New Button()
    Private WithEvents btnUpload As New Button()
    Private WithEvents btnExit As New Button()
    Private WithEvents bw As New BackgroundWorker()
    Private WithEvents bwConn As New BackgroundWorker()
    Private WithEvents tmrClose As New Timer()

    Private _cfg As New AppSettings()
    Private _pending As New List(Of FileItem)()
    Private _busy As Boolean = False
    Private _closeAfter As Boolean = False
    Private _confirming As Boolean = False
    Private _connOk As Boolean = False
    Private _autoStart As Boolean = False

    ' folosite de thread-ul de transfer
    Private _totalBytes As Long
    Private _doneBytes As Long
    Private _curIndex As Integer
    Private _lastPct As Integer = -1
    Private _lastError As String = ""
    Private _fileCount As Integer
    Private _madeDirs As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

    Private ReadOnly DbPath As String = Path.Combine(AppSettings.AppDir, "incarcate.db")

    Public Sub New()
        Me.Text = "FTP Upload"
        Me.ClientSize = New Size(520, 170)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False

        lblConn.SetBounds(15, 8, 490, 32)
        lblConn.AutoSize = False
        lblConn.Font = New Font(Me.Font, FontStyle.Bold)
        lblConn.Text = ""

        lblStatus.SetBounds(15, 44, 490, 32)
        lblStatus.AutoSize = False
        lblStatus.Text = ""

        pb.SetBounds(15, 84, 490, 22)
        pb.Minimum = 0
        pb.Maximum = 100

        btnSettings.Text = "Setari"
        btnSettings.SetBounds(15, 124, 110, 32)
        btnUpload.Text = "Incarca"
        btnUpload.SetBounds(285, 124, 110, 32)
        btnUpload.Enabled = False
        btnExit.Text = "Iesire"
        btnExit.SetBounds(395, 124, 110, 32)

        Me.Controls.AddRange(New Control() {lblConn, lblStatus, pb, btnSettings, btnUpload, btnExit})

        bw.WorkerReportsProgress = True
        bw.WorkerSupportsCancellation = True
        bwConn.WorkerSupportsCancellation = False
        tmrClose.Interval = 30000
    End Sub

    ' ---------- pornire ----------
    Private Sub MainForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Try
            Directory.CreateDirectory(AppSettings.AppDir)
        Catch
        End Try
        ' scurtatura proprie pe Desktop
        Try
            LocalSetup.EnsureAppShortcut()
        Catch
        End Try
        _autoStart = True
        InitState() ' scaneaza fisierele si porneste testul de conexiune
    End Sub

    Private Sub InitState()
        _cfg = AppSettings.Load()
        _pending.Clear()
        pb.Value = 0
        If Not _cfg.IsValid() Then
            lblConn.Text = ""
            lblStatus.Text = "Setarile nu sunt configurate. Apasati butonul Setari."
            btnUpload.Enabled = False
            _autoStart = False
            Return
        End If
        Try
            LocalSetup.EnsureStructure(_cfg.LocalFolder)
            ScanFiles()
        Catch ex As Exception
            lblStatus.Text = "Eroare la citirea folderului local: " & ex.Message
            btnUpload.Enabled = False
            _autoStart = False
            Return
        End Try
        lblStatus.Text = "Sunt disponibile " & _pending.Count & " fisiere pentru actualizare pe server"
        btnUpload.Enabled = False ' se activeaza dupa testul de conexiune
        StartConnTest()
    End Sub

    ' ---------- test conexiune FTP ----------
    Private Sub StartConnTest()
        If bwConn.IsBusy Then Return
        _connOk = False
        lblConn.ForeColor = SystemColors.GrayText
        lblConn.Text = "Se verifica conexiunea FTP..."
        btnSettings.Enabled = False
        bwConn.RunWorkerAsync(_cfg)
    End Sub

    Private Sub bwConn_DoWork(sender As Object, e As DoWorkEventArgs) Handles bwConn.DoWork
        FtpHelper.TestConnection(CType(e.Argument, AppSettings))
    End Sub

    Private Sub bwConn_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles bwConn.RunWorkerCompleted
        btnSettings.Enabled = Not _busy
        If e.Error Is Nothing Then
            _connOk = True
            lblConn.ForeColor = Color.ForestGreen
            lblConn.Text = "OK - conectare reusita"
        Else
            _connOk = False
            lblConn.ForeColor = Color.Firebrick
            lblConn.Text = "Eroare - " & e.Error.Message.Replace(vbCr, " ").Replace(vbLf, " ")
        End If
        If _busy Then Return
        btnUpload.Enabled = _connOk AndAlso _pending.Count > 0

        If _autoStart Then
            _autoStart = False
            If _cfg.AutoAction Then
                If _connOk AndAlso _pending.Count > 0 Then
                    StartUpload()
                Else
                    If _connOk Then lblStatus.Text = "Nu exista fisiere noi de transferat."
                    tmrClose.Start() ' inchidere dupa 30 secunde
                End If
            End If
        End If
    End Sub

    ' ---------- baza de date a fisierelor incarcate ----------
    Private Function DbHeader() As String
        Return "#" & _cfg.Server.Trim().ToLowerInvariant() & "|" & _cfg.FtpFolder
    End Function

    Private Function LoadDb() As Dictionary(Of String, String)
        Dim d As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Dim header As String = DbHeader()
        Dim ok As Boolean = False
        If File.Exists(DbPath) Then
            Dim lines As String() = File.ReadAllLines(DbPath, Encoding.UTF8)
            If lines.Length > 0 AndAlso lines(0) = header Then
                ok = True
                For i As Integer = 1 To lines.Length - 1
                    Dim p As Integer = lines(i).IndexOf("|"c)
                    If p > 0 Then d(lines(i).Substring(0, p)) = lines(i).Substring(p + 1)
                Next
            End If
        End If
        If Not ok Then File.WriteAllText(DbPath, header & Environment.NewLine, Encoding.UTF8) ' alt server/folder -> baza noua
        Return d
    End Function

    Private Sub ScanFiles()
        Dim db As Dictionary(Of String, String) = LoadDb()
        Dim root As String = _cfg.LocalFolder.TrimEnd("\"c)
        For Each f As String In Directory.GetFiles(root, "*", SearchOption.AllDirectories)
            Dim fi As New FileInfo(f)
            Dim it As New FileItem()
            it.Full = f
            it.Rel = f.Substring(root.Length).TrimStart("\"c).Replace("\"c, "/"c)
            it.Size = fi.Length
            it.Ticks = fi.LastWriteTimeUtc.Ticks
            Dim sig As String = Nothing
            If Not db.TryGetValue(it.Rel, sig) OrElse sig <> it.Signature() Then _pending.Add(it)
        Next
    End Sub

    ' ---------- butoane ----------
    Private Sub btnSettings_Click(sender As Object, e As EventArgs) Handles btnSettings.Click
        tmrClose.Stop()
        Using f As New SettingsForm(AppSettings.Load())
            If f.ShowDialog(Me) = DialogResult.OK Then InitState()
        End Using
    End Sub

    Private Sub btnUpload_Click(sender As Object, e As EventArgs) Handles btnUpload.Click
        tmrClose.Stop()
        StartUpload()
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        If _busy Then
            AskCancel()
        Else
            Me.Close()
        End If
    End Sub

    Private Sub AskCancel()
        If _confirming Then Return
        _confirming = True
        Try
            If MessageBox.Show("Renuntati la incarcarea fisierelor?", "Anulare", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                _closeAfter = True
                btnExit.Enabled = False
                lblStatus.Text = "Se anuleaza transferul..."
                bw.CancelAsync()
            End If
        Finally
            _confirming = False
        End Try
    End Sub

    Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If _busy Then
            e.Cancel = True
            AskCancel()
        End If
    End Sub

    Private Sub tmrClose_Tick(sender As Object, e As EventArgs) Handles tmrClose.Tick
        tmrClose.Stop()
        Me.Close()
    End Sub

    ' ---------- transfer ----------
    Private Sub StartUpload()
        If _pending.Count = 0 OrElse _busy Then Return
        _busy = True
        _closeAfter = False
        _totalBytes = 0
        For Each it As FileItem In _pending
            _totalBytes += it.Size
        Next
        If _totalBytes < 1 Then _totalBytes = 1
        _doneBytes = 0
        _fileCount = _pending.Count
        _lastPct = -1
        _lastError = ""
        _madeDirs.Clear()
        pb.Value = 0
        btnUpload.Enabled = False
        btnSettings.Enabled = False
        btnExit.Text = "Anulare"
        lblStatus.Text = "Se transfera fisierul 1 din " & _pending.Count
        bw.RunWorkerAsync()
    End Sub

    Private Sub bw_DoWork(sender As Object, e As DoWorkEventArgs) Handles bw.DoWork
        Dim ok As Integer = 0, fail As Integer = 0
        Dim failed As New List(Of FileItem)()
        For i As Integer = 0 To _pending.Count - 1
            If bw.CancellationPending Then e.Cancel = True : Exit For
            Dim it As FileItem = _pending(i)
            _curIndex = i + 1
            Dim baseBytes As Long = _doneBytes
            _lastPct = -1
            ReportPct()
            Try
                EnsureRemoteDirs(it)
                UploadOne(it)
                AppendDb(it)
                ok += 1
            Catch ex As OperationCanceledException
                FtpHelper.DeleteFile(_cfg, _cfg.FtpFolder & "/" & it.Rel)
                e.Cancel = True
                Exit For
            Catch ex As Exception
                fail += 1
                _lastError = ex.Message
                failed.Add(it)
            End Try
            _doneBytes = baseBytes + it.Size
        Next
        _pending = failed
        e.Result = New Integer() {ok, fail}
    End Sub

    Private Sub ReportPct()
        Dim pct As Integer = CInt(Math.Min(100L, (_doneBytes * 100L) \ _totalBytes))
        If pct <> _lastPct Then
            _lastPct = pct
            bw.ReportProgress(pct, _curIndex)
        End If
    End Sub

    Private Sub EnsureRemoteDirs(it As FileItem)
        Dim parts As String() = it.Rel.Split("/"c)
        Dim acc As String = _cfg.FtpFolder.TrimEnd("/"c)
        For i As Integer = 0 To parts.Length - 2
            acc &= "/" & parts(i)
            If _madeDirs.Add(acc) Then FtpHelper.MakeDir(_cfg, acc)
        Next
    End Sub

    Private Sub UploadOne(it As FileItem)
        Dim url As String = FtpHelper.BuildUrl(_cfg.Server, _cfg.FtpFolder & "/" & it.Rel)
        Dim req As FtpWebRequest = FtpHelper.CreateRequest(url, WebRequestMethods.Ftp.UploadFile, _cfg)
        Using fs As New FileStream(it.Full, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            req.ContentLength = fs.Length
            Using rs As Stream = req.GetRequestStream()
                Dim buf(32767) As Byte
                Dim n As Integer = fs.Read(buf, 0, buf.Length)
                Do While n > 0
                    If bw.CancellationPending Then Throw New OperationCanceledException()
                    rs.Write(buf, 0, n)
                    _doneBytes += n
                    ReportPct()
                    n = fs.Read(buf, 0, buf.Length)
                Loop
            End Using
        End Using
        Using resp As FtpWebResponse = CType(req.GetResponse(), FtpWebResponse)
        End Using
    End Sub

    Private Sub AppendDb(it As FileItem)
        File.AppendAllText(DbPath, it.Rel & "|" & it.Signature() & Environment.NewLine, Encoding.UTF8)
    End Sub

    Private Sub bw_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bw.ProgressChanged
        pb.Value = Math.Max(0, Math.Min(100, e.ProgressPercentage))
        If Not _closeAfter Then
            lblStatus.Text = "Se transfera fisierul " & CInt(e.UserState) & " din " & _fileCount
        End If
    End Sub


    Private Sub bw_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles bw.RunWorkerCompleted
        _busy = False
        btnExit.Enabled = True
        btnExit.Text = "Iesire"
        btnSettings.Enabled = True

        If _closeAfter Then
            Me.Close()
            Return
        End If

        If e.Error IsNot Nothing Then
            lblStatus.Text = "Eroare: " & e.Error.Message
            btnUpload.Enabled = True
            Return
        End If

        If e.Cancelled Then
            lblStatus.Text = "Transfer anulat."
            InitState()
            Return
        End If

        Dim r As Integer() = CType(e.Result, Integer())
        pb.Value = 100
        Dim txt As String = "Au fost transferate " & r(0) & " fisiere"
        If r(1) > 0 Then txt &= " (" & r(1) & " fisiere cu erori: " & _lastError & ")"
        lblStatus.Text = txt
        btnUpload.Enabled = (r(1) > 0)

        If _cfg.AutoAction Then tmrClose.Start() ' inchidere dupa 30 secunde
    End Sub
End Class

' ------------------------------------------------------------
'  Formular Setari
' ------------------------------------------------------------
Public Class SettingsForm
    Inherits Form

    Private WithEvents txtServer As New TextBox()
    Private WithEvents txtUser As New TextBox()
    Private WithEvents txtPass As New TextBox()
    Private WithEvents btnTest As New Button()
    Private lblTest As New Label()
    Private txtFtpFolder As New TextBox()
    Private WithEvents btnBrowseFtp As New Button()
    Private chkAuto As New CheckBox()
    Private WithEvents btnWpUpdate As New Button()
    Private WithEvents btnSave As New Button()
    Private WithEvents btnCancel As New Button()

    Private _cfg As AppSettings

    Public Sub New(cfg As AppSettings)
        _cfg = cfg
        Me.Text = "Setari"
        Me.ClientSize = New Size(500, 300)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        AddLabel("Server FTP:", 15, 18)
        txtServer.SetBounds(110, 15, 378, 22)

        AddLabel("Utilizator:", 15, 53)
        txtUser.SetBounds(110, 50, 170, 22)

        AddLabel("Parola:", 15, 88)
        txtPass.SetBounds(110, 85, 170, 22)
        txtPass.UseSystemPasswordChar = True

        btnTest.Text = "Test"
        btnTest.SetBounds(300, 65, 90, 32)
        lblTest.SetBounds(395, 65, 100, 32)
        lblTest.AutoSize = False
        lblTest.TextAlign = ContentAlignment.MiddleLeft

        AddLabel("Folder FTP:", 15, 128)
        txtFtpFolder.SetBounds(110, 125, 340, 22)
        btnBrowseFtp.Text = "..."
        btnBrowseFtp.SetBounds(458, 124, 30, 24)

        chkAuto.Text = "Actiune automata"
        chkAuto.Checked = False
        chkAuto.SetBounds(110, 160, 200, 22)

        btnWpUpdate.Text = "Actualizare tema WP si Module"
        btnWpUpdate.SetBounds(110, 195, 260, 32)

        btnSave.Text = "Salvare"
        btnSave.SetBounds(300, 255, 90, 32)
        btnCancel.Text = "Renunta"
        btnCancel.SetBounds(398, 255, 90, 32)

        Me.Controls.AddRange(New Control() {txtServer, txtUser, txtPass, btnTest, lblTest,
                                            txtFtpFolder, btnBrowseFtp, chkAuto, btnWpUpdate, btnSave, btnCancel})

        ' valori existente
        txtServer.Text = _cfg.Server
        txtUser.Text = _cfg.User
        txtPass.Text = _cfg.Password
        txtFtpFolder.Text = _cfg.FtpFolder
        chkAuto.Checked = _cfg.AutoAction

        txtFtpFolder.ReadOnly = True
        SetFtpPartEnabled(False) ' se activeaza doar dupa Test OK
    End Sub

    Private Sub AddLabel(text As String, x As Integer, y As Integer)
        Dim l As New Label()
        l.Text = text
        l.SetBounds(x, y, 90, 20)
        Me.Controls.Add(l)
    End Sub

    Private Sub SetFtpPartEnabled(en As Boolean)
        txtFtpFolder.Enabled = en
        btnBrowseFtp.Enabled = en
        btnWpUpdate.Enabled = en
        btnSave.Enabled = en
        If Not en Then lblTest.Text = ""
    End Sub

    Private Function CurrentCfg() As AppSettings
        Dim c As New AppSettings()
        c.Server = txtServer.Text.Trim()
        c.User = txtUser.Text.Trim()
        c.Password = txtPass.Text
        c.FtpFolder = If(txtFtpFolder.Text.Trim() = "", "/", txtFtpFolder.Text.Trim())
        c.AutoAction = chkAuto.Checked
        Return c
    End Function

    ' orice schimbare de server/user/parola invalideaza testul
    Private Sub CredChanged(sender As Object, e As EventArgs) Handles txtServer.TextChanged, txtUser.TextChanged, txtPass.TextChanged
        SetFtpPartEnabled(False)
    End Sub

    Private Sub btnTest_Click(sender As Object, e As EventArgs) Handles btnTest.Click
        If txtServer.Text.Trim() = "" OrElse txtUser.Text.Trim() = "" Then
            lblTest.ForeColor = Color.Firebrick
            lblTest.Text = "Completati serverul si utilizatorul"
            Return
        End If
        Me.Cursor = Cursors.WaitCursor
        Try
            FtpHelper.TestConnection(CurrentCfg())
            SetFtpPartEnabled(True)
            lblTest.ForeColor = Color.ForestGreen
            lblTest.Text = "Test OK"
        Catch ex As Exception
            SetFtpPartEnabled(False)
            lblTest.ForeColor = Color.Firebrick
            lblTest.Text = "Eroare"
            MessageBox.Show("Conectare esuata: " & ex.Message, "Test FTP", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub btnBrowseFtp_Click(sender As Object, e As EventArgs) Handles btnBrowseFtp.Click
        Using f As New FtpBrowseForm(CurrentCfg(), txtFtpFolder.Text)
            If f.ShowDialog(Me) = DialogResult.OK Then txtFtpFolder.Text = f.SelectedFolder
        End Using
    End Sub

    ' folderul parinte al folderului FTP (.../wp-content/uploads -> .../wp-content)
    Private Shared Function ParentFolder(folder As String) As String
        Dim t As String = folder.TrimEnd("/"c)
        Dim idx As Integer = t.LastIndexOf("/"c)
        Return If(idx <= 0, "/", t.Substring(0, idx))
    End Function

    Private Sub btnWpUpdate_Click(sender As Object, e As EventArgs) Handles btnWpUpdate.Click
        Dim c As AppSettings = CurrentCfg()
        Dim wpLocal As String = AppSettings.WpContentFolder
        Dim remote As String = ParentFolder(c.FtpFolder)

        If Not Directory.Exists(wpLocal) Then
            MessageBox.Show("Folderul " & wpLocal & " nu exista.", "Actualizare", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim warn As String = ""
        If Not remote.TrimEnd("/"c).EndsWith("/wp-content", StringComparison.OrdinalIgnoreCase) Then
            warn = Environment.NewLine & Environment.NewLine & "ATENTIE: folderul tinta nu pare sa fie wp-content!"
        End If
        If MessageBox.Show("Se vor incarca folderele themes si plugins din" & Environment.NewLine & wpLocal & Environment.NewLine &
                           "in " & remote & " pe server, inlocuind fisierele existente." & warn & Environment.NewLine & Environment.NewLine &
                           "Continuati?", "Actualizare tema WP si Module", MessageBoxButtons.YesNo,
                           If(warn = "", MessageBoxIcon.Question, MessageBoxIcon.Warning)) <> DialogResult.Yes Then Return

        Using f As New WpUpdateForm(c, wpLocal, remote)
            f.ShowDialog(Me)
        End Using
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim c As AppSettings = CurrentCfg()

        ' creeaza uploads (cu toate subfolderele) si wp-content
        Try
            Directory.CreateDirectory(AppSettings.AppDir)
            LocalSetup.EnsureStructure(c.LocalFolder)
            Directory.CreateDirectory(AppSettings.WpContentFolder)
        Catch ex As Exception
            MessageBox.Show("Nu s-au putut crea folderele din Documente: " & ex.Message, "Setari", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try

        ' scurtatura pe Desktop catre folderul local
        Dim scutErr As String = ""
        Try
            LocalSetup.CreateDesktopShortcut(c.LocalFolder)
        Catch ex As Exception
            scutErr = ex.Message
        End Try

        _cfg.Server = c.Server
        _cfg.User = c.User
        _cfg.Password = c.Password
        _cfg.FtpFolder = c.FtpFolder
        _cfg.AutoAction = c.AutoAction
        Try
            _cfg.Save()
        Catch ex As Exception
            MessageBox.Show("Nu s-a putut scrie setari.ini: " & ex.Message, "Setari", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try
        If scutErr <> "" Then
            MessageBox.Show("Setarile au fost salvate, dar scurtatura de pe Desktop nu a putut fi creata: " & scutErr, "Setari", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class

' ------------------------------------------------------------
'  Browse folder pe serverul FTP
' ------------------------------------------------------------
Public Class FtpBrowseForm
    Inherits Form

    Private lblPath As New Label()
    Private WithEvents lst As New ListBox()
    Private WithEvents btnOk As New Button()
    Private WithEvents btnCancel As New Button()
    Private _cfg As AppSettings
    Private _cur As String

    Public SelectedFolder As String = "/"

    Public Sub New(cfg As AppSettings, startFolder As String)
        _cfg = cfg
        _cur = If(String.IsNullOrWhiteSpace(startFolder), "/", startFolder)
        Me.Text = "Folder pe server FTP"
        Me.ClientSize = New Size(400, 340)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        lblPath.SetBounds(15, 12, 370, 20)
        lst.SetBounds(15, 38, 370, 250)
        btnOk.Text = "Selectare"
        btnOk.SetBounds(205, 298, 85, 30)
        btnCancel.Text = "Renunta"
        btnCancel.SetBounds(300, 298, 85, 30)
        Me.Controls.AddRange(New Control() {lblPath, lst, btnOk, btnCancel})
    End Sub

    Private Sub FtpBrowseForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        LoadDir()
    End Sub

    Private Sub LoadDir()
        Me.Cursor = Cursors.WaitCursor
        Try
            Dim dirs As List(Of String) = FtpHelper.ListDirs(_cfg, _cur)
            lst.Items.Clear()
            If _cur <> "/" Then lst.Items.Add("..")
            For Each d As String In dirs
                lst.Items.Add(d)
            Next
            lblPath.Text = _cur
        Catch ex As Exception
            If _cur <> "/" Then
                _cur = "/"
                LoadDir()
            Else
                MessageBox.Show("Nu se poate citi folderul: " & ex.Message, "FTP", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub lst_DoubleClick(sender As Object, e As EventArgs) Handles lst.DoubleClick
        If lst.SelectedItem Is Nothing Then Return
        Dim name As String = lst.SelectedItem.ToString()
        If name = ".." Then
            Dim t As String = _cur.TrimEnd("/"c)
            Dim idx As Integer = t.LastIndexOf("/"c)
            _cur = If(idx <= 0, "/", t.Substring(0, idx))
        Else
            _cur = _cur.TrimEnd("/"c) & "/" & name
        End If
        LoadDir()
    End Sub

    Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click
        SelectedFolder = _cur
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class

' ------------------------------------------------------------
'  Structura de foldere locale + scurtatura Desktop
' ------------------------------------------------------------
Public Class LocalSetup
    ' foldere/subfoldere create in "uploads" (din lista foldere1.txt)
    Public Shared ReadOnly Folders As String() = New String() {
        "formulare",
        "fotografii",
        "monitorul-oficial-local",
        "formulare\diverse-si-544",
        "formulare\registru-agricol",
        "formulare\stare-civila",
        "formulare\taxe-si-impozite",
        "formulare\urbanism",
        "monitorul-oficial-local\Alte documente",
        "monitorul-oficial-local\Dispozitiile Autoritatii Executive",
        "monitorul-oficial-local\Documente si Informatii Financiare",
        "monitorul-oficial-local\Hotararile Autoritatii Deliberative",
        "monitorul-oficial-local\Regulamentele privind Procedurile Administrative",
        "monitorul-oficial-local\Statutul Unitatii Administrativ-Teritoriale",
        "monitorul-oficial-local\Alte documente\Alte tipuri de documente",
        "monitorul-oficial-local\Alte documente\Declaratii de avere",
        "monitorul-oficial-local\Alte documente\Declaratii de casatorie",
        "monitorul-oficial-local\Alte documente\Declaratii de interese",
        "monitorul-oficial-local\Alte documente\Minutele sedintelor",
        "monitorul-oficial-local\Alte documente\Probleme de interes public care urmeaza sa fie dezbatute",
        "monitorul-oficial-local\Alte documente\Procese verbale ale sedintelor autoritatii deliberative",
        "monitorul-oficial-local\Alte documente\Proiecte de acte normative",
        "monitorul-oficial-local\Alte documente\Registru pentru propuneri, sugestii, opinii la proiectele de hotarare",
        "monitorul-oficial-local\Alte documente\Registrul refuzurilor de a semna-contrasemna-aviza acte administrative",
        "monitorul-oficial-local\Alte documente\Alte tipuri de documente\Concursuri",
        "monitorul-oficial-local\Alte documente\Alte tipuri de documente\Convocator sedinta",
        "monitorul-oficial-local\Alte documente\Alte tipuri de documente\Licitatii",
        "monitorul-oficial-local\Alte documente\Alte tipuri de documente\Oferte vanzare teren",
        "monitorul-oficial-local\Alte documente\Alte tipuri de documente\Ziar local",
        "monitorul-oficial-local\Dispozitiile Autoritatii Executive\Arhiva Dispozitii autoritatii executive",
        "monitorul-oficial-local\Dispozitiile Autoritatii Executive\Arhiva Proiecte de dispozitii ale autoritatii executive",
        "monitorul-oficial-local\Dispozitiile Autoritatii Executive\Dispozitiile Autoritatii Executive",
        "monitorul-oficial-local\Dispozitiile Autoritatii Executive\Proiectele de dispozitii ale autoritatii executive",
        "monitorul-oficial-local\Documente si Informatii Financiare\Bilant contabil",
        "monitorul-oficial-local\Documente si Informatii Financiare\Buget local pe surse de finantare",
        "monitorul-oficial-local\Documente si Informatii Financiare\Datoria publica locala",
        "monitorul-oficial-local\Documente si Informatii Financiare\Proiect buget",
        "monitorul-oficial-local\Hotararile Autoritatii Deliberative\Arhiva Hotarari Autoritatii Deliberative",
        "monitorul-oficial-local\Hotararile Autoritatii Deliberative\Arhiva Proiecte de hotarari ale Autoritatii Deliberative",
        "monitorul-oficial-local\Hotararile Autoritatii Deliberative\Hotararile Autoritatii Deliberative",
        "monitorul-oficial-local\Hotararile Autoritatii Deliberative\Proiecte de hotarari ale Autoritatii Deliberative"
    }

    ' Creeaza folderul radacina (ex. D:\uploads) si toate subfolderele
    Public Shared Sub EnsureStructure(root As String)
        Directory.CreateDirectory(root)
        For Each f As String In Folders
            Directory.CreateDirectory(Path.Combine(root, f))
        Next
    End Sub

    ' Creeaza (sau rescrie) o scurtatura .lnk
    Public Shared Sub MakeShortcut(lnkPath As String, target As String, workDir As String, desc As String, icon As String)
        Dim t As Type = Type.GetTypeFromProgID("WScript.Shell")
        If t Is Nothing Then Throw New InvalidOperationException("WScript.Shell nu este disponibil.")
        Dim sh As Object = Activator.CreateInstance(t)
        Dim lnk As Object = t.InvokeMember("CreateShortcut", BindingFlags.InvokeMethod, Nothing, sh, New Object() {lnkPath})
        Dim lt As Type = lnk.GetType()
        lt.InvokeMember("TargetPath", BindingFlags.SetProperty, Nothing, lnk, New Object() {target})
        lt.InvokeMember("WorkingDirectory", BindingFlags.SetProperty, Nothing, lnk, New Object() {workDir})
        lt.InvokeMember("Description", BindingFlags.SetProperty, Nothing, lnk, New Object() {desc})
        If icon <> "" Then lt.InvokeMember("IconLocation", BindingFlags.SetProperty, Nothing, lnk, New Object() {icon})
        lt.InvokeMember("Save", BindingFlags.InvokeMethod, Nothing, lnk, Nothing)
    End Sub

    ' Intinta unei scurtaturi existente ("" daca nu se poate citi)
    Public Shared Function ShortcutTarget(lnkPath As String) As String
        Try
            Dim t As Type = Type.GetTypeFromProgID("WScript.Shell")
            Dim sh As Object = Activator.CreateInstance(t)
            Dim lnk As Object = t.InvokeMember("CreateShortcut", BindingFlags.InvokeMethod, Nothing, sh, New Object() {lnkPath})
            Dim lt As Type = lnk.GetType()
            Return Convert.ToString(lt.InvokeMember("TargetPath", BindingFlags.GetProperty, Nothing, lnk, Nothing))
        Catch
            Return ""
        End Try
    End Function

    ' Scurtatura pe Desktop catre folderul local
    Public Shared Sub CreateDesktopShortcut(target As String)
        Dim desk As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        Dim icon As String = Path.Combine(Environment.SystemDirectory, "imageres.dll") & ",68"
        MakeShortcut(Path.Combine(desk, "uploads FTP.lnk"), target, target, "Folder local FTP Upload", icon)
    End Sub

    ' Scurtatura pe Desktop catre FTP_Upload.exe (pornit din folderul lui, unde sunt setari.ini si incarcate.db)
    Public Shared Sub EnsureAppShortcut()
        Dim desk As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        Dim lnkPath As String = Path.Combine(desk, "FTP Upload.lnk")
        Dim exe As String = Application.ExecutablePath
        If File.Exists(lnkPath) AndAlso String.Equals(ShortcutTarget(lnkPath), exe, StringComparison.OrdinalIgnoreCase) Then Return
        MakeShortcut(lnkPath, exe, Application.StartupPath, "FTP Upload", exe & ",0")
    End Sub
End Class

' ------------------------------------------------------------
'  Actualizare tema WP si module (themes + plugins -> wp-content)
' ------------------------------------------------------------
Public Class WpUpdateForm
    Inherits Form

    Private lblStatus As New Label()
    Private pb As New ProgressBar()
    Private WithEvents btnClose As New Button()
    Private WithEvents bw As New BackgroundWorker()

    Private _cfg As AppSettings
    Private _local As String
    Private _remote As String
    Private _files As New List(Of FileItem)()
    Private _busy As Boolean = False
    Private _closeAfter As Boolean = False
    Private _confirming As Boolean = False
    Private _total As Long = 1
    Private _done As Long = 0
    Private _idx As Integer = 0
    Private _lastPct As Integer = -1
    Private _lastError As String = ""
    Private _made As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

    Public Sub New(cfg As AppSettings, localWpContent As String, remoteWpContent As String)
        _cfg = cfg
        _local = localWpContent.TrimEnd("\"c)
        _remote = remoteWpContent
        Me.Text = "Actualizare tema WP si Module"
        Me.ClientSize = New Size(520, 140)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        lblStatus.SetBounds(15, 15, 490, 40)
        lblStatus.AutoSize = False
        pb.SetBounds(15, 60, 490, 22)
        btnClose.Text = "Anulare"
        btnClose.SetBounds(395, 95, 110, 32)
        Me.Controls.AddRange(New Control() {lblStatus, pb, btnClose})

        bw.WorkerReportsProgress = True
        bw.WorkerSupportsCancellation = True
    End Sub

    Private Sub WpUpdateForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Try
            For Each sub1 As String In New String() {"themes", "plugins"}
                Dim d As String = Path.Combine(_local, sub1)
                If Directory.Exists(d) Then
                    For Each f As String In Directory.GetFiles(d, "*", SearchOption.AllDirectories)
                        Dim it As New FileItem()
                        it.Full = f
                        it.Rel = f.Substring(_local.Length).TrimStart("\"c).Replace("\"c, "/"c)
                        it.Size = New FileInfo(f).Length
                        _files.Add(it)
                        _total += it.Size
                    Next
                End If
            Next
        Catch ex As Exception
            lblStatus.Text = "Eroare la citirea folderului local: " & ex.Message
            btnClose.Text = "Inchide"
            Return
        End Try

        If _files.Count = 0 Then
            lblStatus.Text = "Nu exista fisiere in subfolderele themes si plugins din " & _local
            btnClose.Text = "Inchide"
            Return
        End If

        _busy = True
        lblStatus.Text = "Se transfera fisierul 1 din " & _files.Count
        bw.RunWorkerAsync()
    End Sub

    Private Sub bw_DoWork(sender As Object, e As DoWorkEventArgs) Handles bw.DoWork
        Dim ok As Integer = 0, fail As Integer = 0
        For i As Integer = 0 To _files.Count - 1
            If bw.CancellationPending Then e.Cancel = True : Exit For
            Dim it As FileItem = _files(i)
            _idx = i + 1
            Dim baseBytes As Long = _done
            _lastPct = -1
            ReportPct()
            Try
                EnsureDirs(it)
                FtpHelper.UploadFile(_cfg, it.Full, _remote.TrimEnd("/"c) & "/" & it.Rel, AddressOf OnBytes, AddressOf IsCancelled)
                ok += 1
            Catch ex As OperationCanceledException
                e.Cancel = True
                Exit For
            Catch ex As Exception
                fail += 1
                _lastError = ex.Message
            End Try
            _done = baseBytes + it.Size
        Next
        e.Result = New Integer() {ok, fail}
    End Sub

    Private Sub OnBytes(n As Integer)
        _done += n
        ReportPct()
    End Sub

    Private Function IsCancelled() As Boolean
        Return bw.CancellationPending
    End Function

    Private Sub ReportPct()
        Dim pct As Integer = CInt(Math.Min(100L, (_done * 100L) \ _total))
        If pct <> _lastPct Then
            _lastPct = pct
            bw.ReportProgress(pct, _idx)
        End If
    End Sub

    Private Sub EnsureDirs(it As FileItem)
        Dim parts As String() = it.Rel.Split("/"c)
        Dim acc As String = _remote.TrimEnd("/"c)
        For i As Integer = 0 To parts.Length - 2
            acc &= "/" & parts(i)
            If _made.Add(acc) Then FtpHelper.MakeDir(_cfg, acc)
        Next
    End Sub

    Private Sub bw_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bw.ProgressChanged
        pb.Value = Math.Max(0, Math.Min(100, e.ProgressPercentage))
        If Not _closeAfter Then
            lblStatus.Text = "Se transfera fisierul " & CInt(e.UserState) & " din " & _files.Count
        End If
    End Sub

    Private Sub bw_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles bw.RunWorkerCompleted
        _busy = False
        btnClose.Enabled = True
        btnClose.Text = "Inchide"
        If _closeAfter Then
            Me.Close()
            Return
        End If
        If e.Error IsNot Nothing Then
            lblStatus.Text = "Eroare: " & e.Error.Message
            Return
        End If
        If e.Cancelled Then
            lblStatus.Text = "Transfer anulat. Reluati actualizarea pentru a completa fisierele."
            Return
        End If
        Dim r As Integer() = CType(e.Result, Integer())
        pb.Value = 100
        Dim txt As String = "Au fost transferate " & r(0) & " fisiere"
        If r(1) > 0 Then txt &= " (" & r(1) & " fisiere cu erori: " & _lastError & ")"
        lblStatus.Text = txt
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        If _busy Then
            AskCancel()
        Else
            Me.Close()
        End If
    End Sub

    Private Sub AskCancel()
        If _confirming Then Return
        _confirming = True
        Try
            If MessageBox.Show("Renuntati la actualizare?", "Anulare", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                _closeAfter = True
                btnClose.Enabled = False
                lblStatus.Text = "Se anuleaza transferul..."
                bw.CancelAsync()
            End If
        Finally
            _confirming = False
        End Try
    End Sub

    Private Sub WpUpdateForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If _busy Then
            e.Cancel = True
            AskCancel()
        End If
    End Sub
End Class
