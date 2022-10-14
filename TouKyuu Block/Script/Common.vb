Imports System.Console
Imports System.ConsoleColor
Imports System.Diagnostics.Process
Imports System.IO
Imports System.IO.Directory
Imports System.Math
Imports System.Net
Imports System.Runtime.CompilerServices
Imports System.Threading.Thread
Imports System.Windows.Forms
Imports System.Windows.Forms.Application
Imports System.Windows.Forms.MessageBox
Imports System.Windows.Forms.MessageBoxButtons
Imports System.Windows.Forms.MessageBoxDefaultButton
Imports System.Windows.Forms.MessageBoxIcon

Friend Module Common
#Region "Helper"
    ''' <summary>
    ''' Check internet connection.
    ''' </summary>
    ''' <returns>Connection state.</returns>
    Private Function IsNetAvail()
        Dim objResp As WebResponse
        Try
            objResp = WebRequest.Create(New Uri(My.Resources.link_base)).GetResponse
            objResp.Close()
            objResp = Nothing
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Check update.
    ''' </summary>
    Private Sub ChkUpd()
        HdrSty("アップデートの確認...")
        If IsNetAvail() AndAlso Not (New WebClient).DownloadString(My.Resources.link_ver).Contains(My.Resources.app_ver) Then
            Show($"「{My.Resources.app_true_name}」新しいバージョンが利用可能！", "更新", OK, Information)
            Run(New FrmUpdate)
        End If
    End Sub

    ''' <summary>
    ''' Update valid license
    ''' </summary>
    Friend Sub UpdVldLic()
        My.Settings.Chk_Key = True
        My.Settings.Save()
    End Sub

    ''' <summary>
    ''' Fade in form
    ''' </summary>
    <Extension()>
    Friend Sub FIFrm(frm As Form)
        While frm.Opacity < 1
            frm.Opacity += 0.05
            frm.Update()
            Sleep(10)
        End While
    End Sub

    ''' <summary>
    ''' Fade out form
    ''' </summary>
    <Extension()>
    Friend Sub FOFrm(frm As Form)
        While frm.Opacity > 0
            frm.Opacity -= 0.05
            frm.Update()
            Sleep(10)
        End While
    End Sub
#End Region

#Region "Master"
    ''' <summary>
    ''' End process.
    ''' </summary>
    ''' <param name="name">Process name.</param>
    Friend Sub KillPrcs(name As String)
        If GetProcessesByName(name).Count > 0 Then
            For Each item In GetProcessesByName(name)
                item.Kill()
            Next
        End If
    End Sub

    ''' <summary>
    ''' Run application.
    ''' </summary>
    Private Sub RunApp()
        ' w
        Dim w = 0D
        Dim sW = "W = "
        For i = 0 To Integer.MaxValue
            Intro()
            If i > 0 Then
                HxSty(sW & vbCrLf)
                Dim wi = InpG(vbCrLf & vbTab & $"w{i + 1} = ")
                If wi > 0 Then
                    w += wi
                    sW += $" + {wi}"
                Else
                    Exit For
                End If
            Else
                Dim wi = InpG(vbTab & $"w{i + 1} = ")
                If wi > 0 Then
                    w += wi
                    sW += wi.ToString()
                Else
                    i = -1
                End If
            End If
        Next
        ' h
        Dim h = 0D
        Dim sH = "H = "
        For i = 0 To Integer.MaxValue
            Intro()
            HxSty(sW & vbCrLf)
            If i > 0 Then
                WriteLine(sH)
                Dim hi = InpG(vbCrLf & vbTab & $"h{i + 1} = ")
                If hi > 0 Then
                    h += hi
                    sH += $" + {hi}"
                Else
                    Exit For
                End If
            Else
                Dim hi = InpG(vbCrLf & vbTab & $"h{i + 1} = ")
                If hi > 0 Then
                    h += hi
                    sH += hi.ToString()
                Else
                    i = -1
                End If
            End If
        Next
        ' Process
        Dim c = Ceiling((w + h) * 2 / Pow(10, 3))
        Dim s = Ceiling(w * h / Pow(10, 6))
        Dim block = c + s
        For i = 10 To Integer.MaxValue
            If (block + i) Mod 30 = 0 Then
                block += i
                Exit For
            End If
        Next
        ' Output
        Dim fmt = FmtNo(w, h, c, s, block)
        Intro()
        HxSty(vbTab & "Ｗ (mm)" & vbTab & vbTab & ": " + String.Format(fmt, w) & vbCrLf)
        HxSty(vbTab & "Ｈ (mm)" & vbTab & vbTab & ": " + String.Format(fmt, h) & vbCrLf)
        RsltSty(vbTab & "Ｃ (m)" & vbTab & vbTab & ": " + String.Format(fmt, c) & vbCrLf)
        RsltSty(vbTab & "Ｓ (m²)" & vbTab & vbTab & ": " + String.Format(fmt, s) & vbCrLf)
        RsltSty(vbTab & "ブロック (個)" & vbTab & ": " + String.Format(fmt, block) & vbCrLf)
        Credit()
    End Sub

    ''' <summary>
    ''' First run.
    ''' </summary>
    Friend Sub FstRunApp()
        ChkUpd()
        RunApp()
    End Sub
#End Region

#Region "Main"
    ''' <summary>
    ''' Create directory advanced.
    ''' </summary>
    ''' <param name="path">Folder path.</param>
    Friend Sub CrtDirAdv(path As String)
        If Not Exists(path) Then
            CreateDirectory(path)
        End If
    End Sub

    ''' <summary>
    ''' Delete file advanced.
    ''' </summary>
    ''' <param name="path">File path.</param>
    Friend Sub DelFileAdv(path As String)
        If File.Exists(path) Then
            File.Delete(path)
        End If
    End Sub

    ''' <summary>
    ''' Convert to G.
    ''' </summary>
    ''' <param name="num">Number.</param>
    ''' <returns>Number converted.</returns>
    Private Function ConvertToG(num As Double)
        Return If(num < 30, num * 910, num)
    End Function

    ''' <summary>
    ''' Format number.
    ''' </summary>
    ''' <param name="w">Width.</param>
    ''' <param name="h">Heigh.</param>
    ''' <param name="c">Circuit.</param>
    ''' <param name="s">Spread.</param>
    ''' <param name="block">Block.</param>
    ''' <returns>Format.</returns>
    Private Function FmtNo(w As Double, h As Double, c As Double, s As Double, block As Double)
        Dim wSize = w.ToString().Length
        Dim hSize = h.ToString().Length
        Dim cSize = c.ToString().Length
        Dim sSize = s.ToString().Length
        Dim blockSize = block.ToString().Length
        Dim maxSize = Max(Max(Max(wSize, hSize), Max(cSize, sSize)), blockSize)
        Return "{0," + maxSize.ToString() + ":####.#}"
    End Function
#End Region

#Region "Timer"
    ''' <summary>
    ''' Start timer advanced.
    ''' </summary>
    <Extension()>
    Friend Sub StrtAdv(tmr As Timer)
        If Not tmr.Enabled Then
            tmr.Start()
        End If
    End Sub

    ''' <summary>
    ''' Stop timer advanced.
    ''' </summary>
    <Extension()>
    Friend Sub StopAdv(tmr As Timer)
        If tmr.Enabled Then
            tmr.Start()
        End If
    End Sub
#End Region

#Region "Actor"
    ''' <summary>
    ''' Header style.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    Private Sub HdrSty(caption As String)
        ForegroundColor = DarkYellow
        Write(caption)
    End Sub

    ''' <summary>
    ''' Intro style.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    Private Sub IntroSty(caption As String)
        ForegroundColor = Blue
        Write(caption)
    End Sub

    ''' <summary>
    ''' Credit style.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    Private Sub CreditSty(caption As String)
        ForegroundColor = DarkGray
        Write(caption)
    End Sub

    ''' <summary>
    ''' Title style.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    Private Sub TitSty(caption As String)
        ForegroundColor = Green
        Write(caption)
    End Sub

    ''' <summary>
    ''' Input style.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    Private Sub InpSty(caption As String)
        ForegroundColor = Cyan
        Write(caption)
    End Sub

    ''' <summary>
    ''' History style.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    Private Sub HxSty(caption As String)
        ForegroundColor = Gray
        Write(caption)
    End Sub

    ''' <summary>
    ''' Result style.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    Private Sub RsltSty(caption As String)
        ForegroundColor = DarkCyan
        Write(caption)
    End Sub

    ''' <summary>
    ''' Error style.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    Friend Sub ErrSty(caption As String)
        ForegroundColor = Red
        Write(caption)
    End Sub

    ''' <summary>
    ''' Prefix input.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    Private Sub PrefInp(caption As String)
        InpSty(caption)
        ForegroundColor = White
    End Sub

    ''' <summary>
    ''' Intro.
    ''' </summary>
    Private Sub Intro()
        Clear()
        IntroSty(My.Resources.gr_name & vbCrLf)
        IntroSty(My.Resources.cc_text & vbCrLf)
        TitSty(vbCrLf & My.Resources.app_true_name & vbCrLf & vbCrLf)
    End Sub

    ''' <summary>
    ''' Credit.
    ''' </summary>
    Private Sub Credit()
        CreditSty(vbCrLf & "続行するには、任意のキーを押してください...")
        ReadKey()
        If Show("続けたいですか？", "質問", YesNo, Question, Button2) = DialogResult.Yes Then
            RunApp()
        End If
    End Sub

    ''' <summary>
    ''' Input G.
    ''' </summary>
    ''' <param name="caption">Caption.</param>
    ''' <returns>Input value.</returns>
    Private Function InpG(caption As String)
        PrefInp(caption)
        Return ConvertToG(Val(ReadLine))
    End Function
#End Region
End Module
