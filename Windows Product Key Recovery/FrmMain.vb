Option Explicit On
Option Infer On

Imports System.IO

Public Class FrmMain

    Private Declare Sub BASSMOD_Free Lib "bassmod.dll" ()
    Private Declare Function BASSMOD_Init Lib "bassmod.dll" (ByVal device As Integer, ByVal freq As Integer, ByVal flags As Integer) As Short
    Private Declare Function BASSMOD_MusicLoad Lib "bassmod.dll" (ByVal mem As Short, ByVal pfile As String, ByVal offset As Integer, ByVal Length As Integer, ByVal flags As Integer) As Short
    Private Declare Function BASSMOD_MusicPlay Lib "bassmod.dll" () As Short
    Private Declare Function BASSMOD_MusicStop Lib "bassmod.dll" () As Short
    Private Declare Function BASSMOD_MusicGetPosition Lib "bassmod.dll" () As Short
    Private Declare Function BASSMOD_MusicIsActive Lib "bassmod.dll" () As Short
    Private Declare Function BASSMOD_MusicSetPosition Lib "bassmod.dll" (ByVal Posição As Integer) As Integer
    Private Declare Function BASSMOD_MusicGetLength Lib "bassmod.dll" (ByVal Tamanho As Integer) As Integer
    Private Declare Function BASSMOD_MusicPause Lib "bassmod.dll" () As Short

    Dim BASSMOD_Dll_Already_There As Boolean = False

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Button1.Enabled = False
        Application.DoEvents()
        LoadPKeys()
        Application.DoEvents()
        Button1.Enabled = True
        If Trim(TextBox1.Text) = "" Then
            MsgBox("No keys were recovered for supported products at this time.", MsgBoxStyle.Information, "Windows Product Key Recovey")
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Close()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        MsgBox("Windows Product Key Recovery 2.0 (22/04/2013)" & vbNewLine & vbNewLine & "Author: Steven Jenkins De Haro" & _
        vbNewLine & "A Steve Creation/Convergence" & vbNewLine & vbNewLine & _
        "Microsoft .NET Framework 3.5", MsgBoxStyle.OkOnly, "Windows Product Key Recovery")
    End Sub

    Private Sub Donate5PaypalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Donate5PaypalToolStripMenuItem.Click
        On Error Resume Next
        Process.Start("https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=8493677")
    End Sub

    Private Sub CopyKeyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CopyKeyToolStripMenuItem.Click
        Clipboard.SetText(TextBox1.Text)
        TextBox1.DeselectAll()
    End Sub

    Private Sub SaveToolStripMenuItem1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem1.Click
        Dim FileWriter As StreamWriter
        Dim Results As DialogResult

        SaveFileDialog1.FileName = "Untitled"
        SaveFileDialog1.Filter = "Text files (*.txt)|*.txt"
        Results = SaveFileDialog1.ShowDialog
        If Results = DialogResult.OK Then
            FileWriter = New StreamWriter(SaveFileDialog1.FileName, False)
            Try
                FileWriter.Write(TextBox1.Text)
            Catch
                MsgBox("An error has occurred while saving the file.", MsgBoxStyle.Critical, "Windows Product Key Recovery")
            Finally
                FileWriter.Close()
                TextBox1.DeselectAll()
            End Try
        Else
            TextBox1.DeselectAll()
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        SaveToolStripMenuItem1.Enabled = CBool(Not Trim(TextBox1.Text) = "")
        CopyKeyToolStripMenuItem.Enabled = CBool(Not Trim(TextBox1.Text) = "")
    End Sub

    Private Sub FrmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        On Error Resume Next

        If getWindowsRunningMode() = 64 Then Exit Sub

        BASSMOD_MusicStop()
        BASSMOD_Free()

        If File.Exists(Environ$("TEMP") & "\" & "GTA_SAtrn.xm") = True Then
            Kill(Environ$("TEMP") & "\" & "GTA_SAtrn.xm")
        End If

        'below will produce an access denied error even with admin rights. not sure why, but not serious with
        'help from the error handler which lets the dll stay there anyways. left code mostly for XP which should not have the error.
        If File.Exists(Environ$("windir") & "\system32\" & "BASSMOD.dll") = True And BASSMOD_Dll_Already_There = False Then
            Kill(Environ$("windir") & "\system32\" & "BASSMOD.dll")
        End If
    End Sub

    Private Sub FrmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        On Error Resume Next
        'Do not add this before creating files. BASSMOD_MusicStop() and BASSMOD_Free()
        If getWindowsRunningMode() = 64 Then
            CheckBox1.Visible = False
            Exit Sub
        End If

        Dim DataArray() As Byte
        DataArray = My.Resources.GTA_SAtrn.ToArray()
        FileOpen(1, Environ$("TEMP") & "\" & "GTA_SAtrn.xm", OpenMode.Binary) 'Puts xm file in temp directory.
        FilePut(1, DataArray)
        FileClose(1)
        Erase DataArray

        If File.Exists(Environ$("windir") & "\system32\" & "BASSMOD.dll") = False Then
            Dim DataArray2() As Byte
            DataArray = My.Resources.BASSMOD.ToArray()
            FileOpen(1, Environ$("windir") & "\system32\" & "BASSMOD.dll", OpenMode.Binary) 'Places the dll in system32... needs admin rights.
            FilePut(1, DataArray)
            FileClose(1)
            Erase DataArray2
        Else
            BASSMOD_Dll_Already_There = True
        End If

        BASSMOD_Init(-1, 44100, 0)
        BASSMOD_MusicLoad(0, Environ$("TEMP") & "\" & "GTA_SAtrn.xm", 0, 0, 4)
        BASSMOD_MusicPlay()
    End Sub

     Private Sub CheckBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox1.Click
        On Error Resume Next
        If CheckBox1.Checked = True Then
            BASSMOD_MusicStop()
        Else
            BASSMOD_MusicPlay()
        End If
    End Sub

End Class
