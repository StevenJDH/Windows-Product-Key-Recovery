Option Explicit On
Option Infer On

Imports Microsoft.Win32
Imports System.Text

Module ModuleSmart

    Public Enum ReleaseVersion
        PreBeta = 0
        Beta1 = 1
        Beta2 = 2
        RC0 = 3
        RC1 = 4
        RTM = 9
    End Enum

    Public Enum ReleaseEditionTypeXP2003
        Enterprise = 0
        Retail = 1
        Trial = 2
    End Enum

    Public Enum ReleaseType2007Plus
        Volume = 0
        Retail = 1
        Trial = 2
        Download = 5
    End Enum

    Public Enum ReleaseArch2007Plus
        x86 = 0
        x64 = 1
    End Enum

    Public Function OfficeIdentification(nGUID As String) As String
        Dim nResult As Int16
        Dim sBuilder As New StringBuilder

        If nGUID.EndsWith("6000-11D3-8CFE-0150048383C9}") = True Then 'Office XP-2003
            Try
                'Release Version       
                Int16.TryParse(nGUID.Substring(1, 1), nResult) 'Converts the string digit to int16 and puts it into the nResult varible for use.
                sBuilder.Append(CType(nResult, ReleaseVersion).ToString).Append(" ")
            Catch
                'If the version is not one of the main 6 which is rare then don't return the version and continue.
            End Try
            Try
                'Release Edition Type
                Int16.TryParse(nGUID.Substring(2, 1), nResult) 'Converts the string digit to int16 and puts it into the nResult varible for use.
                sBuilder.Append(CType(nResult, ReleaseEditionTypeXP2003).ToString).Append(" ")
            Catch
                'If any code errors happen then don't return any part of this and continue.
            End Try
            'Release Architecture is only 32-bit for Office pre 2007. 
            sBuilder.Append("x86")
        ElseIf nGUID.EndsWith("0FF1CE}") = True Then 'Office 2007+
            Try
                'Release Version       
                Int16.TryParse(nGUID.Substring(1, 1), nResult) 'Converts the string digit to int16 and puts it into the nResult varible for use.
                sBuilder.Append(CType(nResult, ReleaseVersion).ToString).Append(" ")
            Catch
                'If the version is not one of the main 6 which is rare then don't return the version and continue.
            End Try
            Try
                'Release Type
                Int16.TryParse(nGUID.Substring(2, 1), nResult) 'Converts the string digit to int16 and puts it into the nResult varible for use.
                sBuilder.Append(CType(nResult, ReleaseType2007Plus).ToString).Append(" ")
            Catch
                'If any code errors happen then don't return any part of this and continue.
            End Try
            Try
                'Release Architecture
                Int16.TryParse(nGUID.Substring(20, 1), nResult) 'Converts the string digit to int16 and puts it into the nResult varible for use.
                sBuilder.Append(CType(nResult, ReleaseArch2007Plus).ToString)
            Catch
                'If any code errors happen then don't return any part of this and continue.
            End Try
        Else
                Return "" 'If something new then just return nothing.
        End If
        Return sBuilder.ToString.Trim
    End Function

    ''' <summary>
    ''' Function to retrieve all the child subkeys of a SubKey in the Windows Registry
    ''' </summary>
    ''' <param name="MainKey">RegistryKey -> One of the 6 main keys to retrieve from</param>
    ''' <param name="sKey">String -> Name of the SubKey to retrieve all child subkeys for</param>
    ''' <returns>An ArrayList of all the child subkeys</returns>
    ''' <remarks>Created 20APR13 -> Steven Jenkins De Haro</remarks>
    Public Function GetAllChildSubKeys(ByVal MainKey As RegistryKey, ByVal sKey As String) As ArrayList
        Dim rkKey As RegistryKey    'RegistryKey to work with
        Dim sSubKeys() As String    'string array to hold the subkeys
        Dim arySubKeys As New ArrayList 'arraylist to return the subkeys in an array
        Try
            'open the given subkey
            rkKey = MainKey.OpenSubKey(sKey)
            'check to see if the subkey exists
            If Not sKey Is Nothing Then 'subkey exists
                'get all the child subkeys
                sSubKeys = rkKey.GetSubKeyNames
                'loop through all the child subkeys
                For Each s As String In sSubKeys
                    'As there is more than just the office version keys in the subkey we just filter out the rest to only what we want.
                    If IsNumeric(s) = True Then
                        'add them to the arraylist
                        arySubKeys.Add(s)
                    End If
                Next
            Else    'subkey doesnt exist
                'The SubKey provided doesn't exist. Please check your entry and try again.
                Return arySubKeys 'We return like this in case of error becaue the other code will recovery from this will out error. Any other way and the other code will break.
            End If
        Catch ex As Exception
            Return arySubKeys 'We return like this in case of error becaue the other code will recovery from this will out error. Any other way and the other code will break.
        End Try
        Return arySubKeys 'return the subkeys arraylist if everything goes good.
    End Function

    ''' <summary>
    ''' Function to retrieves GUID child subkeys of a SubKey in the Windows Registry.
    ''' </summary>
    ''' <param name="MainKey">RegistryKey -> One of the 6 main keys to retrieve from</param>
    ''' <param name="sKey">String -> Name of the SubKey to retrieve all child subkeys for</param>
    ''' <returns>An Office product GUID of the child subkeys</returns>
    ''' <remarks>Created 20APR13 -> Steven Jenkins De Haro</remarks>
    Public Function GetProductChildSubKeys(ByVal MainKey As RegistryKey, ByVal sKey As String) As String
        Dim rkKey As RegistryKey    'RegistryKey to work with
        Dim sSubKeys() As String    'string array to hold the subkeys
        Try
            'open the given subkey
            rkKey = MainKey.OpenSubKey(sKey)
            'check to see if the subkey exists
            If Not sKey Is Nothing Then 'subkey exists
                'get all the child subkeys
                sSubKeys = rkKey.GetSubKeyNames
                'loop through all the child subkeys
                For Each nGUID As String In sSubKeys
                    'We filter out all GUID except the one we want that holds the product name and key. 
                    '6000-11D3-8CFE-0150048383C9} is how office older than 2007 ended with. 0FF1CE} is for 2007+
                    If nGUID.EndsWith("0FF1CE}") = True Or nGUID.EndsWith("6000-11D3-8CFE-0150048383C9}") = True Then
                        'return this office GUID
                        Return nGUID
                    End If
                Next
                'No GUIDs office product.
                Return ""
            Else    'subkey doesnt exist
                'The SubKey provided doesn't exist. Please check your entry and try again.
                Return ""
            End If
        Catch ex As Exception
            'Error retrieving subKeys
            Return ""
        End Try
    End Function

    Public Function GetOSVersion() As String
        Dim nOSFullName As String
        Dim nCSDVersion As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine
        nCSDVersion = nCSDVersion.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion")
        nOSFullName = My.Computer.Info.OSFullName & " " & nCSDVersion.GetValue("CSDVersion", Nothing)
        nCSDVersion.Close()
        nOSFullName = Replace(nOSFullName, "®", "")
        nOSFullName = Replace(nOSFullName, "™", "")
        Return Trim(nOSFullName)
    End Function

    Public Function GetProductKey(ByVal KeyPath As String, ByVal ValueName As String) As String
        Try
            Dim HexBuf As Object = My.Computer.Registry.GetValue(KeyPath, ValueName, 0)
            If HexBuf Is Nothing Then Return "Unable to read key."
            Dim tmp As String = ""
            For l As Integer = LBound(HexBuf) To UBound(HexBuf)
                tmp = tmp & " " & Hex(HexBuf(l))
            Next
            Dim StartOffset As Integer = 52
            Dim EndOffset As Integer = 67
            Dim Digits(24) As String
            Digits(0) = "B" : Digits(1) = "C" : Digits(2) = "D" : Digits(3) = "F"
            Digits(4) = "G" : Digits(5) = "H" : Digits(6) = "J" : Digits(7) = "K"
            Digits(8) = "M" : Digits(9) = "P" : Digits(10) = "Q"
            Digits(11) = "R"
            Digits(12) = "T" : Digits(13) = "V" : Digits(14) = "W"
            Digits(15) = "X"
            Digits(16) = "Y" : Digits(17) = "2" : Digits(18) = "3"
            Digits(19) = "4"
            Digits(20) = "6" : Digits(21) = "7" : Digits(22) = "8"
            Digits(23) = "9"
            Dim dLen As Integer = 29
            Dim sLen As Integer = 15
            Dim HexDigitalPID(15) As String
            Dim Des(30) As String
            Dim tmp2 As String = ""
            For i = StartOffset To EndOffset
                HexDigitalPID(i - StartOffset) = HexBuf(i)
                tmp2 = tmp2 & " " & Hex(HexDigitalPID(i - StartOffset))
            Next
            Dim KEYSTRING As String = ""
            For i As Integer = dLen - 1 To 0 Step -1
                If ((i + 1) Mod 6) = 0 Then
                    Des(i) = "-"
                    KEYSTRING = KEYSTRING & "-"
                Else
                    Dim HN As Integer = 0
                    For N As Integer = (sLen - 1) To 0 Step -1
                        Dim Value As Integer = ((HN * 2 ^ 8) Or HexDigitalPID(N))
                        HexDigitalPID(N) = Value \ 24
                        HN = (Value Mod 24)
                    Next
                    Des(i) = Digits(HN)
                    KEYSTRING = KEYSTRING & Digits(HN)
                End If
            Next
            Return StrReverse(KEYSTRING)
        Catch
            Return "Error while reading key."
        End Try
    End Function


    Public Function getWindowsRunningMode() As Integer
        'Tells you if Windows is 32-bit or 64-bit.
        If IntPtr.Size = 8 Then
            Return 64

        Else
            Return 32
        End If
    End Function

    Public Sub LoadPKeys()
        On Error Resume Next
        Dim nGUID As String
        Dim nCDKey As String
        Dim nOfficeVersion As String
        FrmMain.TextBox1.Text = ""

        'Microsoft OS
        Dim nVersion As Decimal
        nVersion = Environment.OSVersion.Version.Major & "." & Environment.OSVersion.Version.Minor
        If (nVersion >= 5.1 And nVersion <= 6.2) = True Or nVersion = 60 Then
            If getWindowsRunningMode() = 32 Then
                FrmMain.TextBox1.Text = GetOSVersion() & vbNewLine & _
                GetProductKey("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\", "DigitalProductId") & vbNewLine & vbNewLine
            Else
                FrmMain.TextBox1.Text = GetOSVersion() & " 64-bit" & vbNewLine & _
                GetProductKey("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\", "DigitalProductId") & vbNewLine & vbNewLine
            End If
        Else
            FrmMain.TextBox1.Text = "OS Not Supported" & vbNewLine & "OS key recovery only for Microsoft Windows XP/2003/Vista/2008/7/8/2012" & vbNewLine & vbNewLine
        End If
        'If an Office product was installed at some point and the regkey is there then continue.
        If Not Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Office", True) Is Nothing Then
            'Get all stored Office product version subkey numbers for potential CD-Keys. There may still be old keys left behind from older versions that were used.
            For Each nOfficeVersion In GetAllChildSubKeys(Microsoft.Win32.Registry.LocalMachine, "SOFTWARE\Microsoft\Office")
                'If subkey Registration exists then continue.
                If Not Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Office\" & nOfficeVersion & "\Registration", True) Is Nothing Then
                    nGUID = GetProductChildSubKeys(Microsoft.Win32.Registry.LocalMachine, "SOFTWARE\Microsoft\Office\" & nOfficeVersion & "\Registration")
                    'If a subkey of Registration has an Office product GUID then continue.
                    If Not nGUID = "" Then
                        'To store key or error message for checking.
                        nCDKey = GetProductKey("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\" & nOfficeVersion & "\Registration\" & nGUID, "DigitalProductId")
                        'If the GUID key that holds the name of the Office product is there then continue.
                        If Not Microsoft.Win32.Registry.LocalMachine.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Uninstall\" & nGUID, True) Is Nothing Then
                            ' Alternative one but it won't show for example the Plus in Profession Plus, just Professional. My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\" & nOfficeVersion & "\Registration\" & nGUID, "ConvertToEdition", "Error while reading Office edition name.")
                            FrmMain.TextBox1.Text = FrmMain.TextBox1.Text & My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & nGUID, "DisplayName", "Microsoft Office Unknown Edition v" & nOfficeVersion) & " " & OfficeIdentification(nGUID) & vbNewLine & _
                                    nCDKey & vbNewLine & vbNewLine
                        ElseIf Not nCDKey = "Error while reading key." Then 'If no title, but has a valid key then do display unknown. If no title and not a valid key then ignore.
                            FrmMain.TextBox1.Text = FrmMain.TextBox1.Text & "Microsoft Office Unknown Edition v" & nOfficeVersion & " " & OfficeIdentification(nGUID) & vbNewLine & _
                            nCDKey & vbNewLine & vbNewLine
                        End If
                    End If
                End If
            Next
        End If
        'We need to check if Windows is running in 64-bit mode and if a 32-bit office is installed.
        If getWindowsRunningMode() = 64 Then
            'If an Office product was installed at some point and the regkey is there then continue.
            If Not Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\WOW6432Node\Microsoft\Office", True) Is Nothing Then
                'Get all stored Office product version subkey numbers for potential CD-Keys. There may still be old keys left behind from older versions that were used.
                For Each nOfficeVersion In GetAllChildSubKeys(Microsoft.Win32.Registry.LocalMachine, "SOFTWARE\WOW6432Node\Microsoft\Office")
                    'If subkey Registration exists then continue.
                    If Not Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\WOW6432Node\Microsoft\Office\" & nOfficeVersion & "\Registration", True) Is Nothing Then
                        nGUID = GetProductChildSubKeys(Microsoft.Win32.Registry.LocalMachine, "SOFTWARE\WOW6432Node\Microsoft\Office\" & nOfficeVersion & "\Registration")
                        'If a subkey of Registration has an Office product GUID then continue.
                        If Not nGUID = "" Then
                            'To store key or error message for checking.
                            nCDKey = GetProductKey("HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\" & nOfficeVersion & "\Registration\" & nGUID, "DigitalProductId")
                            'If the GUID key that holds the name of the Office product is there then continue.
                            If Not Microsoft.Win32.Registry.LocalMachine.OpenSubKey("Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\" & nGUID, True) Is Nothing Then
                                ' Alternative one but it won't show for example the Plus in Profession Plus, just Professional. My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\" & nOfficeVersion & "\Registration\" & nGUID, "ConvertToEdition", "Error while reading Office edition name.")
                                FrmMain.TextBox1.Text = FrmMain.TextBox1.Text & My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\" & nGUID, "DisplayName", "Microsoft Office Unknown Edition v" & nOfficeVersion) & " " & OfficeIdentification(nGUID) & vbNewLine & _
                                        nCDKey & vbNewLine & vbNewLine
                            ElseIf Not nCDKey = "Error while reading key." Then 'If no title, but has a valid key then do display unknown. If no title and not a valid key then ignore.
                                FrmMain.TextBox1.Text = FrmMain.TextBox1.Text & "Microsoft Office Unknown Edition v" & nOfficeVersion & " " & OfficeIdentification(nGUID) & vbNewLine & _
                                nCDKey & vbNewLine & vbNewLine
                            End If
                        End If
                    End If
                Next
            End If
        End If
    End Sub

End Module
