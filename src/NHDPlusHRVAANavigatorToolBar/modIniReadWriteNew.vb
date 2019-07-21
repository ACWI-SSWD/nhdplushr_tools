Option Strict On

Module INIAccess

#Region "API Calls"
   ' standard API declarations for INI access
   ' changing only "As Long" to "As Int32" (As Integer would work also)
   Private Declare Unicode Function WritePrivateProfileString Lib "kernel32" _
   Alias "WritePrivateProfileStringW" (ByVal lpApplicationName As String, _
   ByVal lpKeyName As String, ByVal lpString As String, _
   ByVal lpFileName As String) As Int32

   Private Declare Unicode Function GetPrivateProfileString Lib "kernel32" _
   Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As String, _
   ByVal lpKeyName As String, ByVal lpDefault As String, _
   ByVal lpReturnedString As String, ByVal nSize As Int32, _
   ByVal lpFileName As String) As Int32
#End Region

    Public Function INIRead(ByVal INIPath As String, ByVal SectionName As String, ByVal KeyName As String, ByVal DefaultValue As String) As String
        ' primary version of call gets single value given all parameters
        Dim n As Int32
        Dim sData As String
        sData = space$(1024) ' allocate some room
        n = GetPrivateProfileString(SectionName, KeyName, DefaultValue, _
        sData, sData.Length, INIPath)
        If n > 0 Then ' return whatever it gave us
            INIRead = sdata.Substring(0, n)
        Else
            iniread = ""
        End If
    End Function

    Public Sub INIWrite(ByVal INIPath As String, ByVal SectionName As String, ByVal KeyName As String, ByVal TheValue As String)
        Call WritePrivateProfileString(SectionName, KeyName, TheValue, INIPath)
    End Sub

    Public Sub INIDeleteLine(ByVal INIPath As String, ByVal SectionName As String, ByVal KeyName As String) ' delete single line from section
        Call WritePrivateProfileString(SectionName, KeyName, Nothing, INIPath)
    End Sub

    Public Sub INIDeleteSection(ByVal INIPath As String, ByVal SectionName As String)
        ' delete section from INI file
        Call WritePrivateProfileString(SectionName, Nothing, Nothing, INIPath)
    End Sub

End Module



