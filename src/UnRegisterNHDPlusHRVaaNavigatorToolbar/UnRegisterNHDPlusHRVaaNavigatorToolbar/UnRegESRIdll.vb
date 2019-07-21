Module UnRegESRIdll

    '---------------------------------------------
    ' Program Name  : UnRegisterNHDPlusHRVaaNavigatorToolBar
    ' Purpose       : Console application called as a custom action during the installation of NHDPlusHRVaaNavigatorToolBar
    '               : used to unregister NHDPlusHRVaaNavigatorToolBar.dll with ESRIRegAsm.exe because life isn't difficult enough
    ' Input         : Location of application folder used during install - passed on the command line
    ' Author        : Bob Deffenbaugh
    ' Organization  : Horizon Systems Corporation
    ' Created       : 7/18/2018
    '---------------------------------------------
    '
    Sub Main()
        'Location of the application installation folder
        Dim strTargetDir As String = Command()
        Dim strX86CommonFolder As String
        If 8 = IntPtr.Size OrElse (Not [String].IsNullOrEmpty(Environment.GetEnvironmentVariable("PROCESSOR_ARCHITEW6432"))) Then
            strX86CommonFolder = Environment.GetEnvironmentVariable("CommonProgramFiles(x86)")
        Else
            strX86CommonFolder = Environment.GetEnvironmentVariable("CommonProgramFiles")
        End If

        Dim oProcess As New Process()
        Dim oStartInfo1 As New ProcessStartInfo(strX86CommonFolder & "\ArcGIS\bin\ESRIRegAsm.exe")
        Dim oStartInfo2 As New ProcessStartInfo(strX86CommonFolder & "\ArcGIS\bin\ESRIRegAsm.exe")

        Try
            If System.Environment.OSVersion.Version.Major >= 6 Then
                ' Run as admin on Windows Vista or higher 
                oStartInfo1.Verb = "runas"
                oStartInfo2.Verb = "runas"
                'MsgBox(oStartInfo2.Verb)
            End If
            'MsgBox(strTargetDir)
            strTargetDir = "C:\NHDPlusHRTools\NHDPlusHRVaaNavigatorToolBar"

            'oStartInfo1.FileName = strX86CommonFolder & "\ArcGIS\bin\ESRIRegAsm.exe"
            oStartInfo1.Arguments = strTargetDir & "\NHDPlusHRVAANavigatorToolbar.dll /p:desktop /s /u"
            oStartInfo1.WindowStyle = ProcessWindowStyle.Hidden
            oStartInfo1.CreateNoWindow = True
            oStartInfo1.UseShellExecute = False
            oProcess.StartInfo = oStartInfo1
            ' Start the process
            oProcess.Start()

            'oStartInfo2.FileName = strX86CommonFolder & "\ArcGIS\bin\ESRIRegAsm.exe"
            oStartInfo2.Arguments = strTargetDir & "\NHDPlusHRVAANavigator.dll /p:desktop /s /u"
            oStartInfo2.WindowStyle = ProcessWindowStyle.Hidden
            oStartInfo2.CreateNoWindow = True
            oStartInfo2.UseShellExecute = False
            oProcess.StartInfo = oStartInfo2
            ' Start the process
            oProcess.Start()

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            If Not (oProcess Is Nothing) Then
                oProcess.Close()
                oProcess.Dispose()
            End If
        End Try
    End Sub
End Module
