Imports System.Data.SqlClient
Imports System.Data

Public Class clsMakeWorkingTable
    Private strSQLDataSource As String
    Private strDatabaseLocation As String
    Private strDatabaseName As String
    Private strSessionID As String
    Private intProcessStatus As Integer
    Private strProcessMessage As String
    Private strTempWorkAreaPath As String
    Private strWorkingTableName As String
    Private intSQLCommandTimeout As Integer
    Private intSQLConnectionTimeout As Integer
    Private strAttrName As String
    Private dblStartNHDPlusID As Double
    Private strHydroseq As String
    Private strTerminalpa As String
    Private strNavType As String

    Public Sub New()
        MyBase.New()
        strSQLDataSource = ""
        strDatabaseLocation = ""
        strDatabaseName = ""
        intProcessStatus = -1
        strProcessMessage = ""
        strTempWorkAreaPath = ""
        strSessionID = ""
        strAttrName = ""
        dblStartNHDPlusID = 0
        intSQLConnectionTimeout = 50000
        intSQLCommandTimeout = 50000
        strNavType = ""

    End Sub

    Protected Overrides Sub Finalize()
        On Error Resume Next
        MyBase.Finalize()
        strSQLDataSource = Nothing
        strDatabaseLocation = Nothing
        strDatabaseName = Nothing
        intProcessStatus = Nothing
        strProcessMessage = Nothing
        strTempWorkAreaPath = Nothing
        strSessionID = Nothing
        strAttrName = Nothing
        dblStartNHDPlusID = Nothing
        strNavType = Nothing
    End Sub

    Public Property SQLCommandTimeout() As Integer
        Get
            SQLCommandTimeout = intSQLCommandTimeout
        End Get
        Set(ByVal Value As Integer)
            intSQLCommandTimeout = Value
        End Set
    End Property

    Public Property SQLConnectionTimeout() As Integer
        Get
            SQLConnectionTimeout = intSQLConnectionTimeout
        End Get
        Set(ByVal Value As Integer)
            intSQLConnectionTimeout = Value
        End Set
    End Property

    Public Property TempWorkAreaPath() As String
        Get
            TempWorkAreaPath = strTempWorkAreaPath
        End Get
        Set(ByVal Value As String)
            strTempWorkAreaPath = Value
        End Set
    End Property

    Public Property SQLDataSource() As String
        Get
            SQLDataSource = strSQLDataSource
        End Get
        Set(ByVal Value As String)
            strSQLDataSource = Value
        End Set
    End Property

    Public Property DatabaseLocation() As String
        Get
            DatabaseLocation = strDatabaseLocation
        End Get
        Set(ByVal Value As String)
            strDatabaseLocation = Value
        End Set
    End Property

    Public Property DatabaseName() As String
        Get
            DatabaseName = strDatabaseName
        End Get
        Set(ByVal Value As String)
            strDatabaseName = Value
        End Set
    End Property

    Public Property AttrName() As String
        Get
            AttrName = strAttrName
        End Get
        Set(ByVal Value As String)
            strAttrName = Value
        End Set
    End Property

    Public Property StartNHDPlusID() As Double
        Get
            StartNHDPlusID = dblStartNHDPlusID
        End Get
        Set(ByVal Value As Double)
            dblStartNHDPlusID = Value
        End Set
    End Property

    Public Property SessionID() As String
        Get
            SessionID = strSessionID
        End Get
        Set(ByVal Value As String)
            strSessionID = Value
        End Set
    End Property

    Public Property Navtype() As String
        Get
            Navtype = strNavType
        End Get
        Set(ByVal Value As String)
            strNavType = Value
        End Set
    End Property

    Public ReadOnly Property WorkingTableName() As String
        'Output Property
        Get
            WorkingTableName = strWorkingTableName
        End Get
        'Set(ByVal Value As String)
        '   strWorkingTableName = Value
        'End Set
    End Property

    Public ReadOnly Property ProcessStatus() As Integer
        'Output Property
        Get
            ProcessStatus = intProcessStatus
        End Get
        'Set(ByVal Value As Integer)
        '   intProcessStatus = Value
        'End Set
    End Property

    Public ReadOnly Property ProcessMessage() As String
        'Output Property
        Get
            ProcessMessage = strProcessMessage
        End Get
        'Set(ByVal Value As String)
        '   strProcessMessage = Value
        'End Set
    End Property

    Public Function MakeWorkingTable() As Integer
        Dim strExMsg As String
        Dim strReturn As String
        Dim varSs_time2 As Object = Nothing
        Dim varEs_time2 As Object = Nothing
        Dim strElapsedTime As String
        Dim strConnectionString As String = ""
        Dim sqlconConnection As SqlConnection = New SqlConnection(strConnectionString)
        Dim strSQL As String

        MakeWorkingTable = 0
        intProcessStatus = 0
        Try
            'Get the localdb datasource with an instance based on a the current user combined with "NHDPLUSTOOLSDBS"
            'GetSQLDatasource will create it if it doesn't exist
            Dim strCurrentUser As String = Environment.UserName.ToString.ToUpper
            'Remove any spaces that might be in the userid so that SQLLocalDB.exe doesn't crash
            strCurrentUser = strCurrentUser.Replace(" ", "")

            If Not strSQLDataSource.ToUpper.Contains("(LOCALDB)\") Then    '& strCurrentUser & "NHDPLUSTOOLSDBS"
                strProcessMessage = "Invalid SQLDataSource: " & strSQLDataSource
                MakeWorkingTable = 1
                intProcessStatus = 900
                Exit Try
            End If
            varSs_time2 = Now() 'Set starting time 
            strProcessMessage = ""
            intProcessStatus = 0

            strReturn = Validate_InputProperties()
            If strReturn <> "" Then
                strProcessMessage = strReturn
                MakeWorkingTable = 1
                intProcessStatus = 900
                Exit Try
            End If

            strConnectionString = "Integrated Security=SSPI;" + "Initial Catalog=" + strDatabaseName + ";" + "Data Source=" + strSQLDataSource + "; Connection Timeout = " + intSQLConnectionTimeout.ToString + "; "
            sqlconConnection.ConnectionString = strConnectionString
            sqlconConnection.Open()

            strReturn = DoMakeWorkingTable(sqlconConnection)
            If strReturn <> "" Then
                strProcessMessage = strReturn
                MakeWorkingTable = 1
                intProcessStatus = 900
                Exit Try
            End If

            strElapsedTime = GetElapsedTime(varSs_time2)

        Catch ex As Exception
            strExMsg = ex.Message.ToString + vbCrLf + _
                       ex.StackTrace.ToString + vbCrLf + _
                       ex.Source.ToString
            strProcessMessage = "MakeWorkingTable Exception: " + strExMsg.ToString
            MakeWorkingTable = 1
           intProcessStatus = 900

        Finally
            'Close the connection, if it was opened
            If sqlconConnection.State = System.Data.ConnectionState.Open Then
                sqlconConnection.Close()
                sqlconConnection.Dispose()
                SqlConnection.ClearPool(sqlconConnection)
            End If

            'Clear variables
            strExMsg = Nothing
            strReturn = Nothing
            varSs_time2 = Nothing
            varEs_time2 = Nothing
            strElapsedTime = Nothing
            strConnectionString = Nothing
            sqlconConnection = Nothing
            strSQL = Nothing

        End Try

    End Function

    Private Function Validate_InputProperties() As String

        Dim strExMsg As String
        Dim strReturn As String
        Dim strMessage As String

        Validate_InputProperties = ""

        Try
            strReturn = ""

            'Database must exist in the provided location
            If Not My.Computer.FileSystem.FileExists(strDatabaseLocation + "\" + strDatabaseName + ".mdf") Then
                strReturn = "DatabaseName must exist in the DatabaseLocation provided. ( " + strDatabaseLocation + "\" + strDatabaseName + ".mdf" + " ) "
            End If

            'DataSource cannot be blank 
            If strSQLDataSource = "" Then
                strMessage = "SQLDataSource cannot be blank."
                If strReturn = "" Then
                    strReturn = strMessage
                Else
                    strReturn = strReturn + vbCrLf + strMessage
                End If
            End If

            'TempWorkAreaPath must exist
            If Not My.Computer.FileSystem.DirectoryExists(strTempWorkAreaPath) Then
                strMessage = "TempWorkAreaPath must exist."
                If strReturn = "" Then
                    strReturn = strMessage
                Else
                    strReturn = strReturn + vbCrLf + strMessage
                End If
            End If

            'SessionID cannot be blank 
            If strSessionID = "" Then
                strMessage = "SessionID cannot be blank."
                If strReturn = "" Then
                    strReturn = strMessage
                Else
                    strReturn = strReturn + vbCrLf + strMessage
                End If
            End If

            Validate_InputProperties = strReturn

        Catch ex As Exception
            strExMsg = ex.Message.ToString + vbCrLf +
                ex.StackTrace.ToString + vbCrLf +
                ex.Source.ToString
            Validate_InputProperties = "Validate_InputProperties Exception: " + strExMsg

        Finally
            strExMsg = Nothing
            strReturn = Nothing
            strMessage = Nothing

        End Try

    End Function

    Private Function DoMakeWorkingTable(ByRef sqlconConnection As SqlConnection) As String
        Dim varSs_time As Object
        Dim varEs_time As Object
        Dim strSQL As String
        Dim strReturn As String
        Dim strElapsedTime As String
        Dim strExMsg As String
        Dim strOperator As String
        Dim strTermpa As String

        DoMakeWorkingTable = ""

        Try
            varSs_time = Now() 'Set starting time

            'Sets the output property
            strWorkingTableName = "t" + strSessionID + "_VAA"

            'LDM - 20140702 - performance mod to limit content of strWorkingTableName by TERMINALPATHID AND HYDROSEQ
            'LDM - 20140702 - performance mod added indexes to strWorkingTableName by LEVELPATHID AND HYDROSEQ

            Dim queryString As String = "SELECT hydroseq, terminalpa FROM PlusFlowlineVAA where nhdplusid = " & StartNHDPlusID
            'MsgBox(strHydroseq & " " & strTerminalpa)

            Dim command As New SqlCommand(queryString, sqlconConnection)

            Dim reader As SqlDataReader = command.ExecuteReader()

            ' Call Read before accessing data. 
            reader.Read()
            ReadSingleRow(CType(reader, IDataRecord))
            reader.Close()

            If Navtype = "UPMAIN" Or Navtype = "UPTRIB" Then
                strOperator = ">="
            Else
                strOperator = "<="
            End If

            If Navtype = "DNDIV" Then
                strTermpa = ""
            Else
                strTermpa = " AND a.terminalpa = " & strTerminalpa
            End If

            If strAttrName <> "" Then
                'Populate working table.  New fields will have dummy values
                strSQL = "SELECT a.nhdplusid,a.reachcode,a.hydroseq,a.levelpathi,a.pathlength,a.terminalpa,a.uplevelpat,a.uphydroseq,a.dnlevelpat, a.dnminorhyd, a.dndraincou, a.divergence,a.frommeas, a.tomeas, a.lengthkm, a.dnhydroseq, a." & strAttrName & " AS attrname, a.nhdplusid AS from1, a.nhdplusid AS to1, 0 AS selected " &
                    " INTO " & strWorkingTableName & " FROM PlusFlowlineVAA AS a WHERE a.hydroseq " & strOperator & " " & strHydroseq 'TODO - Readd this & strTermpa
                strReturn = ExecuteSQL(strSQL, sqlconConnection, intSQLCommandTimeout)
                If strReturn <> "" Then
                    DoMakeWorkingTable = strReturn
                    Exit Try
                End If
            Else
                'Populate working table.  New fields will have dummy values
                strSQL = "SELECT a.nhdplusid,a.reachcode,a.hydroseq,a.levelpathi,a.pathlength,a.terminalpa,a.uplevelpat,a.uphydroseq,a.dnlevelpat, a.dnminorhyd, a.dndraincou, a.divergence,a.frommeas, a.tomeas, a.lengthkm, a.dnhydroseq, NULL AS attrname, a.nhdplusid AS from1, a.nhdplusid AS to1, 0 AS selected " &
                    " INTO " & strWorkingTableName & " FROM PlusFlowlineVAA AS a WHERE a.hydroseq " & strOperator & " " & strHydroseq 'TODO - Readd this & strTermpa
                strReturn = ExecuteSQL(strSQL, sqlconConnection, intSQLCommandTimeout)
                If strReturn <> "" Then
                    DoMakeWorkingTable = strReturn
                    Exit Try
                End If
            End If

            strSQL = "CREATE INDEX iWorkingnhdplusid ON " & strWorkingTableName & " (nhdplusid); " & vbCrLf &
                                     " CREATE INDEX iWorkingHydroseq ON " & strWorkingTableName & " (hydroseq); " & vbCrLf +
                                     " CREATE INDEX iWorkingLevelPathi ON " & strWorkingTableName & " (pathlength); "
            strReturn = ExecuteSQL(strSQL, sqlconConnection, intSQLCommandTimeout)
            If strReturn <> "" Then
                DoMakeWorkingTable = strReturn
                Exit Try
            End If

            'Update the dummy values
            strSQL = "UPDATE " + strWorkingTableName + " SET from1 = NULL, to1 = NULL, Selected = 0"
            strReturn = ExecuteSQL(strSQL, sqlconConnection, intSQLCommandTimeout)
            If strReturn <> "" Then
                DoMakeWorkingTable = strReturn
                Exit Try
            End If

            'Create indices
            strSQL = " CREATE INDEX i" + strSessionID + "nhdplusid ON " + strWorkingTableName + " (nhdplusid); " + vbCrLf +
                     " CREATE INDEX i" + strSessionID + "Selected ON " + strWorkingTableName + " (selected); " + vbCrLf +
                     " CREATE INDEX i" + strSessionID + "Hydroseq ON " + strWorkingTableName + " (hydroseq); " + vbCrLf +
                     " CREATE INDEX i" + strSessionID + "Levelpathi ON " + strWorkingTableName + " (levelpathi); " + vbCrLf +
                     " CREATE INDEX i" + strSessionID + "Terminalpa ON " + strWorkingTableName + " (terminalpa); " + vbCrLf +
                     " CREATE INDEX i" + strSessionID + "Attrname ON " + strWorkingTableName + " (attrname); " + vbCrLf +
                     " CREATE INDEX i" + strSessionID + "DnDrainCou ON " + strWorkingTableName + " (DnDrainCou); "
            strReturn = ExecuteSQL(strSQL, sqlconConnection, intSQLCommandTimeout)
            If strReturn <> "" Then
                DoMakeWorkingTable = strReturn
                Exit Try
            End If

            strElapsedTime = GetElapsedTime(varSs_time)  'Calculate elapsed time

        Catch ex As Exception
            strExMsg = ex.Message.ToString + vbCrLf +
                ex.StackTrace.ToString + vbCrLf +
                ex.Source.ToString
            DoMakeWorkingTable = "DoMakeWorkingTable Exception: " + strExMsg

        Finally
            varSs_time = Nothing
            varEs_time = Nothing
            strSQL = Nothing
            strReturn = Nothing
            strElapsedTime = Nothing
            strExMsg = Nothing

        End Try

    End Function

    Private Sub ReadSingleRow(ByVal record As IDataRecord)
        strHydroseq = record(0)
        strTerminalpa = record(1)
    End Sub

End Class
