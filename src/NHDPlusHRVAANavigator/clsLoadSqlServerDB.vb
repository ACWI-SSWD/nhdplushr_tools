Imports System.Data.SqlClient

Public Class clsLoadSqlServerDB

    Private strSQLDataSource As String
    Private strDatabaseLocation As String
    Private strDatabaseName As String
    Private strInputNHDPlusLocation As String
    Private boolAddToExisting As Boolean
    Private intProcessStatus As Integer
    Private strProcessMessage As String
    Private strTempWorkAreaPath As String
    Private intSQLCommandTimeout As Integer
    Private intSQLConnectionTimeout As Integer
    Private strAttrName As String
    Private intSQLBulkCopyTimeout As Integer
    Private intSQLBatchSize As Integer

    Private gintSQLConnectionTimeout As Integer = 50000
    Private gintSQLCommandTimeOut As Integer = 50000

    Public Sub New()
        MyBase.New()
		
        strSQLDataSource = ""
        strDatabaseLocation = ""
        strDatabaseName = ""
        strInputNHDPlusLocation = ""
        boolAddToExisting = False
        intProcessStatus = -1
        strProcessMessage = ""
        strTempWorkAreaPath = ""
        intSQLConnectionTimeout = 50000
        intSQLCommandTimeout = 50000
        strAttrName = ""
    End Sub

    Protected Overrides Sub Finalize()
        On Error Resume Next
        MyBase.Finalize()

        strSQLDataSource = Nothing
        strDatabaseLocation = Nothing
        strDatabaseName = Nothing
        strInputNHDPlusLocation = Nothing
        boolAddToExisting = Nothing
        intProcessStatus = Nothing
        strProcessMessage = Nothing
        strTempWorkAreaPath = Nothing
        'strAttrFile = Nothing
        strAttrName = Nothing
        intSQLConnectionTimeout = Nothing
        intSQLCommandTimeout = Nothing

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

    Public Property SQLBulkCopyTimeout() As Integer
        Get
            Return intSQLBulkCopyTimeout
        End Get
        Set(ByVal Value As Integer)
            intSQLBulkCopyTimeout = Value
        End Set
    End Property

    Public Property SQLBatchSize() As Integer
        Get
            Return intSQLBatchSize
        End Get
        Set(ByVal Value As Integer)
            intSQLBatchSize = Value
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

    Public Property InputNHDPlusLocation() As String
        Get
            InputNHDPlusLocation = strInputNHDPlusLocation
        End Get
        Set(ByVal Value As String)
            strInputNHDPlusLocation = Value
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

    Public Property AddToExisting() As Boolean
        Get
            AddToExisting = boolAddToExisting
        End Get
        Set(ByVal Value As Boolean)
            boolAddToExisting = Value
        End Set
    End Property

    Public ReadOnly Property ProcessStatus() As Integer
        'Output Property
        Get
            ProcessStatus = intProcessStatus
        End Get
    End Property

    Public ReadOnly Property ProcessMessage() As String
        'Output Property
        Get
            ProcessMessage = strProcessMessage
        End Get
    End Property

    Public Function LoadSQLServerDB() As Integer
        Dim strExMsg As String
        Dim strReturn As String
        Dim varSs_time2 As Object = Nothing
        Dim varEs_time2 As Object = Nothing
        Dim strElapsedTime As String
        Dim strConnectionString As String = ""
        Dim sqlconConnection As SqlConnection = New SqlConnection(strConnectionString)
        Dim strSQL As String
        LoadSQLServerDB = 0
        intProcessStatus = 0

        Try
            varSs_time2 = Now() 'Set starting time 

            'Get the localdb datasource with an instance based on a the current user combined with "NHDPLUSTOOLSDBS"
            'GetSQLDatasource will create it if it doesn't exist
            Dim strCurrentUser As String = Environment.UserName.ToString.ToUpper
            'Remove any spaces that might be in the userid so that SQLLocalDB.exe doesn't crash
            strCurrentUser = strCurrentUser.Replace(" ", "")
            If Not strSQLDataSource.ToUpper.Contains("(LOCALDB)\") Then    '& strCurrentUser & "NHDPLUSTOOLSDBS"
                strProcessMessage = "Invalid SQLDataSource: " & strSQLDataSource
                LoadSQLServerDB = 1
                intProcessStatus = 900
                Exit Try
            End If
            strReturn = Validate_InputProperties()

            If strReturn <> "" Then
                strProcessMessage = strReturn
                LoadSQLServerDB = 1
                intProcessStatus = 900
                Exit Try
            End If

            If boolAddToExisting Then
                'User is choosing to add data to an existing sqlserver database
                'Connect to it.
                'Sql Server will generate errors if the database is not correct 
                'Create a connection to the DB
                strConnectionString = "Integrated Security=SSPI;" + "Initial Catalog=" + strDatabaseName + ";" + "Data Source=" + strSQLDataSource + ";"
                sqlconConnection.ConnectionString = strConnectionString
                sqlconConnection.Open()
            Else
                strConnectionString = "Integrated Security=SSPI;" + "Data Source=" + strSQLDataSource + ";"
                sqlconConnection.ConnectionString = strConnectionString
                sqlconConnection.Open()
                strSQL = " SELECT COUNT(*) FROM sys.databases WHERE name = '" & strDatabaseName & "'"
                Dim cmd As SqlCommand = New SqlCommand(strSQL, sqlconConnection)
                If cmd.ExecuteScalar > 0 Then
                    sqlconConnection.Close()
                Else
                    sqlconConnection.Close()
                    'Working DB doesn't exist. 
                    'User is choosing to create a new database and add data to it.
                    strReturn = CreateDB(sqlconConnection)
                    If strReturn <> "" Then
                        strProcessMessage = strReturn
                        LoadSQLServerDB = 1
                        intProcessStatus = 900
                        Exit Try
                    End If
                End If

            End If
            'create a connection to the newly attached DB
            strReturn = LoadDB(sqlconConnection)

            If strReturn <> "" Then
                'Load failed, but the database was created.  Drop the database unless
                strProcessMessage = strReturn
                LoadSQLServerDB = 1
                intProcessStatus = 900
                If boolAddToExisting = False Then
                    'Create a generic connection
                    sqlconConnection.Close()
                    SqlConnection.ClearPool(sqlconConnection)

                    strReturn = DetachDB(strDatabaseName, intSQLConnectionTimeout)
                    If strReturn <> "" Then
                        strProcessMessage = strReturn
                        LoadSQLServerDB = 1
                        intProcessStatus = 900
                        Exit Try
                    End If
                    strReturn = DropDB(strDatabaseName, intSQLConnectionTimeout)
                    If strReturn <> "" Then
                        strProcessMessage = strReturn
                        LoadSQLServerDB = 1
                        intProcessStatus = 900
                        Exit Try
                    End If
                End If
            End If

            strElapsedTime = GetElapsedTime(varSs_time2)

        Catch ex As Exception
            strExMsg = ex.Message.ToString + vbCrLf +
                       ex.StackTrace.ToString + vbCrLf +
                       ex.Source.ToString
            strProcessMessage = "LoadSqlServerDB Exception: " + strExMsg.ToString
            LoadSQLServerDB = 1

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

    Private Function CreateDB(ByRef sqlconConnection As SqlConnection) As String

        Dim varSs_time As Object
        Dim varEs_time As Object
        Dim strConnectionString As String
        Dim strSQL As String
        Dim strReturn As String
        Dim strElapsedTime As String
        Dim strExMsg As String

        CreateDB = ""

        Try
            varSs_time = Now() 'Set starting time

            'Create working sqlserver db and connect to it
            strReturn = CreateLocalDatabase(strDatabaseName, strDatabaseLocation, intSQLConnectionTimeout, intSQLCommandTimeOut, strSQLDataSource)
            If strReturn <> "" Then
                CreateDB = strReturn
                Exit Try
            End If

            strReturn = CreateArcGisDatabaseConnection(strDatabaseLocation, strDatabaseName, "SQL_SERVER", strSQLDataSource, strDatabaseName)
            If strReturn <> "" Then
                CreateDB = strReturn
                Exit Try
            End If

            'Reconnect to the newly created db
            strConnectionString = "Integrated Security=SSPI;" + "Initial Catalog=" + strDatabaseName + ";" + "Data Source=" + strSQLDataSource + "; Connection Timeout = " + intSQLConnectionTimeout.ToString + ";"
            sqlconConnection.ConnectionString = strConnectionString
            sqlconConnection.Open()

            strElapsedTime = GetElapsedTime(varSs_time)  'Calculate elapsed time

        Catch ex As Exception
            strExMsg = ex.Message.ToString + vbCrLf +
                ex.StackTrace.ToString + vbCrLf +
                ex.Source.ToString
            CreateDB = "CreateDB Exception: " + strExMsg

        Finally
            varSs_time = Nothing
            varEs_time = Nothing
            strConnectionString = Nothing
            strSQL = Nothing
            strReturn = Nothing
            strElapsedTime = Nothing
            strExMsg = Nothing

        End Try

    End Function

    Private Function Validate_InputProperties() As String
        Dim strExMsg As String
        Dim strReturn As String
        Dim strMessage As String
        Dim strAlphaList As String
        Dim strCompareString As String

        Validate_InputProperties = ""

        Try
            strReturn = ""

            'DatabaseLocation must exist 
            If Not My.Computer.FileSystem.DirectoryExists(strDatabaseLocation) Then
                strReturn = "DatabaseLocation " + strDatabaseLocation + " must exist."
            End If

            'DatabaseName cannot be blank 
            If strDatabaseName = "" Then
                strMessage = "DatabaseName cannot be blank."
                If strReturn = "" Then
                    strReturn = strMessage
                Else
                    strReturn = strReturn + vbCrLf + strMessage
                End If
            End If
            strAlphaList = "||A||B||C||D||E||F||G||H||I||J||K||L||M||N||O||P||Q||R||S||T||U||V||W||X||Y||Z||"
            strCompareString = UCase("||" + Trim(Left(strDatabaseName, 1)) + "||")
            If (InStr(strAlphaList, strCompareString) = 0) Then
                strMessage = "DatabaseName must begin with a letter, not a number or other character."
                If strReturn = "" Then
                    strReturn = strMessage
                Else
                    strReturn = strReturn + vbCrLf + strMessage
                End If
            End If

            'SQLDataSource cannot be blank 
            If strSQLDataSource = "" Then
                strMessage = "SQLDataSource cannot be blank."
                If strReturn = "" Then
                    strReturn = strMessage
                Else
                    strReturn = strReturn + vbCrLf + strMessage
                End If
            End If

            'AddToExisting cannot be null.  This actually can't happen
            '    since it defaults to false
            If IsDBNull(boolAddToExisting) Then
                strMessage = "AddToExisting cannot be null."
                If strReturn = "" Then
                    strReturn = strMessage
                Else
                    strReturn = strReturn + vbCrLf + strMessage
                End If
            End If

            If Not My.Computer.FileSystem.DirectoryExists(strInputNHDPlusLocation) Then
                strMessage = "InputNHDPlusLocation must exist."
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
    Private Function LoadDB(ByRef sqlconConnection As SqlConnection) As String

        Dim varSs_time As Object
        Dim varEs_time As Object
        Dim strSQL As String
        Dim strReturn As String
        Dim strElapsedTime As String
        Dim strExMsg As String
        Dim strMegaDivFieldList As String
        Dim strPlusFlowFieldList As String
        Dim strPlusFlowlineVAAFieldList As String

        LoadDB = ""

        Try

            varSs_time = Now() 'Set starting time

            strMegaDivFieldList = " FromNHDPID, ToNHDPID, VPUID "
            strPlusFlowFieldList = " FromNHDPID, ToNHDPID, FromHydseq, ToHydSeq, FromLvlpat, ToLvlpat, " + _
               " nodenumber, deltalevel, direction, " + _
               " gapdistkm, hasgeo, FromVPUID, ToVPUID, FromPermid, ToPermid "
            strPlusFlowlineVAAFieldList = " nhdplusid, streamleve, streamorde, " + _
               " streamcalc, fromnode, tonode, hydroseq, levelpathi, pathlength, " + _
               " terminalpa, arbolatesu, divergence, startflag, terminalfl, " + _
               " uplevelpat, uphydroseq, dnlevel, dnlevelpat, dnhydroseq, dnminorhyd, " + _
               " dndraincou, frommeas, tomeas, rtndiv, thinner, vpuin, vpuout, reachcode, lengthkm, areasqkm, " + _
               " totdasqkm, divdasqkm, gnis_id, " + _
               " MaxElevRaw, MinElevRaw, MaxElevSmo, MinElevSmo, Slope, SlopeLenkm, ElevFixed, HWType, HWNodeSqKM, StatusFlag, VPUID, fcode "

            If boolAddToExisting = True Then
                strSQL = " SELECT *  INTO tVAA FROM PlusFlowlineVAA; " + vbCrLf + _
                         " SELECT *  INTO tFlow FROM PlusFlow; " + vbCrLf + _
                         " SELECT *  INTO tMega FROM Megadiv; "
                strReturn = ExecuteSQL(strSQL, sqlconConnection, gintSQLCommandTimeOut)
                If strReturn <> "" Then
                    LoadDB = strReturn
                    Exit Try
                End If
            End If

            strReturn = CopyArcGisTable(strInputNHDPlusLocation & "\NHDPlusFlowlineVAA", strDatabaseLocation & "\" & strDatabaseName & ".sde\" & strDatabaseName & ".dbo.PlusFlowlineVAA")
            If strReturn <> "" Then
                LoadDB = strReturn
                Exit Try
            End If

            strReturn = CopyArcGisTable(strInputNHDPlusLocation & "\NHDPlusFlow", strDatabaseLocation & "\" & strDatabaseName & ".sde\" & strDatabaseName & ".dbo.PlusFlow")
            If strReturn <> "" Then
                LoadDB = strReturn
                Exit Try
            End If

            strSQL = "DELETE FROM PlusFlow WHERE direction <> 709; "
            strReturn = ExecuteSQL(strSQL, sqlconConnection, gintSQLCommandTimeOut)
            If strReturn <> "" Then
                LoadDB = strReturn
                Exit Try
            End If

            strReturn = CopyArcGisTable(strInputNHDPlusLocation & "\NHDPlusMegaDiv", strDatabaseLocation & "\" & strDatabaseName & ".sde\" & strDatabaseName & ".dbo.MegaDiv")
            If strReturn <> "" Then
                LoadDB = strReturn
                Exit Try
            End If

            strReturn = CopyArcGisTable(strInputNHDPlusLocation & "\Hydrography\NHDFlowline", strDatabaseLocation & "\" & strDatabaseName & ".sde\" & strDatabaseName & ".dbo.NHDFlowline")
            If strReturn <> "" Then
                LoadDB = strReturn
                Exit Try
            End If

            strSQL = " ALTER TABLE PlusFlow ADD fromhydseq decimal(15,0) NULL; " + vbCrLf + _
                     " ALTER TABLE PlusFlow ADD fromlvlpat decimal(15,0) NULL; " + vbCrLf + _
                     " ALTER TABLE PlusFlow ADD tohydseq decimal(15,0) NULL; " + vbCrLf + _
                     " ALTER TABLE PlusFlow ADD tolvlpat decimal(15,0) NULL; " + vbCrLf + _
                     " ALTER TABLE PlusFlowLineVAA ADD gnis_id nvarchar(10) NULL; " + vbCrLf + _
                     " ALTER TABLE PlusFlowlineVAA ADD LengthKM decimal(38,8) NULL; " + vbCrLf + _
                     " ALTER TABLE PlusFlowlineVAA ADD Fcode int NULL; "
            strReturn = ExecuteSQL(strSQL, sqlconConnection, gintSQLCommandTimeOut)
            If strReturn <> "" Then
                LoadDB = strReturn
                Exit Try
            End If

            strSQL = " UPDATE PlusFlow SET fromhydseq = b.hydroseq FROM PlusFlowlinevaa as b WHERE b.nhdplusid = fromnhdpid; " + vbCrLf + _
                     " UPDATE PlusFlow SET fromlvlpat = b.levelpathi FROM PlusFlowlinevaa as b WHERE b.nhdplusid = fromnhdpid; " + vbCrLf + _
                     " UPDATE PlusFlow SET tohydseq = b.hydroseq FROM PlusFlowlinevaa as b WHERE b.nhdplusid = tonhdpid; " + vbCrLf + _
                     " UPDATE PlusFlowlineVAA SET PlusFlowlineVAA.lengthkm = b.lengthkm, PlusFlowlineVAA.gnis_id = b.gnis_id, PlusFlowlineVAA.fcode = b.fcode FROM NHDFlowline as b WHERE b.nhdplusid = PlusFlowlineVAA.nhdplusid; " + vbCrLf + _
                     " UPDATE PlusFlow SET toLvlpat = b.levelpathi FROM PlusFlowlinevaa as b WHERE b.nhdplusid = tonhdpid; "
            strReturn = ExecuteSQL(strSQL, sqlconConnection, gintSQLCommandTimeOut)
            If strReturn <> "" Then
                LoadDB = strReturn
                Exit Try
            End If

            strSQL = " DROP TABLE NHDFlowline; "
            strReturn = ExecuteSQL(strSQL, sqlconConnection, gintSQLCommandTimeOut)
            If strReturn <> "" Then
                LoadDB = strReturn
                Exit Try
            End If

            If boolAddToExisting = False Then
                strSQL = "CREATE PROCEDURE clearresultstable @tblname varchar(50), @isthere INT OUTPUT  AS BEGIN " +
                          "  SELECT @isthere = COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE' AND TABLE_NAME=@tblname END "
                strReturn = ExecuteSQL(strSQL, sqlconConnection, intSQLCommandTimeout)
                If strReturn <> "" Then
                    LoadDB = strReturn
                    Exit Try
                End If
            End If

            If boolAddToExisting = True Then
                strSQL = " INSERT INTO PlusFlowlineVAA  ( " & strPlusFlowlineVAAFieldList & ") SELECT " & strPlusFlowlineVAAFieldList & " FROM tVAA; " + vbCrLf + _
                         " INSERT INTO PlusFlow ( " & strPlusFlowFieldList & ") SELECT " & strPlusFlowFieldList & " FROM tFlow; " + vbCrLf + _
                         " INSERT INTO MegaDiv ( " & strMegaDivFieldList & ") SELECT " & strMegaDivFieldList & " FROM tMega; "
                strReturn = ExecuteSQL(strSQL, sqlconConnection, intSQLCommandTimeOut)
                If strReturn <> "" Then
                    LoadDB = strReturn
                    Exit Try
                End If
				
                strSQL = " DROP TABLE tVAA; " + vbCrLf + _
                         " DROP TABLE tFlow; " + vbCrLf + _
                         " DROP TABLE tMega; "
                strReturn = ExecuteSQL(strSQL, sqlconConnection, intSQLCommandTimeOut)
                If strReturn <> "" Then
                    LoadDB = strReturn
                    Exit Try
                End If
            End If
			
            strSQL = " CREATE INDEX iPlusFlowToNHDPID ON PlusFlow (ToNHDPID); " + vbCrLf + _
                     " CREATE INDEX iMegadivFromNHDPID ON megadiv (FromNHDPID); " + vbCrLf + _
                     " CREATE INDEX iPlusFlowlineVAAnhdplusid ON plusflowlinevaa (nhdplusid); " + vbCrLf + _
                     " CREATE INDEX iPlusFlowlineVAAHydroseq ON plusflowlinevaa (hydroseq); " + vbCrLf + _
                     " CREATE INDEX iPlusFlowlineVAATerminalpa ON plusflowlinevaa (terminalpa); " + vbCrLf + _
                     " CREATE INDEX iPlusFlowlineVAAPathlength ON plusflowlinevaa (pathlength); "
            strReturn = ExecuteSQL(strSQL, sqlconConnection, intSQLCommandTimeout)
            If strReturn <> "" Then
                LoadDB = strReturn
                Exit Try
            End If

            'RMD - 05/17/19 Add call to ClearArcGisWorkspaceCache to release locks
            strReturn = ClearArcGisWorkspaceCache(strDatabaseLocation & "\" & strDatabaseName & ".sde")
            If strReturn <> "" Then
                LoadDB = strReturn
                Exit Try
            End If
            strElapsedTime = GetElapsedTime(varSs_time)  'Calculate elapsed time

        Catch ex As Exception
            strExMsg = ex.Message.ToString + vbCrLf +
                ex.StackTrace.ToString + vbCrLf +
                ex.Source.ToString
            LoadDB = "LoadDB Exception: " + strExMsg

        Finally
            varSs_time = Nothing
            varEs_time = Nothing
            strSQL = Nothing
            strReturn = Nothing
            strElapsedTime = Nothing
            strExMsg = Nothing

        End Try

    End Function

End Class
