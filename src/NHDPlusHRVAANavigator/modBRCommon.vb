' 02/09/2015 V1009 RMD - Added parameter to Function CreateLocalDatabase: ByVal strSQLDataSource As String
' It contains the name of the localdb SQL Server instance (e.g., "(localdb)\v11.0")
' 02/09/2015 V1009 RMD - Removed variable Private strINIFileName As String - no longer needed
' 02/03/2015 V1009 RMD - Move declaration of GP Geoprocessor object to only the functions that needed it
' 01/05/2015 V1003 RMD - Added If gp.MaxSeverity > 1 to Function ReturnMessage
' 12/18/2024 V1002 RMD - Changed Application.StartupPath to System.Environment.CurrentDirectory to work for DLLs
'  Made Function RunTool and Function ReturnMessages Public instead of Private
'  Removed calls to MessageBox()
' 11/26/2014 V1001 RMD - Original Version

Imports System
Imports System.IO
Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System.Data.OleDb
Imports System.Data

Imports ESRI.ArcGIS.DataManagementTools
Imports ESRI.ArcGIS.Geoprocessor
Imports ESRI.ArcGIS.esriSystem

Module modBRCommon
    Private intProcessStatus As String
    Private strProcessMessage As String
    Private blnOverwriteOutput As Boolean
    Private blnTemporaryMapLayers As Boolean
    Private blnAddOutputsToMap As Boolean

    '---------------------------------------------
    ' Function Name : CreateArcGisDatabaseConnection
    ' Purpose       : Creates a SDE connection file that can be used to connect to SQL Server database.
    ' Input         : strOut_Folder_Path - The folder path where the .sde file will be stored. 
    '               : strOut_NameOfSDEfile - The name of the .sde file. The output file extension must end with .sde. 
    '               : strDatabase_Platform - "SQL_SERVER" — For connecting to Microsoft SQL Server
    '               : strInstance - The name of the SQL Server instance e.g., "(localdb)\v11.0"
    '               : strDataBaseName - The name of the database that you will be connecting to
    ' Output        : A sde file that contains the connection information to a specific SQL Server database
    ' Comments      : This function is based on the Create Database Connection (Data Management\Workspace) Geoprocessing tool
    '               : The SQL Server instance must be running and the database must exist
    '               : The database can be created with Function CreateLocalDatabase
    ' Created       : 11/25/2014
    ' Modified      : 
    '---------------------------------------------
    Public Function CreateArcGisDatabaseConnection(ByVal strOut_Folder_Path As String,
                                                   ByVal strOut_NameOfSDEfile As String,
                                                   ByVal strDatabase_Platform As String,
                                                   ByVal strInstance As String,
                                                   ByVal strDataBaseName As String) As String
        Dim GP As Geoprocessor = New Geoprocessor
        Dim CreateDatabaseConnection As CreateDatabaseConnection = New CreateDatabaseConnection()

        Try
            CreateArcGisDatabaseConnection = ""

            'Get the current ESRI environment variables 
            blnOverwriteOutput = GP.OverwriteOutput
            blnTemporaryMapLayers = GP.TemporaryMapLayers
            blnAddOutputsToMap = GP.AddOutputsToMap

            'Set them for this function
            GP.OverwriteOutput = True
            GP.TemporaryMapLayers = True
            GP.AddOutputsToMap = False

            CreateDatabaseConnection.out_folder_path = strOut_Folder_Path
            CreateDatabaseConnection.out_name = strOut_NameOfSDEfile & ".sde"
            CreateDatabaseConnection.database_platform = strDatabase_Platform
            CreateDatabaseConnection.instance = strInstance
            CreateDatabaseConnection.account_authentication = "OPERATING_SYSTEM_AUTH"
            CreateDatabaseConnection.database = strDataBaseName
            If Not RunTool(GP, CreateDatabaseConnection, Nothing) Then
                intProcessStatus = 900
                CreateArcGisDatabaseConnection = "CreateArcGisDatabaseConnection failed with error: " & strProcessMessage
                Exit Try
            End If

        Catch ex As Exception
            CreateArcGisDatabaseConnection = "CreateArcGisDatabaseConnection Exception: " & vbCrLf & _
                ex.ToString & vbCrLf & _
                ex.StackTrace.ToString & vbCrLf & _
                ex.Source.ToString
        Finally
            'Set the ArcGIS parameter back to their original settings
            GP.OverwriteOutput = blnOverwriteOutput
            GP.TemporaryMapLayers = blnTemporaryMapLayers
            GP.AddOutputsToMap = blnAddOutputsToMap

            CreateDatabaseConnection = Nothing
            GP = Nothing
        End Try
    End Function

    '---------------------------------------------
    ' Function Name : CopyArcGisTable
    ' Purpose       : Writes the rows from an input table, table view, feature class, or feature layer to a new table.  See Copy Rows (Data Management\Table) Geoprocessing tool for details
    '               : If a selection is defined on a feature class or feature layer in ArcMap, only the selected rows are copied out.  SELECTION PART HAS NOT BEEN TESTED YET.
    ' Input         : strIn_Table - The table name of the rows to be copied 
    ' Output        : strOut_Table - The name of the output table
    ' Comments      : This function is based on the Copy Rows (Data Management\Table) Geoprocessing tool
    '               : If copying to or from a SDE geodatabase, the path to the SDE file that contains the connection properties of the geodatabase must be specified e.g,
    '               : strTempWorkAreaPath & "\" & strSessionID & "\" & "TestSQLdb" & strSessionID & ".sde\" & "TestSQLdb" & strSessionID & ".dbo.tblNHDPlus"
    '               : 
    '               : Note: This does not work for feature classes.  Use CopyArcGisFeature instead
    '               : 
    ' Created       : 11/25/2014
    ' Modified      : 
    '---------------------------------------------
    Public Function CopyArcGisTable(ByVal strIn_Table As String,
                                   ByVal strOut_Table As String) As String
        Dim GP As Geoprocessor = New Geoprocessor
        Dim CopyRows As CopyRows = New CopyRows()

        Try
            CopyArcGisTable = ""

            'Get the current ESRI environment variables 
            blnOverwriteOutput = GP.OverwriteOutput
            blnTemporaryMapLayers = GP.TemporaryMapLayers
            blnAddOutputsToMap = GP.AddOutputsToMap

            'Set them for this function
            GP.OverwriteOutput = True
            GP.TemporaryMapLayers = True
            GP.AddOutputsToMap = False

            CopyRows.in_rows = strIn_Table
            CopyRows.out_table = strOut_Table
            If Not RunTool(GP, CopyRows, Nothing) Then
                intProcessStatus = 900
                CopyArcGisTable = "CopyArcGisTable failed with error: " & strProcessMessage
                Exit Try
            End If

        Catch ex As Exception
            CopyArcGisTable = "CopyArcGisTable Exception: " & vbCrLf & _
                ex.ToString & vbCrLf & _
                ex.StackTrace.ToString & vbCrLf & _
                ex.Source.ToString
        Finally
            'Set the ArcGIS parameter back to their original settings
            GP.OverwriteOutput = blnOverwriteOutput
            GP.TemporaryMapLayers = blnTemporaryMapLayers
            GP.AddOutputsToMap = blnAddOutputsToMap

            CopyRows = Nothing
            GP = Nothing

        End Try
    End Function

    '---------------------------------------------
    ' Function Name : CopyArcGisFeature
    ' Purpose       : Copies features from the input feature class or layer to a new feature class.  See Copy Features (Data Management\Features) Geoprocessing tool for details
    '               : If the input is a layer which has a selection, only the selected features will be copied. 
    '               : If the input is a geodatabase feature class or shapefile, all features will be copied.  SELECTION PART HAS NOT BEEN TESTED YET.
    ' Input         : strIn_Features - The features to be copied. 
    ' Output        : strOut_Feature_Class - The name of the output feature class
    ' Comments      : This function is based on the Copy Features (Data Management\Features) Geoprocessing tool
    '               : If copying to or from a SDE geodatabase, the path to the SDE file that contains the connection properties of the geodatabase must be specified e.g,
    '               : strTempWorkAreaPath & "\" & strSessionID & "\" & "TestSQLdb" & strSessionID & ".sde\" & "TestSQLdb" & strSessionID & ".dbo.fcNHDFlowline"
    '               : 
    ' Created       : 11/25/2014
    ' Modified      : 
    '---------------------------------------------
    Public Function CopyArcGisFeature(ByVal strIn_Features As String,
                               ByVal strOut_Feature_Class As String) As String
        Dim GP As Geoprocessor = New Geoprocessor
        Dim CopyFeatures As CopyFeatures = New CopyFeatures()

        Try
            CopyArcGisFeature = ""

            'Get the current ESRI environment variables 
            blnOverwriteOutput = GP.OverwriteOutput
            blnTemporaryMapLayers = GP.TemporaryMapLayers
            blnAddOutputsToMap = GP.AddOutputsToMap

            'Set them for this function
            GP.OverwriteOutput = True
            GP.TemporaryMapLayers = True
            GP.AddOutputsToMap = False

            CopyFeatures.in_features = strIn_Features
            CopyFeatures.out_feature_class = strOut_Feature_Class
            If Not RunTool(GP, CopyFeatures, Nothing) Then
                intProcessStatus = 900
                CopyArcGisFeature = "CopyArcGisFeature failed with error: " & strProcessMessage
                Exit Try
            End If

        Catch ex As Exception
            CopyArcGisFeature = "CopyArcGisFeature Exception: " & vbCrLf & _
                ex.ToString & vbCrLf & _
                ex.StackTrace.ToString & vbCrLf & _
                ex.Source.ToString
        Finally
            'Set the ArcGIS parameter back to their original settings
            GP.OverwriteOutput = blnOverwriteOutput
            GP.TemporaryMapLayers = blnTemporaryMapLayers
            GP.AddOutputsToMap = blnAddOutputsToMap

            CopyFeatures = Nothing
            GP = Nothing
        End Try
    End Function

    '---------------------------------------------
    ' Function Name : DeleteArcGisData
    ' Purpose       : Copies features from the input feature class or layer to a new feature class.  See Copy Features (Data Management\Features) Geoprocessing tool for details
    '               : If the input is a layer which has a selection, only the selected features will be copied. 
    '               : If the input is a geodatabase feature class or shapefile, all features will be copied.  SELECTION PART HAS NOT BEEN TESTED YET.
    ' Input         : strIn_data - The input data to be deleted
    ' Output        : None
    ' Comments      : This function is based on the Delete (Data Management\General) Geoprocessing tool
    '               : Permanently deletes data from disk. All types of geographic data supported by ArcGIS, as well as toolboxes and workspaces (folders, geodatabases), can be deleted.
    '               : If the specified item is a workspace, all contained items are also deleted
    '               : 
    ' Created       : 11/25/2014
    ' Modified      : 
    '---------------------------------------------
    Public Function DeleteArcGisData(ByVal strIn_data As String) As String
        Dim GP As Geoprocessor = New Geoprocessor
        Dim Delete As Delete = New Delete()

        Try
            DeleteArcGisData = ""

            'Get the current ESRI environment variables 
            blnOverwriteOutput = GP.OverwriteOutput
            blnTemporaryMapLayers = GP.TemporaryMapLayers
            blnAddOutputsToMap = GP.AddOutputsToMap

            'Set them for this function
            GP.OverwriteOutput = True
            GP.TemporaryMapLayers = True
            GP.AddOutputsToMap = False

            Delete.in_data = strIn_data
            If Not RunTool(GP, Delete, Nothing) Then
                intProcessStatus = 900
                DeleteArcGisData = "DeleteArcGisData failed with error: " & strProcessMessage
                Exit Try
            End If

        Catch ex As Exception
            DeleteArcGisData = "DeleteArcGisData Exception: " & vbCrLf & _
                ex.ToString & vbCrLf & _
                ex.StackTrace.ToString & vbCrLf & _
                ex.Source.ToString
        Finally
            'Set the ArcGIS parameter back to their original settings
            GP.OverwriteOutput = blnOverwriteOutput
            GP.TemporaryMapLayers = blnTemporaryMapLayers
            GP.AddOutputsToMap = blnAddOutputsToMap

            Delete = Nothing
            GP = Nothing
        End Try
    End Function

    '---------------------------------------------
    ' Function Name : ClearArcGisWorkspaceCache
    ' Purpose       : Removes the connection to the 
    ' Input         : strIn_data - The path and name of the .sde file used to create a database connection
    ' Output        : None
    ' Comments      : This function is based on the Clear Workspace Cache (Data Management\Workspace) Geoprocessing tool
    '               : This does not delete the .SDE file
    '               : If strIn_data is empty/blank the function will remove all connections to all SDE geodatabases
    ' Created       : 11/25/2014
    ' Modified      : 
    '---------------------------------------------
    Public Function ClearArcGisWorkspaceCache(ByVal strIn_data As String) As String
        Dim GP As Geoprocessor = New Geoprocessor
        Dim ClearWorkspaceCache As ClearWorkspaceCache = New ClearWorkspaceCache()
        Try
            ClearArcGisWorkspaceCache = ""

            'Get the current ESRI environment variables 
            blnOverwriteOutput = GP.OverwriteOutput
            blnTemporaryMapLayers = GP.TemporaryMapLayers
            blnAddOutputsToMap = GP.AddOutputsToMap

            'Set them for this function
            GP.OverwriteOutput = True
            GP.TemporaryMapLayers = True
            GP.AddOutputsToMap = False

            ClearWorkspaceCache.in_data = strIn_data
            If Not RunTool(GP, ClearWorkspaceCache, Nothing) Then
                intProcessStatus = 900
                ClearArcGisWorkspaceCache = "ClearArcGisWorkspaceCache failed with error: " & strProcessMessage
                Exit Try
            End If

        Catch ex As Exception
            ClearArcGisWorkspaceCache = "ClearArcGisWorkspaceCache Exception: " & vbCrLf & _
                ex.ToString & vbCrLf & _
                ex.StackTrace.ToString & vbCrLf & _
                ex.Source.ToString
        Finally
            'Set the ArcGIS parameter back to their original settings
            GP.OverwriteOutput = blnOverwriteOutput
            GP.TemporaryMapLayers = blnTemporaryMapLayers
            GP.AddOutputsToMap = blnAddOutputsToMap

            ClearWorkspaceCache = Nothing
            GP = Nothing

        End Try
    End Function

    Public Function RunTool(ByRef geoprocessor As Geoprocessor, ByRef process As IGPProcess, ByRef TC As ITrackCancel) As Boolean
        RunTool = False
        Try
            geoprocessor.Execute(process, Nothing)
            If Not ReturnMessages(geoprocessor, process) Then Exit Function
        Catch err As Exception
            Console.WriteLine(err.Message)
            If Not ReturnMessages(geoprocessor, process) Then Exit Function
        End Try

        RunTool = True

    End Function

    'Function for returning the tool messages.
    Public Function ReturnMessages(ByRef gp As Geoprocessor, ByRef nprocess As IGPProcess) As Boolean
        ReturnMessages = False
        Dim Count As Integer
        Dim intCount As Integer

        If gp.MaxSeverity > 1 Then
            If gp.MessageCount > 0 Then
                For Count = 0 To gp.MessageCount - 1
                    'Debug.Print(gp.GetMessage(Count))
                    If Len(gp.GetMessages(2)) > 0 Or gp.GetSeverity(Count) > 0 Then
                        strProcessMessage = nprocess.ToolName & " failed with the following error" & vbCrLf
                        For intCount = 0 To gp.MessageCount - 1
                            strProcessMessage = strProcessMessage & vbCrLf & gp.GetMessage(intCount)
                        Next
                        Exit Function
                    End If
                Next
            End If
        End If

        ReturnMessages = True

    End Function

    Public Function ExecuteSQL(ByVal strSQL As String, ByRef sqlconConnection As SqlConnection, ByVal intTimeout As Integer) As String
        Dim sqlcmdCommand As SqlCommand
        Dim strConnectionString As String
        Try
            sqlcmdCommand = New SqlCommand(strSQL, sqlconConnection)
            sqlcmdCommand.CommandTimeout = intTimeout
            strConnectionString = sqlconConnection.ConnectionString
            sqlcmdCommand.ExecuteNonQuery()
            ExecuteSQL = ""
        Catch ex As Exception
            ExecuteSQL = "ExecuteSQL:  SQL Exception: " & vbCrLf & strSQL & vbCrLf & _
                ex.ToString & vbCrLf & _
                ex.StackTrace.ToString & vbCrLf & _
                ex.Source.ToString
        Finally
            sqlcmdCommand = Nothing
        End Try
    End Function

    Public Function CreateLocalDatabase(ByVal strWorkingDB As String,
                                        ByVal strWorkingPath As String,
                                        ByVal intSQLConnectionTimeout As Integer,
                                        ByVal intSQLCommandTimeOut As Integer,
                                        ByVal strSQLDataSource As String) As String
        Dim strSQL As String
        Dim strReturn As String
        Try
            CreateLocalDatabase = ""
            Using sqlconConnection As SqlConnection = New SqlConnection("Data Source=" & strSQLDataSource & ";Integrated Security=True;Connect Timeout=" & intSQLConnectionTimeout & ";")
                sqlconConnection.Open()

                strSQL = "CREATE DATABASE " & strWorkingDB & " ON PRIMARY" & _
                         "(Name=TestWorking, filename = '" & strWorkingPath & "\" & strWorkingDB & ".mdf', filegrowth=10%)log on" & _
                         "(name=" & strWorkingDB & "_log, filename='" & strWorkingPath & "\" & strWorkingDB & ".ldf',filegrowth=1)"
                strReturn = ExecuteSQL(strSQL, sqlconConnection, intSQLCommandTimeOut)
                If strReturn <> "" Then
                    CreateLocalDatabase = strReturn
                    Exit Try
                End If

                strSQL = "ALTER DATABASE " & strWorkingDB & " SET RECOVERY SIMPLE WITH no_wait "
                strReturn = ExecuteSQL(strSQL, sqlconConnection, intSQLCommandTimeOut)
                If strReturn <> "" Then
                    CreateLocalDatabase = strReturn
                    Exit Try
                End If

                strSQL = "ALTER DATABASE " & strWorkingDB & " SET AUTO_SHRINK ON "
                strReturn = ExecuteSQL(strSQL, sqlconConnection, intSQLCommandTimeOut)
                If strReturn <> "" Then
                    CreateLocalDatabase = strReturn
                    Exit Try
                End If

                sqlconConnection.Close()
            End Using

        Catch ex As Exception
            CreateLocalDatabase = "SQLRunTime Exception in CreateLocalDatabase.  " & ex.Message.ToString & _
                ex.StackTrace.ToString & vbCrLf & _
                ex.Source.ToString
        End Try

    End Function

    Public Function ImportTextFile(ByVal strFullPathFileName As String, ByVal strTableName As String,
                                  ByVal strSQLFields As String, ByVal strSQLGroup As String,
                                  ByVal strSQLWhere As String, ByVal strWorkingDB As String,
                                  ByVal intSQLConnectionTimeout As Integer, intSQLCommandTimeOut As Integer,
                                  ByVal intSQLBatchSize As Integer) As String
        Try

            ImportTextFile = ""
            Dim strReturn As String = ""
            Dim strRetSQL As String
            Dim strQuoteCheckLine As String = ""
            Dim blnCreatNewFile As Boolean = False
            Dim strFileNameNoExt = Path.GetFileNameWithoutExtension(strFullPathFileName)
            Dim strFileName = Path.GetFileName(strFullPathFileName)
            Dim strPathFileName = My.Computer.FileSystem.GetParentPath(strFullPathFileName)
            Dim strNewFullPathFileName As String = strPathFileName & "\" & strFileNameNoExt & "QC.txt"

            'Use OleDbDataReader and a DataTable to create the table in SQL Server.
            'We can't use it to load the rest of the table because it runs out of memory on really large files
            Dim connectionFile As OleDbConnection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strPathFileName & "; Extended Properties='text;HDR=Yes;FMT=Delimited'")
            Dim strSQL = "SELECT TOP 1 " & strSQLFields & " FROM " & strFileName & strSQLWhere & strSQLGroup
            Dim command As OleDbCommand = New OleDbCommand(strSQL, connectionFile)
            connectionFile.Open()
            Dim ImportedDataTable As DataTable = New DataTable()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            ImportedDataTable.Load(reader)
            reader.Close()
            connectionFile.Close()

            'Check to see if the text file to be imported contains any double quotes
            'If so, we need to remove them because BULK INSERT loads the quotes into the table
            Using srQuoteCheck As StreamReader = New StreamReader(strFullPathFileName)
                Using srRemovedQuotes As StreamWriter = New StreamWriter(strNewFullPathFileName)
                    'Skip past the header in srQuoteCheck
                    strQuoteCheckLine = srQuoteCheck.ReadLine()
                    'Check the second line
                    If srQuoteCheck.EndOfStream Then
                        'There is no second line, only a header and we have nothing to process
                        ImportTextFile = "ImportTextFile Failed. There is no data to process in " & strFullPathFileName
                        Exit Try
                    Else
                        'Read the second line
                        strQuoteCheckLine = srQuoteCheck.ReadLine()
                        'We only need to create a new file if the input file contains quotes
                        If strQuoteCheckLine.Contains(Chr(34)) Then
                            'Create a new file without double quotes and no header
                            Do While Not strQuoteCheckLine Is Nothing
                                strQuoteCheckLine = strQuoteCheckLine.Replace(Chr(34), "")
                                srRemovedQuotes.WriteLine(strQuoteCheckLine)
                                strQuoteCheckLine = srQuoteCheck.ReadLine()
                            Loop
                        Else
                            'No quotes, so just use the original file
                            strNewFullPathFileName = strFullPathFileName
                        End If
                    End If
                End Using 'srRemovedQuotes
            End Using 'srQuoteCheck

            Using destinationConnection As SqlConnection = New SqlConnection("Data Source=(LocalDB)\v11.0; Integrated Security=True;Connect Timeout=" & intSQLConnectionTimeout & ";Initial Catalog=" & strWorkingDB)
                destinationConnection.Open()
                strRetSQL = GetCreateFromDataTableSQL(strTableName, ImportedDataTable)
                Dim cmd As SqlCommand = New SqlCommand(strRetSQL, destinationConnection)
                cmd.ExecuteNonQuery()

                'Remove the one record in the SQL Server table that was loaded when the table was created with ImportedDataTable
                strSQL = "TRUNCATE TABLE " & strTableName & ";"
                strReturn = ExecuteSQL(strSQL, destinationConnection, 0)
                If strReturn <> "" Then
                    ImportTextFile = strReturn
                End If

                'Insert the records from the text file
                strSQL = "BULK INSERT " & strTableName & _
                         " FROM '" & strNewFullPathFileName & "'" & _
                         " with (FIELDTERMINATOR  = ',', ROWTERMINATOR  = '\n', BATCHSIZE = " & intSQLBatchSize & ");"
                strReturn = ExecuteSQL(strSQL, destinationConnection, 0)
                'This to cover the possiblity that we loaded the original file and that it contained a header
                If Not strReturn.Contains("Bulk load data conversion error (type mismatch or invalid character for the specified codepage) for row 1, column 1") Then
                    ImportTextFile = strReturn
                End If

                destinationConnection.Close()
            End Using

            reader = Nothing
            connectionFile.Dispose()
            ImportedDataTable.Dispose()

        Catch ex As Exception
            ImportTextFile = "SQLRunTime Exception in ImportTextFile." & vbCrLf & _
               ex.ToString & vbCrLf & _
               ex.StackTrace.ToString & vbCrLf & _
               ex.Source.ToString
        End Try
    End Function

    'This has not been modified to process large DBFs yet.
    Public Function ImportDbfFile(ByVal strFullPathFileName As String, ByVal strWorkingDB As String,
                                  ByVal strTableName As String, ByVal intSQLConnectionTimeout As Integer,
                                  intSQLCommandTimeOut As Integer, ByVal strFields As String,
                                  ByVal strWhereclause As String, ByVal boolAppend As Boolean,
                                  ByVal intSQLBulkCopyTimeout As Integer, ByVal intSQLBatchSize As Integer) As String
        Try

            ImportDbfFile = ""
            Dim strRetSQL As String
            Dim strFileNameNoExt = Path.GetFileNameWithoutExtension(strFullPathFileName)
            Dim strPathFileName = My.Computer.FileSystem.GetParentPath(strFullPathFileName)
            Dim strSQL As String = "SELECT " & strFields & " FROM " & strFileNameNoExt & strWhereclause
            Dim connectionFile As OleDbConnection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strPathFileName & ";Extended Properties=dbase IV;User ID=Admin;")
            Dim command As OleDbCommand = New OleDbCommand(strSQL, connectionFile)
            connectionFile.Open()

            Dim ImportedDataTable As DataTable = New DataTable()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            ImportedDataTable.Load(reader)

            Using destinationConnection As SqlConnection = New SqlConnection("Data Source=(LocalDB)\v11.0;Integrated Security=True;Connect Timeout=" & intSQLConnectionTimeout & ";Initial Catalog=" & strWorkingDB)

                destinationConnection.Open()
                If Not boolAppend Then
                    strRetSQL = GetCreateFromDataTableSQL(strTableName, ImportedDataTable)
                    Dim cmd As SqlCommand = New SqlCommand(strRetSQL, destinationConnection)
                    cmd.ExecuteNonQuery()
                End If

                ' Set up the bulk copy object.  
                ' The column positions in the source data reader  
                ' match the column positions in the destination table,  
                ' so there is no need to map columns. 
                Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(destinationConnection)
                    bulkCopy.DestinationTableName = "dbo." & strTableName

                    ' Set the timeout.
                    bulkCopy.BulkCopyTimeout = intSQLBulkCopyTimeout
                    ' Set the BatchSize.
                    bulkCopy.BatchSize = intSQLBatchSize

                    Try
                        ' Write from the source to the destination.
                        bulkCopy.WriteToServer(ImportedDataTable)
                    Catch ex As Exception
                        ImportDbfFile = "SQLRunTime Exception in Importdbffile BULKCOPY." & vbCrLf & _
                        ex.ToString & vbCrLf & _
                            ex.StackTrace.ToString & vbCrLf & _
                            ex.Source.ToString
                    End Try
                End Using
                destinationConnection.Close()
            End Using

            reader.Close()
            reader = Nothing
            connectionFile.Dispose()
            ImportedDataTable.Dispose()

        Catch ex As Exception
            ImportDbfFile = "SQLRunTime Exception in ImportDbfFile." & vbCrLf & _
               ex.ToString & vbCrLf & _
               ex.StackTrace.ToString & vbCrLf & _
               ex.Source.ToString
        End Try
    End Function

    Public Function GetCreateFromDataTableSQL(tableName As String, table As DataTable) As String
        Dim sql As String = "CREATE TABLE [" & tableName & "] (" & vbLf
        ' columns
        For Each column As DataColumn In table.Columns
            sql &= "[" & column.ColumnName & "] " & SQLGetType(column) & "," & vbLf
        Next
        sql = sql.TrimEnd(New Char() {","c, ControlChars.Lf}) & vbLf
        ' primary keys
        If table.PrimaryKey.Length > 0 Then
            sql &= "CONSTRAINT [PK_" & tableName & "] PRIMARY KEY CLUSTERED ("
            For Each column As DataColumn In table.PrimaryKey
                sql &= "[" & column.ColumnName & "],"
            Next
            sql = sql.TrimEnd(New Char() {","c}) & "))" & vbLf
        End If

        'if not ends with ")"
        If (table.PrimaryKey.Length = 0) AndAlso (Not sql.EndsWith(")")) Then
            sql &= ")"
        End If

        Return sql
    End Function

    Public Function SQLGetType(type As Object, columnSize As Integer, numericPrecision As Integer, numericScale As Integer) As String
        Select Case type.ToString()
            Case "System.Byte[]"
                Return "VARBINARY(MAX)"

            Case "System.Boolean"
                Return "BIT"

            Case "System.DateTime"
                Return "DATETIME"

            Case "System.DateTimeOffset"
                Return "DATETIMEOFFSET"

            Case "System.Decimal"
                If numericPrecision <> -1 AndAlso numericScale <> -1 Then
                    Return "DECIMAL(" & numericPrecision & "," & numericScale & ")"
                Else
                    Return "DECIMAL"
                End If

            Case "System.Double"
                Return "FLOAT"

            Case "System.Single"
                Return "REAL"

            Case "System.Int64"
                Return "BIGINT"

            Case "System.Int32"
                Return "INT"

            Case "System.Int16"
                Return "SMALLINT"

            Case "System.String"
                Return "NVARCHAR(" & (If((columnSize = -1 OrElse columnSize > 8000), "MAX", columnSize.ToString())) & ")"

            Case "System.Byte"
                Return "TINYINT"

            Case "System.Guid"
                Return "UNIQUEIDENTIFIER"
            Case Else

                Throw New Exception(type.ToString() & " not implemented.")
        End Select
    End Function

    ' Overload based on row from schema table
    Public Function SQLGetType(schemaRow As DataRow) As String
        Dim numericPrecision As Integer
        Dim numericScale As Integer

        If Not Integer.TryParse(schemaRow("NumericPrecision").ToString(), numericPrecision) Then
            numericPrecision = -1
        End If
        If Not Integer.TryParse(schemaRow("NumericScale").ToString(), numericScale) Then
            numericScale = -1
        End If

        Return SQLGetType(schemaRow("DataType"), Integer.Parse(schemaRow("ColumnSize").ToString()), numericPrecision, numericScale)
    End Function

    ' Overload based on DataColumn from DataTable type
    Public Function SQLGetType(column As DataColumn) As String
        Return SQLGetType(column.DataType, column.MaxLength, -1, -1)
    End Function

    Public Function AttachDB(ByVal strDBName As String, ByVal intSQLConnectionTimeout As Integer) As String
        Try
            AttachDB = ""
            'Attach a specific database file to the localdb instance
            Using myLocalDB As SqlConnection = New SqlConnection("Data Source=(LocalDB)\v11.0;Integrated Security=True;Connect Timeout=" & intSQLConnectionTimeout & ";AttachDbFilename=" & strDBName & ";")
                myLocalDB.Open() 'You have to open the connection for the database to become attached
                myLocalDB.Close() 'Supposedly you don't have to explicitly close the database if you the open it inside a "Using, End Using" structure but it seems the safer thing to do
                'One of the advantages of the "Using, End Using" structure is that the declared object is automatically disposed of at the end.
            End Using

        Catch ex As Exception
            AttachDB = "Runtime Exception in AttachDB: " & ex.ToString & vbCrLf & _
                ex.StackTrace.ToString & vbCrLf & _
                ex.Source.ToString
        End Try

    End Function

    Public Function DetachDB(ByVal strDBName As String, ByVal intSQLConnectionTimeout As Integer) As String
        Try
            DetachDB = ""
            Dim strSQL As String
            'Detach a specific database file from  the localdb instance
            SqlConnection.ClearAllPools()
            Using myLocalDB As SqlConnection = New SqlConnection("Data Source=(LocalDB)\v11.0;Integrated Security=True;Connect Timeout=" & intSQLConnectionTimeout & ";")
                myLocalDB.Open()
                strSQL = " SELECT COUNT(*) FROM sys.databases WHERE name = '" & strDBName & "'"
                Dim cmd As SqlCommand = New SqlCommand(strSQL, myLocalDB)
                If cmd.ExecuteScalar > 0 Then
                    cmd.CommandText = "EXEC master.dbo.sp_detach_db @dbname = N'" & strDBName & "'"
                    cmd.ExecuteNonQuery()
                End If
                myLocalDB.Close()
            End Using

        Catch ex As Exception
            DetachDB = "Runtime Exception in DetachDB: " & ex.ToString & vbCrLf & _
                ex.StackTrace.ToString & vbCrLf & _
                ex.Source.ToString

        End Try
    End Function

    Public Function DropDB(ByVal strDBName As String, ByVal intSQLConnectionTimeout As Integer) As String
        Try
            DropDB = ""
            'NOTE: SqlConnection.ClearAllPools() is necessary if you have previously connected to a database and performed some action on it (like load a table).
            'Even though you closed the original connection, you will still get an error that says the database is in use.
            'The ClearAllPools method is called directly from the SqlConnection object and not from a object declared as a SqlConnection object as in:

            'Dim sqlconConnection As SqlConnection = New SqlConnection("Data Source=(LocalDB)\v11.0;Integrated Security=True;Connect Timeout=" & intSQLConnectionTimeout & ";") <- example only
            'sqlconConnection.ClearAllPools() <- This doesn't work
            SqlConnection.ClearAllPools() '<-This does..

            Dim strSQL As String

            'Drop a database from  the localdb instance
            SqlConnection.ClearAllPools()
            Using myLocalDB As SqlConnection = New SqlConnection("Data Source=(LocalDB)\v11.0;Integrated Security=True;Connect Timeout=" & intSQLConnectionTimeout & ";")
                myLocalDB.Open()
                strSQL = " SELECT COUNT(*) FROM sys.databases WHERE name = '" & strDBName & "'"
                Dim cmd As SqlCommand = New SqlCommand(strSQL, myLocalDB)
                If cmd.ExecuteScalar > 0 Then
                    cmd.CommandText = "DROP DATABASE " & strDBName
                    cmd.ExecuteNonQuery()
                End If
                myLocalDB.Close()
            End Using

        Catch ex As Exception
            DropDB = "Runtime Exception in DropDB: " & ex.ToString & vbCrLf & _
                ex.StackTrace.ToString & vbCrLf & _
                ex.Source.ToString
        End Try
    End Function

    Public Function Copy_Directory(ByVal strSrc As String, ByVal strDest As String, Optional ByVal strLogName As String = "")
        Dim dirInfo As New System.IO.DirectoryInfo(strSrc)
        Dim fsInfo As System.IO.FileSystemInfo
        Dim intFileCount As Integer = 0
        Dim strExMsg As String = ""

        Try
            If FolderExists(strSrc) Then
                If strLogName = "" Then
                    strLogName = System.Environment.CurrentDirectory & "\DontOpenThis.log"
                End If

                If Not System.IO.Directory.Exists(strDest) Then
                    System.IO.Directory.CreateDirectory(strDest)
                End If

                For Each fsInfo In dirInfo.GetFileSystemInfos
                    Dim strDestFileName As String = System.IO.Path.Combine(strDest, fsInfo.Name)
                    If TypeOf fsInfo Is System.IO.FileInfo Then
                        System.IO.File.Copy(fsInfo.FullName, strDestFileName, True)
                        'This will overwrite files that already exist
                    Else
                        Copy_Directory(fsInfo.FullName, strDestFileName, strLogName)
                    End If
                    intFileCount += 1
                Next
                If intFileCount = 0 Then
                    LogEvent("There are no files in folder: " & strSrc, strLogName, True)
                    Return "There are no files in folder: " & strSrc
                End If
            Else
                LogEvent("The source folder: " & strSrc & " does not exist.", strLogName, True)
                Return "The source folder: " & strSrc & " does not exist."
            End If

        Catch ex As Exception
            strExMsg = "Copy_Directory Exception: " & ex.Message.ToString & vbCrLf & _
                ex.StackTrace.ToString & vbCrLf & _
                ex.Source.ToString
            LogEvent(strExMsg, strLogName, True)
            Return strExMsg
        End Try
        Return ""

    End Function
    '---------------------------------------------
    ' Program Name  : Copy_Files
    ' Purpose       : Copies one or more files from one folder to another
    ' Input         : Source Folder Name, Target Folder Name, (Optional) FileName or Extension
    ' Output        : Copied files in target folder
    ' Comments      : Filename or Extension may include wildcard (*)
    '               : If blank, copies all files from source folder to target folder
    '               : With wildcard for extension: "MyFileName.*" copies all the files named MyFileName with any extension
    '               : With wildcard for filename: "*.DOC" copies all the files with an extension of .DOC
    '               : "MyFileName.DOC" copies only specified file to target folder
    ' Organization  : Horizon Systems Corporation
    ' Created       : 04/07/2009
    ' Modified      : 
    '---------------------------------------------
    Public Function Copy_Files(ByVal strSourceFolderName As String, ByVal strTargetFolderName As String, Optional ByVal strFileNameOrExtension As String = "", Optional ByVal strLogName As String = "") As String
        Dim intFileCount As Integer = 0
        Dim strExMsg As String = ""
        Try
            If strLogName = "" Then
                strLogName = System.Environment.CurrentDirectory & "\DontOpenThis.log"
            End If
            'Check to make sure the source folder exists
            If FolderExists(strSourceFolderName) Then
                'No file name specified: Copy all files in source folders
                If strFileNameOrExtension = "" Then
                    For Each FileToCopy As String In My.Computer.FileSystem.GetFiles(strSourceFolderName)
                        My.Computer.FileSystem.CopyFile(FileToCopy, strTargetFolderName & "\" & System.IO.Path.GetFileName(FileToCopy), True)
                        intFileCount += 1
                    Next
                    If intFileCount = 0 Then
                        LogEvent("There are no files in folder: " & strSourceFolderName, strLogName, True)
                        Return "There are no files in folder: " & strSourceFolderName
                    End If

                    'Wildcard for Extension: Copy all files with supplied filename
                ElseIf strFileNameOrExtension.EndsWith("*") Then
                    For Each FileToCopy As String In My.Computer.FileSystem.GetFiles(strSourceFolderName, FileIO.SearchOption.SearchTopLevelOnly, strFileNameOrExtension)
                        My.Computer.FileSystem.CopyFile(FileToCopy, strTargetFolderName & "\" & System.IO.Path.GetFileName(FileToCopy), True)
                        intFileCount += 1
                    Next
                    If intFileCount = 0 Then
                        LogEvent("There are no files in folder: " & strSourceFolderName & " with a file name beginning with " & strFileNameOrExtension, strLogName, True)
                        Return "There are no files in folder: " & strSourceFolderName & " with a file name beginning with " & strFileNameOrExtension
                    End If

                    'Wildcard for FileName: Copy all files with supplied extension
                ElseIf strFileNameOrExtension.StartsWith("*") Then
                    For Each FileToCopy As String In My.Computer.FileSystem.GetFiles(strSourceFolderName, FileIO.SearchOption.SearchTopLevelOnly, strFileNameOrExtension)
                        My.Computer.FileSystem.CopyFile(FileToCopy, strTargetFolderName & "\" & System.IO.Path.GetFileName(FileToCopy), True)
                        intFileCount += 1
                    Next
                    If intFileCount = 0 Then
                        LogEvent("There are no files in folder: " & strSourceFolderName & " with an extension of " & strFileNameOrExtension, strLogName, True)
                        Return "There are no files in folder: " & strSourceFolderName & " with an extension of " & strFileNameOrExtension
                    End If

                    'Copy specified file name
                Else
                    If My.Computer.FileSystem.FileExists(strSourceFolderName & "\" & strFileNameOrExtension) Then
                        My.Computer.FileSystem.CopyFile(strSourceFolderName & "\" & strFileNameOrExtension, strTargetFolderName & "\" & strFileNameOrExtension, True)
                    Else
                        LogEvent("The file: " & strSourceFolderName & "\" & strFileNameOrExtension & " does not exist!", strLogName, True)
                        Return "The file: " & strSourceFolderName & "\" & strFileNameOrExtension & " does not exist!"
                    End If
                End If
            Else
                LogEvent("The source folder: " & strSourceFolderName & " does not exist.", strLogName, True)
                Return "The source folder: " & strSourceFolderName & " does not exist."
            End If

        Catch ex As Exception
            strExMsg = "Copy_Files Exception: " & ex.Message.ToString & vbCrLf & _
                ex.StackTrace.ToString & vbCrLf & _
                ex.Source.ToString
            LogEvent(strExMsg, strLogName, True)
            Return strExMsg
        Finally
            intFileCount = Nothing
        End Try
        Return ""

    End Function

    Public Function FolderExists(ByVal FolderPath As String) As Boolean
        Dim f As New IO.DirectoryInfo(FolderPath)
        Return f.Exists
    End Function

    Public Function GetElapsedTime(ByVal startTime As DateTime) As String
        Dim elapsedTime As TimeSpan = Date.Now.Subtract(startTime)
        If elapsedTime.Days.ToString.Length <> 0 Then
            Return elapsedTime.Days.ToString & " Days " & elapsedTime.Hours.ToString & " Hours " & elapsedTime.Minutes.ToString & " Minutes " & elapsedTime.Seconds.ToString & " Seconds"
        ElseIf elapsedTime.Hours.ToString.Length <> 0 Then
            Return elapsedTime.Hours.ToString & "Hours " & elapsedTime.Minutes.ToString & " Minutes " & elapsedTime.Seconds.ToString & " Seconds"
        ElseIf elapsedTime.Minutes.ToString.Length <> 0 Then
            Return elapsedTime.Minutes.ToString & "Minutes " & elapsedTime.Seconds.ToString & " Seconds"
        End If
        Return elapsedTime.Seconds.ToString & "Seconds"
    End Function

    Public Sub LogEvent(ByVal logMessage As String, ByVal strLogName As String, ByVal blnHeader As Boolean)
        Try
            Using swLog As StreamWriter = File.AppendText(strLogName)
                If blnHeader Then
                    Dim strAppTitleVersion As String = My.Application.Info.Title & " " & My.Application.Info.Version.Major.ToString & "." & My.Application.Info.Version.Minor.ToString & "." & My.Application.Info.Version.Build.ToString & "." & My.Application.Info.Version.Revision.ToString
                    Dim strMonth As String = DateTime.Now.Month()
                    Dim strDay As String = DateTime.Now.Day()
                    If Len(strMonth) = 1 Then strMonth = "0" & strMonth
                    If Len(strDay) = 1 Then strDay = "0" & strDay
                    swLog.Write(ControlChars.CrLf)
                    swLog.WriteLine("{0}{1}{2} - {3} - {4} ", DateTime.Now.Year(), strMonth, strDay, DateTime.Now.ToLongTimeString(), strAppTitleVersion)
                End If
                swLog.WriteLine("  :{0}", logMessage)
                ' Update the underlying file.
                swLog.Flush()
                swLog.Close()
            End Using
        Catch ex As Exception 'Any exeception here is probably due to DontOpenThis.log being open...So we can't write to it.
            MsgBox(ex.ToString & vbCrLf & vbCrLf & " This error is usually caused because the DontOpenThis.log file is opened with a editor.")
        End Try
    End Sub

    Public Sub MarshalObject(ByRef obj As Object)
        If Not obj Is Nothing Then
            If System.Runtime.InteropServices.Marshal.IsComObject(obj) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            End If
            obj = Nothing
        End If
    End Sub
End Module
