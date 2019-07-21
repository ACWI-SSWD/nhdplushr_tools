Imports System.IO
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Data

Module modSQLLOCALDBCommon

    Public Function LoadDbf(ByVal strDbfname As String,
                            ByVal strTempWorkAreaPath As String,
                            ByVal strFldNames As String,
                            ByVal strSSETablename As String,
                            ByVal strWhereClause As String,
                            ByVal boolAppend As Boolean,
                            ByVal strWorkingDB As String,
                            ByVal intSQLConnectionTimeout As Integer,
                            ByVal intSQLCommandTimeout As Integer) As String

        Dim strReturn As String
        Dim strExMsg As String
        Try
            If File.Exists(strDbfname) Then
                File.Copy(strDbfname, strTempWorkAreaPath + "\ttemp.dbf")
            Else
                LoadDbf = "The file (" & strDbfname + ") is missing."
                Exit Try
            End If
            If strWhereClause <> "" And Not strWhereClause.ToUpper.Trim.Contains("WHERE") Then
                strWhereClause = " WHERE " & strWhereClause
            End If
            strReturn = ImportDbfFile(strTempWorkAreaPath & "\ttemp.dbf", strWorkingDB, strSSETablename, intSQLConnectionTimeout, intSQLCommandTimeout, strFldNames, strWhereClause, boolAppend)
            If strReturn <> "" Then
                LoadDbf = strReturn
                Exit Try
            End If
            LoadDbf = ""

        Catch ex As Exception
            strExMsg = "Loaddbf Exception: " + ex.Message.ToString + vbCrLf +
                ex.StackTrace.ToString + vbCrLf +
                ex.Source.ToString
            LoadDbf = strExMsg

        Finally
            If File.Exists(strTempWorkAreaPath + "\ttemp.dbf") Then
                File.Delete(strTempWorkAreaPath + "\ttemp.dbf")
            End If
            strReturn = Nothing
            strExMsg = Nothing

        End Try

    End Function

    Public Function ImportTextFile(ByVal strFullPathFileName As String, ByVal strTableName As String, ByVal strSQLFields As String, ByVal strSQLGroup As String, ByVal strSQLWhere As String, ByVal strWorkingDB As String, ByVal intSQLConnectionTimeout As Integer, intSQLCommandTimeOut As Integer) As String
        Try
            ImportTextFile = ""
            Dim strRetSQL As String
            Dim strFileNameNoExt = Path.GetFileNameWithoutExtension(strFullPathFileName)
            Dim strFileName = Path.GetFileName(strFullPathFileName)
            Dim strPathFileName = My.Computer.FileSystem.GetParentPath(strFullPathFileName)
            Dim connectionFile As OleDbConnection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strPathFileName & "; Extended Properties='text;HDR=Yes;FMT=Delimited'")
            Dim strSQL = "SELECT " & strSQLFields & " FROM " & strFileName & strSQLWhere & strSQLGroup
            Dim command As OleDbCommand = New OleDbCommand(strSQL, connectionFile)
            connectionFile.Open()

            Dim ImportedDataTable As DataTable = New DataTable()
            Dim reader As OleDbDataReader = command.ExecuteReader()
            ImportedDataTable.Load(reader)

            Using destinationConnection As SqlConnection = New SqlConnection("Data Source=(LocalDB)\v11.0;Integrated Security=True;Connect Timeout=" & intSQLConnectionTimeout & ";Initial Catalog=" & strWorkingDB)
                destinationConnection.Open()
                strRetSQL = GetCreateFromDataTableSQL(strTableName, ImportedDataTable)
                Dim cmd As SqlCommand = New SqlCommand(strRetSQL, destinationConnection)
                cmd.ExecuteNonQuery()

                ' Set up the bulk copy object.  
                ' The column positions in the source data reader  
                ' match the column positions in the destination table,  
                ' so there is no need to map columns. 
                Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(destinationConnection)
                    bulkCopy.DestinationTableName = "dbo." & strTableName
                    Try
                        ' Write from the source to the destination.
                        bulkCopy.WriteToServer(ImportedDataTable)
                    Catch ex As Exception
                        ImportTextFile = "SQLRunTime Exception in ImportTextFile BULKCOPY." & vbCrLf & _
                        ex.ToString & vbCrLf & _
                            ex.StackTrace.ToString & vbCrLf & _
                            ex.Source.ToString
                    End Try
                End Using
                destinationConnection.Close()
            End Using

            reader.Close()
            connectionFile.Close()

        Catch ex As Exception
            ImportTextFile = "SQLRunTime Exception in ImportTextFile." & vbCrLf & _
               ex.ToString & vbCrLf & _
               ex.StackTrace.ToString & vbCrLf & _
               ex.Source.ToString

        End Try

    End Function

    Public Function ImportDbfFile(ByVal strFullPathFileName As String, ByVal strWorkingDB As String, ByVal strTableName As String, ByVal intSQLConnectionTimeout As Integer, intSQLCommandTimeOut As Integer, ByVal strFields As String, ByVal strWhereclause As String, ByVal boolAppend As Boolean) As String
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
            connectionFile.Close()

        Catch ex As Exception
            ImportDbfFile = "SQLRunTime Exception in ImportDbfFile." & vbCrLf & _
               ex.ToString & vbCrLf & _
               ex.StackTrace.ToString & vbCrLf & _
               ex.Source.ToString
        End Try
    End Function

    Public Function GetCreateFromDataTableSQL(tableName As String, table As DataTable) As String
        Dim sql As String = "CREATE TABLE [" + tableName + "] (" & vbLf
        ' columns
        For Each column As DataColumn In table.Columns
            sql += "[" + column.ColumnName + "] " + SQLGetType(column) + "," & vbLf
        Next
        sql = sql.TrimEnd(New Char() {","c, ControlChars.Lf}) + vbLf
        ' primary keys
        If table.PrimaryKey.Length > 0 Then
            sql += "CONSTRAINT [PK_" + tableName + "] PRIMARY KEY CLUSTERED ("
            For Each column As DataColumn In table.PrimaryKey
                sql += "[" + column.ColumnName + "],"
            Next
            sql = sql.TrimEnd(New Char() {","c}) + "))" & vbLf
        End If

        'if not ends with ")"
        If (table.PrimaryKey.Length = 0) AndAlso (Not sql.EndsWith(")")) Then
            sql += ")"
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
                    Return "DECIMAL(" + numericPrecision + "," + numericScale + ")"
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
                Return "NVARCHAR(" + (If((columnSize = -1 OrElse columnSize > 8000), "MAX", columnSize.ToString())) + ")"

            Case "System.Byte"
                Return "TINYINT"

            Case "System.Guid"
                Return "UNIQUEIDENTIFIER"
            Case Else

                Throw New Exception(type.ToString() + " not implemented.")
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

End Module
