'V4003 3/19/19 Clear all commented out code and unused variables?
'
'V4002 3/19/19 Reconfigure projects and solution
'
'TODO comment out special cases for UNTIL?  UpTrib?  What does the toolbar have as 
'possible stop condition comparisons?
'V10018 JRH 20150831
'Fixed a problem in dndiv navigations.  They were stopping prematurely in cases where the two divergent outflows 
'did not return to the same path.
'V10018 - 4/30/2015
'Version to send to Tim - Converted to work with HR NHDPlus Production

'V10015 - 6/15/2014
'RMD - Removed "size=3, maxsize=4000" and "size=3, maxsize=8000" from SQL statement in CreateLocalDatabase that was crashing

'V10014 - 5/31/2014
'  2.  Removed - If stopping on attribute condition, make sure the comparison is valid at the start of the navigation

'V10013 - 5/24/2014
'  1.  Fixed a problem related to DNDIV with a stop distance navigations.  In GetStartPL - needed to marshal object and dispose of tblTemp
'  2.  If stopping on attribute condition, make sure the comparison is valid at the start of the navigation

'V10012 - 3/2/2014
'  Changed CreateLocalDatabase in modSQLLOCALDBCommon to first check for the existence of the database.  (NOT using the filename with _data.mdf only - that part is done
' before calling CreateLocalDatabase because it can be accomplished without connecting to (localdb)v11.0)

'V10011 - 06/06/2013
' Converted to localdb.  DOES NOT WRITE TO DBF AT THE END.  Results are in sql server tbl.  Calling program will need to access from there.
'This will need to be fixed at some point.  I believe Bob needs the navigator asap, so I will provide vb code to accomplish this using arcobjects(?)
'This version is called from the v2navigator toolbar - which has arcobjects available (and RMD will be calling it from the basin delineator)
'change listing for versions 7, 8, 9, 10 (not sure if there was a 10 - I may have incremented incorrectly) must be in svn.  I can't access svn any longer.

'V1006 - 9/30/2010
'  Added connection timeout
'  Added command timeout

'V1005 - Sept 29, 2010
'  Change in clsV02Navigator1.vb - ClearAllPools instead of ClearPool(connection to navdb)
'  Removed Connection timeout = 0 from connect statement

'V1004 - Sept 27, 2010
'  Change in clsLoadSqlServerDB - only call loaddbf for attrname if attrname > ""

'V1003 - Sept 23, 2010
'  Recompile after commenting out a messagebox.

'V1002 - Sept 21, 2010
'  Added the stop based on attribute value feature
'  implemented stop at a distance (ONLY clsV02Navigator1.vb)

'V1001 - April 4, 2009 
'      Used in BuildRefresh Step 10  
'         (clsLoadSQLServerDB, clsMakeWorkingTable, and clsV02Navigator1)
'      Therefore, using this module to contain version info
'      clsV02Navigator2 navigates using datatables not sql updates.  This needs work.
'         There are definite performance issues with large datasets.

'NAVIGATIONS ARE PERFORMED USING SQL SERVER TABLE UPDATES

'OUTPUT IS A SQL SERVER TABLE.

Imports System.Data
Imports System.Data.SqlClient

Public Class clsHRNavigator
    Private strNavType As String = ""
    Private numStartNHDPlusID As Double = -999
    Private dblStartMeasure As Double = -1
    Private dblMaxDistance As Double = 0
    Private strSessionID As String = ""
    Private strSQLDataSource As String
    Private strDatabaseLocation As String
    Private strDatabaseName As String
    Private strWorkingTableName As String
    Private strResultsTableName As String
    Private intProcessStatus As Integer
    Private strProcessMessage As String
    Private intSQLCommandTimeout As Integer
    Private intSQLConnectionTimeout As Integer
    Private strAttrName As String
    Private strAttrComp As String
    Private strAttrValue As String
    Private strAttrSelect As String
    Private strConnString As String
    Private cnnNHDPlusXtend As SqlConnection
    Private daWPlusFlowlineVAA As SqlDataAdapter
    Private dsNHDPlusXtend As DataSet
    Private dtWPlusFlowlineVAA As DataTable
    Private sqlcmdUpdate As SqlCommand
    Private sqlcmdSelectedCount As SqlCommand
    Private sqlcmdDrop As SqlCommand

    'Declare the row(s) array object and row object for WPlusFlowlineVAA and WPlusFlowlineVAA_Selected
    Private row_WPlusFlowlineVAA As DataRow
    Private gnumStartingPL As Double

    Public Sub New()
        MyBase.New()
        strNavType = ""
        numStartNHDPlusID = -999
        dblStartMeasure = -1
        dblMaxDistance = 0
        strSessionID = ""
        strSQLDataSource = ""
        strDatabaseLocation = ""
        strDatabaseName = ""
        strWorkingTableName = ""
        strResultsTableName = ""
        intProcessStatus = 0
        strProcessMessage = ""
        intSQLCommandTimeout = 50000
        intSQLConnectionTimeout = 50000
        strAttrName = ""
        strAttrComp = ""
        strAttrValue = ""
        strAttrSelect = ""

        strConnString = Nothing
        cnnNHDPlusXtend = Nothing
        daWPlusFlowlineVAA = Nothing
        dsNHDPlusXtend = Nothing
        dtWPlusFlowlineVAA = Nothing
        sqlcmdUpdate = Nothing
        sqlcmdSelectedCount = Nothing
        sqlcmdDrop = Nothing
        row_WPlusFlowlineVAA = Nothing
        gnumStartingPL = Nothing

    End Sub

    Protected Overrides Sub Finalize()
        On Error Resume Next
        MyBase.Finalize()
        strNavType = Nothing
        numStartNHDPlusID = Nothing
        dblStartMeasure = Nothing
        dblMaxDistance = Nothing
        strSessionID = Nothing
        strSQLDataSource = Nothing
        strDatabaseLocation = Nothing
        strDatabaseName = Nothing
        strWorkingTableName = Nothing
        strResultsTableName = Nothing
        intProcessStatus = Nothing
        strProcessMessage = Nothing
        intSQLCommandTimeout = Nothing
        intSQLConnectionTimeout = Nothing
        strAttrName = Nothing
        strAttrComp = Nothing
        strAttrValue = Nothing
        strAttrSelect = Nothing
        strConnString = Nothing
        cnnNHDPlusXtend.Dispose()

        MarshalObject(cnnNHDPlusXtend)
        daWPlusFlowlineVAA.Dispose()
        MarshalObject(daWPlusFlowlineVAA)
        dsNHDPlusXtend.Dispose()
        MarshalObject(dsNHDPlusXtend)
        dtWPlusFlowlineVAA.Dispose()
        MarshalObject(dtWPlusFlowlineVAA)

        sqlcmdUpdate.Dispose()
        sqlcmdSelectedCount.Dispose()
        sqlcmdDrop.Dispose()

        MarshalObject(sqlcmdUpdate)
        MarshalObject(sqlcmdSelectedCount)
        MarshalObject(sqlcmdDrop)
        MarshalObject(row_WPlusFlowlineVAA)

        gnumStartingPL = Nothing

    End Sub

    Public Property NavType() As String
        Get
            NavType = strNavType
        End Get
        Set(ByVal Value As String)
            strNavType = Value
        End Set
    End Property

    Public Property StartNHDPlusID() As Double
        Get
            StartNHDPlusID = numStartNHDPlusID
        End Get
        Set(ByVal Value As Double)
            numStartNHDPlusID = Value
        End Set
    End Property

    Public Property StartMeasure() As Double
        Get
            StartMeasure = dblStartMeasure
        End Get
        Set(ByVal Value As Double)
            dblStartMeasure = Value
        End Set
    End Property

    Public Property MaxDistance() As Double
        Get
            MaxDistance = dblMaxDistance
        End Get
        Set(ByVal Value As Double)
            dblMaxDistance = Value
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

    Public Property WorkingTableName() As String
        Get
            WorkingTableName = strWorkingTableName
        End Get
        Set(ByVal Value As String)
            strWorkingTableName = Value
        End Set
    End Property

    Public Property ResultsTableName() As String
        Get
            ResultsTableName = strResultsTableName
        End Get
        Set(ByVal Value As String)
            strResultsTableName = Value
        End Set
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

    Public Property AttrName() As String
        Get
            AttrName = strAttrName
        End Get
        Set(ByVal Value As String)
            strAttrName = Value
        End Set
    End Property

    Public Property AttrComp() As String
        Get
            AttrComp = strAttrComp
        End Get
        Set(ByVal Value As String)
            strAttrComp = Value
        End Set
    End Property

    Public Property AttrValue() As String
        Get
            AttrValue = strAttrValue
        End Get
        Set(ByVal Value As String)
            strAttrValue = Value
        End Set
    End Property

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

    Public Function VAANavigate() As Integer
        Dim strReturn As String
        strReturn = do_navigation()
        If strReturn = "" Then
            intProcessStatus = 0
            strProcessMessage = "Navigation Succeeded.  See Sql Server Table: " + strResultsTableName
            VAANavigate = 0
        Else
            intProcessStatus = 900
            strProcessMessage = strReturn
            VAANavigate = 1
        End If

    End Function

    Private Function do_validateproperties() As String
        Dim strReturn As String
        Dim boolMessage As Boolean
        Dim numProblems As Integer

        numProblems = 0
        boolMessage = True
        strReturn = ""

        do_validateproperties = strReturn

    End Function

    Private Function do_navigation() As String
        Dim numReturn As Integer = 0
        Dim strReturn As String = ""
        Dim strSQL As String = ""
        Dim strConnectionString As String
        Dim strExMsg As String

        Try
            do_navigation = ""
            'Get the localdb datasource with an instance based on a the current user combined with "NHDPLUSTOOLSDBS"
            'GetSQLDatasource will create it if it doesn't exist
            'strSQLDataSource = GetSQLDatasource()  -- JRH 3/9/2019 Input as a property.  No need to find it.
            Dim strCurrentUser As String = Environment.UserName.ToString.ToUpper
            'Remove any spaces that might be in the userid so that SQLLocalDB.exe doesn't crash
            strCurrentUser = strCurrentUser.Replace(" ", "")
            If Not strSQLDataSource.ToUpper.Contains("(LOCALDB)\") Then    '& strCurrentUser & "NHDPLUSTOOLSDBS"
                do_navigation = "Invalid SQLDataSource: " & strSQLDataSource
                Exit Try
            End If
            cnnNHDPlusXtend = New SqlConnection()
            daWPlusFlowlineVAA = New SqlDataAdapter()
            dsNHDPlusXtend = New DataSet()
            dtWPlusFlowlineVAA = dsNHDPlusXtend.Tables.Add("tblWPlusFlowlineVAA")

            'Initialize the SelectCommand
            daWPlusFlowlineVAA.SelectCommand = New SqlCommand
            strConnectionString = "Integrated Security=SSPI;" + "Initial Catalog=" + strDatabaseName + ";" + "Data Source=" + strSQLDataSource + "; Connection Timeout = " + intSQLConnectionTimeout.ToString + ";"

            cnnNHDPlusXtend.ConnectionString = strConnectionString
            cnnNHDPlusXtend.Open()

            'Set the connection for SQLdataAdapter (daWPlusFlowlineVAA)
            daWPlusFlowlineVAA.SelectCommand.Connection = cnnNHDPlusXtend

            'Update the working values to make sure there are no selections
            'from a previous navigation
            strSQL = "UPDATE " + strWorkingTableName + " SET from1 = NULL, to1 = NULL, Selected = 0"
            strReturn = ExecuteSQL(strSQL, daWPlusFlowlineVAA.SelectCommand.Connection, intSQLCommandTimeout)
            If strReturn <> "" Then
                do_navigation = strReturn
                Exit Try
            End If

            'Establish the update command
            sqlcmdUpdate = New SqlCommand("update " + strWorkingTableName + " SET selected = @selected, from1 = @from1, to1 = @to1 WHERE nhdplusid = @nhdplusid", cnnNHDPlusXtend)
            sqlcmdUpdate.Parameters.Add("@selected", SqlDbType.Int, 6, "selected")
            sqlcmdUpdate.Parameters.Add("@from1", SqlDbType.Real, 12, "from1")
            sqlcmdUpdate.Parameters.Add("@to1", SqlDbType.Real, 12, "to1")
            sqlcmdUpdate.Parameters.Add("@nhdplusid", SqlDbType.Decimal, 38, "nhdplusid")
            daWPlusFlowlineVAA.UpdateCommand = sqlcmdUpdate

            If strNavType = "UPTRIB" Then
                strSQL = "CREATE PROCEDURE GetSelected" + strWorkingTableName + "  @outval INT OUTPUT  AS BEGIN " +
                              "  SELECT @outval = COUNT(*) FROM " + strWorkingTableName + " WHERE selected >= 1 END "
                strReturn = ExecuteSQL(strSQL, daWPlusFlowlineVAA.SelectCommand.Connection, intSQLCommandTimeout)
                If strReturn <> "" Then
                    do_navigation = strReturn
                    Exit Try
                End If
                sqlcmdSelectedCount = New SqlCommand("", cnnNHDPlusXtend)
                sqlcmdSelectedCount.CommandText = "GetSelected" + strWorkingTableName
                sqlcmdSelectedCount.CommandType = CommandType.StoredProcedure
                sqlcmdSelectedCount.Parameters.AddWithValue("outval", 0)
                sqlcmdSelectedCount.Parameters("outval").Direction = System.Data.ParameterDirection.Output
            End If

            strSQL = "SELECT nhdplusid,REACHCODE,HYDROSEQ,LEVELPATHI,PATHLENGTH,TERMINALPA," &
                            "UPLEVELPAT,UPHYDROSEQ,DNLEVELPAT,DNMINORHYD,DNDRAINCOU," &
                            "DIVERGENCE,FROMMEAS,TOMEAS,LENGTHKM,DNHYDROSEQ FROM " + strWorkingTableName +
                            " WHERE nhdplusid = " + numStartNHDPlusID.ToString
            daWPlusFlowlineVAA.SelectCommand.CommandText = strSQL

            'Fill the datatable(dtWPlusFlowlineVAA)
            daWPlusFlowlineVAA.Fill(dtWPlusFlowlineVAA)

            'Find the record(row) containing the start nhdplusid
            If dtWPlusFlowlineVAA.Rows.Count = 1 Then
                For Each row As DataRow In dtWPlusFlowlineVAA.Select
                    row_WPlusFlowlineVAA = row
                    Exit For
                Next
            End If
            If row_WPlusFlowlineVAA Is Nothing Then
                do_navigation = "NHDPlusID: " & numStartNHDPlusID.ToString & " is not valid."
                Exit Try
            End If

            'If start measure is -1 , find the real starting measure.
            If (dblStartMeasure = -1) Then
                'This recordset should only return ONE record
                If ((strNavType = "UPMAIN") Or (strNavType = "UPTRIB")) Then
                    If row_WPlusFlowlineVAA.IsNull("FROMMEAS") Then
                        dblStartMeasure = -1
                    Else
                        dblStartMeasure = row_WPlusFlowlineVAA("FROMMEAS")
                    End If
                End If
                If ((strNavType = "DNMAIN") Or (strNavType = "DNDIV")) Then
                    If row_WPlusFlowlineVAA.IsNull("TOMEAS") Then
                        dblStartMeasure = -1
                    Else
                        dblStartMeasure = row_WPlusFlowlineVAA("TOMEAS")
                    End If
                End If
            End If

            numReturn = 0

            'Define the call to run clearresultstable
            sqlcmdDrop = New SqlCommand("", cnnNHDPlusXtend)
            sqlcmdDrop.CommandText = "clearresultstable"
            sqlcmdDrop.CommandType = CommandType.StoredProcedure
            sqlcmdDrop.Parameters.AddWithValue("tblname", strResultsTableName)
            sqlcmdDrop.Parameters.AddWithValue("isthere", 0)
            sqlcmdDrop.Parameters("isthere").Direction = System.Data.ParameterDirection.Output

            'Perform navigation based on type
            If (strNavType = "UPMAIN") Then
                strReturn = Navigate_UpMain(numStartNHDPlusID, dblStartMeasure)
            End If
            If (strNavType = "DNMAIN") Then
                strReturn = Navigate_DnMain(numStartNHDPlusID, dblStartMeasure)
            End If
            If (strNavType = "UPTRIB") Then
                strReturn = Navigate_UpTrib(numStartNHDPlusID, dblStartMeasure)
            End If
            If (strNavType = "DNDIV") Then
                strReturn = Navigate_DnDiv(numStartNHDPlusID, dblStartMeasure)
            End If

            If (strReturn <> "") Then
                intProcessStatus = 1
                strProcessMessage = strReturn
                do_navigation = strReturn
            Else
                strResultsTableName = "t" + strSessionID + "_NavResults"

                'The procedure is now created in clsLoadSqlServerDB
                strSQL = "CREATE PROCEDURE clearresultstable @tblname varchar(50), @isthere INT OUTPUT  AS BEGIN " +
                            "  SELECT @isthere = COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE' AND TABLE_NAME='@tblname' END "
                'strReturn = ExecuteSQL(strSQL, cnnNHDPlusXtend)
                If strReturn <> "" Then
                    do_navigation = strReturn
                    Exit Try
                End If

                sqlcmdDrop.Parameters("tblname").Value = strResultsTableName
                sqlcmdDrop.ExecuteNonQuery()
                If sqlcmdDrop.Parameters("isthere").Value = 1 Then
                    strSQL = "DROP TABLE " + strResultsTableName
                    strReturn = ExecuteSQL(strSQL, cnnNHDPlusXtend, intSQLCommandTimeout)
                    If strReturn <> "" Then
                        do_navigation = strReturn
                        Exit Try
                    End If
                End If

                If strAttrName > "" Then
                    'Attrvalue stop condition
                    strAttrSelect = " and Attrname " + strAttrComp + " " + strAttrValue
                Else
                    strAttrSelect = ""
                End If

                If strAttrName = "DIVERGENCE" Then
                    strAttrSelect = ""
                End If

                'Copy the results into tblNavResults in working mdb
                If (strNavType = "DNDIV" Or strNavType = "DNMAIN") Then
                    strSQL = "SELECT nhdplusid, reachcode, from1 as frommeas, to1 as tomeas, hydroseq INTO " + strResultsTableName + " FROM " + strWorkingTableName + "  WHERE selected >= 1 AND from1 <> to1 " + strAttrSelect + " ORDER BY hydroseq DESC "
                    strReturn = ExecuteSQL(strSQL, daWPlusFlowlineVAA.SelectCommand.Connection, intSQLCommandTimeout)
                    If strReturn <> "" Then
                        do_navigation = strReturn
                        Exit Try
                    End If
                Else
                    strSQL = "SELECT nhdplusid, reachcode, from1 as frommeas, to1 as tomeas, hydroseq INTO " + strResultsTableName + " FROM " + strWorkingTableName + "  WHERE selected >= 1 AND from1 <> to1 " + strAttrSelect + "  ORDER BY hydroseq "
                    strReturn = ExecuteSQL(strSQL, daWPlusFlowlineVAA.SelectCommand.Connection, intSQLCommandTimeout)
                    If strReturn <> "" Then
                        do_navigation = strReturn
                        Exit Try
                    End If
                End If
            End If

            If strNavType = "UPTRIB" Then
                strSQL = "DROP PROCEDURE GetSelected" + strWorkingTableName
                strReturn = ExecuteSQL(strSQL, daWPlusFlowlineVAA.SelectCommand.Connection, intSQLCommandTimeout)
                If strReturn <> "" Then
                    do_navigation = strReturn
                    Exit Try
                End If
            End If

        Catch ex As Exception
            strExMsg = ex.Message.ToString + vbCrLf +
                         ex.StackTrace.ToString + vbCrLf +
                         ex.Source.ToString
            do_navigation = "Do_Navigation Exception: " + strExMsg

        Finally
            dtWPlusFlowlineVAA.Dispose()
            dsNHDPlusXtend.Dispose()
            daWPlusFlowlineVAA.Dispose()
            cnnNHDPlusXtend.Close()
            cnnNHDPlusXtend.Dispose()
            SqlConnection.ClearAllPools()

            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()

        End Try

    End Function

    Private Function GetIncDis(ByVal numFrom As Double, ByVal numTo As Double, ByVal numLength As Double, ByVal numMeasure As Double, ByVal strHalf As String) As Double
        'Calculate Included Distance
        'NHDFlowline C:   A(To)------X------------B(From)
        'Distance from A to X is the 'TOP'
        'Distance from X to B is the 'BOTTOM'

        GetIncDis = -1

        Dim numDistance As Double

        'Changed the calculation for numDistance 7/20/2007
        numDistance = (numLength * ((numMeasure - numFrom) / (numTo - numFrom)))

        If (strHalf = "TOP") Then
            numDistance = numLength - numDistance
        End If
        If (strHalf = "BOTTOM") Then
            numDistance = numDistance
        End If
        GetIncDis = numDistance

    End Function

    Private Function GetMeasure(ByVal numFrom As Double, ByVal numTo As Double, ByVal numLength As Double, ByVal numDistance As Double, ByVal strHalf As String) As Double
        'Calculate Measure
        'NHDFlowline C:   A(To)------X------------B(From)
        'Distance from A to X is the 'TOP'
        'Distance from X to B is the 'BOTTOM'

        GetMeasure = -1

        Dim numMeasure As Double

        If (strHalf = "TOP") Then
            numMeasure = numTo - (((numTo - numFrom) / numLength) * numDistance)
        End If
        If (strHalf = "BOTTOM") Then
            numMeasure = numFrom + (((numTo - numFrom) / numLength) * numDistance)
        End If

        GetMeasure = numMeasure

    End Function

    Private Function GetStartPL(ByVal numFrom As Double, ByVal numTo As Double,
       ByVal numLength As Double, ByVal numMeasure As Double, ByVal numPL As Double,
       ByVal numDivergence As Object, ByVal numUphs As Double, ByVal strDirection As String) As Double

        Dim numIncDis As Double
        Dim numStartPL As Double
        'Calculate Starting pathlength

        'Normally, this is the pathlength + 'bottom included distance'

        'When going up - For flowlines that are div 2,
        '   this is the pathlength of the Upstream flowline - 'top included distance'

        'Where Top and Bottom distances are as follows:

        'NHDFlowline C:   A(To)------X------------B(From)

        'Distance from A to X is the 'TOP'
        'Distance from X to B is the 'BOTTOM'

        GetStartPL = -1
        If (strDirection = "DOWN" Or numDivergence <> 2) Then
            numIncDis = GetIncDis(numFrom, numTo, numLength, numMeasure, "BOTTOM")
            numStartPL = numPL + numIncDis
        Else
            Dim dtTemp As DataTable = dsNHDPlusXtend.Tables.Add("tblTemp")
            daWPlusFlowlineVAA.SelectCommand.CommandText = "SELECT pathlength FROM " & strWorkingTableName & " WHERE hydroseq = " + Str(numUphs)
            daWPlusFlowlineVAA.Fill(dtTemp)
            For Each row As DataRow In dtTemp.Select
                numIncDis = GetIncDis(numFrom, numTo, numLength, numMeasure, "TOP")
                numStartPL = row("pathlength") - numIncDis
            Next
            dtTemp.Clear()
            dtTemp.Dispose()
            'Added these two lines 5/24/2014 JRH
            dsNHDPlusXtend.Tables.Remove("tblTemp")
            MarshalObject(dtTemp)

        End If
        GetStartPL = numStartPL

    End Function

    Private Function Navigate_UpMain(ByVal numStartNHDPlusID As Double, ByVal numStartMeasure As Double) As String
        Dim numTermid As Double
        Dim numHydroseqno As Double
        Dim numLevelpathid As Double

        Dim numMaxHS As Double
        Dim numMaxHSnhdplusid As Double
        Dim numMaxHSUshs As Double
        Dim numMaxHSUslp As Double

        Dim strSQL As String
        Dim strReturn As String

        Dim boolContinue As Boolean
        Dim boolFirst As Boolean
        Dim sqlcmdCommand As System.Data.SqlClient.SqlCommand
        Dim intMinHS As Double

        Dim numTopIncDis As Double
        Dim numBottomIncDis As Double
        Dim numStartingPL As Double
        Dim dtUpMain As DataTable = dsNHDPlusXtend.Tables.Add("tblUpMain")
        Dim strExMsg As String
        Dim intRecCount As Integer
        Dim intRowPos As Integer

        Try
            Navigate_UpMain = ""

            boolContinue = True
            boolFirst = True

            '***GET INFORMATION ABOUT THE STARTING nhdplusid
            Dim rowWPlusFlowlineVAA As DataRow = Nothing
            For Each row As DataRow In dtWPlusFlowlineVAA.Select
                rowWPlusFlowlineVAA = row
            Next

            numTermid = rowWPlusFlowlineVAA("terminalpa")
            numHydroseqno = rowWPlusFlowlineVAA("hydroseq")
            numLevelpathid = rowWPlusFlowlineVAA("levelpathi")

            If (dblMaxDistance > 0) Then
                If (IsDBNull(rowWPlusFlowlineVAA("Frommeas"))) Then
                    'Should only happen when starting nhdplusid has no measures
                    numTopIncDis = rowWPlusFlowlineVAA("lengthkm")
                    numBottomIncDis = 0
                    numStartingPL = rowWPlusFlowlineVAA("pathlength")
                Else
                    numTopIncDis = GetIncDis(rowWPlusFlowlineVAA("Frommeas"), rowWPlusFlowlineVAA("tomeas"),
                       rowWPlusFlowlineVAA("lengthkm"), numStartMeasure, "TOP")
                    numBottomIncDis = GetIncDis(rowWPlusFlowlineVAA("Frommeas"), rowWPlusFlowlineVAA("tomeas"),
                        rowWPlusFlowlineVAA("lengthkm"), numStartMeasure, "BOTTOM")
                    numStartingPL = GetStartPL(rowWPlusFlowlineVAA("Frommeas"), rowWPlusFlowlineVAA("tomeas"),
                        rowWPlusFlowlineVAA("lengthkm"), numStartMeasure, rowWPlusFlowlineVAA("pathlength"),
                        rowWPlusFlowlineVAA("divergence"), rowWPlusFlowlineVAA("uphydroseq"), "UP")
                End If
            End If
            '***END GET INFORMATION ABOUT THE STARTING nhdplusid

            '***MAIN LOOP
            'boolContinue will be false if the starting nhdplusid does not exist
            If (boolContinue) Then

                numMaxHS = 0
                boolContinue = True

                While (boolContinue = True)

                    'Set the starting hydroseqno for upstream levelpaths that are
                    'NOT the initial query
                    If Not boolFirst Then

                        If (numHydroseqno = numMaxHSUshs) Then
                            numHydroseqno = 0
                        Else
                            numHydroseqno = numMaxHSUshs
                            numLevelpathid = numMaxHSUslp
                        End If

                    End If

                    If (numHydroseqno <> 0) Then

                        'There is an upstream levelpath

                        'Upstream levelpaths exist on the initial query and
                        'when a particular levelpath ends on a divergence = 2

                        boolFirst = False

                        If dblMaxDistance > 0 Then
                            'max distance stop condition
                            strSQL = "UPDATE " + strWorkingTableName +
                               " SET selected = 1, from1 = frommeas, to1 = tomeas " +
                               " WHERE levelpathi = " + Str(numLevelpathid) + " and " +
                                  "hydroseq >= " + Str(numHydroseqno) + " and " +
                                  "pathlength <= " + Str(numStartingPL + dblMaxDistance)
                        Else
                            'No stop condition
                            strSQL = "UPDATE " + strWorkingTableName +
                                   " SET selected = 1, from1 = frommeas, to1 = tomeas " +
                                   " WHERE levelpathi = " + Str(numLevelpathid) + " and " +
                                      "hydroseq >= " + Str(numHydroseqno)
                        End If

                        strReturn = ExecuteSQL(strSQL, daWPlusFlowlineVAA.SelectCommand.Connection, intSQLCommandTimeout)
                        If strReturn <> "" Then
                            Navigate_UpMain = strReturn
                            Exit Try
                        End If

                        ''SPECIAL CASE - "="
                        If strAttrComp = "=" Or strAttrComp = "<>" Then   'UNTIL
                            strSQL = "SELECT min(hydroseq) as hs FROM " & strWorkingTableName & " WHERE selected = 1 AND attrname " & strAttrComp & " " & strAttrValue
                            sqlcmdCommand = New System.Data.SqlClient.SqlCommand("", daWPlusFlowlineVAA.SelectCommand.Connection)
                            sqlcmdCommand.CommandText = strSQL
                            intMinHS = IIf(IsDBNull(sqlcmdCommand.ExecuteScalar()), 0, sqlcmdCommand.ExecuteScalar())
                            If Not intMinHS = Nothing Then
                                If intMinHS > 0 Then
                                    strSQL = "UPDATE " + strWorkingTableName +
                                       " SET selected = 0 " +
                                       " WHERE selected = 1 AND " +
                                       " hydroseq > " & Str(intMinHS)
                                    strReturn = ExecuteSQL(strSQL, daWPlusFlowlineVAA.SelectCommand.Connection, intSQLCommandTimeout)
                                    If strReturn <> "" Then
                                        Navigate_UpMain = strReturn
                                        Exit Try
                                    End If
                                End If
                            End If
                        End If

                        'Retrieve the just selected records to modify from/to as needed
                        strSQL = "SELECT * FROM " + strWorkingTableName +
                               " WHERE selected = 1 and " +
                                  "levelpathi = " + Str(numLevelpathid) + " and " +
                                  "hydroseq >= " + Str(numHydroseqno) +
                                  " ORDER BY hydroseq"
                        daWPlusFlowlineVAA.SelectCommand.CommandText = strSQL
                        daWPlusFlowlineVAA.Fill(dtUpMain)
                        intRecCount = 0
                        For Each row As DataRow In dtUpMain.Select
                            intRecCount = intRecCount + 1
                        Next row

                        intRowPos = 0
                        For Each row As DataRow In dtUpMain.Select
                            intRowPos = intRowPos + 1
                            If intRowPos = 1 Then
                                If (row("nhdplusid") = numStartNHDPlusID) Then
                                    row("from1") = numStartMeasure
                                    strSQL = "UPDATE " + strWorkingTableName + " SET from1 = " + numStartMeasure.ToString + " WHERE nhdplusid = " + numStartNHDPlusID.ToString
                                    strReturn = ExecuteSQL(strSQL, daWPlusFlowlineVAA.SelectCommand.Connection, intSQLCommandTimeout)
                                    If strReturn <> "" Then
                                        Navigate_UpMain = strReturn
                                        Exit Try
                                    End If
                                End If
                            End If

                            If intRowPos = intRecCount Then
                                numMaxHS = row("hydroseq")
                                numMaxHSnhdplusid = row("nhdplusid")
                                numMaxHSUshs = row("uphydroseq")
                                numMaxHSUslp = row("uplevelpat")
                                'update stop measure if necessary
                                If (dblMaxDistance > 0) And
                                   ((row("pathlength") + row("lengthkm")) >
                                         (numStartingPL + dblMaxDistance)) Then
                                    If (IsDBNull(row("tomeas"))) Then
                                        'End NHDFlowline has null measures.  DO nothing.
                                    Else

                                        row("to1") = GetMeasure(row("From1"),
                                          row("to1"), row("lengthkm"),
                                         numStartingPL + dblMaxDistance - row("pathlength"), "BOTTOM")
                                        daWPlusFlowlineVAA.Update(dtUpMain)
                                    End If
                                    boolContinue = False
                                End If
                            End If
                        Next
                        dtUpMain.Clear()
                    Else
                        'There are no more upstream queries, so stop
                        boolContinue = False
                    End If
                End While
            End If
            '***END MAIN LOOP

        Catch ex As Exception
            strExMsg = "Description: " + Err.Description + vbCrLf +
                       "Source: " + Err.Source + vbCrLf +
                       "Exception Stack Trace: " & Err.GetException.StackTrace.ToString
            Navigate_UpMain = "Navigate_UpMain Exception: " + strExMsg

        Finally
            dtUpMain.Dispose()
            dtUpMain = Nothing

        End Try

    End Function

    Private Function Navigate_DnMain(ByVal numStartNHDPlusID As Double, ByVal numStartMeasure As Double) As String
        Dim strExMsg As String
        Dim numTermid As Double
        Dim numHydroseqno As Double
        Dim numLevelpathid As Double
        Dim numMinHSDshs As Double
        Dim numMinHSDslp As Double
        Dim numLastLP As Double
        Dim strSQL As String
        Dim boolContinue As Boolean = True
        Dim boolFirst As Boolean = True
        Dim numTopIncDis As Double
        Dim numBottomIncDis As Double
        Dim numStartingPL As Double
        Dim intRecCount As Integer = 0
        Dim intRowPos As Integer = 0
        Dim strReturn As String
        Dim dtDnMain As DataTable = dsNHDPlusXtend.Tables.Add("tblDnMain")

        Try
            Navigate_DnMain = ""

            '***GET INFORMATION ABOUT THE STARTINGNHDPlusID
            'rstStartResults = QueryStartPoint(numStartNHDPlusID, "NHDPlusID")
            Dim rowWPlusFlowlineVAA As DataRow = Nothing
            For Each row As DataRow In dtWPlusFlowlineVAA.Select
                rowWPlusFlowlineVAA = row
            Next

            numTermid = rowWPlusFlowlineVAA("TERMINALPA")
            numHydroseqno = rowWPlusFlowlineVAA("HYDROSEQ")
            numLevelpathid = rowWPlusFlowlineVAA("LEVELPATHI")

            If (dblMaxDistance > 0) Then
                If rowWPlusFlowlineVAA.IsNull("FROMMEAS") = True Then
                    'Should only happen when starting nhdplusid has no measures
                    numTopIncDis = rowWPlusFlowlineVAA("LENGTHKM")
                    numBottomIncDis = 0
                    numStartingPL = rowWPlusFlowlineVAA("PATHLENGTH")
                Else
                    numTopIncDis = GetIncDis(rowWPlusFlowlineVAA("FROMMEAS"), rowWPlusFlowlineVAA("TOMEAS"),
                        rowWPlusFlowlineVAA("LENGTHKM"), numStartMeasure, "TOP")
                    numBottomIncDis = GetIncDis(rowWPlusFlowlineVAA("FROMMEAS"), rowWPlusFlowlineVAA("TOMEAS"),
                        rowWPlusFlowlineVAA("LENGTHKM"), numStartMeasure, "BOTTOM")
                    numStartingPL = GetStartPL(rowWPlusFlowlineVAA("FROMMEAS"), rowWPlusFlowlineVAA("TOMEAS"),
                        rowWPlusFlowlineVAA("LENGTHKM"), numStartMeasure, rowWPlusFlowlineVAA("PATHLENGTH"),
                        rowWPlusFlowlineVAA("DIVERGENCE"), rowWPlusFlowlineVAA("UPHYDROSEQ"), "DOWN")
                End If
            End If
            '***END GET INFORMATION ABOUT THE STARTING NHDPlusID

            gnumStartingPL = numStartingPL

            '***MAIN LOOP
            'boolContinue will be false if the starting nhdplusid  does not exist
            If (boolContinue) Then
                numLastLP = 0
                boolContinue = True

                While (boolContinue = True)

                    'Set the starting hydroseqno for upstream levelpaths that are
                    'NOT the initial query

                    If Not boolFirst Then
                        numLevelpathid = numMinHSDslp
                        'numhslp = numMinHSDshs
                        numHydroseqno = numMinHSDshs
                    End If

                    If (numLastLP <> numLevelpathid And numLevelpathid <> 0) Then

                        'There is a downstream levelpath
                        numLastLP = numLevelpathid

                        'Downstream levelpaths exist on the initial query and
                        'when a particular levelpath ends and there is another below it
                        boolFirst = False
                        If dblMaxDistance > 0 Then
                            'max distance stop condition
                            strSQL = "UPDATE " + strWorkingTableName +
                                  " SET selected = 1, from1 = frommeas, to1 = tomeas " +
                                  " WHERE levelpathi = " + Str(numLevelpathid) + " and " +
                                  " terminalpa = " + Str(numTermid) + " And " +
                                  " hydroseq <= " + Str(numHydroseqno) + " and " +
                                  " hydroseq <> 0 and " +
                                  Str(numStartingPL - dblMaxDistance) + " <= pathlength+lengthkm "
                        Else
                            'No stop condition
                            strSQL = "UPDATE " + strWorkingTableName +
                                  " SET selected = 1, from1 = frommeas, to1 = tomeas " +
                                  " WHERE levelpathi = " + Str(numLevelpathid) + " and " +
                                  " terminalpa = " + Str(numTermid) + " And " +
                                  " hydroseq <= " + Str(numHydroseqno) + " and " +
                                  " hydroseq <> 0 "
                        End If

                        strReturn = ExecuteSQL(strSQL, daWPlusFlowlineVAA.SelectCommand.Connection, intSQLCommandTimeout)
                        If strReturn <> "" Then
                            Navigate_DnMain = strReturn
                            Exit Try
                        End If

                        'Retrieve the just selected records to modify from/to as needed
                        strSQL = "SELECT * FROM " + strWorkingTableName +
                                     " WHERE selected = 1 and " +
                                      "levelpathi = " + Str(numLevelpathid) + " and " +
                                     "hydroseq <= " + Str(numHydroseqno) + " and " +
                                     "hydroseq <> 0 " +
                                     " ORDER BY hydroseq DESC"
                        daWPlusFlowlineVAA.SelectCommand.CommandText = strSQL
                        daWPlusFlowlineVAA.Fill(dtDnMain)
                        intRecCount = 0
                        For Each row As DataRow In dtDnMain.Select
                            intRecCount = intRecCount + 1
                        Next row

                        intRowPos = 0
                        For Each row As DataRow In dtDnMain.Select
                            intRowPos = intRowPos + 1
                            If intRowPos = 1 Then
                                If (row("nhdplusid") = numStartNHDPlusID) Then
                                    row("to1") = numStartMeasure
                                    strSQL = "UPDATE " + strWorkingTableName + " SET to1 = " + numStartMeasure.ToString + " WHERE nhdplusid = " + numStartNHDPlusID.ToString
                                    strReturn = ExecuteSQL(strSQL, daWPlusFlowlineVAA.SelectCommand.Connection, intSQLCommandTimeout)
                                End If
                            End If

                            If intRowPos = intRecCount Then
                                If Not IsDBNull(row("dnlevelpat")) Then
                                    numMinHSDslp = row("dnlevelpat")
                                Else
                                    numMinHSDshs = -1
                                End If
                                If Not IsDBNull(row("dnhydroseq")) Then
                                    numMinHSDshs = row("dnhydroseq")
                                Else
                                    numMinHSDshs = -1
                                End If

                                If (dblMaxDistance > 0) And
                                   ((numStartingPL - dblMaxDistance) > row("pathlength")) Then
                                    If (IsDBNull(row("tomeas"))) Then
                                        'End NHDFlowline has null measures.  DO nothing.
                                    Else
                                        row("From1") = GetMeasure(row("Frommeas"),
                                        row("tomeas"), row("lengthkm"),
                                           (row("pathlength") + row("lengthkm")) - (numStartingPL - dblMaxDistance), "TOP")
                                        daWPlusFlowlineVAA.Update(dtDnMain)
                                    End If
                                    boolContinue = False
                                End If
                            End If
                        Next
                        dtDnMain.Clear()
                    Else
                        'There are no more downstream queries, so stop
                        boolContinue = False
                    End If
                End While
            End If
            '***END MAIN LOOP

        Catch ex As Exception
            strExMsg = "Description: " + Err.Description + vbCrLf +
                       "Source: " + Err.Source + vbCrLf +
                       "Exception Stack Trace: " & Err.GetException.StackTrace.ToString
            Navigate_DnMain = "Navigate_DnMain Exception: " + strExMsg

        Finally
            dtDnMain.Dispose()
            dtDnMain = Nothing

        End Try

    End Function

    Private Function Navigate_UpTrib(ByVal numStartNHDPlusID As Double, ByVal numStartMeasure As Double) As String
        Dim numTermid As Double
        Dim numHydroseqno As Double
        Dim numLevelpathid As Double
        Dim strSQL As String
        Dim strExMsg As String
        Dim strReturn As String
        Dim boolContinue As Boolean
        Dim numTopIncDis As Double
        Dim numBottomIncDis As Double
        Dim numStartingPL As Double
        Dim numSelectedPre As Long
        Dim numSelectedPost As Long
        Dim numIteration As Long

        Dim daFlowX As New SqlDataAdapter()
        Dim dtFlowX As DataTable = dsNHDPlusXtend.Tables.Add("tblflowx")
        Dim strTempFileName = "t" + strSessionID + "_Temp"
        Dim strTempFileName1 = "t" + strSessionID + "_Temp1"

        Try
            Navigate_UpTrib = ""

            '-----------------------------------------------
            'Normal UPMain navigation from the start point
            '-----------------------------------------------
            strReturn = Navigate_UpMain(numStartNHDPlusID, numStartMeasure)
            If (strReturn <> "") Then
                Navigate_UpTrib = strReturn
                Exit Try
            End If

            boolContinue = True

            '***GET INFORMATION ABOUT THE STARTING nhdplusid
            Dim rowWPlusFlowlineVAA As DataRow = Nothing
            For Each row As DataRow In dtWPlusFlowlineVAA.Select
                rowWPlusFlowlineVAA = row
            Next

            numTermid = rowWPlusFlowlineVAA("terminalpa")
            numHydroseqno = rowWPlusFlowlineVAA("hydroseq")
            numLevelpathid = rowWPlusFlowlineVAA("levelpathi")

            If (dblMaxDistance > 0) Then
                If (IsDBNull(rowWPlusFlowlineVAA("Frommeas"))) Then
                    'Should only happen when starting nhdplusid has no measures
                    numTopIncDis = rowWPlusFlowlineVAA("lengthkm")
                    numBottomIncDis = 0
                    numStartingPL = rowWPlusFlowlineVAA("pathlength")
                Else
                    numTopIncDis = GetIncDis(rowWPlusFlowlineVAA("Frommeas"), rowWPlusFlowlineVAA("tomeas"),
                       rowWPlusFlowlineVAA("lengthkm"), numStartMeasure, "TOP")
                    numBottomIncDis = GetIncDis(rowWPlusFlowlineVAA("Frommeas"), rowWPlusFlowlineVAA("tomeas"),
                        rowWPlusFlowlineVAA("lengthkm"), numStartMeasure, "BOTTOM")
                    numStartingPL = GetStartPL(rowWPlusFlowlineVAA("Frommeas"), rowWPlusFlowlineVAA("tomeas"),
                        rowWPlusFlowlineVAA("lengthkm"), numStartMeasure, rowWPlusFlowlineVAA("pathlength"),
                        rowWPlusFlowlineVAA("divergence"), rowWPlusFlowlineVAA("uphydroseq"), "UP")
                End If
            End If
            '***END GET INFORMATION ABOUT THE STARTING nhdplusid

            '-----------------------------------------------
            '***Main UPTRIB Loop - Do until nothing more
            '                   is selected
            '-----------------------------------------------
            numSelectedPre = -1
            numSelectedPost = -1
            numIteration = 1

            While boolContinue

                'GET PRE-UPDATE SELECTED COUNT IF NECESSARY
                '5/31/2006 - changed selected = 1 to selected >= 1
                If (numSelectedPost = -1) Then
                    sqlcmdSelectedCount.ExecuteNonQuery()
                    numSelectedPre = sqlcmdSelectedCount.Parameters("outval").Value
                Else
                    numSelectedPre = numSelectedPost
                End If

                '5/31/2006 - Drop temp table
                sqlcmdDrop.Parameters("tblname").Value = strTempFileName
                sqlcmdDrop.ExecuteNonQuery()
                If sqlcmdDrop.Parameters("isthere").Value = 1 Then
                    strSQL = "DROP TABLE " + strTempFileName
                    strReturn = ExecuteSQL(strSQL, daWPlusFlowlineVAA.SelectCommand.Connection, intSQLCommandTimeout)
                    If strReturn <> "" Then
                        Navigate_UpTrib = strReturn
                        Exit Try
                    End If
                End If

                'Create temp table
                strSQL = "SELECT b.fromlvlpat, MIN(b.fromhydseq) AS minhs INTO " + strTempFileName +
                            " FROM " + strWorkingTableName + " AS a " +
                            " INNER JOIN PlusFlow AS b ON a.nhdplusid = b.ToNHDPID AND a.levelpathi <> b.fromlvlpat " +
                            " WHERE a.selected = " + Str(numIteration) +
                            " GROUP BY b.fromlvlpat"

                strReturn = ExecuteSQL(strSQL, daWPlusFlowlineVAA.SelectCommand.Connection, intSQLCommandTimeout)
                If strReturn <> "" Then
                    Navigate_UpTrib = strReturn
                    Exit Try
                End If

                If strAttrComp = "=" Or strAttrComp = "<>" Then  'until
                    strSQL = "select a.minhs, b.tohydseq, c.hydroseq, c.nhdplusid, c.attrname " +
                             " into " + strTempFileName1 +
                             " from " + strTempFileName + " as a inner join plusflow as b on a.minhs = b.fromhydseq " +
                             " inner join " + strWorkingTableName + " as c on b.tohydseq = c.hydroseq " +
                             " where (c.attrname " & strAttrComp & " " & strAttrValue & ")"
                    strReturn = ExecuteSQL(strSQL, daWPlusFlowlineVAA.SelectCommand.Connection, intSQLCommandTimeout)
                    If strReturn <> "" Then
                        Navigate_UpTrib = strReturn
                        Exit Try
                    End If
                    strSQL = "delete from " + strTempFileName1 + " where exists " +
                             " (select 1 from PlusFlowlineVAA, " + strWorkingTableName + " as w  ,(SELECT tonode FROM PlusflowlineVAA WHERE plusflowlinevaa.hydroseq = " + strTempFileName1 + ".minhs) as d WHERE plusflowlinevaa.fromnode = d.tonode and plusflowlinevaa.nhdplusid = w.nhdplusid AND selected >= 1 AND w.divergence = 1);"

                    strReturn = ExecuteSQL(strSQL, daWPlusFlowlineVAA.SelectCommand.Connection, intSQLCommandTimeout)
                    If strReturn <> "" Then
                        Navigate_UpTrib = strReturn
                        Exit Try
                    End If

                    strSQL = "delete from " + strTempFileName + " where exists " +
                             " (select 1 from " + strTempFileName1 +
                             " where " + strTempFileName + ".minhs = minhs)"

                    strReturn = ExecuteSQL(strSQL, daWPlusFlowlineVAA.SelectCommand.Connection, intSQLCommandTimeout)
                    If strReturn <> "" Then
                        Navigate_UpTrib = strReturn
                        Exit Try
                    End If

                    strSQL = "drop table " + strTempFileName1
                    strReturn = ExecuteSQL(strSQL, daWPlusFlowlineVAA.SelectCommand.Connection, intSQLCommandTimeout)
                    If strReturn <> "" Then
                        Navigate_UpTrib = strReturn
                        Exit Try
                    End If
                End If

                'DO UPDATE
                numIteration = numIteration + 1
                If dblMaxDistance > 0 Then
                    'max distance stop condition
                    strSQL = "UPDATE a " +
                             " SET a.selected = " + Str(numIteration) +
                             ", a.from1 = a.frommeas, a.to1 = a.tomeas " +
                             "  FROM " + strWorkingTableName + " a INNER JOIN " + strTempFileName + " i ON (i.fromlvlpat = a.levelpathi) " +
                             " WHERE a.hydroseq >= i.minhs " + " and " +
                             "pathlength <= " + Str(numStartingPL + dblMaxDistance)
                Else
                    'No stop condition
                    strSQL = "UPDATE a " +
                             " SET a.selected = " + Str(numIteration) +
                             ", a.from1 = a.frommeas, a.to1 = a.tomeas " +
                             "  FROM " + strWorkingTableName + " a INNER JOIN " + strTempFileName + " i ON (i.fromlvlpat = a.levelpathi) " +
                             " WHERE a.hydroseq >= i.minhs "
                End If

                strReturn = ExecuteSQL(strSQL, daWPlusFlowlineVAA.SelectCommand.Connection, intSQLCommandTimeout)
                If strReturn <> "" Then
                    Navigate_UpTrib = strReturn
                    Exit Try
                End If

                'SPECIAL CASE - "UNTIL"
                If strAttrComp = "=" Or strAttrComp = "<>" Then   'UNTIL

                    strSQL = "SELECT min(hydroseq) as hs, levelpathi INTO " + strTempFileName1 + " FROM " & strWorkingTableName &
                             " WHERE selected >= 1 and nhdplusid <> " + numStartNHDPlusID.ToString & " AND attrname " & strAttrComp & " " & strAttrValue + " GROUP BY levelpathi "
                    strReturn = ExecuteSQL(strSQL, daWPlusFlowlineVAA.SelectCommand.Connection, intSQLCommandTimeout)
                    If strReturn <> "" Then
                        Navigate_UpTrib = strReturn
                        Exit Try
                    End If

                    strSQL = "UPDATE " + strWorkingTableName +
                               " SET selected = 0 FROM " + strTempFileName1 + " INNER JOIN " + strWorkingTableName +
                               " ON " + strTempFileName1 + ".levelpathi = " + strWorkingTableName + ".levelpathi " +
                               " WHERE selected  > 0 AND  hydroseq > " + strTempFileName1 + ".hs"
                    strReturn = ExecuteSQL(strSQL, daWPlusFlowlineVAA.SelectCommand.Connection, intSQLCommandTimeout)
                    If strReturn <> "" Then
                        Navigate_UpTrib = strReturn
                        Exit Try
                    End If

                    strSQL = "DROP TABLE " + strTempFileName1
                    strReturn = ExecuteSQL(strSQL, daWPlusFlowlineVAA.SelectCommand.Connection, intSQLCommandTimeout)
                    If strReturn <> "" Then
                        Navigate_UpTrib = strReturn
                        Exit Try
                    End If
                End If

                'GET POST-UPDATE SELECTED COUNT
                '5/31/2006 - changed selected = 1 to selected >= 1
                sqlcmdSelectedCount.ExecuteNonQuery()
                numSelectedPost = sqlcmdSelectedCount.Parameters("outval").Value
                If (numSelectedPre = numSelectedPost) Then
                    boolContinue = False
                End If

            End While
            '***END MAIN LOOP

            'JRH 3/31/2008 - Change for calculations when stopping at a distance along the first nhdplusid
            '   in the navigation.  Add to1 = tomeas to where clause.

            '5/31/2006 - changed selected = 1 to selected >= 1
            'Update measures at the end of paths where necessary if there is a stop distance.
            If (dblMaxDistance > 0) Then
                strSQL = "UPDATE " & strWorkingTableName +
                     " SET  to1 =  (from1 + ((to1-from1) / lengthkm) * (" +
                        Str(numStartingPL + dblMaxDistance) + " - pathlength) )" +
                     " WHERE selected >= 1 and to1 = tomeas and " +
                     " ( pathlength+lengthkm > " + Str(numStartingPL + dblMaxDistance) + " )"
                strReturn = ExecuteSQL(strSQL, daWPlusFlowlineVAA.SelectCommand.Connection, intSQLCommandTimeout)
                If strReturn <> "" Then
                    Navigate_UpTrib = strReturn
                    Exit Try
                End If
            End If

        Catch ex As Exception
            strExMsg = "Description: " + Err.Description + vbCrLf +
                       "Source: " + Err.Source + vbCrLf +
                       "Exception Stack Trace: " & Err.GetException.StackTrace.ToString
            Navigate_UpTrib = "Navigate_UpTrib Exception: " + strExMsg

        Finally
            sqlcmdDrop.Parameters("tblname").Value = strTempFileName
            sqlcmdDrop.ExecuteNonQuery()
            If sqlcmdDrop.Parameters("isthere").Value = 1 Then
                strSQL = "DROP TABLE " + strTempFileName
                strReturn = ExecuteSQL(strSQL, daWPlusFlowlineVAA.SelectCommand.Connection, intSQLCommandTimeout)
            End If

        End Try

    End Function
    Private Function Navigate_DnDiv(ByVal numStartNHDPlusID As Double, ByVal numStartMeasure As Double) As String
        Dim strExMsg As String
        Dim numTermid As Double
        Dim numHydroseqno As Double
        Dim numLevelpathid As Double
        Dim listDivs(10000, 10) As Double
        Dim numDivCount As Long
        Dim strSQL As String
        Dim intRecCount As Integer
        Dim strReturn As String
        Dim numLastLP As Double
        Dim numMinHSdshs As Double
        Dim numMinHSdslp As Double
        Dim numLastLPMain As Double
        Dim numLastHSMain As Double
        Dim numStartingPL As Double
        Dim dblMaxPathDistance As Double
        Dim boolContinue As Boolean
        Dim boolFirst As Boolean
        Dim boolFinished As Boolean
        Dim intRowPos As Integer
        Dim dtDnDiv As DataTable = Nothing
        Dim intMaxHS As Double
        Dim i As Integer

        Try

            Navigate_DnDiv = ""

            boolContinue = True

            '-----------------------------------------------
            'Normal DownMain navigation from the start point
            '-----------------------------------------------
            strReturn = Navigate_DnMain(numStartNHDPlusID, numStartMeasure)
            If (strReturn <> "") Then
                Navigate_DnDiv = strReturn
                Exit Try
            End If

            strSQL = "SELECT * FROM " + strWorkingTableName + " WHERE selected = 1 ORDER BY HYDROSEQ "
            daWPlusFlowlineVAA.SelectCommand.CommandText = strSQL
            dtDnDiv = dsNHDPlusXtend.Tables.Add("tblDnDiv")
            daWPlusFlowlineVAA.Fill(dtDnDiv)
            For Each row As DataRow In dtDnDiv.Select("", "HYDROSEQ")
                numLastLPMain = row("levelpathi")
                numLastHSMain = row("hydroseq")
                Exit For
            Next
            dtDnDiv.Clear()
            dtDnDiv.Dispose()
            dsNHDPlusXtend.Tables.Remove("tblDnDiv")
            MarshalObject(dtDnDiv)

            boolFinished = False
            numDivCount = -1

            intMaxHS = -1
            While Not boolFinished

                boolFinished = True

                strReturn = GetDivs(listDivs, numDivCount)
                If (strReturn <> "") Then
                    Navigate_DnDiv = strReturn
                    Exit Try
                End If

                For i = 0 To numDivCount

                    'If this drain that has multiple outflows has not been processed...
                    If listDivs(i, 5) = 0 Then

                        listDivs(i, 5) = 1
                        boolFinished = False

                        numStartNHDPlusID = listDivs(i, 0)
                        numStartMeasure = listDivs(i, 6)
                        numTermid = listDivs(i, 2)
                        numHydroseqno = listDivs(i, 3)
                        numLevelpathid = listDivs(i, 4)

                        boolFirst = True
                        boolContinue = True
                        If numHydroseqno < intMaxHS Then
                            boolContinue = False
                        End If
                        '-----------------------------------------------
                        'Special DownMain navigation
                        '   Stop navigating if a levelpath has already
                        '   been navigated.
                        '-----------------------------------------------
                        'Adjust StartingPL and maxdistance for this start
                        If (dblMaxDistance > 0) Then
                            numStartingPL = GetStartPL(listDivs(i, 1), listDivs(i, 6),
                              listDivs(i, 7), numStartMeasure, listDivs(i, 8),
                             listDivs(i, 9), listDivs(i, 10), "UP")
                            dblMaxPathDistance = dblMaxDistance - (gnumStartingPL - numStartingPL)
                            numStartingPL = listDivs(i, 8) + listDivs(i, 7)
                            'MsgBox(String.Format("dblMaxPathDistance: {0},{1}numStartingPL: {2}{1}", dblMaxPathDistance, Environment.NewLine, numStartingPL.ToString))
                        End If

                        '***DOWN MAIN LOOP
                        'boolContinue will be false if the starting nhdplusid does not exist
                        If (boolContinue) Then

                            numLastLP = 0
                            boolContinue = True

                            While (boolContinue = True)

                                'Set the starting hydroseqno for downstream levelpaths that are
                                'NOT the initial query

                                If Not boolFirst Then
                                    numLevelpathid = numMinHSdslp
                                    'numHydroseqno = numMinHs
                                    numHydroseqno = numMinHSdshs
                                End If
                                If numHydroseqno < intMaxHS Then
                                    boolContinue = False
                                    Exit While
                                End If

                                If (numLastLP <> numLevelpathid And numLevelpathid <> 0) Then

                                    'There is a downstream levelpath
                                    numLastLP = numLevelpathid

                                    'Downstream levelpaths exist on the initial query and
                                    'when a particular levelpath ends and there is another below it
                                    boolFirst = False

                                    If dblMaxDistance > 0 Then
                                        If numLastLPMain <> numLevelpathid Then
                                            strSQL = "UPDATE " + strWorkingTableName +
                                                  " SET selected = 1, from1 = frommeas, to1 = tomeas " +
                                                  " WHERE selected = 0 AND levelpathi = " + Str(numLevelpathid) + " and " +
                                                  " terminalpa = " + Str(numTermid) + " And " +
                                                  " hydroseq <= " + Str(numHydroseqno) + " and " +
                                                  Str(numStartingPL - dblMaxPathDistance) + " <= pathlength+lengthkm "
                                        Else
                                            'Special case to make sure we don't go
                                            'past the main path end point
                                            strSQL = "UPDATE " + strWorkingTableName +
                                               " SET selected = 1, from1 = frommeas, to1 = tomeas " +
                                               " WHERE selected = 0 and " +
                                               "levelpathi = " + Str(numLevelpathid) + " and " +
                                                  "hydroseq <= " + Str(numHydroseqno) + " and " +
                                                  "hydroseq > " + Str(numLastHSMain) + " and (" +
                                                  Str(numStartingPL - dblMaxDistance) + " <= pathlength+lengthkm) "
                                        End If

                                    Else
                                        'No stop condition
                                        strSQL = "UPDATE " + strWorkingTableName +
                                              " SET selected = 1, from1 = frommeas, to1 = tomeas " +
                                              " WHERE selected = 0 AND levelpathi = " + Str(numLevelpathid) + " and " +
                                              " terminalpa = " + Str(numTermid) + " And " +
                                              " hydroseq <= " + Str(numHydroseqno) 'JRH 20150831 COMMENT THIS CLAUSE + + " and " + _
                                    End If
                                    strReturn = ExecuteSQL(strSQL, daWPlusFlowlineVAA.SelectCommand.Connection, intSQLCommandTimeout)
                                    If strReturn <> "" Then
                                        Navigate_DnDiv = strReturn
                                        Exit Try
                                    End If

                                    'Retrieve the just selected records to modify from/to as needed
                                    strSQL = "SELECT * FROM " + strWorkingTableName +
                                           " WHERE selected = 1 and " +
                                           "levelpathi = " + Str(numLevelpathid) + " and " +
                                           "hydroseq <= " + Str(numHydroseqno) + " and " +
                                           "hydroseq <> 0 " +
                                           " ORDER BY hydroseq DESC"
                                    daWPlusFlowlineVAA.SelectCommand.CommandText = strSQL
                                    dtDnDiv = dsNHDPlusXtend.Tables.Add("tblDnDiv")
                                    daWPlusFlowlineVAA.Fill(dtDnDiv)
                                    intRecCount = 0
                                    For Each row As DataRow In dtDnDiv.Select
                                        intRecCount = intRecCount + 1
                                    Next row
                                    intRowPos = 0
                                    For Each row As DataRow In dtDnDiv.Select
                                        intRowPos = intRowPos + 1
                                        If intRowPos = 1 Then
                                            If (row("nhdplusid") = numStartNHDPlusID) Then
                                                row("from1") = numStartMeasure
                                                strSQL = "UPDATE " + strWorkingTableName + " SET to1 = " + numStartMeasure.ToString + " WHERE nhdplusid = " + numStartNHDPlusID.ToString
                                                strReturn = ExecuteSQL(strSQL, daWPlusFlowlineVAA.SelectCommand.Connection, intSQLCommandTimeout)
                                                If strReturn <> "" Then
                                                    Navigate_DnDiv = strReturn
                                                    Exit Try
                                                End If
                                            End If
                                        End If
                                        If intRowPos = intRecCount Then
                                            If Not IsDBNull(row("dnlevelpat")) Then
                                                numMinHSdslp = row("dnlevelpat")
                                            Else
                                                numMinHSdshs = -1
                                            End If
                                            If Not IsDBNull(row("dnhydroseq")) Then
                                                numMinHSdshs = row("dnhydroseq")
                                            Else
                                                numMinHSdshs = -1
                                            End If
                                            'update stop measure if necessary

                                            If (dblMaxPathDistance > 0) And ((numStartingPL - dblMaxPathDistance) > row("pathlength")) And
                                                (row("levelpathi") <> numLastLPMain) Then
                                                If (IsDBNull(row("tomeas"))) Then
                                                    'End NHDFlowline has null measures.  DO nothing.
                                                Else
                                                    row("From1") = GetMeasure(row("Frommeas"),
                                                    row("tomeas"), row("lengthkm"),
                                                       (row("pathlength") + row("lengthkm")) - (numStartingPL - dblMaxPathDistance), "TOP")
                                                    daWPlusFlowlineVAA.Update(dtDnDiv)

                                                End If
                                                boolContinue = False
                                            End If
                                        End If

                                    Next
                                    dtDnDiv.Clear()
                                    dtDnDiv.Dispose()
                                    dsNHDPlusXtend.Tables.Remove("tblDnDiv")
                                    MarshalObject(dtDnDiv)
                                Else
                                    'There are no more downstream queries, so stop
                                    boolContinue = False
                                End If
                            End While
                        End If
                        '***END MAIN LOOP
                        '-----------------------------------------------
                        'End Special DownMain navigation
                        '-----------------------------------------------
                    End If

                Next i

            End While

        Catch ex As Exception
            'Concurrency violation - The update command affected 0 records... is thrown 
            'incorrectly.  
            strExMsg = "Description: " + Err.Description + vbCrLf +
                       "Source: " + Err.Source + vbCrLf +
                       "Exception Stack Trace: " & Err.GetException.StackTrace.ToString
            Navigate_DnDiv = "Navigate_DnDiv Exception: " + strExMsg

        Finally
            strExMsg = Nothing
            numTermid = Nothing
            numHydroseqno = Nothing
            numLevelpathid = Nothing
            listDivs(10000, 10) = Nothing
            numDivCount = Nothing
            strSQL = Nothing
            intRecCount = Nothing
            strReturn = Nothing
            numLastLP = Nothing
            numMinHSdshs = Nothing
            numMinHSdslp = Nothing
            numLastLPMain = Nothing
            numLastHSMain = Nothing
            numStartingPL = Nothing
            dblMaxPathDistance = Nothing
            boolContinue = Nothing
            boolFirst = Nothing
            boolFinished = Nothing
            intRowPos = Nothing
            If Not dtDnDiv Is Nothing Then
                dtDnDiv.Clear()
                dtDnDiv.Dispose()
                dtDnDiv = Nothing
            End If
            i = Nothing

        End Try

        'KEEPING THIS CODE BECAUSE IT MIGHT BE USEFUL
        'WHEN IMPLEMENTING DISTANCE CALCS AND STOPPING AT A DISTANCE

        '            '-----------------------------------------------
        '            'Special DownMain navigation
        '            '   Stop navigating if a levelpath has already
        '            '   been navigated.
        '            '-----------------------------------------------

        '            'INFO in listdivs
        '            'listDivs(numDivCount, 0) = rstStartResults!nhdplusid
        '            'listDivs(numDivCount, 1) = rstStartResults!frommeas
        '            'listDivs(numDivCount, 2) = rstStartResults!terminalpa
        '            'listDivs(numDivCount, 3) = rstStartResults!hydroseq
        '            'listDivs(numDivCount, 4) = rstStartResults!levelpathi
        '            'listDivs(numDivCount, 5) = 0
        '            'listDivs(numDivCount, 6) = rstStartResults!tomeas
        '            'listDivs(numDivCount, 7) = rstStartResults!lengthkm
        '            'listDivs(numDivCount, 8) = rstStartResults!pathlength
        '            'listDivs(numDivCount, 9) = rstStartResults!divergence
        '            'listDivs(numDivCount, 10) = rstStartResults!uphydroseq

        '            'Adjust StartingPL and maxdistance for this start
        '            If (dblMaxDistance > 0) Then

        '                numStartingPL = GetStartPL(listDivs(i, 2), listDivs(i, 6), _
        '                   listDivs(i, 7), listDivs(i, 2), listDivs(i, 8), _
        '                   listDivs(i, 9), listDivs(i, 10), "UP")
        '                dblMaxDistance = numStartingPL - (gnumStartingPL - dblMaxDistance)
        '                numStartingPL = listDivs(i, 8) + listDivs(i, 7)
        '            End If


        '            'gswLog.Write("New Max Distance: " + Str(dblMaxDistance) + vbCrLf)
        '            'gswLog.Write("New starting pl: " + Str(numStartingPL) + vbCrLf)

        '            '***DOWN MAIN LOOP
        '            'boolContinue will be false if the starting nhdplusid does not exist
        '            If (boolContinue) Then

        '                numLastLP = 0
        '                boolContinue = True

        '                While (boolContinue = True)

        '                    'Set the starting hydroseqno for downstream levelpaths that are
        '                    'NOT the initial query

        '                    If Not boolFirst Then
        '                        numLevelpathid = numMinHSdslp
        '                        'numHydroseqno = numMinHs
        '                        numHydroseqno = numMinHSdshs
        '                    End If

        '                    If (numLastLP <> numLevelpathid And numLevelpathid > 0) Then

        '                        'There is a downstream levelpath

        '                        numLastLP = numLevelpathid

        '                        'Downstream levelpaths exist on the initial query and
        '                        'when a particular levelpath ends and there is another below it

        '                        boolFirst = False

        '                        rstResults = New ADODB.Recordset
        '                        rstResults.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
        '                        rstResults.LockType = ADODB.LockTypeEnum.adLockOptimistic

        '                        strSQL = "SELECT * FROM tblWPlusFlowlineVAA " + _
        '                           " WHERE selected = 1 "

        '                        'Select (set selected = 1 ) for downstream levelpath
        '                        If (dblMaxDistance = 0) Then
        '                            strSQL = "UPDATE tblWPlusFlowlineVAA " + _
        '                               "SET selected = 1, from1 = frommeas, to1 = tomeas " + _
        '                               "WHERE selected = 0 and " + _
        '                                  "levelpathi = " + Str(numLevelpathid) + " and " + _
        '                                  "hydroseq  <= " + Str(numHydroseqno) + " and " + _
        '                                  "hydroseq <> 0 "
        '                        Else
        '                            If (numLastLPMain <> numLevelpathid) Then
        '                                strSQL = "UPDATE tblWPlusFlowlineVAA " + _
        '                                   "SET selected = 1, from1 = frommeas, to1 = tomeas " + _
        '                                   "WHERE selected = 0 and " + _
        '                                   "levelpathi = " + Str(numLevelpathid) + " and " + _
        '                                      "hydroseq <= " + Str(numHydroseqno) + " and " + _
        '                                      "hydroseq <> 0 and (" + _
        '                                      Str(numStartingPL - dblMaxDistance) + " <= pathlength+lengthkm) "
        '                            Else
        '                                'Special case to make sure we don't go
        '                                'past the main path end point
        '                                strSQL = "UPDATE tblWPlusFlowlineVAA " + _
        '                                   "SET selected = 1, from1 = frommeas, to1 = tomeas " + _
        '                                   "WHERE selected = 0 and " + _
        '                                   "levelpathi = " + Str(numLevelpathid) + " and " + _
        '                                      "hydroseq <= " + Str(numHydroseqno) + " and " + _
        '                                      "hydroseq > " + Str(numLastHSMain) + " and (" + _
        '                                      Str(numStartingPL - dblMaxDistance) + " <= pathlength+lengthkm) "
        '                            End If
        '                        End If
        '                        'MsgBox strSQL
        '                        'gswLog.Write(strSQL + vbCrLf)
        '                        gcnnWorking.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

        '                        strSQL = "SELECT * FROM tblWPlusFlowlineVAA " + _
        '                           " WHERE selected = 1 "

        '                        'Retrieve the just selected records to modify from/to as needed
        '                        strSQL = "SELECT * FROM tblWPlusFlowlineVAA " + _
        '                           " WHERE selected = 1 and " + _
        '                           "levelpathi = " + Str(numLevelpathid) + " and " + _
        '                           "hydroseq <= " + Str(numHydroseqno) + " and " + _
        '                           "hydroseq <> 0 " + _
        '                           " ORDER BY hydroseq DESC"
        '                        rstResults.Open(strSQL, gcnnWorking, , , ADODB.CommandTypeEnum.adCmdText)
        '                        If Not rstResults.EOF Then
        '                            'movelast to get minhs, and update stop measure
        '                            rstResults.MoveFirst()
        '                            If (rstResults.Fields("nhdplusid").Value = numStartNHDPlusID) Then
        '                                rstResults.Fields("from1").Value = numStartMeasure
        '                                rstResults.Update()
        '                            End If
        '                            rstResults.MoveLast()
        '                            numMinHs = rstResults.Fields("hydroseq").Value
        '                            numMinHSnhdplusid = rstResults.Fields("nhdplusid").Value
        '                            numMinHSdslp = rstResults.Fields("dnlevelpat").Value
        '                            If Not IsDBNull(rstResults.Fields("dnhydroseq").Value) Then
        '                                numMinHSdshs = rstResults.Fields("dnhydroseq").Value
        '                            Else
        '                                numMinHSdshs = -1
        '                            End If
        '                            'update stop measure if necessary
        '                            If (dblMaxDistance > 0) And _
        '                               ((numStartingPL - dblMaxDistance) > rstResults.Fields("pathlength").Value) Then
        '                                If (rstResults.Fields("hydroseq").Value <> numLastHSMain) Then
        '                                    If (IsDBNull(rstResults.Fields("tomeas").Value)) Then
        '                                        'End NHDFlowline has null measures.  DO nothing.
        '                                    Else
        '                                        rstResults.Fields("to1").Value = GetMeasure(rstResults.Fields("Frommeas").Value, _
        '                                           rstResults.Fields("tomeas").Value, rstResults.Fields("lengthkm").Value, _
        '                                           (rstResults.Fields("pathlength").Value + rstResults.Fields("lengthkm").Value) - (numStartingPL - dblMaxDistance), "TOP")
        '                                        ''gswLog.Write "Updated " + Str(rstResults!hydroseq) + vbCrLf
        '                                        rstResults.Update()
        '                                    End If
        '                                End If
        '                                'MsgBox "Not continuing"
        '                                boolContinue = False
        '                            End If
        '                        Else
        '                            ''gswLog.Write "Downstream Query for HS " + Str(numHydroseqno) + _
        '                            '       " returned 0 records. " + _
        '                            '       "There should be at least one record. " + vbCrLf
        '                            boolContinue = False
        '                        End If
        '                        rstResults.Close()
        '                    Else
        '                        'There are no more downstream queries, so stop
        '                        boolContinue = False
        '                    End If
        '                End While
        '            End If
        '            '***END MAIN LOOP
        '            '-----------------------------------------------
        '            'End Special DownMain navigation
        '            '-----------------------------------------------

    End Function

    Private Function GetDivs(ByRef listDivs As Array, ByRef numDivCount As Long) As String
        Dim i As Long
        Dim numTemp As Double
        Dim numnhdplusid As Double
        Dim boolFound As Integer
        Dim rowDivergence As DataRow = Nothing
        Dim strExMsg As String
        Dim dtMegs As DataTable = dsNHDPlusXtend.Tables.Add("tblMegs")
        Dim dtDivs As DataTable = dsNHDPlusXtend.Tables.Add("tblDivs")
        Dim dtHSnhdplusid As DataTable = dsNHDPlusXtend.Tables.Add("tblHSnhdplusid")

        Try

            GetDivs = ""
            daWPlusFlowlineVAA.SelectCommand.CommandText = "SELECT * FROM " + strWorkingTableName + " WHERE selected = 1 AND dndraincou > 1 ORDER BY HYDROSEQ DESC "
            daWPlusFlowlineVAA.Fill(dtDivs)
            For Each row As DataRow In dtDivs.Select
                If row("dndraincou") = 2 Then
                    '= to 2
                    'See if the minor downstream hs exists in the list 
                    'of divergences already
                    numTemp = row("dnminorhyd")

                    boolFound = False
                    For i = 0 To numDivCount
                        If (listDivs(i, 3) = numTemp) Then
                            boolFound = True
                            Exit For
                        End If
                    Next i
                    'If not - add it
                    If numTemp <> 0 Then

                        If Not boolFound Then
                            daWPlusFlowlineVAA.SelectCommand.CommandText = "SELECT * FROM " + strWorkingTableName + " WHERE hydroseq = " + numTemp.ToString
                            daWPlusFlowlineVAA.Fill(dtHSnhdplusid)
                            numnhdplusid = 0
                            For Each row1 As DataRow In dtHSnhdplusid.Select
                                'THERE SHOULD ONLY BE ONE RECORD 
                                rowDivergence = row1
                                numnhdplusid = row1("nhdplusid")
                                Exit For
                            Next
                            'Make sure a record was found.  
                            If numnhdplusid = 0 Then
                                'This should never happen
                                GetDivs = "Hydroseq from dnminorhyd field for " + row("HYDROSEQ").ToString + " did not exist"
                                Exit Try
                            End If
                            numDivCount = numDivCount + 1
                            listDivs(numDivCount, 0) = rowDivergence("nhdplusid")
                            listDivs(numDivCount, 1) = rowDivergence("Frommeas")
                            listDivs(numDivCount, 2) = rowDivergence("terminalpa")
                            listDivs(numDivCount, 3) = rowDivergence("hydroseq")
                            listDivs(numDivCount, 4) = rowDivergence("levelpathi")
                            listDivs(numDivCount, 5) = 0
                            listDivs(numDivCount, 6) = rowDivergence("tomeas")
                            listDivs(numDivCount, 7) = rowDivergence("lengthkm")
                            listDivs(numDivCount, 8) = rowDivergence("pathlength")
                            listDivs(numDivCount, 9) = rowDivergence("divergence")
                            listDivs(numDivCount, 10) = rowDivergence("uphydroseq")
                            dtHSnhdplusid.Clear()
                        End If
                    End If

                Else
                    'Greater than 2 - query megadivs flow table
                    daWPlusFlowlineVAA.SelectCommand.CommandText = "SELECT * FROM MegaDiv WHERE FromNHDPID = " + row("nhdplusid").ToString
                    daWPlusFlowlineVAA.Fill(dtMegs)

                    For Each rowMegs As DataRow In dtMegs.Select
                        numTemp = rowMegs("ToNHDPID")
                        boolFound = False
                        For i = 0 To numDivCount
                            If (listDivs(i, 0) = numTemp) Then
                                boolFound = True
                                Exit For
                            End If
                        Next i
                        'If not - add it
                        If Not boolFound Then
                            daWPlusFlowlineVAA.SelectCommand.CommandText = "SELECT * FROM " + strWorkingTableName + " WHERE nhdplusid = " + numTemp.ToString
                            daWPlusFlowlineVAA.Fill(dtHSnhdplusid)
                            numnhdplusid = 0
                            For Each row1 As DataRow In dtHSnhdplusid.Select
                                'THERE SHOULD ONLY BE ONE RECORD 
                                rowDivergence = row1
                                numnhdplusid = row1("nhdplusid")
                                Exit For
                            Next
                            'Make sure a record was found.  
                            If numnhdplusid = 0 Then
                                'This should never happen
                                GetDivs = "Outflow " + numTemp.ToString + " from " + row("nhdplusid").ToString + " not found."
                                Exit Try
                            End If
                            numDivCount = numDivCount + 1
                            listDivs(numDivCount, 0) = rowDivergence("nhdplusid")
                            listDivs(numDivCount, 1) = rowDivergence("Frommeas")
                            listDivs(numDivCount, 2) = rowDivergence("terminalpa")
                            listDivs(numDivCount, 3) = rowDivergence("hydroseq")
                            listDivs(numDivCount, 4) = rowDivergence("levelpathi")
                            listDivs(numDivCount, 5) = 0
                            listDivs(numDivCount, 6) = rowDivergence("tomeas")
                            listDivs(numDivCount, 7) = rowDivergence("lengthkm")
                            listDivs(numDivCount, 8) = rowDivergence("pathlength")
                            listDivs(numDivCount, 9) = rowDivergence("divergence")
                            listDivs(numDivCount, 10) = rowDivergence("uphydroseq")
                            dtHSnhdplusid.Clear()
                        End If

                    Next
                    dtMegs.Clear()

                End If

            Next row

        Catch ex As Exception
            'Concurrency violation - The update command affected 0 records... is thrown 
            'incorrectly.  
            strExMsg = "Description: " + Err.Description + vbCrLf +
                       "Source: " + Err.Source + vbCrLf +
                       "Exception Stack Trace: " & Err.GetException.StackTrace.ToString
            GetDivs = "GetDivs: " + strExMsg

        Finally
            dtDivs.Dispose()
            dtDivs = Nothing
            dtHSnhdplusid.Dispose()
            dtHSnhdplusid = Nothing
            dtMegs.Dispose()
            dtMegs = Nothing
            dsNHDPlusXtend.Tables.Remove("tblDivs")
            dsNHDPlusXtend.Tables.Remove("tblHSnhdplusid")
            dsNHDPlusXtend.Tables.Remove("tblMegs")

        End Try

    End Function

End Class
