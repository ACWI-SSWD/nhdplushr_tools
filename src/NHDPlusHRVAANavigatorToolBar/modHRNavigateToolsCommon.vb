'V2008 - 5/31/2014
'  Added validations in Navigation option form related to stopping based on attributes.

'V2007 - 5/24/2014
'  Fixed a display problem related to alternating vaa and flow navigations
'  put a loop in DisplayResults to make sure all the results rows have been written

'V2006 - 3/2/2014
'  Fixed problems related to reading/writing  the navigator working db path  in the ini
'  Fixed problem related to locating the ini properly and calling the navigator with the correct information

'V2004 is a draft.  msgbox to find ini file

'V1003 - 6/13/2013 
' Add NavDBPath to the ini and the Navigation Options dialog
' Default to Windows temp folder
' Use NavDBPath to store the navigation dbs locally

'V1002 - 6/6/2013 - Convert to ArcGIS 10.1
'   Uses navigator V10011 - (localdb - leaves results in sql server table)

Option Strict Off
Option Explicit On
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geoprocessor
Imports ESRI.ArcGIS.esriSystem
Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System.IO
Imports System.Data.OleDb

Module modHRNavigateToolsCommon
    Public gstrNavigationType As String
    Public gnumFcodeValue As Integer
    Private gdblStartNHDPlusID As Double
    Private gstrStartReachcode As String
    Private gdblStartMeasure As Double
    Private gdblFrom As Double
    Private gdblTo As Double
    Private gdblDistance As Double
    Private gstrVPUPath As String
    Private gstrVPUID As String
    Private gstrDllLocation As String
    Private gstrResultsDBF As String
    Private gstrDataSource As String
    Private gstrAttrName As String
    Private gstrAttrCompare As String
    Private gstrAttrValue As String
    Private gstrNavDBPath As String = ""
    Private numResultsCount As Integer = 0
    Private strSessionID As String
    Private strTempWorkAreaPath As String

    Public Function CoordinateNavigation(ByVal mPoint As IPoint, ByVal strNavType As String, ByVal strDllLocation As String,
         ByRef mDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument, ByRef mMap As IMap, ByRef mActiveView As IActiveView, ByRef mApplication As IApplication) As String

        Dim strReturn As String = ""

        Try
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            'RMD - 05/15/19 - Trying to let the user know the program is still running
            Dim appCursor As IMouseCursor = New MouseCursorClass
            appCursor.SetCursor(2) 'Hourglass

            gstrDllLocation = strDllLocation
            gstrNavigationType = strNavType

            'Determines that there is an active, selectable, NHDFlowine layer
            'Selects flowline nearest mouse click
            'displays a marker (black dot)
            'determines measure of mouse click
            'prompts for distance
            'Initializes gstrStartReachcode, gdblStartNHDPlusID, gdblStartMeasure, gdblDistance, gstrVPUPath
            strReturn = SelectNHDFlowline(mPoint, mDoc, mMap, mActiveView)
            If strReturn <> "" Then
                Exit Try
            End If
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            'RMD - 05/15/19 - Trying to let the user know the program is still running
            appCursor.SetCursor(2) 'Hourglass

            'Remove previous results
            strReturn = RemovePreviousResults(mDoc, mMap, mActiveView)
            If strReturn <> "" Then
                Exit Try
            End If

            'Do the navigation
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            'RMD - 05/15/19 - Trying to let the user know the program is still running
            appCursor.SetCursor(2) 'Hourglass

            mApplication.StatusBar.Message(0) = "Calling NHDPlus HR VAA Navigator"
            strReturn = CallNavigator(mApplication)
            If strReturn <> "" Then
                Exit Try
            End If

            'File Maintenance
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            'RMD - 05/15/19 - Trying to let the user know the program is still running
            appCursor.SetCursor(2) 'Hourglass
            strReturn = FileMaintenance()
            If strReturn <> "" Then
                Exit Try
            End If

            'Display the results
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            'RMD - 05/15/19 - Trying to let the user know the program is still running
            appCursor.SetCursor(2) 'Hourglass
            mApplication.StatusBar.Message(0) = "Displaying results"
            strReturn = DisplayResults(mDoc, mMap, mActiveView)
            If strReturn <> "" Then
                Exit Try
            End If

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            appCursor.SetCursor(0) 'Arrow pointer

        Catch ex As Exception
            CoordinateNavigation = "SelectNHDFlowline Exception: " & ex.Message.ToString &
                vbCrLf & ex.StackTrace.ToString & vbCrLf & ex.Source.ToString

        Finally
            CoordinateNavigation = strReturn
            strReturn = Nothing

        End Try

    End Function

    Public Function DropWorkingDB(ByVal mPoint As IPoint, ByVal strNavType As String, ByVal strDllLocation As String,
           ByRef mDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument, ByRef mMap As IMap, ByRef mActiveView As IActiveView, ByRef mApplication As IApplication) As String

        Dim strReturn As String = ""
        Dim strDBName As String = ""
        Dim strSQL As String = ""

        Dim pDataSet As ESRI.ArcGIS.Geodatabase.IDataset
        Dim pLayer As IGeoFeatureLayer

        Try
            strReturn = MsgBox("Press YES to delete the working database for the selected NHDFlowline layer, NO otherwise.", vbQuestion + vbYesNo, "NHDPlus HR VAA Navigator")

            If strReturn = vbYes Then
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                gstrDllLocation = strDllLocation
                gstrNavigationType = strNavType

                If gstrNavigationType = "DROPDB" Then

                    DropWorkingDB = ""

                    'Confirms that only one NHDFlowline layer is active and it is selectable
                    strReturn = ValidateSelectedLayer(mDoc, "NHDFlowline")
                    If strReturn <> "" Then
                        DropWorkingDB = strReturn
                        Exit Try
                    End If

                    pLayer = mDoc.SelectedLayer
                    pDataSet = pLayer

                    'SET GLOBAL VAR
                    gstrVPUPath = Left(pDataSet.Workspace.PathName, Len(pDataSet.Workspace.PathName))
                    gstrVPUID = gstrVPUPath.Substring(gstrVPUPath.LastIndexOf("\") + 1).Replace("HRNHDPlus", "")
                    gstrVPUID = gstrVPUID.Replace(".gdb", "")

                    strDBName = "V03NavDB_" & gstrVPUID

                    'Drop a database from  the localdb instance
                    SqlConnection.ClearAllPools()
                    Using myLocalDB As SqlConnection = New SqlConnection("Data Source=(LocalDB)\v11.0;Integrated Security=True;Connect Timeout= 0;")
                        SqlConnection.ClearAllPools()
                        myLocalDB.Open()
                        strSQL = " SELECT COUNT(*) FROM sys.databases WHERE name = '" & strDBName & "'"
                        Dim cmd As SqlCommand = New SqlCommand(strSQL, myLocalDB)
                        If cmd.ExecuteScalar > 0 Then
                            cmd.CommandText = "DROP DATABASE " & strDBName
                            cmd.ExecuteNonQuery()

                            'Added by JRH 3/22/2019
                            MsgBox("Working database " & strDBName & " dropped successfully")
                        Else
                            'Added else clause JRH 3/22/2019
                            MsgBox("Working database " & strDBName & " not found.  Nothing to drop.")
                        End If
                        myLocalDB.Close()
                    End Using
                End If
            End If
            strReturn = ""

        Catch ex As Exception
            DropWorkingDB = "DropWorkingDB Exception: " & ex.Message.ToString &
                vbCrLf & ex.StackTrace.ToString & vbCrLf & ex.Source.ToString
            If DropWorkingDB.ToUpper.Contains("CANNOT DROP DATABASE") Then
                'Capture the case where it did not work. Make sure first, Display a message (problematic sometimes when trying to delete after navigating, displaying, and removing results)
                SqlConnection.ClearAllPools()
                Using myLocalDB As SqlConnection = New SqlConnection("Data Source=(LocalDB)\v11.0;Integrated Security=True;Connect Timeout= 0;")
                    SqlConnection.ClearAllPools()
                    myLocalDB.Open()
                    strSQL = " SELECT COUNT(*) FROM sys.databases WHERE name = '" & strDBName & "'"
                    Dim cmd As SqlCommand = New SqlCommand(strSQL, myLocalDB)
                    If cmd.ExecuteScalar > 0 Then
                        MsgBox("Unable to drop " & strDBName & " because it is in use by another process.  Please remove all navigation results from your map, exit and restart ArcMap, then try again.  Make sure the correct NHDFlowline layer is loaded into your map.")
                    End If
                    myLocalDB.Close()
                End Using

                DropWorkingDB = ""
            End If

        Finally
            DropWorkingDB = strReturn
            strReturn = Nothing

        End Try

    End Function

    Public Function FileMaintenance() As String
        Dim infoMDF As FileInfo
        Dim infoLog As FileInfo
        Dim intSizeMDF As Integer
        Dim intSizeLog As Integer
        Dim intUpperLimit As Integer = 500000

        Dim strConnectionString As String
        Dim sqlconConnection As SqlConnection
        Dim strSQL As String
        Dim strReturn As String = ""

        Try
            FileMaintenance = ""

            infoMDF = New FileInfo(gstrNavDBPath & "\V03NavDB_" & gstrVPUID & ".mdf")
            infoLog = New FileInfo(gstrNavDBPath & "\V03NavDB_" & gstrVPUID & ".ldf")
            intSizeMDF = infoMDF.Length
            intSizeLog = infoLog.Length

            If Not gstrDataSource = "" And Not gstrDataSource Is Nothing Then
                If intSizeMDF >= intUpperLimit Or intSizeLog >= intUpperLimit Then
                    strConnectionString = "Integrated Security=SSPI;" + "Initial Catalog=V03NavDB_" & gstrVPUID & ";" + "Data Source=" + gstrDataSource + ";"
                    sqlconConnection = New SqlConnection
                    sqlconConnection.ConnectionString = strConnectionString
                    sqlconConnection.Open()
                    strSQL = "DROP TABLE t" & strSessionID & "_VAA;"
                    strReturn = ExecuteSQL(strSQL, sqlconConnection, Nothing)
                    If strReturn <> "" Then
                        FileMaintenance = strReturn
                        Exit Try
                    End If
                    'To shrink all data and log files for a specific database, execute the DBCC SHRINKDATABASE command.
                    strSQL = "DBCC SHRINKDATABASE(0);"
                    strReturn = ExecuteSQL(strSQL, sqlconConnection, Nothing)
                    If strReturn <> "" Then
                        FileMaintenance = strReturn
                        Exit Try
                    End If

                    sqlconConnection.Close()
                    sqlconConnection.Dispose()
                End If
            End If
            SqlConnection.ClearAllPools()

        Catch ex As Exception
            FileMaintenance = "FileMaintenance Exception: " & ex.Message.ToString &
                vbCrLf & ex.StackTrace.ToString & vbCrLf & ex.Source.ToString

        End Try

    End Function

    Public Function SelectNHDFlowline(ByVal mPoint As IPoint,
         ByRef mDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument, ByRef mMap As IMap, ByRef mActiveView As IActiveView) As String

        Dim strReturn As String
        Dim pContView As ESRI.ArcGIS.ArcMapUI.IContentsView
        Dim fcNHDFlowline As ESRI.ArcGIS.Geodatabase.IFeatureClass
        Dim pDataSet As ESRI.ArcGIS.Geodatabase.IDataset
        Dim pLayer As IGeoFeatureLayer
        Dim pEnvelope As ESRI.ArcGIS.Geometry.IEnvelope
        Dim pSpatialFilter As ESRI.ArcGIS.Geodatabase.ISpatialFilter
        Dim featNHDFlowline As ESRI.ArcGIS.Geodatabase.IFeature
        Dim fcurNHDFlowline As ESRI.ArcGIS.Geodatabase.IFeatureCursor
        Dim pPolyLine As IPolyline

        Dim pUid As New ESRI.ArcGIS.esriSystem.UID
        Dim strShapeFieldName As String
        Dim pFeatureSelection As ESRI.ArcGIS.Carto.IFeatureSelection

        Dim numnhdplusidFieldName As Integer
        Dim numReachcodeFieldName As Integer
        Dim numFlowdirFieldName As Integer
        Dim numFcodeFieldName As Integer
        Dim strFlowlineLayerName As String
        Dim numSearchTol As Double

        Dim pGCont As ESRI.ArcGIS.Carto.IGraphicsContainer
        Dim pGraphicsLayer As ESRI.ArcGIS.Carto.IGraphicsLayer
        Dim pMElement As ESRI.ArcGIS.Carto.IMarkerElement
        Dim pElement As ESRI.ArcGIS.Carto.IElement

        Dim pTable As ESRI.ArcGIS.Geodatabase.ITable
        Dim pDisplayTable As ESRI.ArcGIS.Carto.IDisplayTable
        Dim pRelQueryTable As ESRI.ArcGIS.Geodatabase.IRelQueryTable
        Dim pDestTable As ESRI.ArcGIS.Geodatabase.ITable
        Dim strOut As String
        Dim boolJoined As Boolean
        Dim boolFlowlineFound As Boolean
        Dim numSelectedFlowlines As Integer
        Dim strLayerName As String

        Dim pName As ESRI.ArcGIS.esriSystem.IName
        Dim pRtLocName As ESRI.ArcGIS.Geodatabase.IRouteLocatorName
        Dim pRtLocFlowline As ESRI.ArcGIS.Location.IRouteLocator2
        Dim pPointRouteLocation As ESRI.ArcGIS.Location.IRouteMeasurePointLocation = New ESRI.ArcGIS.Location.RouteMeasurePointLocation
        Dim pRouteLocation As ESRI.ArcGIS.Location.IRouteLocation
        Dim pRoute As ESRI.ArcGIS.Geodatabase.IFeature = New ESRI.ArcGIS.Geodatabase.Feature
        Dim i As Integer
        Dim pEnumResultReach As ESRI.ArcGIS.Location.IEnumRouteIdentifyResult
        Dim objNavigationOptions As NHDPlusHRVAANavToolbar.frmNavigationOptions
        Dim strIniFileName As String

        Try
            SelectNHDFlowline = ""

            'Confirms that only one NHDFlowline layer is active and it is selectable
            strReturn = ValidateSelectedLayer(mDoc, "NHDFlowline")
            If strReturn <> "" Then
                SelectNHDFlowline = strReturn
                Exit Try
            End If

            'INITIALIZE PARAMETERS
            gdblStartNHDPlusID = -1
            gstrStartReachcode = ""
            gdblStartMeasure = -1
            gdblDistance = 0

            'Determine if the selected layer is joined
            pDisplayTable = mDoc.SelectedLayer
            pTable = pDisplayTable.DisplayTable
            strOut = ""
            'Get the list of joined tables
            Do While TypeOf pTable Is ESRI.ArcGIS.Geodatabase.IRelQueryTable
                pRelQueryTable = pTable
                pDestTable = pRelQueryTable.DestinationTable
                pDataSet = pDestTable
                strOut = strOut & pDataSet.Name & vbNewLine
                pTable = pRelQueryTable.SourceTable
            Loop
            If (strOut = "") Then
                boolJoined = False
            Else
                boolJoined = True
                '8/30/2006
                strFlowlineLayerName = mDoc.SelectedLayer.Name
            End If

            pContView = mDoc.CurrentContentsView
            pLayer = mDoc.SelectedLayer
            pDataSet = pLayer

            'SET GLOBAL VAR
            gstrVPUPath = Left(pDataSet.Workspace.PathName, Len(pDataSet.Workspace.PathName))
            gstrVPUID = gstrVPUPath.Substring(gstrVPUPath.LastIndexOf("\") + 1).Replace("HRNHDPlus", "")
            gstrVPUID = gstrVPUID.Replace(".gdb", "")

            fcNHDFlowline = pLayer.FeatureClass

            'Get the selected feature in the NHDFlowline layer

            'Establish the search tolerance
            pEnvelope = mPoint.Envelope
            numSearchTol = mDoc.SearchTolerance
            pEnvelope.Expand(numSearchTol, numSearchTol, False)

            'Create a new spatial filter and use the new envelope as the geometry
            pSpatialFilter = New ESRI.ArcGIS.Geodatabase.SpatialFilter
            pSpatialFilter.Geometry = pEnvelope
            pSpatialFilter.SpatialRel = ESRI.ArcGIS.Geodatabase.esriSpatialRelEnum.esriSpatialRelIntersects
            strShapeFieldName = fcNHDFlowline.ShapeFieldName
            pSpatialFilter.OutputSpatialReference(strShapeFieldName) = mMap.SpatialReference
            pSpatialFilter.GeometryField = strShapeFieldName
            'Make a feature selection
            pFeatureSelection = pLayer 'QI
            mActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, Nothing, mActiveView.Extent)
            pFeatureSelection.SelectFeatures(pSpatialFilter, esriSelectionResultEnum.esriSelectionResultNew, True)
            'mActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, Nothing, mActiveView.Extent)

            numSelectedFlowlines = pFeatureSelection.SelectionSet.Count
            If numSelectedFlowlines = 0 Then
                SelectNHDFlowline = "No NHDFlowline feature has been selected"
                Exit Try
            End If
            If numSelectedFlowlines > 1 Then
                SelectNHDFlowline = "More than one NHDFlowline features have been selected"
                Exit Try
            End If
            If numSelectedFlowlines = 1 Then
                boolFlowlineFound = True
                strLayerName = pLayer.Name
            End If

            'Display the point  (black dot)
            pGraphicsLayer = mMap.BasicGraphicsLayer
            pGCont = pGraphicsLayer
            pMElement = New ESRI.ArcGIS.Carto.MarkerElement
            pElement = pMElement
            pMElement.Symbol = New ESRI.ArcGIS.Display.SimpleMarkerSymbol
            pElement.Geometry = mPoint
            pGCont.DeleteAllElements()
            pGCont.AddElement(pElement, 0)
            mActiveView.Refresh()

            'Make sure the selected feature has flow
            fcurNHDFlowline = fcNHDFlowline.Search(pSpatialFilter, False) 'Do the search
            featNHDFlowline = fcurNHDFlowline.NextFeature
            If Not featNHDFlowline Is Nothing Then
                numFlowdirFieldName = featNHDFlowline.Fields.FindField("FLOWDIR")
                numReachcodeFieldName = featNHDFlowline.Fields.FindField("REACHCODE")
                numnhdplusidFieldName = featNHDFlowline.Fields.FindField("nhdplusid")
                numFcodeFieldName = featNHDFlowline.Fields.FindField("FCODE")
                If (featNHDFlowline.Value(numFlowdirFieldName) = 0) Then   '"Uninitialized"
                    boolFlowlineFound = False
                    SelectNHDFlowline = "This NHDFlowline feature has no flow. (Flowdir is Uninitialized)"
                    Exit Try
                End If
                gnumFcodeValue = featNHDFlowline.Value(numFcodeFieldName)

                'SET GLOBAL VAR
                gstrStartReachcode = featNHDFlowline.Value(numReachcodeFieldName)
                gdblStartNHDPlusID = featNHDFlowline.Value(numnhdplusidFieldName)
                pPolyLine = featNHDFlowline.ShapeCopy
                gdblFrom = pPolyLine.FromPoint.M
                gdblTo = pPolyLine.ToPoint.M
                MarshalObject(pPolyLine)
            Else
                SelectNHDFlowline = "No NHDFlowline feature has been selected"
                Exit Try
            End If

            pFeatureSelection.Clear()

            'Implement a route locator. This is the object that knows how to find
            'locations along a route.  (i.e. measures)

            'Use the NHDFlowline feature class to create a route locator
            pName = pDataSet.FullName
            pRtLocName = New ESRI.ArcGIS.Location.RouteMeasureLocatorName
            With pRtLocName
                .RouteFeatureClassName = pName
                .RouteIDFieldName = "reachcode"
                .RouteIDIsUnique = False
                .RouteMeasureUnit = ESRI.ArcGIS.esriSystem.esriUnits.esriDecimalDegrees
                .RouteWhereClause = """reachcode"" = '" & gstrStartReachcode & "'"
            End With
            pName = pRtLocName
            pRtLocFlowline = pName.Open

            'Create a RouteMeasurePointLocation

            'Get the REACHES near the mouse click (within the tolerance envelope)
            'pPointRouteLocation = New ESRI.ArcGIS.Location.RouteMeasurePointLocation
            'The results of IRouteLocator2::Identify get sent to IEnumRouteIdentifyResult
            pEnumResultReach = pRtLocFlowline.Identify(pEnvelope, "")
            'For now - only want the one which nhdplusid matching FIRST one found
            For i = 1 To 1
                pEnumResultReach.Next(pRouteLocation, pRoute)
                pPointRouteLocation = pRouteLocation
                gdblStartMeasure = pPointRouteLocation.Measure
            Next i

            If gdblStartMeasure < 0 Then
                SelectNHDFlowline = "This NHDFlowline is not measured.  Cannot begin a navigation here."
            Else
                strIniFileName = gstrDllLocation & "\" & "NavigatorCaller.ini"
                gstrNavDBPath = INIRead(strIniFileName, "Application", "NavDBPath", "")
                If gstrNavDBPath = "" Then
                    gstrNavDBPath = Environment.GetEnvironmentVariable("Temp")
                End If

                'Show the Navigation Options Form
                objNavigationOptions = New NHDPlusHRVAANavToolbar.frmNavigationOptions
                objNavigationOptions.txtNavDBPath.Text = gstrNavDBPath
                objNavigationOptions.tbStartMeasure.Text = gdblStartMeasure.ToString
                objNavigationOptions.radStartMeasure.Checked = True
                objNavigationOptions.DoCancel = False
                objNavigationOptions.ShowDialog()
                If objNavigationOptions.DoCancel = True Then
                    'Cancel button was pressed
                    SelectNHDFlowline = "Cancel pressed, not navigating"
                    objNavigationOptions.Dispose()
                    MarshalObject(objNavigationOptions)
                    Exit Try
                End If
                gstrNavDBPath = objNavigationOptions.txtNavDBPath.Text.Trim
                INIWrite(strIniFileName, "Application", "NavDBPath", gstrNavDBPath)

                'Get start measure for navigation
                If objNavigationOptions.radWholeReach.Checked Then
                    'if going downstream, set startmeasure to MEASURE AT FROMPOINT OF SELECTED NHDFLOWLINE
                    'if going upstream, set startmeasure to MEASURE AT TOPOINT OF SELECTED NHDFLOWLINE
                    If gstrNavigationType = "UPTRIB" Or gstrNavigationType = "UPMAIN" Then
                        gdblStartMeasure = gdblTo
                    Else
                        gdblStartMeasure = gdblFrom
                    End If
                End If
                If objNavigationOptions.radStartMeasure.Checked Then
                    'Get start Measure from textbox  (it may have been edited)
                    gdblStartMeasure = Val(objNavigationOptions.tbStartMeasure.Text)
                End If

                '0 or valid stop distance
                gdblDistance = Val(objNavigationOptions.txtMaxDistance.Text)

                'Stop Attribute information
                gstrAttrName = objNavigationOptions.cbAttrName.Text.Trim
                gstrAttrCompare = objNavigationOptions.cbOperator.Text.Trim
                gstrAttrValue = Val(objNavigationOptions.tbAttrValue.Text.Trim)

            End If

        Catch ex As Exception
            SelectNHDFlowline = "SelectNHDFlowline Exception: " & ex.Message.ToString &
                vbCrLf & ex.StackTrace.ToString & vbCrLf & ex.Source.ToString

        Finally
            strReturn = Nothing

        End Try

    End Function

    Public Function RemovePreviousResults(ByRef mDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument, ByRef mMap As IMap, ByRef mActiveView As IActiveView) As String
        Dim pLayer As ESRI.ArcGIS.Carto.ILayer = Nothing
        Dim pFlayer As ESRI.ArcGIS.Carto.IFeatureLayer = Nothing
        Dim i As Short

        Try
            RemovePreviousResults = ""

            '-----------------------------------------------
            'REMOVE PREVIOUS NAVIGATION RESULTS FROM THE MAP
            '-----------------------------------------------
            For i = 0 To mMap.LayerCount - 1
                pLayer = mMap.Layer(i)
                If InStr(1, pLayer.Name, "Navigation Results", CompareMethod.Text) > 0 Then
                    pFlayer = pLayer
                    mMap.DeleteLayer(pFlayer)
                    mActiveView.Refresh()
                    mDoc.UpdateContents()
                    'MsgBox("deleted navigation results layer")
                    Exit For
                End If
            Next i
            MarshalObject(pFlayer)
            MarshalObject(pLayer)

        Catch ex As Exception
            RemovePreviousResults = "RemovePreviousResults Exception: " & ex.Message.ToString &
                vbCrLf & ex.StackTrace.ToString & vbCrLf & ex.Source.ToString
        End Try

    End Function

    Public Function DisplayResults(ByRef mDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument, ByRef mMap As IMap, ByRef mActiveView As IActiveView) As String
        Dim pWSFact As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory
        Dim pWS As ESRI.ArcGIS.Geodatabase.IWorkspace
        Dim pFeatws As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace
        Dim tabResults As ESRI.ArcGIS.Geodatabase.ITable
        Dim pLayer As ESRI.ArcGIS.Carto.ILayer = Nothing
        Dim pFlayer As ESRI.ArcGIS.Carto.IFeatureLayer = Nothing
        Dim pDataSet As ESRI.ArcGIS.Geodatabase.IDataset
        Dim pName As ESRI.ArcGIS.esriSystem.IName
        Dim pRMLName As ESRI.ArcGIS.Geodatabase.IRouteLocatorName
        Dim pRouteFc As ESRI.ArcGIS.Geodatabase.IFeatureClass
        Dim pRtProp As ESRI.ArcGIS.Geodatabase.IRouteEventProperties2
        Dim pRMLnProp As ESRI.ArcGIS.Location.IRouteMeasureLineProperties
        Dim pEventFC As ESRI.ArcGIS.Geodatabase.IFeatureClass
        Dim pRESN As ESRI.ArcGIS.Geodatabase.IRouteEventSourceName

        Dim pShpWSFact As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory
        Dim pShpWS As ESRI.ArcGIS.Geodatabase.IWorkspace
        Dim pShpFeatWS As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace
        Dim pOutFCN As ESRI.ArcGIS.Geodatabase.IFeatureClassName
        Dim pOutFeatDSN As ESRI.ArcGIS.Geodatabase.IFeatureDatasetName

        Dim pQFilt As ESRI.ArcGIS.Geodatabase.IQueryFilter
        Dim pFlds As ESRI.ArcGIS.Geodatabase.IFields
        Dim pOutFlds As ESRI.ArcGIS.Geodatabase.IFields
        Dim pFldChk As ESRI.ArcGIS.Geodatabase.IFieldChecker
        Dim pTempWS As ESRI.ArcGIS.Geodatabase.IWorkspace
        Dim pTempName As ESRI.ArcGIS.esriSystem.IName

        Dim pEnum As ESRI.ArcGIS.Geodatabase.IEnumInvalidObject
        Dim pConv As ESRI.ArcGIS.Geodatabase.IFeatureDataConverter2

        Dim pSRend As ESRI.ArcGIS.Carto.ISimpleRenderer
        Dim pLSymbol As ESRI.ArcGIS.Display.ISimpleLineSymbol
        Dim pGeoFL As ESRI.ArcGIS.Carto.IGeoFeatureLayer
        Dim pColor As ESRI.ArcGIS.Display.IRgbColor

        Dim boolReRead As Boolean = True
        Dim intLoops As Integer = 0

        Dim GP As New ESRI.ArcGIS.Geoprocessor.Geoprocessor
        Dim boolOverwriteOutput As Boolean = GP.OverwriteOutput
        Dim boolTemporaryMapLayers As Boolean = GP.TemporaryMapLayers
        Dim boolAddOutputsToMap As Boolean = GP.AddOutputsToMap
        GP.OverwriteOutput = True
        GP.TemporaryMapLayers = True
        GP.AddOutputsToMap = False

        Try
            DisplayResults = ""

            GP.OverwriteOutput = True
            GP.TemporaryMapLayers = True
            GP.AddOutputsToMap = False

            ''Connect to localdb to get navigation results as a table object

            '10.1 Only - Comment these lines out to run in 9.3
            Dim gpCreateDatabaseConnection As New ESRI.ArcGIS.DataManagementTools.CreateDatabaseConnection
            Dim gpUtilities As ESRI.ArcGIS.Geoprocessing.IGPUtilities

            gpCreateDatabaseConnection.database_platform = "SQL_SERVER"
            gpCreateDatabaseConnection.database = "V03NavDB_" & gstrVPUID
            gpCreateDatabaseConnection.account_authentication = "OPERATING_SYSTEM_AUTH"
            gpCreateDatabaseConnection.out_folder_path = strTempWorkAreaPath
            gpCreateDatabaseConnection.instance = "(localdb)\v11.0"
            gpCreateDatabaseConnection.out_name = "tconn.sde"
            If Not RunTool(GP, gpCreateDatabaseConnection, Nothing, Nothing) Then
                DisplayResults = "Problem running CreateDatabaseConnection tool" + vbCrLf
                Exit Try
            End If
            pWSFact = New ESRI.ArcGIS.DataSourcesGDB.SdeWorkspaceFactory   'CType(Activator.CreateInstance(factoryType), IWorkspaceFactory)
            pWS = pWSFact.OpenFromFile(strTempWorkAreaPath + "\tconn.sde", 0)

            'End of 10.1 only section

            '9.3 Only  - Comment these lines out to run in 10.1
            'Dim psLocalDB As IPropertySet = New PropertySetClass()
            'psLocalDB.SetProperty("PROVIDER", "SQLNCLI11)")
            'psLocalDB.SetProperty("SERVER", "(localdb)\v11.0")
            'psLocalDB.SetProperty("DATABASE", "V03NavDB_" & gstrVPUID)
            'psLocalDB.SetProperty("INTEGRATED SECURITY", "SSPI")
            'pWSFact = New ESRI.ArcGIS.DataSourcesOleDB.OLEDBWorkspaceFactory
            'pWS = pWSFact.Open(psLocalDB, 0)
            'End of 9.3 only section

            'Both 10.1 and 9.3
            pFeatws = pWS
            tabResults = pFeatws.OpenTable("V03NavDB_" & gstrVPUID & ".dbo." & "t" & strSessionID & "_NavResults")
            numResultsCount = tabResults.RowCount(Nothing)

            If numResultsCount = 0 Then
                MsgBox("Navigation results are empty.  No NHDFlowline features met the navigation criteria.")
            End If

            '------------------------
            'IDENTIFY THE ROUTE LAYER
            '------------------------
            pLayer = mDoc.SelectedLayer
            pFlayer = pLayer
            pRouteFc = pFlayer.FeatureClass
            If pRouteFc Is Nothing Then
                DisplayResults = "Could not find the route feature class"""
                Exit Try
            End If

            '-------------------------------------------
            'CREATE A ROUTE LOCATOR FROM THE ROUTE LAYER
            '-------------------------------------------
            pDataSet = pRouteFc
            pName = pDataSet.FullName
            pRMLName = New ESRI.ArcGIS.Location.RouteMeasureLocatorName
            With pRMLName
                .RouteFeatureClassName = pName
                .RouteIDFieldName = "reachcode"
                .RouteIDIsUnique = True
                .RouteMeasureUnit = ESRI.ArcGIS.esriSystem.esriUnits.esriDecimalDegrees
                .RouteWhereClause = ""
            End With

            '-----------------------------------------------
            'ESTABLISH THE PROPERTIES OF THE LINE EVENT LAYER
            '-----------------------------------------------
            pRtProp = New ESRI.ArcGIS.Location.RouteMeasureLineProperties
            With pRtProp
                .EventMeasureUnit = ESRI.ArcGIS.esriSystem.esriUnits.esriDecimalDegrees
                .EventRouteIDFieldName = "reachcode"
            End With
            pRMLnProp = pRtProp
            pRMLnProp.FromMeasureFieldName = "frommeas"
            pRMLnProp.ToMeasureFieldName = "tomeas"

            pDataSet = tabResults
            pName = pDataSet.FullName
            pRESN = New ESRI.ArcGIS.Location.RouteEventSourceName
            With pRESN
                .EventTableName = pName
                .EventProperties = pRMLnProp
                .RouteLocatorName = pRMLName
            End With

            '----------------------
            'CREATE THE EVENT LAYER
            '----------------------
            pName = pRESN
            pEventFC = pName.Open
            pFlayer = New FeatureLayer
            pFlayer.Selectable = True
            pFlayer.FeatureClass = pEventFC
            pFlayer.Name = "Navigation Results"

            '----------------------------
            'MAKE THE LAYER RED AND THICK
            '----------------------------
            'Create a color
            pColor = New ESRI.ArcGIS.Display.RgbColor
            pColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
            'Create a renderer
            pSRend = New ESRI.ArcGIS.Carto.SimpleRenderer
            'Create a line symbol object
            pLSymbol = New ESRI.ArcGIS.Display.SimpleLineSymbol
            With pLSymbol
                .Width = 3.4
                .Color = pColor
                .Style = ESRI.ArcGIS.Display.esriSimpleLineStyle.esriSLSSolid
            End With
            'Set the renderer symbol
            pSRend.Symbol = pLSymbol
            pGeoFL = pFlayer
            pGeoFL.Renderer = pSRend
            '------------------------
            'ADD THE LAYER TO THE MAP
            '------------------------
            pFlayer.Selectable = True
            mMap.AddLayer(pFlayer)
            mActiveView.Refresh()
            mDoc.UpdateContents()
            mActiveView.Refresh()
            mDoc.UpdateContents()

            pFeatws = Nothing
            pRouteFc = Nothing
            pLayer = Nothing
            pFlayer = Nothing
            pName = Nothing
            pRMLName = Nothing
            pRtProp = Nothing
            pRMLnProp = Nothing
            pRESN = Nothing
            pEventFC = Nothing
            pSRend = Nothing
            pLSymbol = Nothing
            pGeoFL = Nothing
            pColor = Nothing
            pShpWSFact = Nothing
            pShpWS = Nothing
            pShpFeatWS = Nothing
            pOutFCN = Nothing
            pOutFeatDSN = Nothing
            pQFilt = Nothing
            pEnum = Nothing
            pConv = Nothing
            pOutFCN = Nothing
            pFlds = Nothing
            pOutFlds = Nothing
            pFldChk = Nothing
            pTempWS = Nothing
            pTempName = Nothing

        Catch ex As Exception
            DisplayResults = "DisplayResults Exception: " & ex.Message.ToString &
                vbCrLf & ex.StackTrace.ToString & vbCrLf & ex.Source.ToString

        End Try

    End Function

    Public Function CallNavigator(ByRef mApplication As IApplication) As String
        Dim strINIFilename As String
        Dim intReturn As Integer
        Dim strNHDPlusV02DataLocation As String
        Dim boolReusableSessionID As Boolean = False
        Dim tBegin As DateTime
        Dim tBeginNav As DateTime
        Dim tsRun As TimeSpan
        Dim tsRunNav As TimeSpan
        Dim objLoadDB As NHDPlusHRNavigator.clsLoadSqlServerDB
        Dim objMakeWorking As NHDPlusHRNavigator.clsMakeWorkingTable
        Dim objNavigate As NHDPlusHRNavigator.clsHRNavigator

        Dim strConnectionString As String
        Dim sqlconConnection As SqlConnection
        Dim strReturn As String

        Dim oleConnection As OleDbConnection = Nothing
        Dim oleCommand As OleDbCommand = Nothing
        Dim boolSaveDataSource As Boolean = False

        Dim gpCopyRows As ESRI.ArcGIS.DataManagementTools.CopyRows
        Dim gpDelete As ESRI.ArcGIS.DataManagementTools.Delete

        Dim GP As New ESRI.ArcGIS.Geoprocessor.Geoprocessor
        Dim boolOverwriteOutput As Boolean = GP.OverwriteOutput
        Dim boolTemporaryMapLayers As Boolean = GP.TemporaryMapLayers
        Dim boolAddOutputsToMap As Boolean = GP.AddOutputsToMap
        GP.OverwriteOutput = True
        GP.TemporaryMapLayers = True
        GP.AddOutputsToMap = False
        Try
            GP.OverwriteOutput = True
            GP.TemporaryMapLayers = True
            GP.AddOutputsToMap = False

            mApplication.StatusBar.Message(0) = "Calling NHDPlusHR VAA Navigator"

            'RMD - 05/15/19 - Trying to let the user know the program is still running
            Dim appCursor As IMouseCursor = New MouseCursorClass
            appCursor.SetCursor(2) 'Hourglass

            CallNavigator = ""

            tBegin = Now()

            strINIFilename = gstrDllLocation & "\" & "NavigatorCaller.ini"
            strNHDPlusV02DataLocation = gstrVPUPath
            strSessionID = Format(Now(), "yyyyMMddhhmmssff")

            'Read NavDBPath from the INI file  6/10/2013
            If gstrNavDBPath = "" Then
                gstrNavDBPath = INIRead(strINIFilename, "Application", "NavDBPath", "")
            End If

            'JRH: 11/23/2012 for use with localDB
            gstrDataSource = "(LocalDB)\v11.0"

            '6/10/2013
            If Not gstrNavDBPath.ToUpper.Trim.Contains("NAVIGATORDBS") Then
                gstrNavDBPath = gstrNavDBPath & "\NavigatorDBs"
                INIWrite(strINIFilename, "Application", "NavDBPath", gstrNavDBPath)
            End If
            If Not Directory.Exists(gstrNavDBPath) Then
                Directory.CreateDirectory(gstrNavDBPath)
            End If

            'Modify 6/10/2013 for new NavDBPath as the location for navigation mdfs
            If File.Exists(gstrNavDBPath & "\V03NavDB_" & gstrVPUID & ".mdf") Then
                'Working db does previously exist - automatically use it AND read values from INI
                'Read the variables from the INI file 
                strTempWorkAreaPath = INIRead(strINIFilename, "Application", "TempWorkAreaPath", "Unknown")
                gstrDataSource = INIRead(strINIFilename, "Application", "DataSource", "Unknown")

                'Verify that the ini contained correct information
                'Set defaults and update if possible
                If Not My.Computer.FileSystem.DirectoryExists(strTempWorkAreaPath) Then
                    strTempWorkAreaPath = Environment.GetEnvironmentVariable("Temp")
                    INIWrite(strINIFilename, "Application", "TempWorkAreaPath", strTempWorkAreaPath)
                End If

                strTempWorkAreaPath = strTempWorkAreaPath & "\Nav_" & strSessionID
                My.Computer.FileSystem.CreateDirectory(strTempWorkAreaPath)

            Else
                'Working db does not previously exist - or exists in a different path
                'Establish defaults and write to NavigateCaller.ini
                strTempWorkAreaPath = Environment.GetEnvironmentVariable("Temp")

                INIWrite(strINIFilename, "Application", "DataSource", gstrDataSource)
                INIWrite(strINIFilename, "Application", "TempWorkAreaPath", strTempWorkAreaPath)
                INIWrite(strINIFilename, "Application", "NavDBPath", gstrNavDBPath)

                strTempWorkAreaPath = strTempWorkAreaPath & "\Nav_" & strSessionID
                My.Computer.FileSystem.CreateDirectory(strTempWorkAreaPath)

                'Create NavWorking DB
                objLoadDB = New NHDPlusHRNavigator.clsLoadSqlServerDB
                objLoadDB.SQLDataSource = gstrDataSource
                objLoadDB.AddToExisting = False
                objLoadDB.DatabaseLocation = gstrNavDBPath 'strNHDPlusV02DataLocation 6/10/2013
                objLoadDB.DatabaseName = "V03NavDB_" & gstrVPUID
                objLoadDB.TempWorkAreaPath = strTempWorkAreaPath
                objLoadDB.InputNHDPlusLocation = strNHDPlusV02DataLocation
                objLoadDB.AttrName = "pathlength, arbolatesu, totdasqkm, divdasqkm, divergence, rtndiv "

                intReturn = objLoadDB.LoadSQLServerDB()
                If intReturn > 0 Then
                    CallNavigator = "LoadDB Return Value: " + intReturn.ToString + vbCrLf +
                           "LoadDB ProcessStatus: " + objLoadDB.ProcessStatus.ToString + vbCrLf +
                           "LoadDB ProcessMessage: " + objLoadDB.ProcessMessage
                    Exit Try
                End If
            End If

            intReturn = 0
            If Not IsDBNull(intReturn) Then
                If intReturn = 1 Then
                    'working table already exists - use it
                Else
                    'working table does not exist - create it
                    objMakeWorking = New NHDPlusHRNavigator.clsMakeWorkingTable
                    objMakeWorking.SQLDataSource = gstrDataSource
                    objMakeWorking.DatabaseLocation = gstrNavDBPath 'strNHDPlusV02DataLocation 6/10/2013
                    objMakeWorking.DatabaseName = "V03NavDB_" & gstrVPUID
                    objMakeWorking.TempWorkAreaPath = strTempWorkAreaPath
                    objMakeWorking.SessionID = strSessionID
                    objMakeWorking.AttrName = gstrAttrName
                    objMakeWorking.StartNHDPlusID = gdblStartNHDPlusID
                    objMakeWorking.Navtype = gstrNavigationType
                    intReturn = objMakeWorking.MakeWorkingTable()
                    If intReturn > 0 Then
                        CallNavigator = "Problem making the working table for VPU " + vbCrLf +
                               "MakeWorking Return Value: " + intReturn.ToString + vbCrLf +
                               "MakeWorking ProcessStatus: " + objMakeWorking.ProcessStatus.ToString + vbCrLf +
                               "MakeWorking ProcessMessage: " + objMakeWorking.ProcessMessage + vbCrLf +
                               "Working Table Name: " + objMakeWorking.WorkingTableName
                        Exit Try
                    End If
                End If
            End If

            objNavigate = New NHDPlusHRNavigator.clsHRNavigator
            objNavigate.SQLDataSource = gstrDataSource
            objNavigate.DatabaseLocation = gstrNavDBPath   'strNHDPlusV02DataLocation   6/10/2013
            objNavigate.DatabaseName = "V03NavDB_" & gstrVPUID
            objNavigate.SessionID = strSessionID
            objNavigate.WorkingTableName = "t" & strSessionID & "_VAA"
            objNavigate.StartNHDPlusID = gdblStartNHDPlusID
            objNavigate.StartMeasure = gdblStartMeasure
            objNavigate.NavType = gstrNavigationType
            objNavigate.MaxDistance = gdblDistance
            objNavigate.AttrName = gstrAttrName
            objNavigate.AttrComp = gstrAttrCompare
            objNavigate.AttrValue = gstrAttrValue

            tBeginNav = Now()
            intReturn = objNavigate.VAANavigate()
            tsRunNav = Now.Subtract(tBeginNav)
            If intReturn <> 0 Then
                CallNavigator = "Navigation problem: " + vbCrLf +
                       "VAANavigate Return Value: " + intReturn.ToString + vbCrLf +
                       "VAANavigate ProcessStatus: " + objNavigate.ProcessStatus.ToString + vbCrLf +
                       "VAANavigate ProcessMessage: " + objNavigate.ProcessMessage
                Exit Try
            Else
                Exit Try   '8/11/2016
                'Write sql server results to dbf, then drop the sql server results
                gstrResultsDBF = "TNavResults.dbf"

                If File.Exists(strNHDPlusV02DataLocation + "\" + gstrResultsDBF) Then
                    gpDelete = New ESRI.ArcGIS.DataManagementTools.Delete
                    gpDelete.in_data = strNHDPlusV02DataLocation + "\" + gstrResultsDBF
                    If Not RunTool(GP, gpDelete, Nothing, Nothing) Then
                        CallNavigator = "Problem running Delete tool" + vbCrLf
                        Exit Try
                    End If
                    MarshalObject(gpDelete)
                End If
                If File.Exists(strTempWorkAreaPath + "\ttemp.dbf") Then
                    gpDelete = New ESRI.ArcGIS.DataManagementTools.Delete
                    gpDelete.in_data = strTempWorkAreaPath + "\ttemp.dbf"
                    If Not RunTool(GP, gpDelete, Nothing, Nothing) Then
                        CallNavigator = "Problem running Delete tool" + vbCrLf
                        Exit Try
                    End If
                    MarshalObject(gpDelete)
                End If

                GP.OverwriteOutput = True
                GP.TemporaryMapLayers = True
                GP.AddOutputsToMap = False

                ''Connect to localdb to get navigation results as a table object
                Dim pWSFact As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory
                Dim pWS As ESRI.ArcGIS.Geodatabase.IWorkspace
                Dim pFeatws As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace
                Dim tabResults As ESRI.ArcGIS.Geodatabase.ITable

                '10.1 Only - Comment these lines out to run in 9.3
                Dim gpCreateDatabaseConnection As New ESRI.ArcGIS.DataManagementTools.CreateDatabaseConnection
                gpCreateDatabaseConnection.database_platform = "SQL_SERVER"
                gpCreateDatabaseConnection.database = "V03NavDB_" & gstrVPUID
                gpCreateDatabaseConnection.account_authentication = "OPERATING_SYSTEM_AUTH"
                gpCreateDatabaseConnection.out_folder_path = strTempWorkAreaPath
                gpCreateDatabaseConnection.instance = "(localdb)\v11.0"
                gpCreateDatabaseConnection.out_name = "tconn.sde"
                'MsgBox(strTempWorkAreaPath)
                If Not RunTool(GP, gpCreateDatabaseConnection, Nothing, Nothing) Then
                    CallNavigator = "Problem running CreateDatabaseConnection tool" + vbCrLf
                    'MsgBox("problem with createddatabaseconnection")
                    Exit Try
                End If
                pWSFact = New ESRI.ArcGIS.DataSourcesGDB.SdeWorkspaceFactory   'CType(Activator.CreateInstance(factoryType), IWorkspaceFactory)
                pWS = pWSFact.OpenFromFile(strTempWorkAreaPath + "\tconn.sde", 0)
                'End of 10.1 only section

                '9.3 Only  - Comment these lines out to run in 10.1
                'Dim psLocalDB As IPropertySet = New PropertySetClass()
                'psLocalDB.SetProperty("PROVIDER", "SQLNCLI11)")
                'psLocalDB.SetProperty("SERVER", "(localdb)\v11.0")
                'psLocalDB.SetProperty("DATABASE", "V03NavDB_" & gstrVPUID)
                'psLocalDB.SetProperty("INTEGRATED SECURITY", "SSPI")
                'pWSFact = New ESRI.ArcGIS.DataSourcesOleDB.OLEDBWorkspaceFactory
                'pWS = pWSFact.Open(psLocalDB, 0)
                'End of 9.3 only section

                'Both 10.1 and 9.3
                pFeatws = pWS
                tabResults = pFeatws.OpenTable("V03NavDB_" & gstrVPUID & ".dbo." & "t" & strSessionID & "_NavResults")
                numResultsCount = tabResults.RowCount(Nothing)

                If numResultsCount = 0 Then
                    MsgBox("Navigation results are empty.  No NHDFlowline features met the navigation criteria.")
                End If

                'Couldn't use OPENROWSET with localdb, so can't go directly to dbf from sql server mdf
                'SO....  get sql server table as an arcobjects table object, make tableview from the arcobjects table, and then copy rows to ttemp.dbf
                '  
                'jrh8/11/2016
                Dim gpMakeTableView As ESRI.ArcGIS.DataManagementTools.MakeTableView
                gpMakeTableView = New ESRI.ArcGIS.DataManagementTools.MakeTableView
                gpMakeTableView.in_table = tabResults
                gpMakeTableView.out_view = "temptableview"
                If Not RunTool(GP, gpMakeTableView, Nothing, Nothing) Then
                    CallNavigator = "Problem running MakeTableView tool" + vbCrLf
                    Exit Try
                End If

                gpCopyRows = New ESRI.ArcGIS.DataManagementTools.CopyRows
                gpCopyRows.in_rows = "temptableview"
                gpCopyRows.out_table = strTempWorkAreaPath + "\ttemp.dbf"
                If Not RunTool(GP, gpCopyRows, Nothing, Nothing) Then
                    CallNavigator = "Problem running CopyRows tool" + vbCrLf
                    Exit Try
                End If
                File.Copy(strTempWorkAreaPath + "\ttemp.dbf", strNHDPlusV02DataLocation + "\" + gstrResultsDBF)
                gpDelete = New ESRI.ArcGIS.DataManagementTools.Delete
                gpDelete.in_data = strTempWorkAreaPath + "\ttemp.dbf"
                If Not RunTool(GP, gpDelete, Nothing, Nothing) Then
                    CallNavigator = "Problem running Delete tool" + vbCrLf
                    Exit Try
                End If
                MarshalObject(gpDelete)

                Dim gpUtilities As ESRI.ArcGIS.Geoprocessing.IGPUtilities
                gpUtilities = New ESRI.ArcGIS.Geoprocessing.GPUtilities
                gpUtilities.RemoveInternalTable("temptableview")

                gpUtilities = Nothing
                MarshalObject(gpCopyRows)
                MarshalObject(gpMakeTableView)
                GP.OverwriteOutput = boolOverwriteOutput
                GP.TemporaryMapLayers = boolTemporaryMapLayers
                GP.AddOutputsToMap = boolAddOutputsToMap
                MarshalObject(GP)

                strConnectionString = "Integrated Security=SSPI;" + "Initial Catalog=V03NavDB_" & gstrVPUID & ";" + "Data Source=" + gstrDataSource + ";"
                sqlconConnection = New SqlConnection
                sqlconConnection.ConnectionString = strConnectionString
                sqlconConnection.Open()

                strReturn = ExecuteSQL("DROP TABLE t" & strSessionID & "_NavResults", sqlconConnection, 0)
                If strReturn <> "" Then
                    CallNavigator = strReturn
                    Exit Try
                End If
                strReturn = ExecuteSQL("DROP TABLE t" & strSessionID & "_VAA", sqlconConnection, 0)
                If strReturn <> "" Then
                    CallNavigator = strReturn
                    Exit Try
                End If

                sqlconConnection.Close()
                sqlconConnection.Dispose()

                If File.Exists(strTempWorkAreaPath + "\ttemp.dbf") Then
                    File.Delete(strTempWorkAreaPath + "\ttemp.dbf")

                End If
                If File.Exists(strTempWorkAreaPath + "\ttemp.dbf.xml") Then

                    File.Delete(strTempWorkAreaPath + "\ttemp.dbf.xml")
                End If
                If File.Exists(strTempWorkAreaPath + "\tconn.sde") Then
                    File.Delete(strTempWorkAreaPath + "\tconn.sde")
                End If
                My.Computer.FileSystem.DeleteDirectory(strTempWorkAreaPath, FileIO.DeleteDirectoryOption.DeleteAllContents)

                tsRun = Now.Subtract(tBegin)
                If boolSaveDataSource Then
                    INIWrite(strINIFilename, "Application", "DataSource", gstrDataSource)
                End If

            End If


        Catch ex As Exception
            CallNavigator = "CallNavigator Exception: " & ex.Message.ToString &
                vbCrLf & ex.StackTrace.ToString & vbCrLf & ex.Source.ToString
        Finally
            If File.Exists(strTempWorkAreaPath + "\ttemp.dbf") Then
                File.Delete(strTempWorkAreaPath + "\ttemp.dbf")
            End If

        End Try

    End Function

    Public Function ValidateSelectedLayer(ByRef pMxDoc As IMxDocument, ByVal strLayerName As String) As String
        Dim pContView As ESRI.ArcGIS.ArcMapUI.IContentsView
        Dim pVarSelectedItem As Object
        Dim pTLayer As ESRI.ArcGIS.Carto.ILayer
        Dim pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
        Dim pFeatureClass As ESRI.ArcGIS.Geodatabase.IFeatureClass
        Dim pDataSet As ESRI.ArcGIS.Geodatabase.IDataset
        Dim pLayer As IGeoFeatureLayer
        Dim strDataName As String

        Try
            ValidateSelectedLayer = ""

            pContView = pMxDoc.CurrentContentsView
            pLayer = pMxDoc.SelectedLayer

            If TypeOf pContView Is ESRI.ArcGIS.ArcMapUI.TOCDisplayView Then
                If IsDBNull(pContView.SelectedItem) Then
                    ValidateSelectedLayer = "One " & strLayerName & " layer must be active."
                    Exit Try
                Else
                    pVarSelectedItem = pContView.SelectedItem
                    If TypeOf pVarSelectedItem Is ESRI.ArcGIS.Carto.ILayer Then
                        pTLayer = pVarSelectedItem
                        If pTLayer.Name.ToUpper.Trim <> strLayerName.ToUpper.Trim Then
                            ValidateSelectedLayer = "One " & strLayerName & " layer must be active."
                            Exit Try
                        End If

                        'Make sure selections are possible for the layer
                        pFeatureLayer = pTLayer
                        If pFeatureLayer.Selectable Then
                            pFeatureClass = pFeatureLayer.FeatureClass
                            pDataSet = pFeatureClass
                            strDataName = pDataSet.Name
                            'Make sure the layer is a NHDFLOWLINE featureclass
                            If Not (IsNHDFlowline(pFeatureClass, strDataName)) Then
                                ValidateSelectedLayer = "Selected layer is not an " & strLayerName & " layer"
                                Exit Try
                            End If
                        Else
                            ValidateSelectedLayer = strLayerName & " layer is not selectable"
                            Exit Try
                        End If
                    Else
                        If TypeOf pVarSelectedItem Is ESRI.ArcGIS.esriSystem.ISet Then
                            ValidateSelectedLayer = "Only one " & strLayerName & " layer may be active."
                            Exit Try
                        Else
                            ValidateSelectedLayer = "No " + strLayerName + " layer is active."
                            Exit Try
                        End If
                    End If
                End If
            Else
                ValidateSelectedLayer = "No " + strLayerName + " layer is active."
                Exit Try
            End If

        Catch ex As Exception
            ValidateSelectedLayer = "ValidateSelectedLayer Exception: " & ex.Message.ToString & vbCrLf &
                                              ex.StackTrace.ToString & vbCrLf &
                                              ex.Source.ToString
        Finally
            pContView = Nothing
            pVarSelectedItem = Nothing
            pTLayer = Nothing
            pFeatureLayer = Nothing
            pFeatureClass = Nothing
            pDataSet = Nothing
            pLayer = Nothing
            strDataName = Nothing

        End Try

    End Function

    Public Function IsNHDFlowline(ByVal pFC As IFeatureClass, ByVal strDataName As String) As Boolean
        Dim CheckFeatureType As Boolean

        'Determine if the selected layer is a NHDFlowline layer by
        'verifying the underlying data for the layer contains NHDFlowline and
        'that the layer
        'contains a nhdplusid field.

        CheckFeatureType = True
        If (pFC.FindField("nhdplusid") = -1) Then
            MsgBox("no nhdplusid")
            CheckFeatureType = False
        End If

        If (InStr(1, strDataName, "NHDFlowline", vbTextCompare) <= 0) Then
            MsgBox(strDataName)
            CheckFeatureType = False
        End If

        IsNHDFlowline = CheckFeatureType
    End Function

    Public Function ExecuteSQL(ByVal strSQL As String, ByRef sqlconConnection As SqlConnection, ByVal intTimeout As Integer) As String
        Dim strExMsg As String
        Dim sqlcmdCommand As SqlCommand
        Try
            sqlcmdCommand = New SqlCommand(strSQL, sqlconConnection)
            sqlcmdCommand.CommandTimeout = intTimeout
            sqlcmdCommand.ExecuteNonQuery()
            ExecuteSQL = ""
        Catch ex As Exception
            strExMsg = ex.Message.ToString + vbCrLf +
                ex.StackTrace.ToString + vbCrLf +
                ex.Source.ToString
            ExecuteSQL = "SQL Exception: " + vbCrLf + strSQL + vbCrLf + strExMsg
        Finally
            sqlcmdCommand = Nothing
            strExMsg = Nothing
        End Try
    End Function

    Public Sub MarshalObject(ByRef obj As Object)
        If Not obj Is Nothing Then
            If System.Runtime.InteropServices.Marshal.IsComObject(obj) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            End If
            obj = Nothing
        End If
    End Sub

    Public Function FindDataSource(ByRef strDataSource As String) As String
        Dim dtTemp As DataTable = New DataTable
        Dim strMachineName As String
        Dim strExMsg As String

        Try
            FindDataSource = ""
            strDataSource = ""

            strMachineName = UCase(Environment.MachineName)

            dtTemp = SqlDataSourceEnumerator.Instance.GetDataSources

            'For Each column In dtTemp.Columns
            ' Next
            For Each row As DataRow In dtTemp.Select
                If Not IsDBNull(row("servername")) And Not IsDBNull(row("instancename")) Then
                    If UCase(row("InstanceName")) = "SQLEXPRESS" And UCase(row("ServerName")) = strMachineName Then
                        Exit For
                    End If
                End If
            Next row
            If strDataSource = "" Then
                For Each row As DataRow In dtTemp.Select
                    If Not IsDBNull(row("servername")) And Not IsDBNull(row("instancename")) Then
                        If row("InstanceName") <> "" And UCase(row("ServerName")) = strMachineName Then
                            strDataSource = row("ServerName") + "\" + row("InstanceName")
                            Exit For
                        End If
                    End If
                Next row
            End If

        Catch ex As Exception
            strExMsg = ex.Message.ToString + vbCrLf +
                  ex.StackTrace.ToString + vbCrLf +
                  ex.Source.ToString
            FindDataSource = "FindDataSource Exception: " + strExMsg

        Finally
            dtTemp.Dispose()
            dtTemp = Nothing
            strMachineName = Nothing
            strExMsg = Nothing

        End Try

    End Function

    Public Function RunTool(ByRef geoprocessor As Geoprocessor, ByVal process As IGPProcess, ByVal TC As ITrackCancel, ByVal swMessage As StreamWriter) As Boolean
        RunTool = False
        Try
            geoprocessor.Execute(process, Nothing)
            If Not ReturnMessages(geoprocessor, swMessage) Then Exit Function
        Catch err As Exception
            If Not ReturnMessages(geoprocessor, swMessage) Then Exit Function

        End Try
        RunTool = True

    End Function

    Public Function ReturnMessages(ByRef gp As Geoprocessor, ByVal swMessage As StreamWriter) As Boolean
        ReturnMessages = False
        Dim strMessage As String = ""
        Dim Count As Integer
        If gp.MessageCount > 0 Then
            For Count = 0 To gp.MessageCount - 1
                strMessage = gp.GetMessage(Count)
            Next
        End If
        ReturnMessages = True

    End Function

End Module