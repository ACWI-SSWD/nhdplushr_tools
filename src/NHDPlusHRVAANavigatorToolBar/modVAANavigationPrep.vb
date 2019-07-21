Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module modVAANavigationPrep
	'3/11/2008 - V103
	'  No change in code.  Recompiling due to a change in the navigator dll.
	
	'10/12/2007 - V102
	'  Changed TNavWorkSkel.mdb version to 4
	
	
	'07/24/2007 - V101 of COM object toolkit.
	'  -Recompiled with new reference to the vaanavigator.
	
	'TODO - Next version, if a user is on a nhdflowline layer - make sure it is selectable.
	'  If not - display a message.
	
	'01/05/2007 - 1/24/2007 V101
	'   Make it so that Navigation Results shapes get rendered correctly when
	'   there are multiple NHDFlowline workspaces loaded.
	'   The change involves adding pTlayer as a parameter to RunAddResultsToMap and use it
	'   to set the route layer.
	'
	'   Changed color of the tool graphics and added the text (COM) to the toolbar and
	'   tool names/captions.
	'
	'   Changed schema version of the TNavWork.mdb to 3
	'
	'
	'9/20/2006  - V 101 Beta3
	'   There has been a change to TNavWorkSkel.mdb.
	'   Updating the code to look for 2 in tblVersion
	'9/14-15/2006 - V 101 Beta2
	'   The VAA Navigator added tblVersion to TNavWorkSkel to contain one
	'      field (version) and one record,
	'      representing the version of TNavWork.mdb.  Setting the value
	'      to 1.
	'      Now run preprocessing if:
	'         TNavWork does not exist OR
	'         TNavWork.mdb does exist and tblVersion does not exist OR
	'         TNavWork.mdb does exist and tblVersion does exist and
	'            tblVersion.version for first record in tblVersion
	'            does not equal 1
	'      Changed the VAA Navigator toolbar to display messages
	'         if the preprocessor will need to be run.
	'   Found a problem with theESRI fix for deleteing navigation results
	'     It only works when TNavigation_Events is the only shapefile
	'     in the uppermost directory of the NHDPlus workspace.
	'     Fixed it by looping through ALL shapefiles in the uppermost
	'     directory of the workspace.
	
	'8/29/2006  - 8/30/2006 V 101 Beta1
	'  ESRI fix for completely removing previous results
	'  Add message when no ContentsView
	'  Add message when no layer selected
	'  Add message when no feature selected
	'  Add code to be able to find comid/flowdir fields
	'     when there are joins
	'  Add message and do not call navigator when start nhdflowline
	'     has flowdir of uninitialized
	'  Add message saying that TNavWork.mdb must be created when it is not there.
	'  Add mouse hourglass commands when calling navigator
	
	'Public Declare Function VAANavigate Lib "d:\hsc\nhdplus\vaanavigator\vaanavigator.dll" (ByVal strNavType As String, ByVal strDataPath As String, numStartComid As Long, numStartMeasure As Double, numStopComid As Long, numStopMeasure As Double, numMaxDistance As Double, numMaxTime As Double, ByVal strAppPath As String, ByVal strLog As String) As Integer
	'Public Declare Function VAANavigate Lib "vaanavigator.dll" (ByVal strNavType As String, ByVal strDataPath As String, numStartComid As Long, numStartMeasure As Double, numStopComid As Long, numStopMeasure As Double, numMaxDistance As Double, numMaxTime As Double, ByVal strAppPath As String, ByVal strLog As String) As Integer
	
	
	Sub VAANavigationPrep(ByRef strCheck As String, ByRef pStartPoint As ESRI.ArcGIS.Geometry.IPoint, ByRef pDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument, ByRef pmap As ESRI.ArcGIS.Carto.IMap, ByRef pActiveView As ESRI.ArcGIS.Carto.IActiveView)
		
		On Error GoTo ErrHandler
		
		'MsgBox "Navigator toolbar currently using development vaanavigator.dll"
		
		'Dim pDoc As IMxDocument
		'Dim pActiveView As IActiveView
		'Dim pmap As IMap
		Dim SearchTol As Double
		
		Dim fso As Object
		
		Dim varMinRunTime As Object
		Dim varSecRunTime As Object
		Dim varSs_time As Object
		Dim varEs_time As Object
		
		'Set pDoc = Application.Document
		'Set pActiveView = pDoc.FocusMap
		'Set pmap = pDoc.FocusMap
		
		Dim pEnvelope As ESRI.ArcGIS.Geometry.IEnvelope
		Dim pSpatialFilter As ESRI.ArcGIS.Geodatabase.ISpatialFilter
		Dim pEnumLayer As ESRI.ArcGIS.Carto.IEnumLayer
		Dim pFeature As ESRI.ArcGIS.Geodatabase.IFeature
		Dim pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim pFeatureClass As ESRI.ArcGIS.Geodatabase.IFeatureClass
		Dim pFeatureCursor As ESRI.ArcGIS.Geodatabase.IFeatureCursor
		Dim pEnumRow As ESRI.ArcGIS.Editor.IEnumRow
		Dim pUid As New ESRI.ArcGIS.esriSystem.UID
		Dim ShapeFieldName As String
		Dim pContView As ESRI.ArcGIS.ArcMapUI.IContentsView
		Dim pVarSelectedItem As Object
		Dim pTLayer As ESRI.ArcGIS.Carto.ILayer
		Dim pFeatureSelection As ESRI.ArcGIS.Carto.IFeatureSelection
		Dim pQueryFilter As ESRI.ArcGIS.Geodatabase.IQueryFilter
		Dim strQuery As String
		
		Dim boolLayerFound As Boolean
		Dim boolFlowlineFound As Boolean
		Dim numSelectedFlowlines As Integer
		Dim numStartComid As Integer
		Dim strLayerName As String
		
		Dim pGCont As ESRI.ArcGIS.Carto.IGraphicsContainer
		Dim pGraphicsLayer As ESRI.ArcGIS.Carto.IGraphicsLayer
		Dim pMElement As ESRI.ArcGIS.Carto.IMarkerElement
		Dim pElement As ESRI.ArcGIS.Carto.IElement
		Dim pPC As ESRI.ArcGIS.Geometry.IPointCollection
		
		'8/29/2006
		Dim numComidFieldName As Integer
		Dim numFlowdirFieldName As Integer
		Dim strFlowlineLayerName As String
		
		Dim numMaxDistance As Double
		
		If pmap.LayerCount = 0 Then
			MsgBox("No layers have been added to the map.")
			Exit Sub
		End If
		
		fso = CreateObject("Scripting.FileSystemObject")
		
		boolLayerFound = False
		boolFlowlineFound = False
		
		'   'Get the WS Name and location
		'   Dim pInfo As Variant
		'   pInfo = GetWSInfo(pDoc)
		'   Dim strWSName As String
		'   Dim strFileName As String
		'   strWSName = pInfo(0)
		'   strFileName = pInfo(1)
		
		'Expand the points envelope to give better search results
		'UPGRADE_WARNING: Couldn't resolve default property of object pStartPoint.Envelope. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pEnvelope = pStartPoint.Envelope
		'Establish the search tolerance
		SearchTol = pDoc.SearchTolerance
		pEnvelope.Expand(SearchTol, SearchTol, False)
		
		'Create a new spatial filter and use the new envelope as the geometry
		pSpatialFilter = New ESRI.ArcGIS.Geodatabase.SpatialFilter
		pSpatialFilter.Geometry = pEnvelope
		pSpatialFilter.SpatialRel = ESRI.ArcGIS.Geodatabase.esriSpatialRelEnum.esriSpatialRelIntersects
		
		pContView = pDoc.CurrentContentsView
		
		'If we have a DisplayView active
		'Find the selected layer
		'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Dim pTable As ESRI.ArcGIS.Geodatabase.ITable
		Dim pDisplayTable As ESRI.ArcGIS.Carto.IDisplayTable
		Dim pRelQueryTable As ESRI.ArcGIS.Geodatabase.IRelQueryTable
		Dim pDestTable As ESRI.ArcGIS.Geodatabase.ITable
		Dim pDataSet As ESRI.ArcGIS.Geodatabase.IDataset
		Dim strOut As String
		Dim boolJoined As Boolean
		If TypeOf pContView Is ESRI.ArcGIS.ArcMapUI.TOCDisplayView Then
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If IsDbNull(pContView.SelectedItem) Then
				MsgBox("No layer is selected.")
			Else
				pVarSelectedItem = pContView.SelectedItem
				If TypeOf pVarSelectedItem Is ESRI.ArcGIS.Carto.ILayer Then
					pTLayer = pVarSelectedItem
					'MsgBox "Found a selected layer " & pTLayer.Name
					boolLayerFound = True
					
					'Determine if the layer has joined.
					'If joined, the field names are different for queries.
					
					pDisplayTable = pTLayer
					pTable = pDisplayTable.DisplayTable
					'Get the list of joined tables
					'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					Do While TypeOf pTable Is ESRI.ArcGIS.Geodatabase.IRelQueryTable
						pRelQueryTable = pTable
						pDestTable = pRelQueryTable.DestinationTable
						pDataSet = pDestTable
						strOut = strOut & pDataSet.Name & vbNewLine
						pTable = pRelQueryTable.SourceTable
					Loop 
					If (strOut = "") Then
						'MsgBox "Not Joined"
						boolJoined = False
					Else
						'MsgBox "The Joined tables include:" & vbNewLine & strOut
						boolJoined = True
						'8/30/2006
						strFlowlineLayerName = pTLayer.Name
					End If
					
				Else
					'8/29/2006
					If TypeOf pVarSelectedItem Is ESRI.ArcGIS.esriSystem.ISet Then
						MsgBox("Multiple feature layers are selected.")
					Else
						MsgBox("No feature layer is selected.")
					End If
				End If
			End If
		Else
			'8/29/2006
			MsgBox("No feature layer is selected.")
		End If
		
		Dim strDataName As String
		Dim strDataPath As String
		Dim strWorkingPath As String
		Dim strOutputFile As String
		Dim pWorkspace As ESRI.ArcGIS.Geodatabase.IWorkspace
		Dim pName As ESRI.ArcGIS.esriSystem.IName
		Dim pRtLocName As ESRI.ArcGIS.Geodatabase.IRouteLocatorName
		Dim pRtLocFlowline As ESRI.ArcGIS.Location.IRouteLocator2
		Dim pPointRouteLocation As ESRI.ArcGIS.Location.IRouteMeasurePointLocation
		Dim pRouteLocation As ESRI.ArcGIS.Location.IRouteLocation
		Dim pRoute As ESRI.ArcGIS.Geodatabase.IFeature
		Dim i As Integer
		Dim strStartRchCode As String
		Dim numStartMeasure As Double
		Dim pEnumResultReach As ESRI.ArcGIS.Location.IEnumRouteIdentifyResult
		Dim pWS As ESRI.ArcGIS.Geodatabase.IWorkspace
		Dim pWorkspaceFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory
		Dim pDS As ESRI.ArcGIS.Geodatabase.IDataset
		Dim pEnumDS As ESRI.ArcGIS.Geodatabase.IEnumDataset
		Dim boolGO As Boolean
		Dim oDB As DAO.Database
		Dim td As DAO.TableDef
		Dim boolTableExists As Object
		Dim adoCn As ADODB.Connection
		Dim rstResults As ADODB.Recordset
		Dim numReturn As Short
		Dim objNav As VAANavigatorCOM.clsVAANavigate
		Dim strConn As Object
		Dim strSQL As String
		If (boolLayerFound) Then
			
			'MsgBox "layer found"
			
			'Make sure selections are possible for the layer
			pFeatureLayer = pTLayer
			If pFeatureLayer.Selectable Then
				
				pFeatureClass = pFeatureLayer.FeatureClass
				pDataSet = pFeatureClass
				
				strDataName = pDataSet.Name
				
				'Make sure the layer is a NHDFLOWLINE featureclass
				If (IsNHDFlowline(pFeatureClass, strDataName)) Then
					
					'Get the selected feature in the NHDFlowline layer
					ShapeFieldName = pFeatureClass.ShapeFieldName
					'UPGRADE_WARNING: Couldn't resolve default property of object pSpatialFilter.OutputSpatialReference. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					pSpatialFilter.OutputSpatialReference(ShapeFieldName) = pmap.SpatialReference
					pSpatialFilter.GeometryField = ShapeFieldName
					'UPGRADE_WARNING: Couldn't resolve default property of object pSpatialFilter. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					pFeatureCursor = pFeatureClass.Search(pSpatialFilter, False) 'Do the search
					pFeature = pFeatureCursor.NextFeature
					If Not pFeature Is Nothing Then
						
						'8/30/2006
						'UPGRADE_WARNING: Couldn't resolve default property of object pFeature.Fields. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						numFlowdirFieldName = pFeature.Fields.FindField("FLOWDIR")
						'UPGRADE_WARNING: Couldn't resolve default property of object pFeature.Fields. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						numComidFieldName = pFeature.Fields.FindField("COMID")
						
						'Display the point
						pGraphicsLayer = pmap.BasicGraphicsLayer
						pGCont = pGraphicsLayer
						pMElement = New ESRI.ArcGIS.Carto.MarkerElement
						pElement = pMElement
						'UPGRADE_WARNING: Couldn't resolve default property of object New (esriDisplay.SimpleMarkerSymbol). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						pMElement.Symbol = New ESRI.ArcGIS.Display.SimpleMarkerSymbol
						'UPGRADE_WARNING: Couldn't resolve default property of object pStartPoint. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						pElement.Geometry = pStartPoint
						pGCont.AddElement(pElement, 0)
						pActiveView.Refresh()
						
						'Make a feature selection
						pFeatureSelection = pFeatureLayer 'QI
						'Create the query filter
						pQueryFilter = New ESRI.ArcGIS.Geodatabase.QueryFilter
						'Determine the query needed to select the feature near the mouse click
						'MsgBox pFeature.Fields.Field(2).Name
						If (boolJoined) Then
							'8/30/2006 Modified this line to get prefix based on layer name
							'UPGRADE_WARNING: Couldn't resolve default property of object pFeature.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strQuery = strFlowlineLayerName & ".COMID = " & Str(pFeature.Value(numComidFieldName))
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object pFeature.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strQuery = "COMID = " & Str(pFeature.Value(numComidFieldName))
						End If
						'MsgBox strQuery
						'UPGRADE_WARNING: Couldn't resolve default property of object pFeature.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						numStartComid = pFeature.Value(numComidFieldName)
						
						'MsgBox strQuery
						pQueryFilter.WhereClause = strQuery
						'Refresh the original selection
						'pActiveView.PartialRefresh esriViewGeoSelection, Nothing, Nothing
						'Perform the selection
						pFeatureSelection.SelectFeatures(pQueryFilter, ESRI.ArcGIS.Carto.esriSelectionResultEnum.esriSelectionResultNew, False)
						'Refresh the new selection
						'pActiveView.PartialRefresh esriViewGeoSelection, Nothing, Nothing
						
						numSelectedFlowlines = pFeatureSelection.SelectionSet.Count
						'MsgBox "Count: " + Str(numSelectedFlowlines)
						If numSelectedFlowlines = 1 Then
							boolFlowlineFound = True
							strLayerName = pTLayer.Name
						End If
						If numSelectedFlowlines = 0 Then
							MsgBox("No NHDFlowline feature has been selected")
						End If
						If numSelectedFlowlines > 1 Then
							MsgBox("More than one NHDFlowline features have been selected")
						End If
						pFeatureSelection.SelectFeatures(Nothing, ESRI.ArcGIS.Carto.esriSelectionResultEnum.esriSelectionResultSubtract, False)
						'8/30/2006
						'UPGRADE_WARNING: Couldn't resolve default property of object pFeature.Value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If (pFeature.Value(numFlowdirFieldName) = "Uninitialized") Then
							boolFlowlineFound = False
							MsgBox("This NHDFlowline feature has no flow. (Flowdir is Uninitialized)")
						End If
						
					Else
						MsgBox("No NHDFlowline feature has been selected")
					End If
				Else
					MsgBox("Selected layer is not an NHDFlowline layer")
				End If
			End If
			
			If (boolFlowlineFound) Then
				
				
				'Gather information to pass to SEPARATE navigation dll
				
				'Navigation direction/type (UPMAIN,DNMAIN,UPTRIB,DNDIV,point2point)
				'NHDGEOinSHP workspace location (complete drive, path to openme.txt - not including openme.txt, including trailing slash)
				'Desired event table name - complete with path name
				'Working directory - path to a working directory - i.e. location for messages, etc.
				'Start NHDFlowline Comid
				'Start Measure
				'Stop NHDFlowline Comid
				'Stop Measure
				'Maximum Distance
				'Maximum Time of Travel in hours (only for NHDPlus navigations)
				
				'Dim pDataSet As IDataset
				
				pDataSet = pFeatureClass
				pWorkspace = pDataSet.Workspace
				
				strDataPath = Left(pWorkspace.PathName, Len(pWorkspace.PathName) - 11)
				strOutputFile = Environ("TEMP") & "\results.txt"
				'strWorkingPath = Environ("TEMP") + "\"
				strWorkingPath = My.Application.Info.DirectoryPath
				
				'MsgBox strDataPath
				'MsgBox strOutputFile
				
				'Create a route locator. This is the object that knows how to find
				'locations along a route.
				
				'Use the NHDFlowline feature class
				'to create a route locator
				pName = pDataSet.FullName
				pRtLocName = New ESRI.ArcGIS.Location.RouteMeasureLocatorName
				With pRtLocName
					.RouteFeatureClassName = pName
					.RouteIDFieldName = "reachcode"
					.RouteIDIsUnique = False
					.RouteMeasureUnit = ESRI.ArcGIS.esriSystem.esriUnits.esriDecimalDegrees
					.RouteWhereClause = ""
				End With
				pName = pRtLocName
				pRtLocFlowline = pName.Open
				
				'Create a RouteMeasurePointLocation
				
				
				'Get the REACHES near the mouse click (within the tolerance envelope)
				pPointRouteLocation = New ESRI.ArcGIS.Location.RouteMeasurePointLocation
				'The results of IRouteLocator2::Identify get sent to IEnumRouteIdentifyResult
				pEnumResultReach = pRtLocFlowline.Identify(pEnvelope, "")
				'Step through all the route/measures found
				'For i = 1 To pEnumResultReach.Count
				'For now - only want the FIRST one found
				For i = 1 To 1
					pEnumResultReach.Next(pRouteLocation, pRoute)
					pPointRouteLocation = pRouteLocation
					'UPGRADE_WARNING: Couldn't resolve default property of object pRouteLocation.RouteID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strStartRchCode = pRouteLocation.RouteID
					numStartMeasure = pPointRouteLocation.Measure
				Next i
				
				If numStartMeasure < 0 Then
					MsgBox("This NHDFlowline is not measured.  Cannot begin a navigation here.")
				Else
					frmNavigationOptions.ShowDialog()
					numMaxDistance = Val(frmNavigationOptions.txtMaxDistance.Text)
					'MsgBox numMaxDistance
					
					
					
					'Display the results
					'MsgBox strCheck + "   " + strStartRchCode + "   " + Str(numStartComid) + "   " + Str(numStartMeasure)
					'MsgBox App.Path
					
					'----------------------------------
					'REMOVE PREVIOUS NAVIGATION RESULTS
					'----------------------------------
					RemoveLayer("Navigation Results", pmap, pActiveView, pDoc)
					
					'8/29/2006
					'Comment out the manual deletion of results shapefiles
					'If (fso.FileExists(strDataPath + "TNavigation_events.shp")) Then
					'   Kill strDataPath + "TNavigation_Events.*"
					'End If
					'
					'8/29/2006
					'Add ESRI code to delete previous results completely
					
					pWorkspaceFactory = New ESRI.ArcGIS.DataSourcesFile.ShapefileWorkspaceFactory
					pWS = pWorkspaceFactory.OpenFromFile(strDataPath, 0)
					'9/14 - Add pEnumDataset - Problem when TNavigation Events
					'  is not the only shapefile in the uppermost directory of the
					'  NHDPlus workspace.  Use pEnumDataset in a loop to
					'  remove the previous navigation results.
					pEnumDS = pWS.Datasets(ESRI.ArcGIS.Geodatabase.esriDatasetType.esriDTFeatureClass)
					pDS = pEnumDS.Next
					'9/14/2006 - added this loop (instead of an if statement)
					While Not pDS Is Nothing
						If StrComp(pDS.Name, "TNavigation_events", CompareMethod.Text) = 0 Then
							pDS.Delete()
						End If
						pDS = pEnumDS.Next
					End While
					'UPGRADE_NOTE: Object pWS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					pWS = Nothing
					'UPGRADE_NOTE: Object pWorkspaceFactory may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					pWorkspaceFactory = Nothing
					'UPGRADE_NOTE: Object pDS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					pDS = Nothing
					
					'Comment out the manual deletion of results shapefiles
					'9/14/2006 - reactivate these lines to MAKE SURE
					'   the previous results are gone.
					'UPGRADE_WARNING: Couldn't resolve default property of object fso.FileExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (fso.FileExists(strDataPath & "TNavigation_events.shp")) Then
						Kill(strDataPath & "TNavigation_Events.*")
					End If
					
					
					'9/14/2006
					'Set boolDoPre to true if preprocessing is necessary
					'   Will need to run preprocessing if:
					'      TNavWork does not exist OR
					'      TNavWork.mdb does exist and tblVersion does not exist OR
					'       TNavWork.mdb does exist and tblVersion does exist and
					'         tblVersion.version for first record in tblVersion
					'         does not equal 3
					boolGO = False
					
					'UPGRADE_WARNING: Couldn't resolve default property of object fso.FileExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If Not fso.FileExists(strDataPath & "TNavWork.mdb") Then
						'8/30/2006
						MsgBox("The VAA Navigator needs to perform a preprocessing step to create a working database named TNavWork.mdb.  This must happen one time per NHD workspace.  This may take some time. Please press OK to continue.")
					Else
						'See if tblVersion exists
						oDB = DAODBEngine_definst.Workspaces(0).OpenDatabase(strDataPath & "TNavWork.mdb")
						On Error Resume Next
						td = oDB.TableDefs("tblVersion")
						On Error GoTo ErrHandler
						'UPGRADE_WARNING: Couldn't resolve default property of object boolTableExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						boolTableExists = Err.Number = 0
						oDB.Close()
						
						If (boolTableExists) Then
							'If tblVersion exists ...
							
							adoCn = New ADODB.Connection
							'UPGRADE_WARNING: Couldn't resolve default property of object strConn. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Password=;" & "User ID=Admin;" & "Data Source=" & strDataPath & "TNavWork.mdb" & ";"
							'UPGRADE_WARNING: Couldn't resolve default property of object strConn. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							adoCn.Open(strConn)
							'MsgBox "Connected"
							
							rstResults = New ADODB.Recordset
							rstResults.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
							rstResults.LockType = ADODB.LockTypeEnum.adLockOptimistic
							
							strSQL = "SELECT * FROM tblVersion"
							'MsgBox strSQL
							rstResults.Open(strSQL, adoCn,  ,  , ADODB.CommandTypeEnum.adCmdText)
							
							'This recordset should only return ONE record
							If Not rstResults.EOF Then
								'NOTE:  IF SKEL CHANGES - INCREASE THE VALUE IN
								'VERSION AND CHANGE THIS LINE AS WELL AS THE MATCHING LINE
								'IN THE NHD VAA TOOLBAR
								rstResults.MoveFirst()
								'9/20/2006 - changed from version 1 to version 2
								'1/24/2007 - changed from version 2 to version 3
								'10/12/2007 - changed from version 3 to version 4
								If (rstResults.Fields("Version")).Value = 4 Then
									'GOOD
									boolGO = True
								Else
									MsgBox("The VAA Navigator needs to perform a preprocessing step to create a working database named TNavWork.mdb because the existing TNavWork is not the right format.  This must happen one time per NHD workspace.  This may take some time. Please press OK to continue.")
								End If
							Else
								'tblVersion contains no records
								MsgBox("The VAA Navigator needs to perform a preprocessing step to create a working database named TNavWork.mdb because the existing TNavWork is not the correct version.  This must happen one time per NHD workspace.  This may take some time. Please press OK to continue.")
							End If
							rstResults.Close()
							'UPGRADE_NOTE: Object rstResults may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
							rstResults = Nothing
							If (boolGO = False) Then
								adoCn.Close()
							End If
						Else
							'If tblVersion does not exist ...
							MsgBox("The VAA Navigator needs to perform a preprocessing step to create a working database named TNavWork.mdb because the existing TNavWork does not contain tblVersion.  This must happen one time per NHD workspace.  This may take some time. Please press OK to continue.")
						End If
						
					End If
					'  ------ End of part added 9/14/2006
					
					
					If (boolGO) Then
						'Dim adoCn As ADODB.Connection
						'Dim strConn, strSQL As String
						'Set adoCn = New ADODB.Connection
						'strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + _
						''          "Password=;" + _
						''          "User ID=Admin;" + _
						''          "Data Source=" + strDataPath + "TNavWork.mdb" + ";"
						'adoCn.Open strConn
						'MsgBox "Connected"
						'Remove any existing records in tblNavResults
						strSQL = "DELETE * FROM tblNavResults"
						adoCn.Execute(strSQL,  , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
						adoCn.Close()
					End If
					
					
					
					
					'----------------------------------
					'CALL NAVIGATION DLL
					'----------------------------------
					ChDir(My.Application.Info.DirectoryPath)
					
					'8/30/2006
					'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
					System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
					
					'numReturn = VAANavigate(strCheck, strDataPath, _
					''                        numStartComid, numStartMeasure, _
					''                        0, 0, numMaxDistance, 0, App.Path, "")
					'                        0, 0, numMaxDistance, 0, App.Path, "KEEPLOG")
					'                        numStartMeasure, 0, 0, numMaxDistance, 0, "d:\hsc\nhdplus\navigator")
					
					'Call the navigator as a COM object
					
					'MsgBox "0"
					'MsgBox "1"
					
					objNav = New VAANavigatorCOM.clsVAANavigate
					'MsgBox "2"
					
					'Dim numReturn As Integer
					
					'Set properties for the navigation parameters
					objNav.Navtype = strCheck
					'MsgBox "3"
					objNav.Startcomid = numStartComid
					'MsgBox "4"
					objNav.Startmeas = numStartMeasure
					'MsgBox "5"
					objNav.Maxdistance = numMaxDistance
					'MsgBox "6"
					objNav.Datapath = strDataPath
					'MsgBox "7"
					objNav.Apppath = My.Application.Info.DirectoryPath
					'MsgBox "8"
					
					numReturn = objNav.VAANavigate
					'MsgBox "9"
					'MsgBox Str(numReturn)
					'UPGRADE_NOTE: Object objNav may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					objNav = Nothing
					
					
					'8/30/2006
					'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
					System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
					'----------------------------------
					'ADD RESULTS TO MAP
					'----------------------------------
					'UPGRADE_WARNING: Couldn't resolve default property of object varSs_time. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					varSs_time = VB.Timer()
					RunAddResultsToMap(strLayerName, strDataPath, pmap, pActiveView, pDoc, strCheck, pTLayer)
					
					'Set the variable to mark the end of the navigation
					'UPGRADE_WARNING: Couldn't resolve default property of object varEs_time. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					varEs_time = VB.Timer()
					'Calculate elapsed time for the navigation
					'UPGRADE_WARNING: Couldn't resolve default property of object varSs_time. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object varEs_time. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object varMinRunTime. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					varMinRunTime = Int(System.Math.Round((varEs_time - varSs_time) / 60, 5))
					'UPGRADE_WARNING: Couldn't resolve default property of object varMinRunTime. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object varSs_time. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object varEs_time. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object varSecRunTime. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					varSecRunTime = System.Math.Round((varEs_time - varSs_time) / 60, 5) - varMinRunTime
					'UPGRADE_WARNING: Couldn't resolve default property of object varSecRunTime. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					varSecRunTime = System.Math.Round(varSecRunTime * 60)
					
					'MsgBox "Display completed in " & varMinRunTime & " minutes and " & varSecRunTime & " seconds"
				End If
				
				pGCont.DeleteAllElements()
				pActiveView.Refresh()
				
			End If 'Selected layer is a NHDFlowline layer
			
		End If 'There is a selected layer
		
		Exit Sub
		
ErrHandler: 
		
		MsgBox(Err.Description)
		
	End Sub
	Function IsNHDFlowline(ByVal pFC As ESRI.ArcGIS.Geodatabase.IFeatureClass, ByVal strDataName As String) As Boolean
		Dim boolReturn As Boolean
		
		'Determine if the selected layer is a NHDFlowline layer by
		'verifying the underlying data for the layer contains NHDFlowline and
		'that the layer
		'contains a comid field.
		
		boolReturn = True
		'UPGRADE_WARNING: Couldn't resolve default property of object pFC.FindField. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (pFC.FindField("comid") = -1) Then
			boolReturn = False
		End If
		
		If (InStr(1, strDataName, "NHDFlowline", CompareMethod.Text) <= 0) Then
			boolReturn = False
		End If
		
		
		'MsgBox pFC.FeatureDataset.BrowseName
		
		
		'   If (pFC.FindField("fdate") = -1) Then
		'      boolReturn = False
		'   End If
		'   If (pFC.FindField("resolution") = -1) Then
		'      boolReturn = False
		'   End If
		'   If (pFC.FindField("gnis_id") = -1) Then
		'      boolReturn = False
		'   End If
		'   If (pFC.FindField("gnis_name") = -1) Then
		'      boolReturn = False
		'   End If
		'   If (pFC.FindField("lengthkm") = -1) Then
		'      boolReturn = False
		'   End If
		'   If (pFC.FindField("reachcode") = -1) Then
		'      boolReturn = False
		'   End If
		'   If (pFC.FindField("flowdir") = -1) Then
		'      boolReturn = False
		'   End If
		'   If (pFC.FindField("wbareacomi") = -1) Then
		'      boolReturn = False
		'   End If
		'   If (pFC.FindField("ftype") = -1) Then
		'      boolReturn = False
		'   End If
		'   If (pFC.FindField("fcode") = -1) Then
		'      boolReturn = False
		'   End If
		'   If (pFC.FindField("shape_leng") = -1) Then
		'      boolReturn = False
		'   End If
		'   If (pFC.FindField("enabled") = -1) Then
		'      boolReturn = False
		'   End If
		
		IsNHDFlowline = boolReturn
		
	End Function
	
	Public Sub RunAddResultsToMap(ByRef strLayerName As String, ByRef strDataPath As String, ByRef pmap As ESRI.ArcGIS.Carto.IMap, ByRef pActiveView As ESRI.ArcGIS.Carto.IActiveView, ByRef pDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument, ByVal strNavType As String, ByRef pTLayer As ESRI.ArcGIS.Carto.ILayer)
		
		Dim strNameDrain As String
		Dim i As Short
		
		'MsgBox "into add results"
		'MsgBox strDataPath
		
		strNameDrain = strLayerName
		
		'---------------------------
		'OPEN THE RESULTS MDB
		'---------------------------
		Dim pFact As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory
		Dim pWorkspace As ESRI.ArcGIS.Geodatabase.IWorkspace
		Dim pFeatws As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace
		Dim pTable As ESRI.ArcGIS.Geodatabase.ITable
		pFact = New ESRI.ArcGIS.DataSourcesGDB.AccessWorkspaceFactory
		pWorkspace = pFact.OpenFromFile(strDataPath & "TNavWork.mdb", 0)
		pFeatws = pWorkspace
		pTable = pFeatws.OpenTable("tblNavResults")
		'Set pTable = pFeatws.OpenTable("tblVaaFromTo")
		
		'---------------------
		'ADD TABLE TO TOC
		'---------------------
		'   'Create a new standalone table and add it
		'   'to the collection of the focus map
		'   Dim pStTab As IStandaloneTable
		'   Set pStTab = New StandaloneTable
		'   Set pStTab.Table = pTable
		'   Dim pStTabColl As IStandaloneTableCollection
		'   Set pStTabColl = pmap
		'   pStTabColl.AddStandaloneTable pStTab
		'    'Refresh the TOC
		'   pDoc.UpdateContents
		
		'------------------------
		'IDENTIFY THE ROUTE LAYER
		'------------------------
		Dim pRouteFc As ESRI.ArcGIS.Geodatabase.IFeatureClass
		Dim pLayer As ESRI.ArcGIS.Carto.ILayer
		Dim pFlayer As ESRI.ArcGIS.Carto.IFeatureLayer
		For i = 0 To pmap.LayerCount - 1
			pLayer = pmap.Layer(i)
			If InStr(1, pLayer.Name, strNameDrain, CompareMethod.Text) > 0 Then
				'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				If TypeOf pLayer Is ESRI.ArcGIS.Carto.IFeatureLayer Then
					'01/05/2007
					'Change here - Not just using layer name.
					'Set the layer to the layer OBJECT corresponding to the mouse click
					'The layer object is now passed to this procedure.
					'So shapes get rendered properly if there are multiple nhdflowline workspaces
					'loaded.
					pFlayer = pTLayer
					pRouteFc = pFlayer.FeatureClass
					'MsgBox "Route Found"
					Exit For
				End If
			End If
		Next i
		If pRouteFc Is Nothing Then
			MsgBox("Could not find the route feature class", MsgBoxStyle.Exclamation, "AddResultsToMap")
			Exit Sub
		End If
		
		'-------------------------------------------
		'CREATE A ROUTE LOCATOR FROM THE ROUTE LAYER
		'-------------------------------------------
		Dim pName As ESRI.ArcGIS.esriSystem.IName
		Dim pRMLName As ESRI.ArcGIS.Geodatabase.IRouteLocatorName
		Dim pDS As ESRI.ArcGIS.Geodatabase.IDataset
		pDS = pRouteFc
		pName = pDS.FullName
		pRMLName = New ESRI.ArcGIS.Location.RouteMeasureLocatorName
		With pRMLName
			.RouteFeatureClassName = pName
			.RouteIDFieldName = "comid"
			.RouteIDIsUnique = True
			.RouteMeasureUnit = ESRI.ArcGIS.esriSystem.esriUnits.esriDecimalDegrees
			.RouteWhereClause = ""
		End With
		
		'-----------------------------------------------
		'ESTABLISH THE PROPERTIES OF THE LINE EVENT LAYER
		'-----------------------------------------------
		Dim pRtProp As ESRI.ArcGIS.Geodatabase.IRouteEventProperties2
		Dim pRMLnProp As ESRI.ArcGIS.Location.IRouteMeasureLineProperties
		pRtProp = New ESRI.ArcGIS.Location.RouteMeasureLineProperties
		With pRtProp
			'UPGRADE_WARNING: Couldn't resolve default property of object pRtProp.EventMeasureUnit. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.EventMeasureUnit = ESRI.ArcGIS.esriSystem.esriUnits.esriDecimalDegrees
			'UPGRADE_WARNING: Couldn't resolve default property of object pRtProp.EventRouteIDFieldName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.EventRouteIDFieldName = "comid"
		End With
		pRMLnProp = pRtProp
		pRMLnProp.FromMeasureFieldName = "frommeas"
		pRMLnProp.ToMeasureFieldName = "tomeas"
		
		'--------------------------------------------------------
		'ASSOCIATE THE EVENT TABLE WITH THE LINE EVENT PROPERTIES
		'--------------------------------------------------------
		pDS = pTable
		pName = pDS.FullName
		Dim pRESN As ESRI.ArcGIS.Geodatabase.IRouteEventSourceName
		pRESN = New ESRI.ArcGIS.Location.RouteEventSourceName
		With pRESN
			.EventTableName = pName
			.EventProperties = pRMLnProp
			.RouteLocatorName = pRMLName
		End With
		
		'----------------------
		'CREATE THE EVENT LAYER
		'----------------------
		Dim pEventFC As ESRI.ArcGIS.Geodatabase.IFeatureClass
		pName = pRESN
		pEventFC = pName.Open
		
		
		Dim pOutDSN As ESRI.ArcGIS.Geodatabase.IDatasetName
		Dim pOutWSN As ESRI.ArcGIS.Geodatabase.IWorkspaceName
		Dim pOutFeatDSN As ESRI.ArcGIS.Geodatabase.IFeatureDatasetName
		
		'MsgBox strDataPath
		
		'-------------------------------------------
		'CREATE THE OUTPUT FEATUREDATASETNAME OBJECT
		'  Will be a shapefile named TNavigation_Events.shp, etc in the upper
		'  level directory of the workspace
		'-------------------------------------------
		
		Dim pShpWSFact As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory
		Dim pShpWS As ESRI.ArcGIS.Geodatabase.IWorkspace
		Dim pShpFeatWS As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace
		pShpWSFact = New ESRI.ArcGIS.DataSourcesFile.ShapefileWorkspaceFactory
		pShpWS = pShpWSFact.OpenFromFile(strDataPath, 0)
		pShpFeatWS = pShpWS
		pOutWSN = New ESRI.ArcGIS.Geodatabase.WorkspaceName
		
		pOutWSN.ConnectionProperties = pShpWS.ConnectionProperties
		pOutWSN.WorkspaceFactoryProgID = "esriDataSourcesFile.ShapefileWorkspaceFactory.1"
		
		'-----------------------------------------
		'CREATE THE OUTPUT FEATURECLASSNAME OBJECT
		'-----------------------------------------
		Dim pOutFCN As ESRI.ArcGIS.Geodatabase.IFeatureClassName
		pOutFCN = New ESRI.ArcGIS.Geodatabase.FeatureClassName
		pOutDSN = pOutFCN
		pOutDSN.WorkspaceName = pOutWSN
		'  If (BoolOne) Then
		'     pOutDSN.Name = "TNavigation_Events1"
		'  Else
		pOutDSN.Name = "TNavigation_Events"
		'  End If
		
		'-----------------------------------------
		'CHECK THE EVENT TABLE's FIELDS
		'-----------------------------------------
		Dim pFlds As ESRI.ArcGIS.Geodatabase.IFields
		Dim pOutFlds As ESRI.ArcGIS.Geodatabase.IFields
		Dim pFldChk As ESRI.ArcGIS.Geodatabase.IFieldChecker
		Dim pTempWS As ESRI.ArcGIS.Geodatabase.IWorkspace
		Dim pTempName As ESRI.ArcGIS.esriSystem.IName
		'UPGRADE_WARNING: Couldn't resolve default property of object pEventFC.Fields. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pFlds = pEventFC.Fields
		pFldChk = New ESRI.ArcGIS.Geodatabase.FieldChecker
		pTempName = pOutWSN
		pTempWS = pTempName.Open
		pFldChk.ValidateWorkspace = pTempWS
		pFldChk.Validate(pFlds, Nothing, pOutFlds)
		If Not pFlds.FieldCount = pOutFlds.FieldCount Then
			MsgBox("The number of fields returned by the field checker is less than the input." & vbCrLf & "Cannot create output feature class", MsgBoxStyle.Exclamation, "ConvertEvents")
			Exit Sub
		End If
		
		'-----------------------------------------------------------
		'CONVERT THE EVENT TABLE (the RouteEventSourceMame)TO SHAPES
		'-----------------------------------------------------------
		Dim pQFilt As ESRI.ArcGIS.Geodatabase.IQueryFilter
		pQFilt = New ESRI.ArcGIS.Geodatabase.QueryFilter
		'pQFilt.WhereClause = "selected=1"
		'UPGRADE_NOTE: Object pQFilt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pQFilt = Nothing
		
		
		Dim pEnum As ESRI.ArcGIS.Geodatabase.IEnumInvalidObject
		Dim pConv As ESRI.ArcGIS.Geodatabase.IFeatureDataConverter2
		'Private WithEvents pConv As IFeatureDataConverter
		pConv = New ESRI.ArcGIS.Geodatabase.FeatureDataConverter
		'MsgBox "calling converter"
		'UPGRADE_WARNING: Couldn't resolve default property of object pRESN. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pEnum = pConv.ConvertFeatureClass(pRESN, pQFilt, Nothing, pOutFeatDSN, pOutFCN, Nothing, pOutFlds, "", 1000, 0)
		'MsgBox "done calling converter"
		
		'------------------------------
		'MAKE SURE EVERYTHING CONVERTED
		'------------------------------
		Dim pInvalidInfo As ESRI.ArcGIS.Geodatabase.IInvalidObjectInfo
		pEnum.Reset()
		pInvalidInfo = pEnum.Next
		While Not pInvalidInfo Is Nothing
			MsgBox(pInvalidInfo.InvalidObjectID & ": " & pInvalidInfo.ErrorDescription)
			pInvalidInfo = pEnum.Next
		End While
		
		'--------------------------------------------------
		'OPEN TNAVIGATION_EVENTS.SHP IN ORDER TO DISPLAY IT
		'--------------------------------------------------
		
		pFlayer = New ESRI.ArcGIS.Carto.FeatureLayer
		pFlayer.FeatureClass = pShpFeatWS.OpenFeatureClass("TNavigation_Events.shp")
		
		'UPGRADE_WARNING: Couldn't resolve default property of object pFlayer.Name. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pFlayer.Name = "Navigation Results"
		
		'----------------------------
		'MAKE THE LAYER RED AND THICK
		'----------------------------
		Dim pSRend As ESRI.ArcGIS.Carto.ISimpleRenderer
		Dim pLSymbol As ESRI.ArcGIS.Display.ISimpleLineSymbol
		Dim pGeoFL As ESRI.ArcGIS.Carto.IGeoFeatureLayer
		'Create a color
		Dim pColor As ESRI.ArcGIS.Display.IRgbColor
		pColor = New ESRI.ArcGIS.Display.RgbColor
		'UPGRADE_WARNING: Couldn't resolve default property of object pColor.RGB. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
		'Create a renderer
		pSRend = New ESRI.ArcGIS.Carto.SimpleRenderer
		'Create a line symbol object
		pLSymbol = New ESRI.ArcGIS.Display.SimpleLineSymbol
		With pLSymbol
			'UPGRADE_WARNING: Couldn't resolve default property of object pLSymbol.Width. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.Width = 3.4
			'UPGRADE_WARNING: Couldn't resolve default property of object pLSymbol.Color. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object pColor. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
		
		AddLayer(pFlayer, pmap, pActiveView, pDoc)
		
		'Set pMap = pDoc.FocusMap
		'Set pActiveView = pMap
		'pMap.AddLayer pFLayer
		pActiveView.Refresh()
		pDoc.UpdateContents()
		
		'   Set pMxDoc = Nothing
		'   Set pMxApp = Nothing
		'   Set pMap = Nothing
		'   Set pActiveView = Nothing
		'UPGRADE_NOTE: Object pFact may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFact = Nothing
		'UPGRADE_NOTE: Object pWorkspace may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pWorkspace = Nothing
		'UPGRADE_NOTE: Object pFeatws may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFeatws = Nothing
		'UPGRADE_NOTE: Object pTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pTable = Nothing
		'UPGRADE_NOTE: Object pRouteFc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRouteFc = Nothing
		'UPGRADE_NOTE: Object pLayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pLayer = Nothing
		'UPGRADE_NOTE: Object pFlayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFlayer = Nothing
		'UPGRADE_NOTE: Object pName may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pName = Nothing
		'UPGRADE_NOTE: Object pRMLName may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRMLName = Nothing
		'UPGRADE_NOTE: Object pDS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pDS = Nothing
		'UPGRADE_NOTE: Object pRtProp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRtProp = Nothing
		'UPGRADE_NOTE: Object pRMLnProp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRMLnProp = Nothing
		'UPGRADE_NOTE: Object pRESN may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRESN = Nothing
		'UPGRADE_NOTE: Object pEventFC may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pEventFC = Nothing
		'UPGRADE_NOTE: Object pSRend may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pSRend = Nothing
		'UPGRADE_NOTE: Object pLSymbol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pLSymbol = Nothing
		'UPGRADE_NOTE: Object pGeoFL may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pGeoFL = Nothing
		'UPGRADE_NOTE: Object pColor may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pColor = Nothing
		'UPGRADE_NOTE: Object pShpWSFact may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pShpWSFact = Nothing
		'UPGRADE_NOTE: Object pShpWS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pShpWS = Nothing
		'UPGRADE_NOTE: Object pShpFeatWS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pShpFeatWS = Nothing
		'UPGRADE_NOTE: Object pOutFCN may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pOutFCN = Nothing
		'UPGRADE_NOTE: Object pOutFeatDSN may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pOutFeatDSN = Nothing
		'UPGRADE_NOTE: Object pQFilt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pQFilt = Nothing
		'UPGRADE_NOTE: Object pEnum may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pEnum = Nothing
		'UPGRADE_NOTE: Object pConv may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pConv = Nothing
		'UPGRADE_NOTE: Object pOutFCN may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pOutFCN = Nothing
		'UPGRADE_NOTE: Object pFlds may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFlds = Nothing
		'UPGRADE_NOTE: Object pOutFlds may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pOutFlds = Nothing
		'UPGRADE_NOTE: Object pFldChk may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pFldChk = Nothing
		'UPGRADE_NOTE: Object pTempWS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pTempWS = Nothing
		'UPGRADE_NOTE: Object pTempName may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pTempName = Nothing
		
		
		'MsgBox "leaving add layer proc"
		
	End Sub
	
	Private Sub RemoveLayer(ByRef strName As String, ByRef pmap As ESRI.ArcGIS.Carto.IMap, ByRef pActiveView As ESRI.ArcGIS.Carto.IActiveView, ByRef pDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument)
		
		Dim pLayer As ESRI.ArcGIS.Carto.ILayer
		Dim pFlayer As ESRI.ArcGIS.Carto.IFeatureLayer
		Dim i As Short
		
		
		For i = 0 To pmap.LayerCount - 1
			pLayer = pmap.Layer(i)
			If InStr(1, pLayer.Name, strName, CompareMethod.Text) > 0 Then
				pFlayer = pLayer
				'UPGRADE_WARNING: Couldn't resolve default property of object pFlayer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pmap.DeleteLayer(pFlayer)
				pActiveView.Refresh()
				pDoc.UpdateContents()
				Exit For
			End If
		Next i
		
	End Sub
	
	Private Sub AddLayer(ByRef pFlayer As ESRI.ArcGIS.Carto.IFeatureLayer, ByRef pmap As ESRI.ArcGIS.Carto.IMap, ByRef pActiveView As ESRI.ArcGIS.Carto.IActiveView, ByRef pDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument)
		
		pFlayer.Selectable = True
		'UPGRADE_WARNING: Couldn't resolve default property of object pFlayer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pmap.AddLayer(pFlayer)
		pActiveView.Refresh()
		pDoc.UpdateContents()
		
	End Sub
End Module