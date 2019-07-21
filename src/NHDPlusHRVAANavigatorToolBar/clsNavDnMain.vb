Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geodatabase
Imports System.Windows.Forms

<ComClass(clsNavDnMain.ClassId, clsNavDnMain.InterfaceId, clsNavDnMain.EventsId),
 ProgId("NHDPlusHRVAANavToolbar.clsNavDnMain")>
Public NotInheritable Class clsNavDnMain
    Inherits ESRI.ArcGIS.ADF.BaseClasses.BaseTool

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "5d1bfd46-a3c1-4e54-81e6-ab0a72702506"
    Public Const InterfaceId As String = "4532e21d-e668-4ae1-b3ca-c49f51855b9a"
    Public Const EventsId As String = "a4bc282e-f5e8-4044-a573-be8d1257f60b"
#End Region

#Region "COM Registration Function(s)"
    <ComRegisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub RegisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryRegistration(registerType)

        'Add any COM registration code after the ArcGISCategoryRegistration() call

    End Sub

    <ComUnregisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub UnregisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryUnregistration(registerType)

        'Add any COM unregistration code after the ArcGISCategoryUnregistration() call

    End Sub

#Region "ArcGIS Component Category Registrar generated code"
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Register(regKey)

    End Sub
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Unregister(regKey)

    End Sub

#End Region
#End Region

    Private mApplication As IApplication
    Private mDoc As IMxDocument
    Private mMap As IMap
    Private mActiveView As IActiveView
    Private strDllLocation As String

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
        MyBase.m_category = ""  'localizable text 
        MyBase.m_caption = "NHDPlusHR VAA Navigator Down Mainstem"   'localizable text 
        MyBase.m_message = "Navigate down stream following the main path from the NHDFlowline feature selected via 'Point and Click'(Zoom as needed)"   'localizable text 
        MyBase.m_toolTip = "NHDPlusHR VAA Navigator Down Mainstem" 'localizable text 
        MyBase.m_name = ""  'unique id, non-localizable (e.g. "MyCategory_ArcMapTool")

        Try
            Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
            MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
            MyBase.m_cursor = New System.Windows.Forms.Cursor(Me.GetType(), Me.GetType().Name + ".cur")

            strDllLocation = System.Reflection.Assembly.GetAssembly(Me.GetType()).Location.ToUpper.Replace("NHDPLUSHRVAANAVTOOLBAR.DLL", "")

            MyBase.m_helpFile = strDllLocation & "\nhdv2vaanavhelp.chm"
            MyBase.m_helpID = 3

        Catch ex As Exception
            System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap")
        End Try
    End Sub

    Public Overrides Sub OnCreate(ByVal hook As Object)
        If Not hook Is Nothing Then
            mApplication = CType(hook, IApplication)

            'Disable if it is not ArcMap
            If TypeOf hook Is IMxApplication Then
                MyBase.m_enabled = True
            Else
                MyBase.m_enabled = False
            End If
        End If
    End Sub

    Public Overrides Sub OnClick()
        'TODO: Add clsNavDnMain.OnClick implementation
    End Sub

    Public Overrides Sub OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)

        Dim mPoint As IPoint
        Dim strReturn As String

        Try
            'Make sure it is the left mouse button
            If Button = 1 Then

                mDoc = mApplication.Document
                mMap = mDoc.FocusMap
                mActiveView = mMap

                ' mApplication is set in OnCreate.
                ' Convert x and y to map units. 
                ' Create a search point
                mPoint = mActiveView.ScreenDisplay.DisplayTransformation.ToMapPoint(X, Y)
                mApplication.StatusBar.Message(0) = "Mouse clicked at coordinates: " + mPoint.X.ToString & "," & mPoint.Y.ToString
                strReturn = CoordinateNavigation(mPoint, "DNMAIN", strDllLocation, mDoc, mMap, mActiveView, mApplication)
                If strReturn <> "" Then
                    MsgBox(strReturn)
                End If

                'This message gets cleared too quickly by ArcMap, so it is not very
                'useful.
                'mApplication.StatusBar.Message(0) = "Number of selected features is: " + intSelectedCount.ToString

            End If

        Catch ex As Exception
            MsgBox("OnMouseDown Exception: " & ex.Message.ToString & vbCrLf & _
                                   ex.StackTrace.ToString & vbCrLf & _
                                   ex.Source.ToString)
        Finally
            mPoint = Nothing
            strReturn = Nothing
            mDoc = Nothing
            mMap = Nothing
            mActiveView = Nothing

        End Try

    End Sub

    Public Overrides Sub OnMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        'TODO: Add clsNavDnMain.OnMouseMove implementation
    End Sub

    Public Overrides Sub OnMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        'TODO: Add clsNavDnMain.OnMouseUp implementation
    End Sub
End Class

