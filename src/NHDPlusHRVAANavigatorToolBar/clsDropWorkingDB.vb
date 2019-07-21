Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto

<ComClass(clsDropWorkingDB.ClassId, clsDropWorkingDB.InterfaceId, clsDropWorkingDB.EventsId),
 ProgId("NHDPlusHRVAANavToolbar.clsDropWorkingDB")>
Public NotInheritable Class clsDropWorkingDB
    Inherits ESRI.ArcGIS.ADF.BaseClasses.BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "6cfc334d-ed29-4577-b9a1-0a14ee881c50"
    Public Const InterfaceId As String = "751b7f05-5497-41c6-859d-f59f2b066e89"
    Public Const EventsId As String = "4582d822-5320-4381-9d03-24395a51aa7c"
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

        ' Define values for the public properties
        MyBase.m_category = ""  'localizable text 
        MyBase.m_caption = "NHDPlusHR VAA Navigator Delete Working Database"   'localizable text 
        MyBase.m_message = "Delete the SQL Server working database"   'localizable text 
        MyBase.m_toolTip = "NHDPlusHR VAA Navigator Delete Working Database" 'localizable text 
        MyBase.m_name = ""  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

        strDllLocation = System.Reflection.Assembly.GetAssembly(Me.GetType()).Location.ToUpper.Replace("NHDPlusHRVAANAVTOOLBAR.DLL", "")

        Try
            'Change bitmap name if necessary
            Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
            MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
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

        ' TODO:  Add other initialization code
    End Sub

    Public Overrides Sub OnClick()
        Dim strReturn As String

        Try
            mDoc = mApplication.Document
            mMap = mDoc.FocusMap
            mActiveView = mMap
            strReturn = DropWorkingDB(Nothing, "DROPDB", strDllLocation, mDoc, mMap, mActiveView, mApplication)
            If strReturn <> "" Then
                MsgBox(strReturn)
            End If

        Catch ex As Exception
            MsgBox("OnClick Exception: " & ex.Message.ToString & vbCrLf &
                                   ex.StackTrace.ToString & vbCrLf &
                                   ex.Source.ToString)
        Finally
            strReturn = Nothing
            mDoc = Nothing
            mMap = Nothing
            mActiveView = Nothing

        End Try

    End Sub

End Class



