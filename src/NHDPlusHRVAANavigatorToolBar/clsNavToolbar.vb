Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports System.Runtime.InteropServices

<ComClass(clsNavToolbar.ClassId, clsNavToolbar.InterfaceId, clsNavToolbar.EventsId),
 ProgId("NHDPlusHRNavigatorToolbar.clsNavToolbar")>
Public NotInheritable Class clsNavToolbar
    Inherits BaseToolbar

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
    ''' <summary>
    ''' Required method for ArcGIS Component Category registration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommandBars.Register(regKey)

    End Sub
    ''' <summary>
    ''' Required method for ArcGIS Component Category unregistration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommandBars.Unregister(regKey)

    End Sub

#End Region
#End Region

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "8864ae1c-e16e-4bfc-908c-36d699233d29"
    Public Const InterfaceId As String = "aeb2d1d7-4eca-4d6d-a21c-0166ebabc3a5"
    Public Const EventsId As String = "2eb7e1f7-53ac-40ae-9a27-d8a3067988df"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()

        'BeginGroup() 'Separator
        AddItem("NHDPlusHRVAANavToolbar.clsNavUpMain")
        AddItem("NHDPlusHRVAANavToolbar.clsNavUptrib")
        AddItem("NHDPlusHRVAANavToolbar.clsNavDnMain")
        AddItem("NHDPlusHRVAANavToolbar.clsNavDnDiv")
        AddItem("NHDPlusHRVAANavToolbar.clsDropWorkingDB")
    End Sub

    Public Overrides ReadOnly Property Caption() As String
        Get
            Return "NHDPlusHR VAA Nav Toolbar"
        End Get
    End Property

    Public Overrides ReadOnly Property Name() As String
        Get
            Return "clsNavToolbar"
        End Get
    End Property
End Class
