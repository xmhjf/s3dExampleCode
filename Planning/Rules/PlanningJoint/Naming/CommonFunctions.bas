Attribute VB_Name = "CommonFunctions"

' ***************************************************************************
'
' Property
'   GetAssemblyTreeRoot
'
' Abstract
'   Returns a reference to the assembly tree root object
' Koushik  26 Apr 07  TR-CP·106203    Compilation of planning name rule doesn't fail due to missing "Option Explicit"
' ***************************************************************************
Option Explicit
Private Const Module = "CommonFunctions: "
Private Const MODELDATABASE = "Model"
Private Const ERRORPROGID As String = "IMSErrorLog.ServerErrors"

Public Const E_ACCESSDENIED As Long = 70
Public Function GetAssemblyTreeRoot() As GSCADAsmHlpers.IJAssembly
    Const METHOD = "GetAssemblyTreeRoot"
    On Error GoTo ErrorHandler
    
    Dim jContext As IJContext
    Dim oModelResourceMgr As IUnknown
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim oConnectMiddle As IJDAccessMiddle
    Dim strModelDBID As String
    Dim oErrors         As IJEditErrors
    Dim oError          As IJEditError
    
    Set oErrors = CreateObject(ERRORPROGID)

    '  Get active connection using the middle context

    Set jContext = GetJContext()
    
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType(MODELDATABASE)
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)
    If (oModelResourceMgr Is Nothing) Then GoTo ErrorHandler
    
    ' Get top-level block
    Dim oPlnIntHelper As GSCADPlnIntHelper.IJDPlnIntHelper
    Set oPlnIntHelper = New GSCADPlnIntHelper.CPlnIntHelper
    
    Dim oBlock As GSCADBlock.IJBlock
    Set oBlock = oPlnIntHelper.GetTopLevelBlock(oModelResourceMgr)
   
    ' Return block
    Set GetAssemblyTreeRoot = oBlock
    
    Set oErrors = Nothing
    Set oError = Nothing
    
    Exit Function
ErrorHandler:
    oError = oErrors.AddFromErr(Err, Module & " - " & METHOD)
    oError.Raise
    Set oErrors = Nothing
    Set oError = Nothing
End Function

Private Function GetStringFromId(ByVal TableName As String, ByVal ValueID As Long) As String

Const METHOD = "GetStringFromId"
On Error GoTo ErrorHandler

    Dim oIJDCodeListMetaData      As IJDCodeListMetaData
    Dim oErrors         As IJEditErrors
    Dim oError          As IJEditError
    
    Set oIJDCodeListMetaData = GetPOM("Model")
    GetStringFromId = oIJDCodeListMetaData.DisplayStringByValueID(TableName, ValueID)
    
Exit Function
ErrorHandler:
    oError = oErrors.AddFromErr(Err, Module & " - " & METHOD)
    oError.Raise
    Set oErrors = Nothing
    Set oError = Nothing
End Function

Public Function GetDisplayString(oObj As Object, strAttName As String, lVal As Long) As String
Const METHOD As String = "FillPartCoatingAreaInfo"
On Error GoTo ErrorHandler
   
    Dim oAttributeMetadata          As IJDAttributeMetaData
    Dim oAttrHelper                 As IJDAttributes
    Dim oAttribute                  As IJDAttribute
    Dim oInterfaceInfo              As IJDInterfaceInfo
    Dim oCodelist                   As IJDCodeListMetaData
    Dim oUserAttrColl               As IJDAttributesCol
    Dim varInterfaceID              As Variant
    Dim oErrors                     As IJEditErrors
    Dim oError                      As IJEditError
    
    Set oCodelist = GetPOM("Model")
    Set oAttributeMetadata = oObj
    Set oAttrHelper = oObj
    
    Dim iID As Variant
    Dim ii As Long
    
    If oAttributeMetadata Is Nothing Then Exit Function
    If oAttrHelper Is Nothing Then Exit Function
    
    For Each varInterfaceID In oAttrHelper
    
        Set oInterfaceInfo = oAttributeMetadata.InterfaceInfo(varInterfaceID)
        iID = oInterfaceInfo.Type
        Set oUserAttrColl = oAttrHelper.CollectionOfAttributes(iID)

        For ii = 1 To oUserAttrColl.Count
            Set oAttribute = oUserAttrColl.Item(ii)
            If Not oAttribute Is Nothing Then
                If oAttribute.AttributeInfo.Name = strAttName Then
                    GetDisplayString = GetStringFromId(oAttribute.AttributeInfo.CodeListTableName, lVal)
                    GoTo CleanUp
                End If
            End If
         Next ii
    Next
                           
CleanUp:

    Set oAttributeMetadata = Nothing
    Set oAttrHelper = Nothing
    Set oAttribute = Nothing
    Set oInterfaceInfo = Nothing
    Set oCodelist = Nothing
        
Exit Function
ErrorHandler:
oError = oErrors.AddFromErr(Err, Module & " - " & METHOD)
    oError.Raise
    Set oErrors = Nothing
    Set oError = Nothing
End Function

Private Function GetPOM(strDbType As String) As IJDPOM
Const METHOD = "GetPOM"
On Error GoTo ErrHandler
    
    Dim oContext            As IJContext
    Dim oAccessMiddle       As IJDAccessMiddle
    Dim oErrors                     As IJEditErrors
    Dim oError                      As IJEditError
    
    Set oContext = GetJContext()
    Set oAccessMiddle = oContext.GetService("ConnectMiddle")
    Set GetPOM = oAccessMiddle.GetResourceManagerFromType(strDbType)
    
    Set oContext = Nothing
    Set oAccessMiddle = Nothing
    
Exit Function
ErrHandler:
   oError = oErrors.AddFromErr(Err, Module & " - " & METHOD)
    oError.Raise
    Set oErrors = Nothing
    Set oError = Nothing
End Function
 
