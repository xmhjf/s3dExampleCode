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

 
