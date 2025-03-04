VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConnectionNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJNameRule
Private Const MODULE = "MemberNameRules::"
Private Const MODELDATABASE = "Model"
Private Const strCountFormat = "0000"
Private m_oErrors As IJEditErrors
Private Const E_FAIL = -2147467259


Private Sub Class_Initialize()
    
    Set m_oErrors = New IMSErrorLog.JServerErrors

End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

Private Sub IJNameRule_ComputeName(ByVal oObject As Object, ByVal elements As IJElements, ByVal oActiveEntity As Object)
    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo ErrHandler
    
    Dim jContext As IJContext
    Dim oModelResourceMgr As IUnknown
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim oConnectMiddle As IJDAccessMiddle
    Dim strModelDBID As String
    
    'Get the middle context
    Set jContext = GetJContext()
      
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType(MODELDATABASE)
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)

    Dim oNameCounter As IJNameCounter
    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
    Dim oParent As Object
    Dim oParentNamedItem As IJNamedItem
    Dim oChildNamedItem As IJNamedItem
  
    Dim strParentName As String
    Dim strChildName As String
    Dim strNamedParentsString As String
    Dim nCount As Long
    Dim strPartname As String
       
    Set oChildNamedItem = oObject

   If elements.Count > 0 Then
        For Each oParent In elements
            On Error Resume Next
            Set oParentNamedItem = oParent
            On Error GoTo ErrHandler:

            strParentName = oParentNamedItem.Name

            If (Len(strChildName) = 0) Then
                strChildName = strParentName
            Else
                strChildName = strChildName + "-" + strParentName
            End If

            Set oParentNamedItem = Nothing
        Next oParent

        strNamedParentsString = oActiveEntity.NamingParentsString

        'Check if New parent name string constructed and old parent name string existing are same
        If strChildName <> strNamedParentsString Then
            oActiveEntity.NamingParentsString = strChildName

            nCount = oNameCounter.GetCount(oModelResourceMgr, strPartname)

            strChildName = strChildName + "-" + strPartname + "-" + Format(nCount, strCountFormat)
            oChildNamedItem.Name = strChildName
        End If
    Else
        ' no parents - use just the strPartName as the base
        strNamedParentsString = oActiveEntity.NamingParentsString

        'Check if New parent name string constructed and old parent name string existing are same
        If strPartname <> strNamedParentsString Then
            oActiveEntity.NamingParentsString = strPartname

            nCount = oNameCounter.GetCount(oModelResourceMgr, strPartname)

            strChildName = strPartname + "-" + Format(nCount, strCountFormat)
            oChildNamedItem.Name = strChildName
        End If
    End If
        
    Set jContext = Nothing
    Set oModelResourceMgr = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
    
    Set oNameCounter = Nothing
    Set oChildNamedItem = Nothing
    Set oParent = Nothing
    Set oParentNamedItem = Nothing
    

Exit Sub

ErrHandler:
    m_oErrors.Add Err.Number, "IJNameRule_ComputeName", Err.Description
    Err.Raise E_FAIL
    
End Sub
 Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements

    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo ErrHandler

    Dim oSysParent As IJSystem
    Dim oSysChild As IJSystemChild

    On Error Resume Next
    Set oSysChild = oEntity
    Set oSysParent = oSysChild.GetParent

    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection

    If Not (oSysParent Is Nothing) Then
        Call IJNameRule_GetNamingParents.Add(oSysParent)
    End If


    Set oSysParent = Nothing
    Set oSysChild = Nothing
Exit Function
ErrHandler:
 
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IJNameRule_GetNamingParents", Err.Description
    Err.Raise E_FAIL
End Function



