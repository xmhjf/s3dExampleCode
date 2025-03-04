VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemberTypeNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJNameRule
Private Const MODULE = "MemberNameRules::"
Private Const DEFAULTBASE = "Defaultname"
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

    Dim oMembPart As SPSMembers.ISPSMemberPartPrismatic
    Dim oMembSys As SPSMembers.ISPSMemberSystem
    
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
    Dim MembType As Long
    Dim strlocation As String
    
    Set oChildNamedItem = oObject
    
    strParentName = ""
    
    If TypeOf oObject Is ISPSMemberSystem Then
        Set oMembSys = oObject
        MembType = oMembSys.MemberType.Type
        strPartname = GetMemberTypeName(MembType)
    ElseIf TypeOf oObject Is ISPSMemberPartPrismatic Then
        Set oMembPart = oObject
        MembType = oMembPart.MemberType.Type
        strPartname = GetMemberTypeName(MembType)
    End If
        
    strNamedParentsString = oActiveEntity.NamingParentsString

    'Check if New parent name string constructed and old parent name string existing are same
    If strPartname <> strNamedParentsString Then
        oActiveEntity.NamingParentsString = strPartname
        strlocation = vbNullString
        
        nCount = oNameCounter.GetCountEx(oModelResourceMgr, strPartname, strlocation)

        If Not (strlocation = vbNullString) Then
            strChildName = strPartname + "-" + strlocation + "-" + Format(nCount, strCountFormat)
        Else
            strChildName = strPartname + "-" + Format(nCount, strCountFormat)
        End If
        
        oChildNamedItem.Name = strChildName
    End If
        
           
    Set oMembPart = Nothing
    Set oMembSys = Nothing
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Description
' All the Naming Parents that need to participate in an objects naming are added here to the
' IJElements collection. The parents added here are used in computing the name of the object in
' ComputeName() of the same interface. Both these methods are called from naming rule semantic.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements

    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo ErrHandler
' Commented the code to get the NamingParent Name to be used to prefix for the name being generated
'    Dim oSysParent As IJSystem
'    Dim oSysChild As IJSystemChild

'    On Error Resume Next
'    Set oSysChild = oEntity
'    Set oSysParent = oSysChild.GetParent

'    On Error GoTo label

'    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection

'    If Not (oSysParent Is Nothing) Then
'        Call IJNameRule_GetNamingParents.Add(oSysParent)
'    End If
    Set IJNameRule_GetNamingParents = Nothing

'    Set oSysParent = Nothing
'    Set oSysChild = Nothing
Exit Function
ErrHandler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IJNameRule_GetNamingParents", Err.Description
    Err.Raise E_FAIL
End Function


