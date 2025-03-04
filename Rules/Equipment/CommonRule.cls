VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommonRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
'   Copyright (C) 2000, Intergraph Corporation.  All rights reserved.
'
'   History:
'       Sadish Sadasivam Oct/07/2003
'           TR CP49909: fixed the usage of LocationID
'***************************************************************************

Option Explicit

Implements IJNameRule
                                                        
Const vbInvalidArg = &H80070057

Private Const Module = "CommonNameRule: "
Private Const DEFAULTBASE = "Defaultname"
Private Const MODELDATABASE = "Model"
Private Const strCountFormat = "0000"       ' define fixed-width number field for name
Private Const PARTOCCINTERFACE = "IJPartOcc"
Private Const SMARTOCCINTERFACE = "IJSmartOccurrence"

Private Const PARTROLE = "part"
Private Const SMARTITEMROLE = "toSI_ORIG"
Private IID_IJPartOcc As Variant

Dim m_oErrors As IJEditErrors

Private Const E_FAIL = -2147467259



Private Sub Class_Initialize()
    Set m_oErrors = New IMSErrorLog.JServerErrors
End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

'********************************************************************
' Description:
' Creates a name for the object passed in. The name is based on the parents
' name and object name.It is something like this: "Base Name" + "Object Name" + Index.
' "Base Name" is derived from the Naming Parents and the Index is unique for the "Base Name"
' It is assumed that all Naming Parents and the Object implement IJPRodModelItem.
' The Naming Parents are added in AddNamingParents() of the same interface.
' Both these methods are called from naming rule semantic.
'********************************************************************
Private Sub IJNameRule_ComputeName(ByVal oObject As Object, ByVal elements As IJElements, ByVal oActiveEntity As Object)

    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo label
    
    Dim jContext As IJContext
    Dim oModelResourceMgr As IUnknown
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim oConnectMiddle As IJDAccessMiddle
    Dim strModelDBID As String
    Dim strLocation As String
    
    'Get the middle context
    Set jContext = GetJContext()
    
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType(MODELDATABASE)
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)

    Dim oNameCounter As IJNameCounter
    Dim oChildNamedItem As IJNamedItem
    Dim oParent As Object
    Dim oParentNamedItem As IJNamedItem

    Dim strChildName As String
    Dim strParentName As String
    Dim strNamedParentsString As String
    Dim nCount As Long
    Dim strPartname As String

    If oObject Is Nothing Then
        Err.Raise vbInvalidArg, Module, METHOD
    End If

    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService

    Set oChildNamedItem = oObject
    
    ' Get name of part
    strPartname = GetPartName(oObject)
    If strPartname = "" Then
        ' if we cannot get the part name, then use the type string
        strPartname = oChildNamedItem.TypeString
    End If

    ' Use part Name for basename, remove blanks:
    strPartname = Join(Split(strPartname), "")

    If elements.Count > 0 Then
        For Each oParent In elements
            On Error Resume Next
            Set oParentNamedItem = oParent
            On Error GoTo label

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
        If (strChildName + strPartname) <> strNamedParentsString Then
            oActiveEntity.NamingParentsString = strChildName + strPartname
        
            strLocation = vbNullString
            nCount = oNameCounter.GetCountEx(oModelResourceMgr, strChildName + strPartname, strLocation)
            
            If strLocation <> vbNullString Then
                strChildName = strChildName & "-" & strPartname & "-" & strLocation & "-" & Format(nCount, strCountFormat)
            Else
                strChildName = strChildName & "-" & strPartname & "-" & Format(nCount, strCountFormat)
            End If
            oChildNamedItem.Name = strChildName
        End If
    Else
        ' no parents - use just the strPartName as the base
        strNamedParentsString = oActiveEntity.NamingParentsString

        'Check if New parent name string constructed and old parent name string existing are same
        If strPartname <> strNamedParentsString Then
            oActiveEntity.NamingParentsString = strPartname

            strLocation = vbNullString
            nCount = oNameCounter.GetCountEx(oModelResourceMgr, strPartname, strLocation)

            If strLocation <> vbNullString Then
                strChildName = strPartname & "-" & strLocation & "-" & Format(nCount, strCountFormat)
            Else
                strChildName = strPartname & "-" & Format(nCount, strCountFormat)
            End If
            
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

label:
    ' log the error in middle tier and propagate the error code to the caller
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
    On Error GoTo label

    Dim oSysParent As IJDesignParent
    Dim oSysChild As IJDesignChild

    On Error Resume Next
    Set oSysChild = oEntity
    
    If Not oSysChild Is Nothing Then
        Set oSysParent = oSysChild.GetParent
    End If
    
    On Error GoTo label

    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection

    If Not (oSysParent Is Nothing) Then
        Call IJNameRule_GetNamingParents.Add(oSysParent)
    End If


    Set oSysParent = Nothing
    Set oSysChild = Nothing
Exit Function

label:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IJNameRule_GetNamingParents", Err.Description
    Err.Raise E_FAIL
End Function



Private Function GetPartName(ByVal oEntity As Object) As String

    Const METHOD = "GetPartName"
    On Error GoTo label
    
    Dim oPart As IJDPart
    
    Set oPart = GetPart(oEntity)
           
    If Not oPart Is Nothing Then
        GetPartName = oPart.PartNumber
    End If

    Exit Function
label:
    ' log error to middle tier
    GetPartName = ""
    m_oErrors.Add Err.Number, "CommonNameRule:GetPartName", Err.Description
    Err.Raise E_FAIL
End Function

Private Function GetPart(ByVal oEntity As Object) As IJDPart

    Const METHOD = "GetPart"
    On Error GoTo label
    
    Dim oRelationHelper As IMSRelation.DRelationHelper
    Dim oCollection As IMSRelation.DCollectionHelper 'IJDrelatiopnshipCol
    Dim oPart As IJDPart
    
    Set oRelationHelper = oEntity
    
    If TypeOf oEntity Is IJPartOcc Then
        Set oCollection = oRelationHelper.CollectionRelations(PARTOCCINTERFACE, PARTROLE)
    ElseIf TypeOf oEntity Is IJSmartOccurrence Then
        Set oCollection = oRelationHelper.CollectionRelations(SMARTOCCINTERFACE, SMARTITEMROLE)
    End If
    
    If Not oCollection Is Nothing Then
        If (oCollection.Count <> 0) Then
            Set oPart = oCollection.Item(1)
            If Not oPart Is Nothing Then
                Set GetPart = oPart
            End If
        End If
    End If

    Exit Function
label:
    ' log error to middle tier
    Set GetPart = Nothing
    m_oErrors.Add Err.Number, "CommonNameRule:GetPart", Err.Description
    Err.Raise E_FAIL
End Function

