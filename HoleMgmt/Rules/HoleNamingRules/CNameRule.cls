VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2001, Intergraph Corporation.  All rights reserved.
'
'  Project: GSCADHoleNamingRules
'
'  Abstract: The file contains a implementation of a AsmSystemNameRule for Hole Management
'
'  Author: Jang Choi
'
'******************************************************************

Option Explicit

Implements IJNameRule
                                                        
Const vbInvalidArg = &H80070057

Private Const Module = "CNameRule: "
Private Const DEFAULTBASE = "Defaultname"
Private Const MODELDATABASE = "Model"
Private Const strCountFormat = "0000"       ' define fixed-width number field for name
Private Const PARTROLE = "part"
Private IID_IJPartOcc As Variant

Dim m_oErrors As IJEditErrors

Private Const E_FAIL = -2147467259

Private Sub Class_Initialize()
    '{1146CF94-6B33-11D1-A300-080036409103}
    IID_IJPartOcc = InitGuid(&H11, &H46, &HCF, &H94, &H6B, &H33, &H11, &HD1, &HA3, &H0, &H8, &H0, &H36, &H40, &H91, &H3)

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
' It is assumed that all Naming Parents and the Object implement IJNamedItem.
' The Naming Parents are added in AddNamingParents() of the same interface.
' Both these methods are called from naming rule semantic.
'********************************************************************
Private Sub IJNameRule_ComputeName(ByVal oObject As Object, ByVal elements As IJElements, ByVal oActiveEntity As Object)

    Const Method = "IJNameRule_ComputeName"
    On Error GoTo label
    
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
    Dim oChildProdModel As IJNamedItem
''    Dim oParent As Object
''    Dim oParentProdModel As IJNamedItem

    Dim strChildName As String
''    Dim strParentName As String
    Dim strNamedParentsString As String
    Dim nCount As Long
    Dim strPartname As String
    Dim strLocationID As String

    If oObject Is Nothing Then
        Err.Raise vbInvalidArg, Module, Method
    End If

    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService

    Set oChildProdModel = oObject
    
    ' Get name of part
    strPartname = GetPartName(oObject)
    If strPartname = "" Then
        ' if we cannot get the part name, then use the type string
        strPartname = oChildProdModel.TypeString
    End If

    ' Use part Name for basename, remove blanks:
    strPartname = Join(Split(strPartname), "")

    If elements.Count > 0 Then
''        For Each oParent In elements
''            On Error Resume Next
''            Set oParentProdModel = oParent
''            On Error GoTo label
''
''            strParentName = oParentProdModel.Name
''
''            If (Len(strChildName) = 0) Then
''                strChildName = strParentName
''            Else
''                strChildName = strChildName + "-" + strParentName
''            End If
''
''            Set oParentProdModel = Nothing
''        Next oParent

        strNamedParentsString = oActiveEntity.NamingParentsString

        'Check if New parent name string constructed and old parent name string existing are same
        If (strChildName + strPartname) <> strNamedParentsString Then
            oActiveEntity.NamingParentsString = strPartname
        
            nCount = oNameCounter.GetCountEx(oModelResourceMgr, strPartname, strLocationID)

            If Not strLocationID = vbNullString Then
                strChildName = strPartname + "-" + strLocationID + "-" + Format(nCount, strCountFormat)
            Else
                strChildName = strPartname + "-" + Format(nCount, strCountFormat)
            End If
            oChildProdModel.Name = strChildName
        End If
    Else
        ' no parents - use just the strPartName as the base
        strNamedParentsString = oActiveEntity.NamingParentsString

        'Check if New parent name string constructed and old parent name string existing are same
        If strPartname <> strNamedParentsString Then
            oActiveEntity.NamingParentsString = strPartname

            nCount = oNameCounter.GetCountEx(oModelResourceMgr, strPartname, strLocationID)

            If Not strLocationID = vbNullString Then
                strChildName = strPartname + "-" + strLocationID + "-" + Format(nCount, strCountFormat)
            Else
                strChildName = strPartname + "-" + Format(nCount, strCountFormat)
            End If
            oChildProdModel.Name = strChildName
        End If
    End If
   
    Set jContext = Nothing
    Set oModelResourceMgr = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
    
    Set oNameCounter = Nothing
    Set oChildProdModel = Nothing
''    Set oParent = Nothing
''    Set oParentProdModel = Nothing
    
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

    Const Method = "IJNameRule_GetNamingParents"
    On Error GoTo label

    Dim oSysParent As IJSystem
    Dim oSysChild As IJSystemChild

    On Error Resume Next
    Set oSysChild = oEntity
    Set oSysParent = oSysChild.GetParent
    
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

    Const Method = "GetPartName"
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
    m_oErrors.Add Err.Number, "CNameRule:GetPartName", Err.Description
    Err.Raise E_FAIL
End Function

Private Function GetPart(ByVal oEntity As Object) As IJDPart

    Const Method = "GetPart"
    On Error GoTo label
    
    Dim oRelationHelper As IMSRelation.DRelationHelper
    Dim oCollection As IMSRelation.DCollectionHelper 'IJDrelatiopnshipCol
    Dim oPart As IJDPart
    
    Set oRelationHelper = oEntity
    
    Set oCollection = oRelationHelper.CollectionRelations(IID_IJPartOcc, PARTROLE)
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
    m_oErrors.Add Err.Number, "CNameRule:GetPart", Err.Description
    Err.Raise E_FAIL
End Function

''//////////////////////////////////////////////////////////////////////////
'' InitGuid
''      Initialize a GUID
''//////////////////////////////////////////////////////////////////////////
Public Function InitGuid(a As Byte, b As Byte, c As Byte, d As Byte, e As Byte, f As Byte, _
                          g As Byte, h As Byte, i As Byte, j As Byte, k As Byte, l As Byte, _
                          m As Byte, n As Byte, o As Byte, p As Byte) As Variant

    Dim Guid(0 To 15) As Byte
    
    '' NOTE that the bytes are not in straight order : they're reversed
    '' based on the GUID 4byte/2byte/2byte start
    Guid(0) = d:  Guid(1) = c:  Guid(2) = b:  Guid(3) = a:
    Guid(4) = f:  Guid(5) = e:  Guid(6) = h:  Guid(7) = g:
    Guid(8) = i:  Guid(9) = j:  Guid(10) = k: Guid(11) = l:
    Guid(12) = m: Guid(13) = n: Guid(14) = o: Guid(15) = p
    
    InitGuid = Guid
End Function
 