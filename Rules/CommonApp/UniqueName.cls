VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UniqueName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 1999, Intergraph Corporation.  All rights reserved.
'
'  Project: NameRules
'
'  Abstract: The file contains a sample implementation of a Unique naming rule
'
'  Author: MUR
'
'   History:
'       SS Oct/07/2003
'           TR CP49909: fixed the usage of LocationID
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Implements IJNameRule
                                                        
Const vbInvalidArg = &H80070057

Private Const Module = "CommonNameRule: "
Private Const DEFAULTBASE = "Defaultname"
Private Const MODELDATABASE = "Model"
Private Const strCountFormat = "0000"       ' define fixed-width number field for name

Dim m_oErrors As IJEditErrors

Private Sub Class_Initialize()

    'DI-CP-164764    Prevent DLL unloading for high usage DLLs
    If m_ForceDllInProcess = 0 Then
        m_ForceDllInProcess = ForceDllInProcess(AddressOf ForceDllInProcess)
    End If

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
    Dim oNamingRulesHlpr As GSCADNameRuleSemantics.IJDNamingRulesHelper
    Dim oFilterFactory As GSCADMiddleFilters.MiddleFiltersFactory
    Dim oObjectTypeFilter As IJObjTypeFilter
    Dim oSimpleFilter As IJSimpleFilter
    Dim oModelPOM As IJDPOM
    Dim strCLSID As String
    Dim unkMoniker As IUnknown
    Dim bNameUnique As Boolean
    Dim strTempChildName As String
    Dim loopCount As Long
    If oObject Is Nothing Then
        Err.Raise vbInvalidArg, Module, METHOD
    End If
    
    Set oFilterFactory = New GSCADMiddleFilters.MiddleFiltersFactory

    If Not oFilterFactory Is Nothing Then
        Set oObjectTypeFilter = oFilterFactory.CreateEntity(ObjTypeFilter, Nothing)
        If Not oObjectTypeFilter Is Nothing Then
            Set oModelPOM = oModelResourceMgr
            Set unkMoniker = oModelPOM.GetObjectMoniker(oObject)
            If Not unkMoniker Is Nothing Then
                strCLSID = oModelPOM.ClsidOf(unkMoniker)
                Dim IClassInfo As IJDClassInfo
                Dim IAttributeMetadata As IJDAttributeMetaData
                Dim IBOClassificationNode As IJDBOClassificationNode
                Dim GDIInterfaces As IJDInfosCol
                
                Set IAttributeMetadata = oModelPOM
                If Not IAttributeMetadata Is Nothing Then
                    Set IClassInfo = IAttributeMetadata.ClassInfo(strCLSID)
                    If Not IClassInfo Is Nothing Then
                        Set IBOClassificationNode = IClassInfo.BOClassificationLeafNode
                        If Not IBOClassificationNode Is Nothing Then
                            Set GDIInterfaces = IBOClassificationNode.ImplementedGDInterfaces
                            Dim stringOfInterfaces As String
                            If GDIInterfaces.Count = 1 Then
                                stringOfInterfaces = GDIInterfaces.Item(1).Name
                            ElseIf GDIInterfaces.Count > 1 Then
                                'Enumate to build an array of IIDs or concatenated string of interface names with “;” delimiters
                                Dim i As Long
                                For i = 1 To GDIInterfaces.Count
                                    If stringOfInterfaces = vbNullString Then
                                        stringOfInterfaces = GDIInterfaces.Item(i).Name
                                    Else
                                        If stringOfInterfaces = vbNullString Then
                                            stringOfInterfaces = GDIInterfaces.Item(i).Name
                                        Else
                                            stringOfInterfaces = GDIInterfaces.Item(i).Name + ";" + stringOfInterfaces
                                        End If
                                    End If
                                Next i
                            Else
                                'Error
                            End If
                            
                            oObjectTypeFilter.IID = stringOfInterfaces
                        End If
                    End If
                End If
                
                Set oSimpleFilter = oObjectTypeFilter
            End If
        End If
    End If
    
    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
    Set oNamingRulesHlpr = New GSCADNameRuleHlpr.NamingRulesHelper
    Set oChildNamedItem = oObject
    bNameUnique = False
    ' Get name of part
    strPartname = GetPartName(oObject)
    If strPartname = "" Then
        ' if we cannot get the part name, then use the type string
        strPartname = oChildNamedItem.TypeString
    End If

    ' Use part Name for basename, remove blanks:
    strPartname = Join(Split(strPartname), "")
    'Variable to maintain the loopcount , to avoid infinite loop situation
    loopCount = 1

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
        strTempChildName = strChildName
        strNamedParentsString = oActiveEntity.NamingParentsString
        
        'Check if New parent name string constructed and old parent name string existing are same
        If (strChildName + strPartname) <> strNamedParentsString Then
            oActiveEntity.NamingParentsString = strChildName + strPartname
            Do While (Not bNameUnique)
                strChildName = strTempChildName
                
                strLocation = vbNullString
                nCount = oNameCounter.GetCountEx(oModelResourceMgr, strChildName + strPartname, strLocation)
                
                If strLocation <> vbNullString Then
                    strChildName = strChildName & "-" & strPartname & "-" & strLocation & "-" & Format(nCount, strCountFormat)
                Else
                    strChildName = strChildName & "-" & strPartname & "-" & Format(nCount, strCountFormat)
                End If
                bNameUnique = oNamingRulesHlpr.IsGeneratedNameUnique(oObject, oSimpleFilter, strChildName)
                loopCount = loopCount + 1
                ' this is just to avoid infinite loop situation
                If loopCount > 50 Then
                    Exit Do
                End If
            Loop
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

    Dim oSysParent As IJSystem
    Dim oSysChild As IJSystemChild

    If (TypeOf oEntity Is IJSystemChild) Then
        Set oSysChild = oEntity
    End If
    
    If (Not oSysChild Is Nothing) Then
        Set oSysParent = oSysChild.GetParent
    End If
    

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

