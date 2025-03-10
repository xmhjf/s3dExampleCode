VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BoardMirrorAssyNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
' Copyright (C) 2002, Intergraph Corporation.  All rights reserved.
'
' Project:
'   NameRules
'
' Abstract:
'   The file contains the VB-namerule for Assemblies
'   It sets the same name as the symmetrically related object.
'   If no Symmetrically related object exists, it sets the name in same process as Assembly Name Rule.
'
' Author:
'   Arnab Biswas
'
' History:
'   12th Mar 2006        Kishore           Created - DM-CP�111010  Mirror Assembly cmd
'                                                  - provide option to copy assembly name
'                                                    (copy of v6.1 fix)
'   Apr 21 2009          Kishore           DM-CP-163399  Performance problem when creating a new assembly
'******************************************************************
Private Const IID_IJAssemblyBase As String = "{DC16272D-8B0E-4079-A5C9-795D97E05584}"

Option Explicit

Implements IJNameRule
                                                        
Private Const Module = "BoardMirrorAssyNameRule: "
Private Const MODELDATABASE = "Model"
Private Const strCountFormat = "000000"       ' define fixed-width number field for name
Private Const ERRORPROGID As String = "IMSErrorLog.ServerErrors"
Private m_oErrors As IJEditErrors       ' To collect and propagate the errors.
Private m_oError As IJEditError         ' The error to raise.

Private Sub Class_Initialize()
   Set m_oErrors = CreateObject(ERRORPROGID)
End Sub

Private Sub Class_Terminate()
   Set m_oErrors = Nothing
   Set m_oError = Nothing
End Sub

'********************************************************************
' Description:
' Creates a name for the object passed in.
'   It sets the same name as the symmetrically related object.
'   If no Symmetrically related object exists, it sets the name in same process as Assembly Name Rule as follows.
'The name is based on the  string "A" and an Index.
' The Index is unique for the Asembly.
' It is assumed that all Naming Parents and the Object implement IJNamedItem.
' The Naming Parents are added in AddNamingParents() of the same interface.
' Both these methods are called from naming rule semantic.
'********************************************************************
Private Sub IJNameRule_ComputeName(ByVal oObject As Object, ByVal elements As IJElements, ByVal oActiveEntity As Object)
    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo ErrorHandler
    
    Dim jContext As IJContext
    Dim oModelResourceMgr As IUnknown
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim oConnectMiddle As IJDAccessMiddle
    Dim strModelDBID As String
    Dim strNamedParentsString As String
    Dim oNameCounter As IJNameCounter
    Dim oChildName As IJNamedItem
    Dim strChildName As String
    Dim strParentName As String
    Dim nCount As Long
    Dim strPartname As String
    Dim oParentName As IJNamedItem
    Dim oSourceAssembly As IJAssembly
    
    'Get the middle context
    Set jContext = GetJContext()
    
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType(MODELDATABASE)
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)

       
    If oObject Is Nothing Then
        Err.Raise E_ACCESSDENIED, Module, METHOD
    End If
    
    Set oChildName = oObject
    
    'First Check whether the Object is the destination of PlnMirrorAssembly.
    Set oSourceAssembly = GetSymmetricalAssembly(oObject)
    If Not oSourceAssembly Is Nothing Then
        Dim oSourceAssemblyNamedObject As IJNamedItem
        Set oSourceAssemblyNamedObject = oSourceAssembly
        oChildName.Name = CalculateSymmetricalAssemblyName(oSourceAssemblyNamedObject.Name)
        Set oSourceAssemblyNamedObject = Nothing
        Set oSourceAssembly = Nothing
        GoTo Wrapup
    End If
    
    'Else do the same as like Assembly Naming Rule.
    
    strNamedParentsString = oActiveEntity.NamingParentsString
         
    strPartname = oChildName.TypeString
    ' Use part Name for basename, remove blanks:
    strPartname = Join(Split(strPartname), "")
    Dim oEntity As Object
    Dim strType As String
    Dim strLocationID As String

    If TypeOf oObject Is IJPlanningAssembly Then
        strType = "A"
    ElseIf TypeOf oObject Is IJAssemblyBlock Then
        strType = "AB"
    End If
    
    If strPartname <> strNamedParentsString Then
        If oChildName.Name = "New Assembly" Or oChildName.Name = "" Then
            Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
        nCount = oNameCounter.GetCountEx(oModelResourceMgr, strPartname, strLocationID)
                
        ' If a Name Rule ID was given at the creation of the database it will be returned
        ' into strLocationID.  If not it will be a null string and not print out
        strChildName = strType & strLocationID & Format(nCount, strCountFormat)
            oChildName.Name = strChildName
        End If
        oActiveEntity.NamingParentsString = strPartname
    End If
        
Wrapup:
    Set jContext = Nothing
    Set oModelResourceMgr = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
    Set oEntity = Nothing

    Set oNameCounter = Nothing
    Set oChildName = Nothing
    Set oParentName = Nothing
    
Exit Sub
ErrorHandler:
    m_oError = m_oErrors.AddFromErr(Err, Module & " - " & METHOD)
    m_oError.Raise
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Description
' All the Naming Parents that need to participate in an objects naming are added here to the
' IJElements collection. The parent added here is root of assembly hierarchy and is used in computing
' the name of the object in ComputeName() of the same interface. Both these methods are called from
' naming rule semantic.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements
    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo ErrorHandler

    Dim oNamingParent As Object
    
    Set oNamingParent = GetSymmetricalAssembly(oEntity)
    
    If oNamingParent Is Nothing Then
        Set oNamingParent = GetAssemblyTreeRoot()
    End If
    
    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection

    If Not (oNamingParent Is Nothing) Then
        Call IJNameRule_GetNamingParents.Add(oNamingParent)
    End If

Exit Function
ErrorHandler:
   m_oError = m_oErrors.AddFromErr(Err, Module & " - " & METHOD)
    m_oError.Raise
End Function

' ***************************************************************************
'
' Function
'   AsIJNamedItem()
'
' Abstract
'   Returns the IJNamedItem interface for the given entity
'
' Usage
'   Use to quickly get access to the IJNamedItem interface without the
'   need for instanciating a variable.
'
' ***************************************************************************

Public Function AsIJNamedItem(ByVal Entity As Object) As IJNamedItem
    Set AsIJNamedItem = Entity
End Function

' ***************************************************************************
'
' Function
'   GetAssemblyName()
'
' Abstract
'   Returns the name of an assembly
'
' ***************************************************************************

Public Function GetAssemblyName(ByVal oAssembly As GSCADAsmHlpers.IJAssembly) As String
    GetAssemblyName = AsIJNamedItem(oAssembly).Name
End Function



'**************************************************************************
' Method : GetSymmetricalAssembly
'
' Description:
'   Gets the assembly reelated to the passed in by the relation PlnMirrorAsm. Returns NOTHING
'   if there is no mirror assembly.
'
' Arguments:
'   [in] ByVal oAssy As GSCADAsmHlpers.IJAssembly, the assembly to get the relationhelper from
'
' Return Values:
'   the symmetrical assembly,
'   if no symmetrical assembly returns "nothing"
'**************************************************************************
Private Function GetSymmetricalAssembly(ByVal oAssy As GSCADAsmHlpers.IJAssembly) As GSCADAsmHlpers.IJAssembly
    Const METHOD As String = "GetSymmetricalAssembly"
    On Error GoTo ErrorHandler
    
    Dim oRelationHelper As IMSRelation.DRelationHelper      'IJDAssocRelation
    Dim oCollectionHelperOrig As IMSRelation.DCollectionHelper  'IJDTargetObjectCollection
    Dim oMirrorObjUnk As IUnknown
    
    'Get the relation helper from the passed in assembly
    Set oRelationHelper = oAssy
    
    'Get the collection of relations
    Set oCollectionHelperOrig = oRelationHelper.CollectionRelations(IID_IJAssemblyBase, "PlnMirrorAsm_DEST")
    'check that there is only one relationship
    If oCollectionHelperOrig.Count > 0 Then
        Set oMirrorObjUnk = oCollectionHelperOrig.Item(1)
    Else
        'test relations the other way
        Set oMirrorObjUnk = Nothing
    End If
    
    'return the symmetrical assembly, might be nothing
    Set GetSymmetricalAssembly = oMirrorObjUnk
    
CleanUp:
    Set oMirrorObjUnk = Nothing
    Set oRelationHelper = Nothing
    Set oCollectionHelperOrig = Nothing
    Exit Function

ErrorHandler:
    Err.Clear
    Set GetSymmetricalAssembly = Nothing
End Function 'GetSymmetricalAssembly

'**************************************************************************
' Function : CalculateSymmetricalAssemblyName
'
' Abstract
'   Returns the name of an assembly by replacing the positional denotation last letter of the 1st token of
'   sSourceAssemblyName with "S" (or "P") if sSourceAssemblyName contains only one '-' delimiter.
'
'   Example: sSourceAssemblyName    CalculateSymmetricalAssemblyName
'            -------------------------------------------------------
'            B11P-F101A             B11S-F101A  (replaces P with S)
'            B11S-G10B              B11P-G10B   (replaces S with P)
'            F32C-F101A             F32C-F101A  (returns the same name since the last letter is not "P" or "S")
'            F101A                  F101A       (returns the same name since no '-' delimiter)
'**************************************************************************
Private Function CalculateSymmetricalAssemblyName(ByVal sSourceAssemblyName As String) As String
    Const METHOD As String = "CalculateSymmetricalAssemblyName"
    On Error GoTo ErrorHandler
    
    Dim arrTmp() As String, sToken1 As String, sToken2 As String, sToken1LastChar As String
    
    arrTmp = Split(sSourceAssemblyName, "-")
    
    If UBound(arrTmp) = 1 Then
        sToken1 = arrTmp(0)
        sToken2 = arrTmp(1)
        sToken1LastChar = UCase(Right(sToken1, 1))
        
        If sToken1LastChar = "P" Or sToken1LastChar = "S" Then
            Dim sToken1ExcludeLastChar As String
            
            sToken1ExcludeLastChar = Left(sToken1, Len(sToken1) - 1)
            
            If sToken1LastChar = "P" Then
                CalculateSymmetricalAssemblyName = sToken1ExcludeLastChar & "S-" & sToken2
            Else
                CalculateSymmetricalAssemblyName = sToken1ExcludeLastChar & "P-" & sToken2
            End If
        Else
            CalculateSymmetricalAssemblyName = sSourceAssemblyName
        End If
    Else
        CalculateSymmetricalAssemblyName = sSourceAssemblyName
    End If
    
    Exit Function

ErrorHandler:
    CalculateSymmetricalAssemblyName = sSourceAssemblyName
End Function

