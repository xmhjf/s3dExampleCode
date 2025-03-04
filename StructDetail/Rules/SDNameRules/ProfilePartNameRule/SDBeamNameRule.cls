VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SDBeamNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2001-2002 Global Research and Development, Inc.
'  All rights reserved.
'
'  Project: SDProfilePartNameRules
'
'  Abstract: The file contains name rules for Beam Profile Parts
'
'******************************************************************
Option Explicit

Implements IJNameRule
                                                        
Private Const MODULE = "SDProfilePartNameRules.SDBeamNameRule."
          
Private m_oErrors As IJEditErrors
Private m_oNameUtil As SDNameRulesUtilHelper
Private m_oNameHelp As SDNameRuleHelper
Private m_oSDOHelper As StructDetailObjects.Helper

Private Sub Class_Initialize()
    On Error GoTo label
   
    'set globals
    Set m_oNameUtil = New GSCADSDNameRulesUtil.SDNameRulesUtilHelper
    Set m_oNameHelp = New SDNameRulesHelper.SDNameRuleHelper
    Set m_oErrors = New IMSErrorLog.JServerErrors
    Set m_oSDOHelper = New StructDetailObjects.Helper
Exit Sub

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "Class_Initialize", Err.Description
    Err.Raise E_FAIL
End Sub

Private Sub Class_Terminate()
    'clear globals
    Set m_oErrors = Nothing
    Set m_oNameUtil = Nothing
    Set m_oNameHelp = Nothing
    Set m_oSDOHelper = Nothing
End Sub

'********************************************************************
' Function:
'   IJNameRule_ComputeName
' Description:
'   Refer to Naming.doc command spec for naming rules
'********************************************************************
Private Sub IJNameRule_ComputeName(ByVal oEntity As Object, _
                                   ByVal oParents As IJElements, _
                                   ByVal oActiveEntity As Object)
    Const METHOD = "IJNameRule_ComputeName"
    Dim oNamedItem As IJNamedItem 'stores name
    Dim oBeamPart As StructDetailObjects.BeamPart 'wrapper
    Dim oResourceMgr As IUnknown
    Dim oSysParent As Object
    Dim oAsmParent As Object
    Dim strBlkParentName As String
    Dim strCategoryName As String
    Dim strNamedParentsString As String 'detect if name change has occured
    Dim strLocationID As String
    On Error GoTo errLabel
                                                 
    If oEntity Is Nothing Then 'error, no entity to name
        Err.Raise E_INVALIDARG, MODULE, METHOD
    Else 'entity valid, set up wrapper and name interface
        Set oBeamPart = New StructDetailObjects.BeamPart
        Set oBeamPart.object = oEntity
        Set oNamedItem = oEntity
    End If
    
    'use index naming fields to construct name
    Call SetEnviron(oParents, oAsmParent)
    Set oResourceMgr = m_oSDOHelper.GetResourceManagerFromObject(oBeamPart.object)
    Call GetNonpositionNamingFields(oBeamPart, _
                                    oResourceMgr, _
                                    oAsmParent, _
                                    strBlkParentName, _
                                    strCategoryName)
    strNamedParentsString = strBlkParentName + strCategoryName
   
    'if naming parents string different, get index and update name
    If (strNamedParentsString <> oActiveEntity.NamingParentsString) Then
        Dim iPos As Long
        
        Call m_oNameHelp.SDIndexByCounter(strNamedParentsString, iPos, oResourceMgr)
        oActiveEntity.NamingParentsString = strNamedParentsString
        
        m_oNameHelp.GetWorkShareLocationID oResourceMgr, strLocationID
        
        If Len(strLocationID) > 0 Then
            oNamedItem.Name = "<" + strBlkParentName + ">-" + _
                          strCategoryName + "." + Format(iPos) + "-" + strLocationID
        Else
        
            oNamedItem.Name = "<" + strBlkParentName + ">-" + _
                          strCategoryName + "." + Format(iPos)
        End If
    End If
    
    Set oBeamPart = Nothing
Exit Sub

errLabel:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IJNameRule_ComputeName", Err.Description
    Err.Raise E_FAIL
End Sub

'********************************************************************
' Function:
'   IJNameRule_GetNamingParents
' Description
'   All the Naming Parents that need to participate in an objects naming are
'   added here to the IJElements collection. The parents added here are used
'   in computing the name of the object in ComputeName() of the same
'   interface. Both these methods are called from naming rule semantic.
'********************************************************************
Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements
    Const METHOD = "IJNameRule_GetNamingParents"
    Dim oAsmParent As IJAssembly
    Dim oAsmChild As IJAssemblyChild
    Dim oWorkColl As IJElements
    On Error GoTo label
    
    Set oAsmChild = oEntity
    Set oAsmParent = oAsmChild.Parent
    Set oWorkColl = New IMSCoreCollections.JObjectCollection

    If Not (oAsmParent Is Nothing) Then
         oWorkColl.Add oAsmParent
    End If
        
    Set IJNameRule_GetNamingParents = oWorkColl
    Set oWorkColl = Nothing
Exit Function

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IJNameRule_GetNamingParents", Err.Description
    Err.Raise E_FAIL
End Function

'********************************************************************
' Function:
'   GetBeamCategory
' Description
'   get the category for the profile part
'********************************************************************
Private Function GetBeamCategory(ByVal oBeamPart As StructDetailObjects.BeamPart, _
                                 ByVal oModelResourceMgr As IUnknown, _
                                 Optional ProfileCatPriority As Long) As String
    Dim objCategory As GSCADMetaDataHelper.CMetaDataCategoryQuery
    Dim LongNames() As String
    Dim ShortNames() As String
    Dim PriorityValues() As Long
    Dim eType As StructProfileType
    Dim lCat As Long
    Dim status As Long
    On Error GoTo label
    
    GetBeamCategory = "" 'make sure its blank in case no category
    
    Set objCategory = New GSCADMetaDataHelper.CMetaDataCategoryQuery
    eType = oBeamPart.BeamType
    lCat = oBeamPart.NamingCategory
    If lCat < 0 Then 'check for no category returned (negative not valid)
        GoTo CatNotFound
    End If
    
    Select Case eType
        Case sptVertical
            status = objCategory.GetVerticalBeamCategoryInfo(oModelResourceMgr, LongNames, ShortNames, PriorityValues)
        Case sptTransversal
            status = objCategory.GetTransverseBeamCategoryInfo(oModelResourceMgr, LongNames, ShortNames, PriorityValues)
        Case sptLongitudinal
            status = objCategory.GetLongitudinalBeamCategoryInfo(oModelResourceMgr, LongNames, ShortNames, PriorityValues)
       Case Else
            GoTo CatNotFound
    End Select
    
    ' lCat is the priorityvalue, NOT the index.  So search for the value to get the appropriate index
    ProfileCatPriority = lCat
    
    Dim i As Long
    
    For i = LBound(PriorityValues) To UBound(PriorityValues)
        If PriorityValues(i) = lCat Then
            GetBeamCategory = ShortNames(i)
            Exit For
        End If
    Next
    
CatNotFound:
    Set objCategory = Nothing
Exit Function

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "GetProfileCategory", Err.Description
    Err.Raise E_FAIL
End Function

Private Sub GetNonpositionNamingFields(ByVal oBeamPart As StructDetailObjects.BeamPart, _
                                       ByVal oResourceMgr As IUnknown, _
                                       ByVal oAsmParent As Object, _
                                       strBlkParentName As String, _
                                       strCategoryName As String)
    Dim ProfileCatPriority As Long
    On Error GoTo label
    
    If Not (oAsmParent Is Nothing) Then
        strBlkParentName = m_oNameHelp.GetBlockName(oAsmParent, m_oNameUtil)
    Else
        strBlkParentName = "NoBlk"
    End If
    
    strCategoryName = GetBeamCategory(oBeamPart, oResourceMgr, ProfileCatPriority)
Exit Sub

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "GetNonpositionNamingFields", Err.Description
    Err.Raise E_FAIL
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Extract the key environment objects
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetEnviron(ByVal oParents As IJElements, oAsmParent As Object)
    Dim oElement As Object
    On Error GoTo label
    
    ' get all the relevant parent objects used for naming
    For Each oElement In oParents
        If (TypeOf oElement Is IJAssembly) Then
            Set oAsmParent = oElement
        End If
            
        Set oElement = Nothing
    Next
Exit Sub

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "SetEnviron", Err.Description
    Err.Raise E_FAIL
End Sub
