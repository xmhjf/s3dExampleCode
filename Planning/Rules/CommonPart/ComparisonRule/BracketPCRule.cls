VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketPCRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************************
'  Copyright (C) 2008, Intergraph Corporation. All rights reserved.
'
'  Project     : SharedContent\Src\Planning\Rules\CommonPartRules\ComparisonRule\
'  File        : BracketPCRule.cls
'
'  Description : This rule compares the brackets & then compares the PhysicalConnections(PCs) on lateral ports of Brackets by default.
'                This can be extended for comparing PCs on base and offset ports also.
'                The defualt property of PCs compared is 'EdgeShape', this can be modified.
'
'  History     :
'   Dec 2012      Chaitanya     Initial creation
'**************************************************************************************

Option Explicit

Implements IJCmnPartComparison

Private Const sSOURCEFILE As String = "BracketPCRule.cls"
Private Const ERRORPROGID As String = "IMSErrorLog.ServerErrors"
Private m_oErrors As IJEditErrors   ' To collect and propagate the errors.
Private m_oError As IJEditError     ' Defined here for convenience

Private Sub Class_Initialize()
    If m_oErrors Is Nothing Then
        Set m_oErrors = CreateObject(ERRORPROGID)
    End If
End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub


Private Function IJCmnPartComparison_IsCommon(ByVal oSrcObj As Object, ByVal oDestObj As Object, _
                                              ByVal oCmnPartMgr As Object) As Boolean
Const METHOD = "IJCmnPartComparison_IsCommon"
On Error GoTo ErrorHandler
   
    Dim bIsCommon           As Boolean
    Dim oPlnCompareHelper   As IJPlnCompareHelper
    Dim oMatchingPCs        As IJElements
    
    Dim totalPCPairsCount As Integer
    Dim iIndex As Integer
    Dim eachPair As IJPlnPhysConnPair
    
    bIsCommon = False
    Set oPlnCompareHelper = New CPlnCommonPartHelper
    
    oPlnCompareHelper.Candidate = oSrcObj
    oPlnCompareHelper.Target = oDestObj
                
    oPlnCompareHelper.DistanceTolerance = 0.001
    
    oPlnCompareHelper.Checks MaterialTypeAndGrade
    oPlnCompareHelper.Checks Definition
    oPlnCompareHelper.Checks OuterContours
    oPlnCompareHelper.Checks InnerContours
    
    'For getting the matching PCs on lateral ports
    oPlnCompareHelper.RelationChecks LateralPhysicalConnections
    
    'For getting the matching PCs on Base and offset ports
    'oPlnCompareHelper.RelationChecks PhysicalConnections
    
    bIsCommon = oPlnCompareHelper.AreCommon(oMatchingPCs)
    
    'If not common exit.
    If bIsCommon = False Then
        GoTo Wrapup
    Else
        
        'All pairs shoudl be full match, else return false.
        totalPCPairsCount = oMatchingPCs.Count
                
        'If no mapping PCs found, then return false.
        If totalPCPairsCount = 0 Then
            bIsCommon = False
            GoTo Wrapup
        End If
    
        For iIndex = 1 To totalPCPairsCount
    
            Set eachPair = oMatchingPCs.Item(iIndex)
    
            If eachPair.TypeOfMatch <> PlnMatch_Full Then
                bIsCommon = False
                GoTo Wrapup
            End If
            
            Set eachPair = Nothing
        Next
    End If
    
    'Compare the PC properties.
    Dim oCandAttrCol                As IJDAttributesCol
    Dim oTargAttrCol                As IJDAttributesCol
    
    'User can change this property name.
    'All the attributes on weld parameter interface can be compared.
    Dim strPropertyToCompare As String
    strPropertyToCompare = "EdgeShape"
        
    For iIndex = 1 To totalPCPairsCount
    
        Set eachPair = oMatchingPCs.Item(iIndex)
                
        Set oCandAttrCol = GetAttributes(eachPair.CandidatePhysConn)
        Set oTargAttrCol = GetAttributes(eachPair.TargetPhysConn)
        
        If CompareProperty(oCandAttrCol, oTargAttrCol, strPropertyToCompare) = False Then
            bIsCommon = False
            GoTo Wrapup
        End If
        
        Set oCandAttrCol = Nothing
        Set oTargAttrCol = Nothing
        
    Next
    
    'If the control is here, PC properties are same
    bIsCommon = True
    
Wrapup:
    IJCmnPartComparison_IsCommon = bIsCommon
    
Exit Function

ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Function

'If the property exists and both Candidate and Target has same values, returns true.
Private Function CompareProperty(oCandAttribs As IJDAttributesCol, oTargAttribs As IJDAttributesCol, _
                                                            sPropertyName As String) As Boolean


    CompareProperty = False
    
    Dim eachAttrib As IJDAttribute
    Dim strCandValue As String
    Dim strTargValue As String
        
    If Not oCandAttribs Is Nothing Then
        For Each eachAttrib In oCandAttribs
            If eachAttrib.AttributeInfo.Name = sPropertyName Then
                
                strCandValue = eachAttrib.Value
                Exit For
            End If
        Next
    End If
    
    Set eachAttrib = Nothing
    
    If Not oTargAttribs Is Nothing Then
        For Each eachAttrib In oTargAttribs
            If eachAttrib.AttributeInfo.Name = sPropertyName Then
            
                strTargValue = eachAttrib.Value
                Exit For
            End If
        Next
    End If
    
    If strCandValue <> vbNullString And strTargValue <> vbNullString Then
        If StrComp(strCandValue, strTargValue, vbTextCompare) = 0 Then
                        
            CompareProperty = True
        End If
    End If
    
End Function

Private Function GetAttributes(oPhysicalConnection As Object) As IJDAttributesCol
    
    Dim oSmartOcc                   As IJSmartOccurrence
    Dim oSmartItem                  As IJSmartItem
    Dim oAttributes                 As IJDAttributes
    Dim oSymbDef                    As IJDSymbolDefinition
    Dim oUserCmds                   As IJDUserCommands
    Dim oCmd                        As IJDCommandDescription
    
    Set oSmartOcc = oPhysicalConnection
    Set oAttributes = oPhysicalConnection
    
    Set oSmartItem = oSmartOcc.ItemObject
    Set oSymbDef = oSmartItem.ParameterRuleDef
    Set oUserCmds = oSymbDef.IJDUserCommands
    Set oCmd = oUserCmds.GetCommand("AnswersIID")
    Set GetAttributes = oAttributes.CollectionOfAttributes(oCmd.Source)
    
    Set oSmartOcc = Nothing
    Set oAttributes = Nothing
    Set oSmartItem = Nothing
    Set oSymbDef = Nothing
    Set oUserCmds = Nothing
    Set oCmd = Nothing
   
End Function

