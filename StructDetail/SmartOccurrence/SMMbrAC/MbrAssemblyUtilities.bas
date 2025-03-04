Attribute VB_Name = "MbrAssemblyUtilities"

Option Explicit

'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : MbrAssemblyUtilities.cls
'
'  Description :
'
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'    26/APR/2011 - CM : Added additional methods to support creation of Web/Flange Cut
'                       on Split/Seam Angle cases
'
'*********************************************************************************************

Private Const MODULE = "StructDetail\Data\SmartOccurrence\SMMbrAC\MbrAssemblyUtilities"
Public Const m_sProjectName As String = CUSTOMERID + "MbrAC"
Public Const m_sProjectPath As String = "S:\StructDetail\Data\SmartOccurrence\" + m_sProjectName + "\"

'*********************************************************************************************
' Method      : Set_WebCutQuestions
' Description :
'       Copy Answers  to End Cut given :
'           Given the MemberDescription and
'           EndCut Selection/Item Rule Prog Id and
'
'
'*********************************************************************************************
Public Sub Set_WebCutQuestions(pMemberDescription As IJDMemberDescription, _
                               sEndCutProgId As String, _
                               bSwitchWeldPart As Boolean, _
                               bForceUpdate As Boolean)
    Const METHOD = "MbrAssemblyUtilities::Set_WebCutQuestions"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim sWeldPart As String

    Dim oCopyAnswerHelper As DefinitionHlprs.CopyAnswerHelper
    
    sMsg = "Creating CopyAnswerHelper object"
    Set oCopyAnswerHelper = New DefinitionHlprs.CopyAnswerHelper
    Set oCopyAnswerHelper.MemberDescription = pMemberDescription
    
    If bSwitchWeldPart Then
        sWeldPart = "Second"
    Else
        sWeldPart = "First"
    End If
    
    sMsg = "Setting the Answer on End Cut Selection rule..."
    oCopyAnswerHelper.PutAnswer sEndCutProgId, "WeldPart", sWeldPart
    
    If bForceUpdate Then
        WebCut_ForceUpdate pMemberDescription
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub

'*********************************************************************************************
' Method      : WebCut_ForceUpdate
' Description : Force an Update on the WebCut using the same interface,IJStructGeometry,
'       as is used when placing the WebCut as an input to the FlangeCuts
'       This allows Assoc to always recompute the WebCut before FlangeCuts
'
'*********************************************************************************************
Public Sub WebCut_ForceUpdate(pMemberDescription As IJDMemberDescription)
    Const METHOD = "MbrAssemblyUtilities::WebCut_ForceUpdate"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    Dim iIndex As Long
    Dim jIndex As Long
    Dim lDispId As Long
    
    Dim oWebCut As Object
    Dim oMemberItems As IJElements
    Dim oMemberObjects As IJDMemberObjects
    Dim oStructFeature As IJStructFeature
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    
    ' Force an Update on the WebCut using the same interface,IJStructGeometry,
    ' as is used when placing the WebCut as an input to the FlangeCuts
    ' This allows Assoc to always recompute the WebCut before FlangeCuts
    sMsg = "Calling Structdetailobjects.WebCut::ForceUpdateForFlangeCuts"
    lDispId = pMemberDescription.dispid
    Set oMemberObjects = pMemberDescription.CAO
    
    For iIndex = 1 To oMemberObjects.Count
        If Not oMemberObjects.Item(iIndex) Is Nothing Then
            If iIndex = lDispId Then
                If TypeOf oMemberObjects.Item(iIndex) Is IJStructFeature Then
                    Set oStructFeature = oMemberObjects.Item(iIndex)
                    If oStructFeature.get_StructFeatureType = SF_WebCut Then
                        Set oWebCut = oStructFeature
                        Exit For
                    End If
                End If
            End If
        End If
    Next iIndex
    
    If Not oWebCut Is Nothing Then
        Set oSDO_WebCut = New StructDetailObjects.WebCut
        Set oSDO_WebCut.object = oWebCut
        oSDO_WebCut.ForceUpdateForFlangeCuts
    
        Set oSDO_WebCut = Nothing
        Set oMemberObjects = Nothing
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub

'*********************************************************************************************
' Method      : Set_FlangeCutQuestions
' Description :
'       Copy Answers to End Cut given :
'           the Given the MemberDescription and
'           EndCut Selection/Item Rule Prog Id and
'
'
'*********************************************************************************************
Public Sub Set_FlangeCutQuestions(pMemberDescription As IJDMemberDescription, _
                                  sEndCutProgId As String, _
                                  bSwitchWeldPart As Boolean, _
                                  bBottomFlange As Boolean)
    Const METHOD = "MbrAssemblyUtilities::Set_FlangeCutQuestions"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim oCopyAnswerHelper As DefinitionHlprs.CopyAnswerHelper
        
    Set_WebCutQuestions pMemberDescription, sEndCutProgId, bSwitchWeldPart, False

    sMsg = "Creating CopyAnswerHelper object"
    Set oCopyAnswerHelper = New DefinitionHlprs.CopyAnswerHelper
    Set oCopyAnswerHelper.MemberDescription = pMemberDescription
    
    sMsg = "Setting the Bottom Flange Answer on Flange Selector Rules..."
    If bBottomFlange Then
        oCopyAnswerHelper.PutAnswer sEndCutProgId, "BottomFlange", "Yes"
    Else
        oCopyAnswerHelper.PutAnswer sEndCutProgId, "BottomFlange", "No"
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub
