VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndCutSelCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : EndCutSelCM.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'    06/Sep/2011 - pnalugol
'           For supporting Generic ACs for Tube standard(Bounding)Vs non-tube Standard member(Bounded)
'    21/Sep/2011 - mpulikol
'           DI-CP-200263 Improve performance by caching measurement symbol results
'    19/Oct/2011 - mpulikol
'           CR-CP-203633 Performance: Increase speed of generic member assembly connections
'    7/4/ 2013 - gkharrin\dsmamidi
'               DI-CP-235071 Improve performance by caching edge mapping data
'               Reduced the number of calls to edgemapping rule
'               Replaced GetEdgeMapping with GetEdgeMap, From GetEdgeMap,getting edgemap data if already cached, if not call the edgemapping rule to get the same
'*********************************************************************************************

Const m_sClassName As String = "EndCutSelCM"
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

'*********************************************************************************************
' Method      : CMSetConnectTopAnswer
' Description : Sets the answer to the question
'
'*********************************************************************************************
Public Sub CMSetConnectTopAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "::CMSetConnectTopAnswer"
    On Error GoTo ErrorHandler
    
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)
    
    ' -----------------------------------------------------
    ' Special-case generic ACs predefine the desired answer
    ' -----------------------------------------------------
    ' Set the values, even if the default would match what is desired by the AC
    ' 1) to guard against changes in the default logic affecting the special AC
    ' 2) to avoid recomputing a lot of information, such as where the bounding flange intersects
    Dim sACItemName As String
    Dim oAppConnection As Object
    AssemblyConnection_SmartItemName pSL.SmartOccurrence, sACItemName, oAppConnection
    
    If sACItemName = "WPEF_PF_1" Or sACItemName = "Generic_Stiff_WPEF_PF_1" Then
        pSL.answer("ConnectTop") = "No"
        Exit Sub
    End If
    
    ' ----------------------------------------------------
    ' If not bounded by a member, then the default is "no"
    ' ----------------------------------------------------
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pSL.SmartOccurrence
    
    If (Not TypeOf oSDO_WebCut.Bounding Is ISPSMemberPartPrismatic) And (Not TypeOf oSDO_WebCut.Bounding Is IJProfile) Then
        pSL.answer("ConnectTop") = "No"
        Exit Sub
    End If
    
    ' ---------------------------------------------------------------------
    ' Find out where on the bounding member each bounded edge makes contact
    ' ---------------------------------------------------------------------
    Dim eTopOrWL As ConnectedEdgeInfo
    Dim eBottomOrWR As ConnectedEdgeInfo
    Dim eInsideTFOrTFL As ConnectedEdgeInfo
    Dim eInsideBFOrTFR As ConnectedEdgeInfo
    Dim oMeasurements As Collection
        
    GetConnectedEdgeInfo oSDO_WebCut.object, oSDO_WebCut.BoundedPort, oSDO_WebCut.BoundingPort, eTopOrWL, eBottomOrWR, eInsideTFOrTFL, eInsideBFOrTFR, oMeasurements

    ' ---------------------------------------------
    ' If the bounded extends above the bounding top
    ' ---------------------------------------------
    Dim sAnswer As String
    sAnswer = "No"
    
    If eTopOrWL.IntersectingEdge = Above Then
        ' --------------------------------------------------------------------------------------------------------
        ' Assume that a top cut will remove overlapping material unless more than half the bounded object is above
        ' --------------------------------------------------------------------------------------------------------
        Dim dTop As Double
        Dim dBtm As Double
        
        dTop = oMeasurements.Item("DimPt15ToTop")
        dBtm = oMeasurements.Item("DimPt15ToBottom")
    
        If dTop > dBtm Then
            sAnswer = "Yes"
        End If
    End If
        
    pSL.answer("ConnectTop") = sAnswer
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
      
End Sub

'*********************************************************************************************
' Method      : CMSetConnectFaceOutsideAnswer
' Description : Sets the answer to the question
'
'*********************************************************************************************
Public Sub CMSetConnectFaceOutsideAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "::CMSetConnectFaceOutsideAnswer"
    On Error GoTo ErrorHandler
    
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)
    
    pSL.answer("ConnectFaceOutside") = "No"

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
          
End Sub

'*********************************************************************************************
' Method      : CMSetConnectTopEdgeAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetConnectTopEdgeAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "::CMSetConnectTopEdgeAnswer"
    On Error GoTo ErrorHandler
    
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)
        
    pSL.answer("ConnectTopEdge") = ComputeConnectTopEdgeAnswer(pInput)

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
              
End Sub

'*********************************************************************************************
' Method      : CMSetConnectTopInsideAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetConnectTopInsideAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "::CMSetConnectTopInsideAnswer"
    On Error GoTo ErrorHandler
    
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)
                                           
    pSL.answer("ConnectTopInside") = ComputeConnectTopInsideAnswer(pInput)
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
     
End Sub

'*********************************************************************************************
' Method      : CMSetConnectFaceAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetConnectFaceAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "::CMSetConnectFaceAnswer"
    On Error GoTo ErrorHandler
    
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)
                                  
    pSL.answer("ConnectFace") = ComputeConnectFaceAnswer(pInput)

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
    
End Sub

'*********************************************************************************************
' Method      : CMSetConnectBottomInsideAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetConnectBottomInsideAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                              ByRef ppArgument As Object)
    Const METHOD As String = "::CMSetConnectBottomInsideAnswer"
    On Error GoTo ErrorHandler
    
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)

    pSL.answer("ConnectBottomInside") = ComputeConnectBottomInsideAnswer(pInput)

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
 
End Sub

'*********************************************************************************************
' Method      : CMSetConnectBottomEdgeAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetConnectBottomEdgeAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                        ByRef ppArgument As Object)
    Const METHOD As String = "::CMSetConnectBottomEdgeAnswer"
    On Error GoTo ErrorHandler
    
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)

    pSL.answer("ConnectBottomEdge") = ComputeConnectBottomEdgeAnswer(pInput)

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
 
End Sub

'*********************************************************************************************
' Method      : CMSetConnectBottomAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetConnectBottomAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                    ByRef ppArgument As Object)
    Const METHOD As String = "::CMSetConnectBottomAnswer"
    On Error GoTo ErrorHandler
    
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)
    
    ' ----------------------------------------------------
    ' If not bounded by a member, then the default is "no"
    ' ----------------------------------------------------
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pSL.SmartOccurrence
    
    If (Not TypeOf oSDO_WebCut.Bounding Is ISPSMemberPartPrismatic) And (Not TypeOf oSDO_WebCut.Bounding Is IJProfile) Then
        pSL.answer("ConnectBottom") = "No"
        Exit Sub
    End If
    
    ' ---------------------------------------------------------------------
    ' Find out where on the bounding member each bounded edge makes contact
    ' ---------------------------------------------------------------------
    Dim eTopOrWL As ConnectedEdgeInfo
    Dim eBottomOrWR As ConnectedEdgeInfo
    Dim eInsideTFOrTFL As ConnectedEdgeInfo
    Dim eInsideBFOrTFR As ConnectedEdgeInfo
    Dim oMeasurements As Collection
        
    GetConnectedEdgeInfo oSDO_WebCut.object, oSDO_WebCut.BoundedPort, oSDO_WebCut.BoundingPort, eTopOrWL, eBottomOrWR, eInsideTFOrTFL, eInsideBFOrTFR, oMeasurements

    ' ------------------------------------------------
    ' If the bounded extends below the bounding bottom
    ' ------------------------------------------------
    Dim sAnswer As String
    sAnswer = "No"
    
    If eBottomOrWR.IntersectingEdge = Below Then
        ' -----------------------------------------------------------------------------------------------------------
        ' Assume that a bottom cut will remove overlapping material unless more than half the bounded object is below
        ' -----------------------------------------------------------------------------------------------------------
        Dim dTop As Double
        Dim dBtm As Double
        
        dTop = oMeasurements.Item("DimPt23ToTop")
        dBtm = oMeasurements.Item("DimPt23ToBottom")
    
        If dBtm > dTop Then
            sAnswer = "Yes"
        End If
    End If
    
    pSL.answer("ConnectBottom") = sAnswer

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
 
End Sub

'*********************************************************************************************
' Method      : CMSetConnectInsideAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetConnectInsideAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                    ByRef ppArgument As Object)
    Const METHOD As String = "::CMSetConnectInsideAnswer"
    On Error GoTo ErrorHandler
  
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)
    
    ' -----------------------------------------------------
    ' Special-case generic ACs predefine the desired answer
    ' -----------------------------------------------------
    ' Set the values, even if the default would match what is desired by the AC
    ' 1) to guard against changes in the default logic affecting the special AC
    ' 2) to avoid recomputing a lot of information, such as where the bounding flange intersects
    Dim sACItemName As String
    AssemblyConnection_SmartItemName pSL.SmartOccurrence, sACItemName
    
    If sACItemName = "WPEF_PF_1" Or sACItemName = "Generic_Stiff_WPEF_PF_1" Then
        pSL.answer("ConnectInside") = "Yes"
        Exit Sub
    End If
        
    ' ----------------------------------------------------
    ' If not bounded by a member, then the default is "no"
    ' ----------------------------------------------------
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pSL.SmartOccurrence
    
    If (Not TypeOf oSDO_WebCut.Bounding Is ISPSMemberPartPrismatic) And (Not TypeOf oSDO_WebCut.Bounding Is IJProfile) Then
        pSL.answer("ConnectInside") = "No"
        Exit Sub
    End If
    
    ' ----------------------------------------
    ' Determine if the inside is top or bottom
    ' ----------------------------------------
    
    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    
    Set oEdgeMapColl = New Collection
    
    Set oEdgeMapColl = GetEdgeMap(pSL.SmartOccurrence, oSDO_WebCut.BoundingPort, oSDO_WebCut.BoundedPort, lSectionAlias, bPenetratesWeb)
    Dim sAnswer As String
    sAnswer = "No"
    
    Select Case lSectionAlias
        ' ----------------------------------------------------------------------------
        ' If there is a top or bottom flange only, get the answer from a helper method
        ' ----------------------------------------------------------------------------
        Case 1, 2, 6
            'WebTopFlangeRight, WebBuiltUpTopFlangeRight, WebTopFlange,
            sAnswer = ComputeConnectTopInsideAnswer(pInput)
        Case 3, 4, 7
            'WebBottomFlangeRight, WebBuiltUpBottomFlangeRight, WebBottomFlange
            sAnswer = ComputeConnectBottomInsideAnswer(pInput)
        ' -----------------------------------------------------
        ' For unexpected cases we will default to no connection
        ' -----------------------------------------------------
        Case Else
    End Select

    pSL.answer("ConnectInside") = sAnswer

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
 
End Sub

'*********************************************************************************************
' Method      : CMSetConnectOutsideAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetConnectOutsideAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                    ByRef ppArgument As Object)
    Const METHOD As String = "::CMSetConnectOutsideAnswer"
    On Error GoTo ErrorHandler
  
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)

    pSL.answer("ConnectOutside") = "No"

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
 
End Sub

'*********************************************************************************************
' Method      : CMSetConnectEdgeAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetConnectEdgeAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                    ByRef ppArgument As Object)
    Const METHOD As String = "::CMSetConnectEdgeAnswer"
    On Error GoTo ErrorHandler
  
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)
    
    ' -----------------------------------------------------
    ' Special-case generic ACs predefine the desired answer
    ' -----------------------------------------------------
    ' Set the values, even if the default would match what is desired by the AC
    ' 1) to guard against changes in the default logic affecting the special AC
    ' 2) to avoid recomputing a lot of information, such as where the bounding flange intersects
    Dim sACItemName As String
    AssemblyConnection_SmartItemName pSL.SmartOccurrence, sACItemName
    
    If sACItemName = "WPEF_PF_1" Or sACItemName = "Generic_Stiff_WPEF_PF_1" Then
        pSL.answer("ConnectEdge") = "No"
        Exit Sub
    End If
    
    ' ----------------------------------------------------
    ' If not bounded by a member, then the default is "no"
    ' ----------------------------------------------------
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pSL.SmartOccurrence
    
    If (Not TypeOf oSDO_WebCut.Bounding Is ISPSMemberPartPrismatic) And (Not TypeOf oSDO_WebCut.Bounding Is IJProfile) Then
        pSL.answer("ConnectEdge") = "No"
        Exit Sub
    End If
    
    ' ----------------------------------------
    ' Determine if the inside is top or bottom
    ' ----------------------------------------
    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    
    Set oEdgeMapColl = New Collection
    
    Set oEdgeMapColl = GetEdgeMap(pSL.SmartOccurrence, oSDO_WebCut.BoundingPort, oSDO_WebCut.BoundedPort, lSectionAlias, bPenetratesWeb)
    Dim sAnswer As String
    sAnswer = "No"
    
    Select Case lSectionAlias
        ' ----------------------------------------------------------------------------
        ' If there is a top or bottom flange only, get the answer from a helper method
        ' ----------------------------------------------------------------------------
        Case 1, 2, 6
            'WebTopFlangeRight, WebBuiltUpTopFlangeRight, WebTopFlange,
            sAnswer = ComputeConnectTopEdgeAnswer(pInput)
        Case 3, 4, 7
            'WebBottomFlangeRight, WebBuiltUpBottomFlangeRight, WebBottomFlange
            sAnswer = ComputeConnectBottomEdgeAnswer(pInput)
        ' -----------------------------------------------------
        ' For unexpected cases we will default to no connection
        ' -----------------------------------------------------
        Case Else
    End Select

    pSL.answer("ConnectEdge") = sAnswer

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
 
End Sub

'*********************************************************************************************
' Method      : CMSetFaceTopOutsideCornerAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetFaceTopOutsideCornerAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                ByRef ppArgument As Object)
    Const METHOD As String = "::CMSetFaceTopOutsideCornerAnswer"
    On Error GoTo ErrorHandler
  
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)

    pSL.answer("FaceTopOutsideCorner") = "No"

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
 
End Sub

'*********************************************************************************************
' Method      : CMSetFaceBottomOutsideCornerAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetFaceBottomOutsideCornerAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                ByRef ppArgument As Object)
    Const METHOD As String = "::CMSetFaceBottomOutsideCornerAnswer"
    On Error GoTo ErrorHandler
  
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)

    pSL.answer("FaceBottomOutsideCorner") = "No"

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
 
End Sub

'*********************************************************************************************
' Method      : CMSetFaceInsideCornerAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetFaceInsideCornerAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                   ByRef ppArgument As Object)
    Const METHOD As String = "::CMSetFaceInsideCornerAnswer"
    On Error GoTo ErrorHandler
  
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)
    
    ' -----------------------------------------------------
    ' Special-case generic ACs predefine the desired answer
    ' -----------------------------------------------------
    ' Set the values, even if the default would match what is desired by the AC
    ' 1) to guard against changes in the default logic affecting the special AC
    ' 2) to avoid recomputing a lot of information, such as where the bounding flange intersects
    Dim sACItemName As String
    AssemblyConnection_SmartItemName pSL.SmartOccurrence, sACItemName
    
    If sACItemName = "WPEF_PF_1" Or sACItemName = "Generic_Stiff_WPEF_PF_1" Then
        pSL.answer("FaceInsideCorner") = "Yes"
        Exit Sub
    End If
    
    ' ----------------------------------------------------
    ' If not bounded by a member, then the default is "no"
    ' ----------------------------------------------------
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pSL.SmartOccurrence
    
    If (Not TypeOf oSDO_WebCut.Bounding Is ISPSMemberPartPrismatic) And (Not TypeOf oSDO_WebCut.Bounding Is IJProfile) Then
        pSL.answer("FaceInsideCorner") = "No"
        Exit Sub
    End If
    
    ' ----------------------------------------
    ' Determine if the inside is top or bottom
    ' ----------------------------------------
    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    
    Set oEdgeMapColl = New Collection
    Set oEdgeMapColl = GetEdgeMap(pSL.SmartOccurrence, oSDO_WebCut.BoundingPort, oSDO_WebCut.BoundedPort, lSectionAlias, bPenetratesWeb)
    Dim sAnswer As String
    sAnswer = "No"

    ' ---------------------------------------------------------------------------------------------------------
    ' Call the corrsponding method for only top or bottom bounding edge.  Unexpected cases will default to "no"
    ' ---------------------------------------------------------------------------------------------------------
    Select Case lSectionAlias
        Case 1, 2, 6
            sAnswer = ComputeFaceTopInsideCornerAnswer(pInput)
        Case 3, 4, 7
            sAnswer = ComputeFaceBottomInsideCornerAnswer(pInput)
        Case Else
    End Select

    pSL.answer("FaceInsideCorner") = sAnswer

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
 
End Sub
'*********************************************************************************************
' Method      : CMSetEdgeInsideCornerAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetEdgeInsideCornerAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                   ByRef ppArgument As Object)
    Const METHOD As String = "::CMSetEdgeInsideCornerAnswer"
    On Error GoTo ErrorHandler
 
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)

    ' -----------------------------------------------------
    ' Special-case generic ACs predefine the desired answer
    ' -----------------------------------------------------
    ' Set the values, even if the default would match what is desired by the AC
    ' 1) to guard against changes in the default logic affecting the special AC
    ' 2) to avoid recomputing a lot of information, such as where the bounding flange intersects
    Dim sACItemName As String
    AssemblyConnection_SmartItemName pSL.SmartOccurrence, sACItemName
    
    If sACItemName = "WPEF_PF_1" Or sACItemName = "Generic_Stiff_WPEF_PF_1" Then
        pSL.answer("EdgeInsideCorner") = "Yes"
        Exit Sub
    End If
    
    ' -----------------------------------------
    ' If not a special case AC, default to "No"
    ' -----------------------------------------
    pSL.answer("EdgeInsideCorner") = "No"

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
 
End Sub
'*********************************************************************************************
' Method      : CMSetEdgeOutsideCornerAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetEdgeOutsideCornerAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                   ByRef ppArgument As Object)
    Const METHOD As String = "::CMSetEdgeOutsideCornerAnswer"
    On Error GoTo ErrorHandler
 
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)

    ' -----------------------------------------------------
    ' Special-case generic ACs predefine the desired answer
    ' -----------------------------------------------------
    ' Set the values, even if the default would match what is desired by the AC
    ' 1) to guard against changes in the default logic affecting the special AC
    ' 2) to avoid recomputing a lot of information, such as where the bounding flange intersects
    Dim sACItemName As String
    AssemblyConnection_SmartItemName pSL.SmartOccurrence, sACItemName
    
    If sACItemName = "WPEF_PF_1" Or sACItemName = "Generic_Stiff_WPEF_PF_1" Then
        pSL.answer("EdgeOutsideCorner") = "No"
        Exit Sub
    End If
    
    pSL.answer("EdgeOutsideCorner") = "No"

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
 
End Sub
'*********************************************************************************************
' Method      : CMSetFaceOutsideCornerAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetFaceOutsideCornerAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                   ByRef ppArgument As Object)
    Const METHOD As String = "::CMSetFaceOutsideCornerAnswer"
    On Error GoTo ErrorHandler
 
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)

    pSL.answer("FaceOutsideCorner") = "No"

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
 
End Sub
'*********************************************************************************************
' Method      : CMSetTopEdgeOutsideCornerAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetTopEdgeOutsideCornerAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                   ByRef ppArgument As Object)
    Const METHOD As String = "::CMSetTopEdgeOutsideCornerAnswer"
    On Error GoTo ErrorHandler
 
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)

    pSL.answer("TopEdgeOutsideCorner") = "No"

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
 
End Sub
'*********************************************************************************************
' Method      : CMSetTopEdgeInsideCornerAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetTopEdgeInsideCornerAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                   ByRef ppArgument As Object)
    Const METHOD As String = "::CMSetTopEdgeInsideCornerAnswer"
    On Error GoTo ErrorHandler
 
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)

    pSL.answer("TopEdgeInsideCorner") = "No"

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
 
End Sub
'*********************************************************************************************
' Method      : CMSetFaceTopInsideCornerAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetFaceTopInsideCornerAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                   ByRef ppArgument As Object)
    Const METHOD As String = "::CMSetFaceTopInsideCornerAnswer"
    On Error GoTo ErrorHandler
    
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)
    
    pSL.answer("FaceTopInsideCorner") = ComputeFaceTopInsideCornerAnswer(pInput)

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
     
End Sub
'*********************************************************************************************
' Method      : CMSetFaceBottomInsideCornerAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetFaceBottomInsideCornerAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                   ByRef ppArgument As Object)
    Const METHOD As String = "::CMSetFaceBottomInsideCornerAnswer"
    On Error GoTo ErrorHandler
    
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)
    
    pSL.answer("FaceBottomInsideCorner") = ComputeFaceBottomInsideCornerAnswer(pInput)
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
 
End Sub

'*********************************************************************************************
' Method      : CMSetBottomEdgeInsideCornerAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetBottomEdgeInsideCornerAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                   ByRef ppArgument As Object)
    Const METHOD As String = "::CMSetBottomEdgeInsideCornerAnswer"
    On Error GoTo ErrorHandler
      
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)

    pSL.answer("BottomEdgeInsideCorner") = "No"

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
 
End Sub
'*********************************************************************************************
' Method      : CMSetBottomEdgeOutsideCornerAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetBottomEdgeOutsideCornerAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                   ByRef ppArgument As Object)
    Const METHOD As String = "::CMSetBottomEdgeOutsideCornerAnswer"
    On Error GoTo ErrorHandler
    
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)

    pSL.answer("BottomEdgeOutsideCorner") = "No"

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
 
End Sub

'*********************************************************************************************
' Method      : ComputeConnectTopInsideAnswer
' Description :
'
'*********************************************************************************************
Public Function ComputeConnectTopInsideAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod) As String
    Const METHOD As String = "::ComputeConnectTopInsideAnswer"
    On Error GoTo ErrorHandler
    
    ComputeConnectTopInsideAnswer = "No"

    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)
    
    ' ----------------------------------------------------
    ' If not bounded by a member, then the default is "no"
    ' ----------------------------------------------------
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pSL.SmartOccurrence
    
    If (Not TypeOf oSDO_WebCut.Bounding Is ISPSMemberPartPrismatic) And (Not TypeOf oSDO_WebCut.Bounding Is IJProfile) Then
        Exit Function
    End If
    
    ' --------------------------------------------------
    ' If there is a top cut by default, no PC by default
    ' --------------------------------------------------
    Dim strTopCutByDefault As String
    strTopCutByDefault = ComputeTopWebCutAnswer(pSL.SmartOccurrence)
    
    If strTopCutByDefault = "Yes" Then
        Exit Function
    End If
    
    Dim oAppConnection As Object
    AssemblyConnection_SmartItemName pSL.SmartOccurrence, , oAppConnection
    
    ' ---------------------------------------------------------------------
    ' Find out where on the bounding member each bounded edge makes contact
    ' ---------------------------------------------------------------------
    Dim eTopOrWL As ConnectedEdgeInfo
    Dim eBottomOrWR As ConnectedEdgeInfo
    Dim eInsideTFOrTFL As ConnectedEdgeInfo
    Dim eInsideBFOrTFR As ConnectedEdgeInfo
    Dim oMeasurements As Collection
        
    GetConnectedEdgeInfo pSL.SmartOccurrence, oSDO_WebCut.BoundedPort, oSDO_WebCut.BoundingPort, eTopOrWL, eBottomOrWR, eInsideTFOrTFL, eInsideBFOrTFR, oMeasurements

    ' ------------------------------------
    ' Determine what bounded flanges exist
    ' ------------------------------------
    Dim bTopFlangeLeft As Boolean
    Dim bBottomFlangeLeft As Boolean
    Dim bTopFlangeRight As Boolean
    Dim bBottomFlangeRight As Boolean
    
    CrossSection_Flanges oSDO_WebCut.Bounded, bTopFlangeLeft, bBottomFlangeLeft, bTopFlangeRight, bBottomFlangeRight
    
    ' -------------------------------------------------------------------------------------------------------------
    ' If the top intersects or is colinear with the edge then we want a connection by default if there is no flange
    ' -------------------------------------------------------------------------------------------------------------
    ' If there is a flange, the flange cut should create the PC
    If (eTopOrWL.IntersectingEdge = Top_Flange_Right_Bottom Or eTopOrWL.CoplanarEdge = Top_Flange_Right_Bottom) And Not _
       (bTopFlangeLeft Or bTopFlangeRight) Then
        ComputeConnectTopInsideAnswer = "Yes"
    End If
    
    ' ----------------------------
    ' If the top is above the edge
    ' ----------------------------
    Select Case eTopOrWL.IntersectingEdge
        Case Above, Top, Web_Right_Top, Top_Flange_Right_Top, Top_Flange_Right
            ' ----------------------------------------------------------------------
            ' If the bottom is at or below the edge, we want a connection by default
            ' ----------------------------------------------------------------------
            Select Case eBottomOrWR.IntersectingEdge
                Case Above, Top, Web_Right_Top, Top_Flange_Right_Top, Top_Flange_Right
                Case Else
                    ' -------------------------------------------------------------------------------
                    ' Verify that the reference collection contains valid supporting/bounding objects
                    ' -------------------------------------------------------------------------------
                    Dim oReferencesCollection As IJDReferencesCollection
                    Dim oEditJDArgument As IJDEditJDArgument
                    Dim oBoundingObjectColl As IJElements
                    
                    'Get Bounding Object Collection Count
                    Set oReferencesCollection = GetRefCollFromSmartOccurrence(oAppConnection)
                    Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument
                    Set oBoundingObjectColl = GetBoundingObjectsFromPorts(oEditJDArgument)
                        
                    'Check the Number of Boundings
                    If oBoundingObjectColl.Count > 1 Then
                        Dim oNearestBounding As Object
                        'Get Nearest Bounding to Top Flange Right
                        Set oNearestBounding = GetNearestBoundingToPort(oBoundingObjectColl, oEditJDArgument, oAppConnection, _
                                    GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounded, JXSEC_TOP_FLANGE_RIGHT))
                        'Check if Bounding is plate
                        If TypeOf oNearestBounding Is IJPlate Then
                            Dim bHasintersection As Boolean
                            Dim oHelper As GSCADShipGeomOps.SGOModelBodyUtilities
                            Set oHelper = New GSCADShipGeomOps.SGOModelBodyUtilities
                            
                            'Check the Intesection BTW Plate and Bounding Web
                            bHasintersection = oHelper.HasIntersectingGeometry(oNearestBounding, _
                            GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounding, JXSEC_WEB_LEFT))
                            
                            'Exit with No, if Bounding Web is Intersecting with Plate
                            If bHasintersection Then
                                ComputeConnectTopInsideAnswer = "No"
                                Exit Function
                            Else
                            'Exit with No, if the inside of member flange is collinear with plate
                                Dim lSectionAlias As Long
                                Dim bPenetratesWeb As Boolean
                                Dim oEdgeMapColl As New Collection
                                
                                'Get the port id of the member that is connected to bounded member inside top flange
                                Set oEdgeMapColl = GetEdgeMap(pSL.SmartOccurrence, oSDO_WebCut.BoundingPort, oSDO_WebCut.BoundedPort, lSectionAlias, bPenetratesWeb)
                                Dim eSubPortID As JXSEC_CODE
                                eSubPortID = ReverseMap(JXSEC_TOP_FLANGE_RIGHT_BOTTOM, oEdgeMapColl)
                                
                                Dim oBndingPort As IJPort
                                Dim bIntersectingPort As Boolean
                                Dim oSGOModelBodyUtils As GSCADShipGeomOps.SGOModelBodyUtilities
                                
                                Set oSGOModelBodyUtils = New GSCADShipGeomOps.SGOModelBodyUtilities
                                Set oBndingPort = GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounding, eSubPortID)
                                
                                'Check whether the inside port of the member is intersecting with plate part.
                                'If the inside port of the member is collinear with the plate then the conditional is set to No
                                'Because there exist no material  on member flange to create the physical connection.
                                bIntersectingPort = oSGOModelBodyUtils.HasIntersectingGeometry(oBndingPort, oNearestBounding)
								
                                If bIntersectingPort Then
                                    ComputeConnectTopInsideAnswer = "No"
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                ComputeConnectTopInsideAnswer = "Yes"
            End Select
        Case Else
    
    End Select
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
 
End Function

'*********************************************************************************************
' Method      : ComputeConnectBottomInsideAnswer
' Description :
'
'*********************************************************************************************
Public Function ComputeConnectBottomInsideAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod) As String
    Const METHOD As String = "::ComputeConnectBottomInsideAnswer"
    On Error GoTo ErrorHandler
    
    ComputeConnectBottomInsideAnswer = "No"

    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)
    
    ' ----------------------------------------------------
    ' If not bounded by a member, then the default is "no"
    ' ----------------------------------------------------
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pSL.SmartOccurrence
    
    If (Not TypeOf oSDO_WebCut.Bounding Is ISPSMemberPartPrismatic) And (Not TypeOf oSDO_WebCut.Bounding Is IJProfile) Then
        Exit Function
    End If
    
    ' -----------------------------------------------------
    ' If there is a bottom cut by default, no PC by default
    ' -----------------------------------------------------
    Dim strBottomCutByDefault As String
    strBottomCutByDefault = ComputeBottomWebCutAnswer(pSL.SmartOccurrence)
    
    If strBottomCutByDefault = "Yes" Then
        Exit Function
    End If
    
    ' ---------------------------------------------------------------------
    ' Find out where on the bounding member each bounded edge makes contact
    ' ---------------------------------------------------------------------
    Dim oAppConnection As Object
    AssemblyConnection_SmartItemName pSL.SmartOccurrence, , oAppConnection
    
    Dim eTopOrWL As ConnectedEdgeInfo
    Dim eBottomOrWR As ConnectedEdgeInfo
    Dim eInsideTFOrTFL As ConnectedEdgeInfo
    Dim eInsideBFOrTFR As ConnectedEdgeInfo
        
    GetConnectedEdgeInfo pSL.SmartOccurrence, oSDO_WebCut.BoundedPort, oSDO_WebCut.BoundingPort, eTopOrWL, eBottomOrWR, eInsideTFOrTFL, eInsideBFOrTFR

    ' ------------------------------------
    ' Determine what bounded flanges exist
    ' ------------------------------------
    Dim bTopFlangeLeft As Boolean
    Dim bBottomFlangeLeft As Boolean
    Dim bTopFlangeRight As Boolean
    Dim bBottomFlangeRight As Boolean
    
    CrossSection_Flanges oSDO_WebCut.Bounded, bTopFlangeLeft, bBottomFlangeLeft, bTopFlangeRight, bBottomFlangeRight
 
    ' ----------------------------------------------------------------------------------------------------------------
    ' If the bottom intersects or is colinear with the edge then we want a connection by default if there is no flange
    ' ----------------------------------------------------------------------------------------------------------------
    ' If there is a flange, the flange cut should create the PC
    If (eBottomOrWR.IntersectingEdge = Bottom_Flange_Right_Bottom Or eBottomOrWR.CoplanarEdge = Bottom_Flange_Right_Bottom) And Not _
       (bBottomFlangeLeft Or bBottomFlangeRight) Then
        ComputeConnectBottomInsideAnswer = "Yes"
    End If
    
    ' -------------------------------
    ' If the bottom is below the edge
    ' -------------------------------
    Select Case eBottomOrWR.IntersectingEdge
        Case Below, Bottom, Web_Right_Bottom, Bottom_Flange_Right_Bottom, Bottom_Flange_Right
            ' -------------------------------------------------------------------
            ' If the top is at or above the edge, we want a connection by default
            ' -------------------------------------------------------------------
            Select Case eTopOrWL.IntersectingEdge
                Case Below, Bottom, Web_Right_Bottom, Bottom_Flange_Right_Bottom, Bottom_Flange_Right
                Case Else
                    ' -------------------------------------------------------------------------------
                    ' Verify that the reference collection contains valid supporting/bounding objects
                    ' -------------------------------------------------------------------------------
                    Dim oReferencesCollection As IJDReferencesCollection
                    Dim oEditJDArgument As IJDEditJDArgument
                    Dim oBoundingObjectColl As IJElements
                    
                    'Get Bounding Object Collection Count
                    Set oReferencesCollection = GetRefCollFromSmartOccurrence(oAppConnection)
                    Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument
                    Set oBoundingObjectColl = GetBoundingObjectsFromPorts(oEditJDArgument)
                        
                    'Check the Number of Boundings
                    If oBoundingObjectColl.Count > 1 Then
                        Dim oNearestBounding As Object
                        'Get Nearest Bounding to Bottom Flange Right
                        Set oNearestBounding = GetNearestBoundingToPort(oBoundingObjectColl, oEditJDArgument, oAppConnection, _
                                    GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounded, JXSEC_BOTTOM_FLANGE_RIGHT))
                        'Check if Bounding is plate
                        If TypeOf oNearestBounding Is IJPlate Then
                            Dim bHasintersection As Boolean
                            Dim oHelper As GSCADShipGeomOps.SGOModelBodyUtilities
                            Set oHelper = New GSCADShipGeomOps.SGOModelBodyUtilities
                            
                            'Check the Intesection BTW Plate and Bounding Web
                            bHasintersection = oHelper.HasIntersectingGeometry(oNearestBounding, _
                            GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounding, JXSEC_WEB_LEFT))
                            
                            'Exit with No, if Bounding Web is Intersecting with Plate
                            If bHasintersection Then
                                ComputeConnectBottomInsideAnswer = "No"
                                Exit Function
                            Else
                            'Exit with No, if the inside of member flange is collinear with plate
                                Dim lSectionAlias As Long
                                Dim bPenetratesWeb As Boolean
                                Dim oEdgeMapColl As New Collection
                                
                                 'Get the port id of the member that is connected to bounded member inside bottom flange
                                Set oEdgeMapColl = GetEdgeMap(pSL.SmartOccurrence, oSDO_WebCut.BoundingPort, oSDO_WebCut.BoundedPort, lSectionAlias, bPenetratesWeb)
                                Dim eSubPortID As JXSEC_CODE
                                eSubPortID = ReverseMap(JXSEC_BOTTOM_FLANGE_RIGHT_TOP, oEdgeMapColl)
                                
                                Dim oBndingPort As IJPort
                                Dim bIntersectingPort As Boolean
                                Dim oSGOModelBodyUtils As GSCADShipGeomOps.SGOModelBodyUtilities
                                
                                Set oSGOModelBodyUtils = New GSCADShipGeomOps.SGOModelBodyUtilities
                                Set oBndingPort = GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounding, eSubPortID)
                                
                                'Check whether the inside port of the member is intersecting with plate part.
                                'If the inside port of the member is collinear with the plate then the conditional is set to No
                                'Because there exist no material  on member flange to create the physical connection.
                                bIntersectingPort = oSGOModelBodyUtils.HasIntersectingGeometry(oBndingPort, oNearestBounding)
                                
                                If bIntersectingPort Then
                                    ComputeConnectBottomInsideAnswer = "No"
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                    ComputeConnectBottomInsideAnswer = "Yes"
            End Select
        Case Else
    End Select
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
 
End Function

'*********************************************************************************************
' Method      : ComputeConnectTopEdgeAnswer
' Description :
'
'*********************************************************************************************
Public Function ComputeConnectTopEdgeAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod) As String
    Const METHOD As String = "::ComputeConnectTopEdgeAnswer"
    On Error GoTo ErrorHandler
    
    ComputeConnectTopEdgeAnswer = "No"
    
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)
    
    ' ----------------------------------------------------
    ' If not bounded by a member, then the default is "no"
    ' ----------------------------------------------------
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pSL.SmartOccurrence
    
    If (Not TypeOf oSDO_WebCut.Bounding Is ISPSMemberPartPrismatic) And (Not TypeOf oSDO_WebCut.Bounding Is IJProfile) Then
        Exit Function
    End If
    
    ' --------------------------------------------------
    ' If there is a top cut by default, no PC by default
    ' --------------------------------------------------
    Dim strTopCutByDefault As String
    strTopCutByDefault = ComputeTopWebCutAnswer(pSL.SmartOccurrence)
    
    If strTopCutByDefault = "Yes" Then
        Exit Function
    End If
    
    ' ---------------------------------------------------------------------
    ' Find out where on the bounding member each bounded edge makes contact
    ' ---------------------------------------------------------------------
    Dim oAppConnection As Object
    AssemblyConnection_SmartItemName pSL.SmartOccurrence, , oAppConnection
    
    Dim eTopOrWL As ConnectedEdgeInfo
    Dim eBottomOrWR As ConnectedEdgeInfo
    Dim eInsideTFOrTFL As ConnectedEdgeInfo
    Dim eInsideBFOrTFR As ConnectedEdgeInfo
        
    GetConnectedEdgeInfo pSL.SmartOccurrence, oSDO_WebCut.BoundedPort, oSDO_WebCut.BoundingPort, eTopOrWL, eBottomOrWR, eInsideTFOrTFL, eInsideBFOrTFR

    ' ------------------------------------
    ' Determine what bounded flanges exist
    ' ------------------------------------
    Dim bTopFlangeLeft As Boolean
    Dim bBottomFlangeLeft As Boolean
    Dim bTopFlangeRight As Boolean
    Dim bBottomFlangeRight As Boolean
    
    CrossSection_Flanges oSDO_WebCut.Bounded, bTopFlangeLeft, bBottomFlangeLeft, bTopFlangeRight, bBottomFlangeRight
    
    ' -----------------------------------------------------------------------------------------
    ' If the top intersects the edge then we want a connection by default if there is no flange
    ' -----------------------------------------------------------------------------------------
    ' If there is a flange, the flange cut should create the PC
    If (eTopOrWL.IntersectingEdge = Top_Flange_Right And Not bTopFlangeLeft And Not bTopFlangeRight) Then
        ComputeConnectTopEdgeAnswer = "Yes"
    End If
    
    ' ----------------------------
    ' If the top is above the edge
    ' ----------------------------
    Select Case eTopOrWL.IntersectingEdge
        Case Above, Top, Web_Right_Top, Top_Flange_Right_Top
            ' ------------------------------------------------------------------------------------------------------------------
            ' If the inside of the top flange intersects here, presume the remaining overlap will be removed by a corner feature
            ' ------------------------------------------------------------------------------------------------------------------
            If ((bTopFlangeLeft Or bTopFlangeRight) And (eInsideTFOrTFL.IntersectingEdge = Top_Flange_Right)) Then
                ' do nothing, default is "no"
            ' -------------------------------------------------------------------------------------
            ' Otherwise, if the bottom is at or below the top edge, we want a connection by default
            ' -------------------------------------------------------------------------------------
            Else
                Select Case eInsideBFOrTFR.IntersectingEdge
                    Case Above, Top, Web_Right_Top, Top_Flange_Right_Top
                    Case Else
                        ComputeConnectTopEdgeAnswer = "Yes"
                End Select
            End If
        Case Else
    End Select
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
     
End Function

'*********************************************************************************************
' Method      : ComputeConnectBottomEdgeAnswer
' Description :
'
'*********************************************************************************************
Public Function ComputeConnectBottomEdgeAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod) As String
    Const METHOD As String = "::ComputeConnectBottomEdgeAnswer"
    On Error GoTo ErrorHandler
    
    ComputeConnectBottomEdgeAnswer = "No"
    
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)
    
    ' ----------------------------------------------------
    ' If not bounded by a member, then the default is "no"
    ' ----------------------------------------------------
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pSL.SmartOccurrence
    
    If (Not TypeOf oSDO_WebCut.Bounding Is ISPSMemberPartPrismatic) And (Not TypeOf oSDO_WebCut.Bounding Is IJProfile) Then
        Exit Function
    End If
    
    ' --------------------------------------------------
    ' If there is a bottom cut by default, no PC by default
    ' --------------------------------------------------
    Dim strBottomCutByDefault As String
    strBottomCutByDefault = ComputeBottomWebCutAnswer(pSL.SmartOccurrence)
    
    If strBottomCutByDefault = "Yes" Then
        Exit Function
    End If
    
    ' ---------------------------------------------------------------------
    ' Find out where on the bounding member each bounded edge makes contact
    ' ---------------------------------------------------------------------
    Dim oAppConnection As Object
    AssemblyConnection_SmartItemName pSL.SmartOccurrence, , oAppConnection
    
    Dim eTopOrWL As ConnectedEdgeInfo
    Dim eBottomOrWR As ConnectedEdgeInfo
    Dim eInsideTFOrTFL As ConnectedEdgeInfo
    Dim eInsideBFOrTFR As ConnectedEdgeInfo
        
    GetConnectedEdgeInfo pSL.SmartOccurrence, oSDO_WebCut.BoundedPort, oSDO_WebCut.BoundingPort, eTopOrWL, eBottomOrWR, eInsideTFOrTFL, eInsideBFOrTFR

    ' ------------------------------------
    ' Determine what bounded flanges exist
    ' ------------------------------------
    Dim bTopFlangeLeft As Boolean
    Dim bBottomFlangeLeft As Boolean
    Dim bTopFlangeRight As Boolean
    Dim bBottomFlangeRight As Boolean
    
    CrossSection_Flanges oSDO_WebCut.Bounded, bTopFlangeLeft, bBottomFlangeLeft, bTopFlangeRight, bBottomFlangeRight
    
    ' --------------------------------------------------------------------------------------------
    ' If the bottom intersects the edge then we want a connection by default if there is no flange
    ' --------------------------------------------------------------------------------------------
    ' If there is a flange, the flange cut should create the PC
    If (eBottomOrWR.IntersectingEdge = Bottom_Flange_Right And Not bBottomFlangeLeft And Not bBottomFlangeRight) Then
        ComputeConnectBottomEdgeAnswer = "Yes"
    End If
    
    ' -------------------------------
    ' If the bottom is below the edge
    ' -------------------------------
    Select Case eBottomOrWR.IntersectingEdge
        Case Below, Bottom, Web_Right_Bottom, Bottom_Flange_Right_Bottom
            ' ---------------------------------------------------------------------------------------------------------------------
            ' If the inside of the bottom flange intersects here, presume the remaining overlap will be removed by a corner feature
            ' ---------------------------------------------------------------------------------------------------------------------
            If (bBottomFlangeLeft Or bBottomFlangeRight) And (eInsideBFOrTFR.IntersectingEdge = Bottom_Flange_Right) Then
                ' do nothing, default is "no"
            ' ----------------------------------------------------------------------------------
            ' Otherwise, if the top is at or above the top edge, we want a connection by default
            ' ----------------------------------------------------------------------------------
            Else
                Select Case eInsideTFOrTFL.IntersectingEdge
                    Case Below, Bottom, Web_Right_Bottom, Bottom_Flange_Right_Bottom
                    Case Else
                        ComputeConnectBottomEdgeAnswer = "Yes"
                End Select
            End If
        Case Else
    End Select
 
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
   
End Function

'*********************************************************************************************
' Method      : ComputeConnectFaceAnswer
' Description :
'
'*********************************************************************************************
Public Function ComputeConnectFaceAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod) As String
    Const METHOD As String = "::ComputeConnectFaceAnswer"
    On Error GoTo ErrorHandler
                                      
    ComputeConnectFaceAnswer = "No"
                                  
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)
    
    ' -----------------------------------------------------
    ' Special-case generic ACs predefine the desired answer
    ' -----------------------------------------------------
    ' Set the values, even if the default would match what is desired by the AC
    ' 1) to guard against changes in the default logic affecting the special AC
    ' 2) to avoid recomputing a lot of information, such as where the bounding flange intersects
    Dim sACItemName As String
    Dim oAppConnection As Object
    AssemblyConnection_SmartItemName pSL.SmartOccurrence, sACItemName, oAppConnection
    
    If sACItemName = "WPEF_PF_1" Or sACItemName = "Generic_Stiff_WPEF_PF_1" Then
        ComputeConnectFaceAnswer = "Yes"
        Exit Function
    End If
    
    ' ----------------------------------------------------
    ' If not bounded by a member, then the default is "no"
    ' ----------------------------------------------------
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pSL.SmartOccurrence
    
    If (Not TypeOf oSDO_WebCut.Bounding Is ISPSMemberPartPrismatic) And (Not TypeOf oSDO_WebCut.Bounding Is IJProfile) Then
        Exit Function
    End If
    
    ' ---------------------------------------------------------------------
    ' Find out where on the bounding member each bounded edge makes contact
    ' ---------------------------------------------------------------------
    Dim eTopOrWL As ConnectedEdgeInfo
    Dim eBottomOrWR As ConnectedEdgeInfo
    Dim eInsideTFOrTFL As ConnectedEdgeInfo
    Dim eInsideBFOrTFR As ConnectedEdgeInfo
        
    GetConnectedEdgeInfo pSL.SmartOccurrence, oSDO_WebCut.BoundedPort, oSDO_WebCut.BoundingPort, eTopOrWL, eBottomOrWR, eInsideTFOrTFL, eInsideBFOrTFR

    ' -------------------------------------------------------------------
    ' If the top intersects the edge then we want a connection by default
    ' -------------------------------------------------------------------
    If (eTopOrWL.IntersectingEdge = Web_Right Or eBottomOrWR.IntersectingEdge = Web_Right) Then
        ComputeConnectFaceAnswer = "Yes"
    End If
    
    ' ----------------------------
    ' If the top is above the face
    ' ----------------------------
    Select Case eTopOrWL.IntersectingEdge
        Case Above, Top, Web_Right_Top, Top_Flange_Right_Top, Top_Flange_Right, Top_Flange_Right_Bottom
            ' ----------------------------------------------------------------------
            ' If the bottom is at or below the face, we want a connection by default
            ' ----------------------------------------------------------------------
            Select Case eBottomOrWR.IntersectingEdge
                Case Above, Top, Web_Right_Top, Top_Flange_Right_Top, Top_Flange_Right, Top_Flange_Right_Bottom
                Case Else
                    ComputeConnectFaceAnswer = "Yes"
            End Select
        Case Else
    End Select
    
    ' -------------------------------
    ' If the bottom is below the face
    ' -------------------------------
    Select Case eBottomOrWR.IntersectingEdge
        Case Below, Bottom, Web_Right_Bottom, Bottom_Flange_Right_Bottom, Bottom_Flange_Right, Bottom_Flange_Right_Top
            ' -------------------------------------------------------------------
            ' If the top is at or above the face, we want a connection by default
            ' -------------------------------------------------------------------
            Select Case eTopOrWL.IntersectingEdge
                Case Below, Bottom, Web_Right_Bottom, Bottom_Flange_Right_Bottom, Bottom_Flange_Right, Bottom_Flange_Right_Top
                Case Else
                    ComputeConnectFaceAnswer = "Yes"
            End Select
        Case Else
    End Select
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
     
End Function

'*********************************************************************************************
' Method      : ComputeTopWebCutAnswer
' Description :
'
'*********************************************************************************************
Public Function ComputeTopWebCutAnswer(oCenterWebCut As IJSmartOccurrence) As String
    Const METHOD As String = "::ComputeTopWebCutAnswer"
    On Error GoTo ErrorHandler
   
    ComputeTopWebCutAnswer = "No"
    
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = oCenterWebCut
    
    If (Not TypeOf oSDO_WebCut.Bounding Is ISPSMemberPartPrismatic) And (Not TypeOf oSDO_WebCut.Bounding Is IJProfile) Then
        Exit Function
    End If
    
    ' -----------------------------------------------------
    ' Special-case generic ACs predefine the desired answer
    ' -----------------------------------------------------
    ' Set the values, even if the default would match what is desired by the AC
    ' 1) to guard against changes in the default logic affecting the special AC
    ' 2) to avoid recomputing a lot of information, such as where the bounding flange intersects
    Dim sACItemName As String
    Dim oACObject As Object
    AssemblyConnection_SmartItemName oCenterWebCut, sACItemName, oACObject
    
    If sACItemName = "WPEF_PF_1" Or sACItemName = "Generic_Stiff_WPEF_PF_1" Then
        ComputeTopWebCutAnswer = "No"
        Exit Function
    End If
    
    ' -------------------------------------
    ' Find out the number of bounding objects
    ' -------------------------------------
    Dim itemCount As Long
    itemCount = 1

    Dim oEditJDArgument As IJDEditJDArgument
    Dim oReferencesCollection As IJDReferencesCollection

    Set oReferencesCollection = GetRefCollFromSmartOccurrence(oACObject)
    Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument

    Dim oBoundingObjectColl As IJElements
    Set oBoundingObjectColl = GetBoundingObjectsFromPorts(oEditJDArgument)
    
    itemCount = oBoundingObjectColl.Count
    
    'consider cases of auxiliary ports
    If TypeOf oACObject Is IJAppConnection Then
        If GetMbrAssemblyConnectionType(oACObject) = eACType.ACType_Stiff_Generic Then
            ComputeTopWebCutAnswer = "No"
            Exit Function
        End If
    End If
    
'    ' -------------------------------------
'    ' Find out how many cuts the AC created
'    ' -------------------------------------
'    Dim itemCount As Long
'    itemCount = 1
'
'    If TypeOf oACObject Is IJDMemberObjects Then
'        Dim oMemObjs As IJDMemberObjects
'        Set oMemObjs = oACObject
'        itemCount = oMemObjs.Count
'    Else
'        ' MsgBox "no support"
'    End If

    ' -----------------------------------------------------------------------------------------------
    ' If only one bounding object, and it is a member, find out where each bounded edge makes contact
    ' -----------------------------------------------------------------------------------------------
    If itemCount = 1 And (TypeOf oSDO_WebCut.Bounding Is ISPSMemberPartPrismatic Or TypeOf oSDO_WebCut.Bounding Is IJProfile) Then
        'to support Generic ACs when tube is Bounding
        If IsTubularMember(oSDO_WebCut.Bounding) Then
            
            Dim bPenetratesWeb As Boolean
            Dim distance1 As Double
            Dim distance2 As Double
            Dim bTopFlangeLeft As Boolean
            Dim bBottomFlangeLeft As Boolean
            Dim bTopFlangeRight As Boolean
            Dim bBottomFlangeRight As Boolean
            bPenetratesWeb = IsWebPenetrated(oSDO_WebCut.BoundingPort, oSDO_WebCut.BoundedPort)
            CrossSection_Flanges oSDO_WebCut.Bounded, bTopFlangeLeft, bBottomFlangeLeft, bTopFlangeRight, bBottomFlangeRight
            If bPenetratesWeb Then
                If bTopFlangeRight Then
                    distance1 = GetDistanceFromBounding(oSDO_WebCut.Bounding, oSDO_WebCut.Bounded, JXSEC_TOP_FLANGE_RIGHT_BOTTOM)
                    distance2 = GetDistanceFromBounding(oSDO_WebCut.Bounding, oSDO_WebCut.Bounded, JXSEC_TOP)
                    If distance1 > 0.005 And distance2 > 0.005 Then
                        ComputeTopWebCutAnswer = "Yes"
                    Else
                        ComputeTopWebCutAnswer = "No"
                    End If
                End If
            Else
               ComputeTopWebCutAnswer = "No"
            End If
        Else
            Dim eTopOrWL As ConnectedEdgeInfo
            Dim eBottomOrWR As ConnectedEdgeInfo
            Dim eInsideTFOrTFL As ConnectedEdgeInfo
            Dim eInsideBFOrTFR As ConnectedEdgeInfo
                
            GetConnectedEdgeInfo oCenterWebCut, oSDO_WebCut.BoundedPort, oSDO_WebCut.BoundingPort, eTopOrWL, eBottomOrWR, eInsideTFOrTFL, eInsideBFOrTFR
            
            ' --------------------------------------------------------------------------------------------------------------------------------------
            ' If it is above the top edge, including the top of the flange or the portion of the web above the flange, then add a top cut by default
            ' --------------------------------------------------------------------------------------------------------------------------------------
            Select Case eTopOrWL.IntersectingEdge
                Case Above, Top, Top_Flange_Right_Top, Web_Right_Top
                    ComputeTopWebCutAnswer = "Yes"
            End Select
        End If
    ' -------------------------------------------------------------------------------------------------------
    ' If there is more than one bounding object, assume another object will trim any portion outside this cut
    ' -------------------------------------------------------------------------------------------------------
    Else
        ComputeTopWebCutAnswer = "No"
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
     
End Function



'*********************************************************************************************
' Method      : ComputeBottomWebCutAnswer
' Description :
'
'*********************************************************************************************
Public Function ComputeBottomWebCutAnswer(oCenterWebCut As IJSmartOccurrence) As String
    Const METHOD As String = "::ComputeBottomWebCutAnswer"
    On Error GoTo ErrorHandler
       
    ComputeBottomWebCutAnswer = "No"
    
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = oCenterWebCut
    
    If (Not TypeOf oSDO_WebCut.Bounding Is ISPSMemberPartPrismatic) And (Not TypeOf oSDO_WebCut.Bounding Is IJProfile) Then
        Exit Function
    End If
    
    ' -----------------------------------------------------
    ' Special-case generic ACs predefine the desired answer
    ' -----------------------------------------------------
    ' Set the values, even if the default would match what is desired by the AC
    ' 1) to guard against changes in the default logic affecting the special AC
    ' 2) to avoid recomputing a lot of information, such as where the bounding flange intersects
    Dim sACItemName As String
    Dim oACObject As Object
    AssemblyConnection_SmartItemName oCenterWebCut, sACItemName, oACObject
    
    If sACItemName = "WPEF_PF_1" Or sACItemName = "Generic_Stiff_WPEF_PF_1" Then
        ComputeBottomWebCutAnswer = "No"
        Exit Function
    End If
    
    ' -------------------------------------
    ' Find out the number of bounding objects
    ' -------------------------------------
    Dim itemCount As Long
    itemCount = 1

    Dim oEditJDArgument As IJDEditJDArgument
    Dim oReferencesCollection As IJDReferencesCollection

    Set oReferencesCollection = GetRefCollFromSmartOccurrence(oACObject)
    Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument

    Dim oBoundingObjectColl As IJElements
    Set oBoundingObjectColl = GetBoundingObjectsFromPorts(oEditJDArgument)
    
    itemCount = oBoundingObjectColl.Count
    
    'consider cases of auxiliary ports
    If TypeOf oACObject Is IJAppConnection Then
        If GetMbrAssemblyConnectionType(oACObject) = eACType.ACType_Stiff_Generic Then
            ComputeBottomWebCutAnswer = "No"
            Exit Function
        End If
    End If
    
'    ' -------------------------------------
'    ' Find out how many cuts the AC created
'    ' -------------------------------------
'    Dim itemCount As Long
'    itemCount = 1
'
'    If TypeOf oACObject Is IJDMemberObjects Then
'        Dim oMemObjs As IJDMemberObjects
'        Set oMemObjs = oACObject
'        itemCount = oMemObjs.Count
'    Else
'        ' MsgBox "no support"
'    End If
    
   ' -----------------------------------------------------------------------------------------------
    ' If only one bounding object, and it is a member or stiffner , find out where each bounded edge makes contact
    ' -----------------------------------------------------------------------------------------------
    If itemCount = 1 And (TypeOf oSDO_WebCut.Bounding Is ISPSMemberPartPrismatic Or TypeOf oSDO_WebCut.Bounding Is IJProfile) Then
        If IsTubularMember(oSDO_WebCut.Bounding) Then
            Dim bPenetratesWeb As Boolean
            Dim distance1 As Double
            Dim distance2 As Double
            Dim bTopFlangeLeft As Boolean
            Dim bBottomFlangeLeft As Boolean
            Dim bTopFlangeRight As Boolean
            Dim bBottomFlangeRight As Boolean
            
            bPenetratesWeb = IsWebPenetrated(oSDO_WebCut.BoundingPort, oSDO_WebCut.BoundedPort)
            CrossSection_Flanges oSDO_WebCut.Bounded, bTopFlangeLeft, bBottomFlangeLeft, bTopFlangeRight, bBottomFlangeRight
            If bPenetratesWeb Then
                If bBottomFlangeRight Then
                    distance1 = GetDistanceFromBounding(oSDO_WebCut.Bounding, oSDO_WebCut.Bounded, JXSEC_BOTTOM_FLANGE_RIGHT_TOP)
                    distance2 = GetDistanceFromBounding(oSDO_WebCut.Bounding, oSDO_WebCut.Bounded, JXSEC_BOTTOM)
                    If distance1 > 0.005 And distance2 > 0.005 Then
                        ComputeBottomWebCutAnswer = "Yes"
                    Else
                        ComputeBottomWebCutAnswer = "No"
                    End If
                End If
            Else
                ComputeBottomWebCutAnswer = "No"
            End If
        Else
            Dim eTopOrWL As ConnectedEdgeInfo
            Dim eBottomOrWR As ConnectedEdgeInfo
            Dim eInsideTFOrTFL As ConnectedEdgeInfo
            Dim eInsideBFOrTFR As ConnectedEdgeInfo
                
            GetConnectedEdgeInfo oCenterWebCut, oSDO_WebCut.BoundedPort, oSDO_WebCut.BoundingPort, eTopOrWL, eBottomOrWR, eInsideTFOrTFL, eInsideBFOrTFR
            
            ' -----------------------------------------------------------------------------------------------------------------------------------------------
            ' If it is below the bottom edge, including the bottom of the flange or the portion of the web below the flange, then add a bottom cut by default
            ' -----------------------------------------------------------------------------------------------------------------------------------------------
            Select Case eBottomOrWR.IntersectingEdge
                Case Below, Bottom, Bottom_Flange_Right_Bottom, Web_Right_Bottom
                    ComputeBottomWebCutAnswer = "Yes"
            End Select
        End If
    ' -------------------------------------------------------------------------------------------------------
    ' If there is more than one bounding object, assume another object will trim any portion outside this cut
    ' -------------------------------------------------------------------------------------------------------
    Else
        ComputeBottomWebCutAnswer = "No"
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
 
End Function




'*********************************************************************************************
' Method      : ComputeFaceTopInsideCornerAnswer
' Description :
'
'*********************************************************************************************
Public Function ComputeFaceTopInsideCornerAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod)
    Const METHOD As String = "::ComputeFaceTopInsideCornerAnswer"
    On Error GoTo ErrorHandler
  
    ComputeFaceTopInsideCornerAnswer = "No"

    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)
    
    ' ----------------------------------------------------
    ' If not bounded by a member, then the default is "no"
    ' ----------------------------------------------------
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pSL.SmartOccurrence
    
    If (Not TypeOf oSDO_WebCut.Bounding Is ISPSMemberPartPrismatic) And (Not TypeOf oSDO_WebCut.Bounding Is IJProfile) Then
        pSL.answer("FaceTopInsideCorner") = "No"
        Exit Function
    End If
    
    ' ---------------------------------------------------------------------------------
    ' If would be connected to both face and edge by default, place a corner by default
    ' ---------------------------------------------------------------------------------
    Dim sConnectToFace As String
    Dim sConnectToInside As String
            
    sConnectToFace = ComputeConnectFaceAnswer(pInput)
    sConnectToInside = ComputeConnectTopInsideAnswer(pInput)
            
    If sConnectToFace = "Yes" And sConnectToInside = "Yes" Then
        'ComputeFaceTopInsideCornerAnswer = "Yes"
    End If

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
 
End Function

'*********************************************************************************************
' Method      : ComputeFaceBottomInsideCornerAnswer
' Description :
'
'*********************************************************************************************
Public Function ComputeFaceBottomInsideCornerAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod)
    Const METHOD As String = "::ComputeFaceBottomInsideCornerAnswer"
    On Error GoTo ErrorHandler
  
    ComputeFaceBottomInsideCornerAnswer = "No"

    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)
    
    ' ----------------------------------------------------
    ' If not bounded by a member, then the default is "no"
    ' ----------------------------------------------------
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pSL.SmartOccurrence
    
    If (Not TypeOf oSDO_WebCut.Bounding Is ISPSMemberPartPrismatic) And (Not TypeOf oSDO_WebCut.Bounding Is IJProfile) Then
        pSL.answer("FaceBottomInsideCorner") = "No"
        Exit Function
    End If
    
    ' ---------------------------------------------------------------------------------
    ' If would be connected to both face and edge by default, place a corner by default
    ' ---------------------------------------------------------------------------------
    Dim sConnectToFace As String
    Dim sConnectToInside As String
            
    sConnectToFace = ComputeConnectFaceAnswer(pInput)
    sConnectToInside = ComputeConnectBottomInsideAnswer(pInput)
            
    If sConnectToFace = "Yes" And sConnectToInside = "Yes" Then
        'ComputeFaceBottomInsideCornerAnswer = "Yes"
    End If

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
 
End Function

'*********************************************************************************************
' Method      : CMSetConnectTopAnswer
' Description : Sets the answer to the question
'
'*********************************************************************************************
Public Sub CMSetConnectTubeAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "::CMSetConnectTopAnswer"
    On Error GoTo ErrorHandler
    
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)
    
    pSL.answer("ConnectTube") = "No"
  
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
      
End Sub
'*********************************************************************************************
' Method      : ComputeTopFlangeCutAnswer
' Description :
'
'*********************************************************************************************
Public Function ComputeTopFlangeCutAnswer(oCenterFlangeCut As IJSmartOccurrence) As String
    Const METHOD As String = "::ComputeTopFlangeCutAnswer"
    On Error GoTo ErrorHandler
   
    ComputeTopFlangeCutAnswer = "No"
    
    Dim oSDO_FlangeCut As New StructDetailObjects.FlangeCut
    Set oSDO_FlangeCut.object = oCenterFlangeCut
    
    If (Not TypeOf oSDO_FlangeCut.Bounding Is ISPSMemberPartPrismatic) And (Not TypeOf oSDO_FlangeCut.Bounding Is IJProfile) Then
        Exit Function
    End If
    
    Dim sACItemName As String
    Dim oACObject As Object
    AssemblyConnection_SmartItemName oCenterFlangeCut, sACItemName, oACObject
    
    ' -------------------------------------
    ' Find out the number of bounding objects
    ' -------------------------------------
    Dim itemCount As Long
    itemCount = 1

    Dim oEditJDArgument As IJDEditJDArgument
    Dim oReferencesCollection As IJDReferencesCollection

    Set oReferencesCollection = GetRefCollFromSmartOccurrence(oACObject)
    Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument

    Dim oBoundingObjectColl As IJElements
    Set oBoundingObjectColl = GetBoundingObjectsFromPorts(oEditJDArgument)
    
    itemCount = oBoundingObjectColl.Count

    ' -----------------------------------------------------------------------------------------------
    ' If only one bounding object, and it is a member, find out where each bounded edge makes contact
    ' -----------------------------------------------------------------------------------------------
    If itemCount = 1 And (TypeOf oSDO_FlangeCut.Bounding Is ISPSMemberPartPrismatic Or TypeOf oSDO_FlangeCut.Bounding Is IJProfile) Then
        'to support Generic ACs when tube is Bounding
        If IsTubularMember(oSDO_FlangeCut.Bounding) Then
            
            Dim bPenetratesWeb As Boolean
            bPenetratesWeb = IsWebPenetrated(oSDO_FlangeCut.BoundingPort, oSDO_FlangeCut.BoundedPort)
            
            If bPenetratesWeb = False Then
                ComputeTopFlangeCutAnswer = "Yes"
            Else
                ComputeTopFlangeCutAnswer = "No"
            End If
        Else
            Dim eTopOrWL As ConnectedEdgeInfo
            Dim eBottomOrWR As ConnectedEdgeInfo
            Dim eInsideTFOrTFL As ConnectedEdgeInfo
            Dim eInsideBFOrTFR As ConnectedEdgeInfo
                
            GetConnectedEdgeInfo oCenterFlangeCut, oSDO_FlangeCut.BoundedPort, oSDO_FlangeCut.BoundingPort, eTopOrWL, eBottomOrWR, eInsideTFOrTFL, eInsideBFOrTFR
            
            ' --------------------------------------------------------------------------------------------------------------------------------------
            ' If it is above flange left edge, then add a top cut by default
            ' --------------------------------------------------------------------------------------------------------------------------------------
            Select Case eInsideTFOrTFL.IntersectingEdge
                Case Above, Top, Top_Flange_Right_Top, Web_Right_Top
                    ComputeTopFlangeCutAnswer = "Yes"
            End Select
        End If
    ' -------------------------------------------------------------------------------------------------------
    ' If there is more than one bounding object, assume another object will trim any portion outside this cut
    ' -------------------------------------------------------------------------------------------------------
    Else
        ComputeTopFlangeCutAnswer = "No"
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
     
End Function

'*********************************************************************************************
' Method      : ComputeBottomFlangeCutAnswer
' Description :
'
'*********************************************************************************************
Public Function ComputeBottomFlangeCutAnswer(oCenterFlangeCut As IJSmartOccurrence) As String
    Const METHOD As String = "::ComputeBottomFlangeCutAnswer"
    On Error GoTo ErrorHandler
       
    ComputeBottomFlangeCutAnswer = "No"
    
    Dim oSDO_FlangeCut As New StructDetailObjects.FlangeCut
    Set oSDO_FlangeCut.object = oCenterFlangeCut
    
    If (Not TypeOf oSDO_FlangeCut.Bounding Is ISPSMemberPartPrismatic) And (Not TypeOf oSDO_FlangeCut.Bounding Is IJProfile) Then
        Exit Function
    End If
    
    Dim sACItemName As String
    Dim oACObject As Object
    AssemblyConnection_SmartItemName oCenterFlangeCut, sACItemName, oACObject
    
    ' -------------------------------------
    ' Find out the number of bounding objects
    ' -------------------------------------
    Dim itemCount As Long
    itemCount = 1
    
    Dim oEditJDArgument As IJDEditJDArgument
    Dim oReferencesCollection As IJDReferencesCollection

    Set oReferencesCollection = GetRefCollFromSmartOccurrence(oACObject)
    Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument

    Dim oBoundingObjectColl As IJElements
    Set oBoundingObjectColl = GetBoundingObjectsFromPorts(oEditJDArgument)
    
    itemCount = oBoundingObjectColl.Count
    
   ' -----------------------------------------------------------------------------------------------
    ' If only one bounding object, and it is a member, find out where each bounded edge makes contact
    ' -----------------------------------------------------------------------------------------------
    If itemCount = 1 And (TypeOf oSDO_FlangeCut.Bounding Is ISPSMemberPartPrismatic Or TypeOf oSDO_FlangeCut.Bounding Is IJProfile) Then
        If IsTubularMember(oSDO_FlangeCut.Bounding) Then
            Dim bPenetratesWeb As Boolean
            bPenetratesWeb = IsWebPenetrated(oSDO_FlangeCut.BoundingPort, oSDO_FlangeCut.BoundedPort)
            
            If bPenetratesWeb = False Then
                ComputeBottomFlangeCutAnswer = "Yes"
            Else
                ComputeBottomFlangeCutAnswer = "No"
            End If
        Else
            Dim eTopOrWL As ConnectedEdgeInfo
            Dim eBottomOrWR As ConnectedEdgeInfo
            Dim eInsideTFOrTFL As ConnectedEdgeInfo
            Dim eInsideBFOrTFR As ConnectedEdgeInfo
                
            GetConnectedEdgeInfo oCenterFlangeCut, oSDO_FlangeCut.BoundedPort, oSDO_FlangeCut.BoundingPort, eTopOrWL, eBottomOrWR, eInsideTFOrTFL, eInsideBFOrTFR
            
            ' -----------------------------------------------------------------------------------------------------------------------------------------------
            ' If it is below the bottom edge, including the bottom of the flange or the portion of the web below the flange, then add a bottom cut by default
            ' -----------------------------------------------------------------------------------------------------------------------------------------------
            Select Case eInsideBFOrTFR.IntersectingEdge
                Case Below, Bottom, Bottom_Flange_Right_Bottom, Web_Right_Bottom
                    ComputeBottomFlangeCutAnswer = "Yes"
            End Select
        End If
    ' -------------------------------------------------------------------------------------------------------
    ' If there is more than one bounding object, assume another object will trim any portion outside this cut
    ' -------------------------------------------------------------------------------------------------------
    Else
        ComputeBottomFlangeCutAnswer = "No"
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
 
End Function

