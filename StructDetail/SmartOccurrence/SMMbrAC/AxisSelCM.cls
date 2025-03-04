VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AxisSelCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : AxisSelCM.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    08APR2011 - Created
'    25Aug2011 - Added new methods to support Inset brace (B-52232)
'    15Sep2011 - Modified/Added methods to support Braces on Flange penetrated cases
'    21Sep2011 - mpulikol
'           DI-CP-200263  Improve performance by caching measurement symbol results
'    1Dec2011  - Bug Fix done to get rid of warinings in Error log file.
'           Changes are done in CMSetShapeAtFaceAnswer, CMSetTopShapeAtFaceAnswer and CMSetBottomShapeAtFaceAnswer()
'*********************************************************************************************

Const MODULE = CUSTOMERID & "MbrAC.AxisSelCM:: "

'*********************************************************************************************
' Method      : CMSetTopFlangeOrLeftWebShapeAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetTopFlangeOrLeftWebShapeAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetTopFlangeOrLeftWebShapeAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
        
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim sMsg As String
    Dim lStatus As Long

    Set oAppConnection = oSelectorLogic.SmartOccurrence
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    ' ----------------------------------------
    ' Determine if web or flange is penetrated
    ' ----------------------------------------
    Dim bPenetratesWeb As Boolean
    
    sMsg = "Determining if web or flange is penetrated.."
    bPenetratesWeb = IsWebPenetrated(oBoundingData.AxisPort, oBoundedData.AxisPort)
        
    ' ------------------------------------------
    ' Determine the configuration of Flanges/Web
    ' ------------------------------------------
    Dim bTFL As Boolean
    Dim bBFL As Boolean
    Dim bTFR As Boolean
    Dim bBFR As Boolean
    
    sMsg = "Determining the configuration of Flanges/Web.."
    If bPenetratesWeb Then
        ' Consider the shape to be given on Flange
        CrossSection_Flanges oBoundedData.MemberPart, bTFL, bBFL, bTFR, bBFR
        
        If (bTFL) Or (bTFR) Then
            oSelectorLogic.Answer("TopFlangeOr(Left)Web") = gsStraight
        Else
            oSelectorLogic.Answer("TopFlangeOr(Left)Web") = gsNotApplied
        End If
        
    Else
        ' Consider the shape to be given on Web
        If IsTubularMember(oBoundedData.MemberPart) Then
            oSelectorLogic.Answer("TopFlangeOr(Left)Web") = gsNotApplied
        Else
            ' As of now there are no cases where we expect two Webs on cross section
            ' and hence no need to check about Right(or)Left Web but if exixt we
            ' might need to handle the case accordingly here
            oSelectorLogic.Answer("TopFlangeOr(Left)Web") = gsStraight
        End If
    End If
    
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : CMSetBottomFlangeOrRightWebShapeAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetBottomFlangeOrRightWebShapeAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetBottomFlangeOrRightWebShapeAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
        
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim sMsg As String
    Dim lStatus As Long

    Set oAppConnection = oSelectorLogic.SmartOccurrence
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    ' ----------------------------------------
    ' Determine if web or flange is penetrated
    ' ----------------------------------------
    Dim bPenetratesWeb As Boolean
    
    sMsg = "Determining if web or flange is penetrated.."
    bPenetratesWeb = IsWebPenetrated(oBoundingData.AxisPort, oBoundedData.AxisPort)
    
    ' ------------------------------------------
    ' Determine the configuration of Flanges/Web
    ' ------------------------------------------
    Dim bTFL As Boolean
    Dim bBFL As Boolean
    Dim bTFR As Boolean
    Dim bBFR As Boolean

    sMsg = "Determining the configuration of Flanges/Web.."
    If bPenetratesWeb Then
        'Consider the shape to be given on Flange
        CrossSection_Flanges oBoundedData.MemberPart, bTFL, bBFL, bTFR, bBFR
        
        If (bBFL) Or (bBFR) Then
            oSelectorLogic.Answer("BottomFlangeOrRightWeb") = gsStraight
        Else
            oSelectorLogic.Answer("BottomFlangeOrRightWeb") = gsNotApplied
        End If
        
    Else
        'Consider the shape to be given on Web
        If IsTubularMember(oBoundedData.MemberPart) Then
            oSelectorLogic.Answer("BottomFlangeOrRightWeb") = gsNotApplied
        Else
            'As of now there are no cases where we expect two Webs on cross section
            'and hence no need to check about Right(or)Left Web but if exixt we
            'might need to handle the case accordingly here
            oSelectorLogic.Answer("BottomFlangeOrRightWeb") = gsStraight
        End If
    End If
              
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError sMsg, METHOD

End Sub

'*********************************************************************************************
' Method      : CMSetExTendOrSetBackAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetExTendOrSetBackAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetExTendOrSetBackAnswer"
    On Error GoTo ErrorHandler
    
    Dim selString As String
    Dim oSMOcc As IJSmartOccurrence
    
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
    
    If TypeOf oSelectorLogic.SmartOccurrence Is IJSmartOccurrence Then
        Set oSMOcc = oSelectorLogic.SmartOccurrence
        selString = oSMOcc.RootSelection
    End If
       
    Select Case selString
        Case "MbrAxisByFCAndOS1Edge"
            oSelectorLogic.Answer("ExTendOrSetBack") = gsOffsetNearCorner
        Case Else
            oSelectorLogic.Answer("ExTendOrSetBack") = gsOffsetFarCorner
    End Select
    
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub

'*********************************************************************************************
' Method      : CMSetShapeAtFaceAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetShapeAtFaceAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetShapeAtFaceAnswer"
    On Error GoTo ErrorHandler
    
      ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
    
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim sMsg As String
    Dim lStatus As Long

    Set oAppConnection = oSelectorLogic.SmartOccurrence
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    Dim bPenetratesWeb As Boolean
   
    sMsg = "Determining if web or flange is penetrated.."
    bPenetratesWeb = IsWebPenetrated(oBoundingData.AxisPort, oBoundedData.AxisPort)

    Dim oTopORWL As ConnectedEdgeInfo
    Dim oBottomOrWR As ConnectedEdgeInfo
    Dim oTFL As ConnectedEdgeInfo
    Dim oTFR As ConnectedEdgeInfo
    
    GetConnectedEdgeInfo oAppConnection, oBoundedData.AxisPort, oBoundingData.AxisPort, oTopORWL, oBottomOrWR, oTFL, oTFR
    
    Dim bTFL As Boolean
    Dim bBFL As Boolean
    Dim bTFR As Boolean
    Dim bBFR As Boolean
    
    CrossSection_Flanges oBoundedData.MemberPart, bTFL, bBFL, bTFR, bBFR
                
    Dim oEdge1 As ConnectedEdgeInfo
    Dim oEdge2 As ConnectedEdgeInfo

    If bPenetratesWeb Then
        oEdge1.IntersectingEdge = oTopORWL.IntersectingEdge
        oEdge2.IntersectingEdge = oBottomOrWR.IntersectingEdge
    Else
        If (bTFL And bTFR) Or (bBFR And bBFL) Then
             oEdge1.IntersectingEdge = oTFL.IntersectingEdge
             oEdge2.IntersectingEdge = oTFR.IntersectingEdge
        ElseIf (bTFL Or bBFL) And Not bBFR And Not bTFR Then
            oEdge1.IntersectingEdge = oTFL.IntersectingEdge
            oEdge2.IntersectingEdge = oBottomOrWR.IntersectingEdge
        ElseIf (bTFR Or bBFR) And Not bTFL And Not bBFL Then
            oEdge1.IntersectingEdge = oTopORWL.IntersectingEdge
            oEdge2.IntersectingEdge = oTFR.IntersectingEdge
        ElseIf Not bTFL And Not bBFL And Not bTFR And Not bBFR Then
            oEdge1.IntersectingEdge = oTopORWL.IntersectingEdge
            oEdge2.IntersectingEdge = oBottomOrWR.IntersectingEdge
        Else
            '??
        End If
    End If
    
    Dim selString As String, bIsTopBrace As Boolean
    Dim oSMOcc As IJSmartOccurrence
    If TypeOf oSelectorLogic.SmartOccurrence Is IJSmartOccurrence Then
        Set oSMOcc = oSelectorLogic.SmartOccurrence
        selString = oSMOcc.RootSelection
    End If
       
    Dim sInsetBrace As String
    sInsetBrace = ""
    If bPenetratesWeb Then
        Select Case selString
            Case "MbrAxisByFCAndEdge", "MbrAxisByFCAndOS1Edge"
                'see if top brace is valid
                ' Determines if a Top/Bottom Brace is needed - A top brace is one which is connected to the Top of the Bounded
                bIsTopBrace = IsTopOrBottomInsetBrace(selString, oEdge1, oEdge2)
                If bIsTopBrace Then  'Top brace
                    sInsetBrace = oSelectorLogic.Answer("TopBraceType")
                Else
                    sInsetBrace = oSelectorLogic.Answer("BottomBraceType")
                End If
        End Select
    Else
        oSelectorLogic.Answer("ShapeAtFace") = gsNone
    End If
    If bPenetratesWeb Then
        If sInsetBrace = "InsetMember" Then
            oSelectorLogic.Answer("ShapeAtFace") = gsCope
        Else
            oSelectorLogic.Answer("ShapeAtFace") = gsNone
        End If
    Else
        oSelectorLogic.Answer("ShapeAtFace") = gsNone
    End If
    
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub

'*********************************************************************************************
' Method      : CMSetTopShapeAtFaceAnswer
' Description :
'*********************************************************************************************
Public Sub CMSetTopShapeAtFaceAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetTopShapeAtFaceAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
    
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim sMsg As String
    Dim lStatus As Long

    Set oAppConnection = oSelectorLogic.SmartOccurrence
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    Dim bPenetratesWeb As Boolean
    sMsg = "Determining if web or flange is penetrated.."
    bPenetratesWeb = IsWebPenetrated(oBoundingData.AxisPort, oBoundedData.AxisPort)

    Dim sTopInsetBrace As String
    GetSelectorAnswer oSelectorLogic, "TopBraceType", sTopInsetBrace
    
    If bPenetratesWeb Then
        If sTopInsetBrace = "InsetMember" Then
            oSelectorLogic.Answer("TopShapeAtFace") = gsCope
        Else
            oSelectorLogic.Answer("TopShapeAtFace") = gsNone
        End If
    Else
        oSelectorLogic.Answer("TopShapeAtFace") = gsNone
    End If
    
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub

'*********************************************************************************************
' Method      : CMSetBottomShapeAtFaceAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetBottomShapeAtFaceAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetBottomShapeAtFaceAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
    
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim sMsg As String
    Dim lStatus As Long

    Set oAppConnection = oSelectorLogic.SmartOccurrence
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    Dim bPenetratesWeb As Boolean
    sMsg = "Determining if web or flange is penetrated.."
    bPenetratesWeb = IsWebPenetrated(oBoundingData.AxisPort, oBoundedData.AxisPort)

    Dim sBtmInsetBrace As String
    GetSelectorAnswer oSelectorLogic, "BottomBraceType", sBtmInsetBrace
    
    If bPenetratesWeb Then
        If sBtmInsetBrace = "InsetMember" Then
            oSelectorLogic.Answer("BottomShapeAtFace") = gsCope
        Else
            oSelectorLogic.Answer("BottomShapeAtFace") = gsNone
        End If
    Else
        oSelectorLogic.Answer("BottomShapeAtFace") = gsNone
    End If
    
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub

'*********************************************************************************************
' Method      : CMSetShapeAtTopEdgeAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetShapeAtTopEdgeAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetShapeAtTopEdgeAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
    oSelectorLogic.Answer("ShapeAtTopEdge") = gsNone
              
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub

'*********************************************************************************************
' Method      : CMSetShapeAtBottomEdgeAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetShapeAtBottomEdgeAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetShapeAtBottomEdgeAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
    oSelectorLogic.Answer("ShapeAtBottomEdge") = gsNone
              
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub

'*********************************************************************************************
' Method      : CMSetShapeAtEdgeAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetShapeAtEdgeAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetShapeAtEdgeAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
    oSelectorLogic.Answer("ShapeAtEdge") = gsNone
              
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub

'*********************************************************************************************
' Method      : CMSetTopInsideCornerAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetTopInsideCornerAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetTopInsideCornerAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
    oSelectorLogic.Answer("TopInsideCorner") = gsNone
              
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub

'*********************************************************************************************
' Method      : CMSetBottomInsideCornerAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetBottomInsideCornerAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetBottomInsideCornerAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
    oSelectorLogic.Answer("BottomInsideCorner") = gsNone
              
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub

'*********************************************************************************************
' Method      : CMSetShapeOutsideAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetShapeOutsideAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetShapeOutsideAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
    oSelectorLogic.Answer("ShapeOutside") = gsStraight
              
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub

'*********************************************************************************************
' Method      : CMSetOffsetAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetOffsetAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetOffsetAnswer"
    On Error GoTo ErrorHandler
    
    Dim selString As String
    Dim oSMOcc As IJSmartOccurrence
    
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
    
    If TypeOf oSelectorLogic.SmartOccurrence Is IJSmartOccurrence Then
        Set oSMOcc = oSelectorLogic.SmartOccurrence
        selString = oSMOcc.RootSelection
    End If
       
    Select Case selString
        Case "MbrAxisByFCAndOS1Edge"
            oSelectorLogic.Answer("Offset") = 0.00003
        Case Else
            oSelectorLogic.Answer("Offset") = 0.02
    End Select
              
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub

'*********************************************************************************************
' Method      : CMSetShapeAtEdgeOverlapAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetShapeAtEdgeOverlapAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetShapeAtEdgeOverlapAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
    oSelectorLogic.Answer("ShapeAtEdgeOverlap") = gsNone
              
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub

'*********************************************************************************************
' Method      : CMSetShapeAtOutsideEdgeAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetShapeAtOutsideEdgeAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetShapeAtOutsideEdgeAnswer"
    On Error GoTo ErrorHandler

    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
    
    Dim sTopBraceType As String
    Dim sBottomBraceType As String
    sTopBraceType = oSelectorLogic.Answer("TopBraceType")
    sBottomBraceType = oSelectorLogic.Answer("BottomBraceType")
    
    If sTopBraceType <> "NotApplicable" Then
        If sTopBraceType = "InsetMember" Then
            oSelectorLogic.Answer("ShapeAtEdge") = gsCornerToOutside
        Else
            oSelectorLogic.Answer("ShapeAtEdge") = gsNone
        End If
    ElseIf sBottomBraceType <> "NotApplicable" Then
        If sBottomBraceType = "InsetMember" Then
            oSelectorLogic.Answer("ShapeAtEdge") = gsCornerToOutside
        Else
            oSelectorLogic.Answer("ShapeAtEdge") = gsNone
        End If
    Else
        oSelectorLogic.Answer("ShapeAtEdge") = gsNone
    End If

    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub

'*********************************************************************************************
' Method      : CMSetTopShapeOutsideAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetTopShapeOutsideAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetTopShapeOutsideAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
    oSelectorLogic.Answer("TopShapeOutside") = gsStraight
              
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub

'*********************************************************************************************
' Method      : CMSetBottomShapeOutsideAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetBottomShapeOutsideAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetBottomShapeOutsideAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
    oSelectorLogic.Answer("BottomShapeOutside") = gsStraight
              
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub

'*********************************************************************************************
' Method      : CMSetInsideCornerAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetInsideCornerAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetInsideCornerAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
    oSelectorLogic.Answer("InsideCorner") = gsNone
              
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub



'*********************************************************************************************
' Method      : CMSetExTensionAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetExTensionAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetExTensionAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
    oSelectorLogic.Answer("Extension") = 0.02
              
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub

'*********************************************************************************************
' Method      : CMSetInsetBraceAnswer
' Description : BraceType is a question to enable creation of InsetBrace between two members connected axis-along using a member AC
'*********************************************************************************************
Public Sub CMSetTopBraceTypeAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetInsetBraceAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
    
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim sMsg As String
    Dim lStatus As Long

    Set oAppConnection = oSelectorLogic.SmartOccurrence
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    If IsTubularMember(oBoundingData.MemberPart) Then
        oSelectorLogic.Answer("TopBraceType") = "NotApplicable"
        Exit Sub
    End If
    
    ' ----------------------------------------
    ' Determine if web or flange is penetrated
    ' ----------------------------------------
    Dim bPenetratesWeb As Boolean
    sMsg = "Determining if web or flange is penetrated.."
    bPenetratesWeb = IsWebPenetrated(oBoundingData.AxisPort, oBoundedData.AxisPort)

    Dim oTopORWL As ConnectedEdgeInfo
    Dim oBottomOrWR As ConnectedEdgeInfo
    Dim oTFL As ConnectedEdgeInfo
    Dim oTFR As ConnectedEdgeInfo
    
    GetConnectedEdgeInfo oAppConnection, oBoundedData.AxisPort, oBoundingData.AxisPort, oTopORWL, oBottomOrWR, oTFL, oTFR
    
    Dim bTFL As Boolean
    Dim bBFL As Boolean
    Dim bTFR As Boolean
    Dim bBFR As Boolean
    
    CrossSection_Flanges oBoundedData.MemberPart, bTFL, bBFL, bTFR, bBFR
                
    Dim oEdge1 As ConnectedEdgeInfo
    Dim oEdge2 As ConnectedEdgeInfo

    If bPenetratesWeb Then
        oEdge1.IntersectingEdge = oTopORWL.IntersectingEdge
        oEdge2.IntersectingEdge = oBottomOrWR.IntersectingEdge
    Else
        If (bTFL And bTFR) Or (bBFR And bBFL) Then
             oEdge1.IntersectingEdge = oTFL.IntersectingEdge
             oEdge2.IntersectingEdge = oTFR.IntersectingEdge
        ElseIf (bTFL Or bBFL) And Not bBFR And Not bTFR Then
            oEdge1.IntersectingEdge = oTFL.IntersectingEdge
            oEdge2.IntersectingEdge = oBottomOrWR.IntersectingEdge
        ElseIf (bTFR Or bBFR) And Not bTFL And Not bBFL Then
            oEdge1.IntersectingEdge = oTopORWL.IntersectingEdge
            oEdge2.IntersectingEdge = oTFR.IntersectingEdge
        ElseIf Not bTFL And Not bBFL And Not bTFR And Not bBFR Then
            oEdge1.IntersectingEdge = oTopORWL.IntersectingEdge
            oEdge2.IntersectingEdge = oBottomOrWR.IntersectingEdge
        Else
            '??
        End If
    End If
    
    Dim selString As String, bIsTopBrace As Boolean
    Dim oSMOcc As IJSmartOccurrence
    If TypeOf oSelectorLogic.SmartOccurrence Is IJSmartOccurrence Then
        Set oSMOcc = oSelectorLogic.SmartOccurrence
        selString = oSMOcc.RootSelection
    End If
    
    If selString = "" Then
        GetSelFrmBoundedToBoundingOrien oAppConnection, oBoundedData, oBoundingData, selString, , False
    End If
    
    If bPenetratesWeb Then
        Select Case selString
            Case "MbrAxisByCenter"
                oSelectorLogic.Answer("TopBraceType") = "None"
            Case "MbrAxisByOnMember", "MbrAxisByEdge", "MbrAxisByFCAndEdge", "MbrAxisByFCAndOS1Edge", "MbrAxisByFCAndOSNoEdge"
                'see if top brace is valid
                ' Determines if a Top/Bottom Brace is needed - A top brace is one which is connected to the Top of the Bounded
                bIsTopBrace = IsTopOrBottomInsetBrace(selString, oEdge1, oEdge2)
                If bIsTopBrace Then
                    oSelectorLogic.Answer("TopBraceType") = "None"
                Else
                    oSelectorLogic.Answer("TopBraceType") = "NotApplicable"
                End If
                
            Case Else
                oSelectorLogic.Answer("TopBraceType") = "NotApplicable"
        End Select
    Else
        oSelectorLogic.Answer("TopBraceType") = "None"
    End If
    
              
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub

'*********************************************************************************************
' Method      : CMSetTopInsetBraceAnswer
' Description : TopBraceType is a question to enable creation of TopInsetBrace between two members connected axis-along using a member AC
'When two braces are possible to be placed, we use these questions
'*********************************************************************************************
Public Sub CMSetTopBraceBoundedEndTypeAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetTopInsetBraceAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
    oSelectorLogic.Answer("TopBraceBoundedEndType") = "Offset"
              
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub
'*********************************************************************************************
' Method      : CMSetBtmInsetBraceAnswer
' Description : BottomBraceType is a question to enable creation of BottomInsetBrace between two members connected axis-along using a member AC
'When two braces are possible to be placed, we use these questions
'*********************************************************************************************
Public Sub CMSetTopBraceBoundedEndValueAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetBtmInsetBraceAnswer"
    On Error GoTo ErrorHandler
    
     ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
     
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim sMsg As String
    Dim lStatus As Long

    Set oAppConnection = oSelectorLogic.SmartOccurrence
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    If IsTubularMember(oBoundingData.MemberPart) Then
        oSelectorLogic.Answer("TopBraceBoundedEndValue") = "0"
        Exit Sub
    End If
    
    ' ----------------------------------------
    ' Determine if web or flange is penetrated
    ' ----------------------------------------
    Dim bPenetratesWeb As Boolean
    sMsg = "Determining if web or flange is penetrated.."
    bPenetratesWeb = IsWebPenetrated(oBoundingData.AxisPort, oBoundedData.AxisPort)

    Dim oSDO_Bounded As New StructDetailObjects.MemberPart
    Set oSDO_Bounded.object = oBoundedData.MemberPart
    
    If bPenetratesWeb Then
        Dim oTopORWL As ConnectedEdgeInfo
        Dim oBottomOrWR As ConnectedEdgeInfo
        Dim oTFL As ConnectedEdgeInfo
        Dim oTFR As ConnectedEdgeInfo
        Dim oMeasurements As Collection, dDistTopToTopOfBounding As Double
        
        GetConnectedEdgeInfo oAppConnection, oBoundedData.AxisPort, oBoundingData.AxisPort, oTopORWL, oBottomOrWR, oTFL, oTFR, oMeasurements
        dDistTopToTopOfBounding = oMeasurements.Item("DimPt15ToTop")
        oSelectorLogic.Answer("TopBraceBoundedEndValue") = (dDistTopToTopOfBounding * 4)
    Else
        ' From the QA related input get the current Symbol Def.
        oSelectorLogic.Answer("TopBraceBoundedEndValue") = (((oSDO_Bounded.WebLength - oSDO_Bounded.flangeThickness) / Cos(Atn(0.25))) - 0.1) * 4 'default 1/4 slope maintained to position brace
    End If
              
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub
'*********************************************************************************************
' Method      : CMSetBtmInsetBraceAnswer
' Description : BottomBraceType is a question to enable creation of BottomInsetBrace between two members connected axis-along using a member AC
'When two braces are possible to be placed, we use these questions
'*********************************************************************************************
Public Sub CMSetTopBraceBoundingEndTypeAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetBtmInsetBraceAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
    oSelectorLogic.Answer("TopBraceBoundingEndType") = "Offset"
              
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub

'*********************************************************************************************
' Method      : CMSetBtmInsetBraceAnswer
' Description : BottomBraceType is a question to enable creation of BottomInsetBrace between two members connected axis-along using a member AC
'When two braces are possible to be placed, we use these questions
'*********************************************************************************************
Public Sub CMSetTopBraceBoundingEndValueAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetBtmInsetBraceAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
     
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim sMsg As String
    Dim lStatus As Long

    Set oAppConnection = oSelectorLogic.SmartOccurrence
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    If IsTubularMember(oBoundingData.MemberPart) Then
        oSelectorLogic.Answer("TopBraceBoundingEndValue") = "0"
        Exit Sub
    End If
    
    ' ----------------------------------------
    ' Determine if web or flange is penetrated
    ' ----------------------------------------
    Dim bPenetratesWeb As Boolean
    sMsg = "Determining if web or flange is penetrated.."
    bPenetratesWeb = IsWebPenetrated(oBoundingData.AxisPort, oBoundedData.AxisPort)

    Dim oSDO_Bounded As New StructDetailObjects.MemberPart
    Set oSDO_Bounded.object = oBoundedData.MemberPart
    
    If bPenetratesWeb Then
        Dim oTopORWL As ConnectedEdgeInfo
        Dim oBottomOrWR As ConnectedEdgeInfo
        Dim oTFL As ConnectedEdgeInfo
        Dim oTFR As ConnectedEdgeInfo
        Dim oMeasurements As Collection, dDistTopToTopOfBounding As Double
        
        GetConnectedEdgeInfo oAppConnection, oBoundedData.AxisPort, oBoundingData.AxisPort, oTopORWL, oBottomOrWR, oTFL, oTFR, oMeasurements
        dDistTopToTopOfBounding = oMeasurements.Item("DimPt15ToTop")
        oSelectorLogic.Answer("TopBraceBoundingEndValue") = dDistTopToTopOfBounding
    Else
        ' From the QA related input get the current Symbol Def.
        oSelectorLogic.Answer("TopBraceBoundingEndValue") = (((oSDO_Bounded.WebLength - oSDO_Bounded.flangeThickness) / Cos(Atn(0.25))) - 0.1) 'default 1/4 slope maintained to position brace
    End If
    
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub

'*********************************************************************************************
' Method      : CMSetInsetBraceAnswer
' Description : BraceType is a question to enable creation of InsetBrace between two members connected axis-along using a member AC
'*********************************************************************************************
Public Sub CMSetBottomBraceTypeAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetInsetBraceAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
    
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim sMsg As String
    Dim lStatus As Long

    Set oAppConnection = oSelectorLogic.SmartOccurrence
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    If IsTubularMember(oBoundingData.MemberPart) Then
        oSelectorLogic.Answer("BottomBraceType") = "NotApplicable"
        Exit Sub
    End If
    
    ' ----------------------------------------
    ' Determine if web or flange is penetrated
    ' ----------------------------------------
    Dim bPenetratesWeb As Boolean
    sMsg = "Determining if web or flange is penetrated.."
    bPenetratesWeb = IsWebPenetrated(oBoundingData.AxisPort, oBoundedData.AxisPort)

    Dim oTopORWL As ConnectedEdgeInfo
    Dim oBottomOrWR As ConnectedEdgeInfo
    Dim oTFL As ConnectedEdgeInfo
    Dim oTFR As ConnectedEdgeInfo
    
    GetConnectedEdgeInfo oAppConnection, oBoundedData.AxisPort, oBoundingData.AxisPort, oTopORWL, oBottomOrWR, oTFL, oTFR
    
    Dim bTFL As Boolean
    Dim bBFL As Boolean
    Dim bTFR As Boolean
    Dim bBFR As Boolean
    
    CrossSection_Flanges oBoundedData.MemberPart, bTFL, bBFL, bTFR, bBFR
                
    Dim oEdge1 As ConnectedEdgeInfo
    Dim oEdge2 As ConnectedEdgeInfo

    If bPenetratesWeb Then
        oEdge1.IntersectingEdge = oTopORWL.IntersectingEdge
        oEdge2.IntersectingEdge = oBottomOrWR.IntersectingEdge
    Else
        If (bTFL And bTFR) Or (bBFR And bBFL) Then
             oEdge1.IntersectingEdge = oTFL.IntersectingEdge
             oEdge2.IntersectingEdge = oTFR.IntersectingEdge
        ElseIf (bTFL Or bBFL) And Not bBFR And Not bTFR Then
            oEdge1.IntersectingEdge = oTFL.IntersectingEdge
            oEdge2.IntersectingEdge = oBottomOrWR.IntersectingEdge
        ElseIf (bTFR Or bBFR) And Not bTFL And Not bBFL Then
            oEdge1.IntersectingEdge = oTopORWL.IntersectingEdge
            oEdge2.IntersectingEdge = oTFR.IntersectingEdge
        ElseIf Not bTFL And Not bBFL And Not bTFR And Not bBFR Then
            oEdge1.IntersectingEdge = oTopORWL.IntersectingEdge
            oEdge2.IntersectingEdge = oBottomOrWR.IntersectingEdge
        Else
            '??
        End If
    End If
    
    Dim selString As String, bIsTopBrace As Boolean
    Dim oSMOcc As IJSmartOccurrence
    If TypeOf oSelectorLogic.SmartOccurrence Is IJSmartOccurrence Then
        Set oSMOcc = oSelectorLogic.SmartOccurrence
        selString = oSMOcc.RootSelection
    End If
    
    If selString = "" Then
        GetSelFrmBoundedToBoundingOrien oAppConnection, oBoundedData, oBoundingData, selString, , False
    End If
    
    If bPenetratesWeb Then
        Select Case selString
            Case "MbrAxisByCenter"
                oSelectorLogic.Answer("BottomBraceType") = "None"
            Case "MbrAxisByOnMember", "MbrAxisByEdge", "MbrAxisByFCAndEdge", "MbrAxisByFCAndOS1Edge", "MbrAxisByFCAndOSNoEdge"
                'see if top brace is valid
                ' Determines if a Top/Bottom Brace is needed - A top brace is one which is connected to the Top of the Bounded
                bIsTopBrace = IsTopOrBottomInsetBrace(selString, oEdge1, oEdge2)
                If Not bIsTopBrace Then 'Bottom brace
                    oSelectorLogic.Answer("BottomBraceType") = "None"
                Else
                    oSelectorLogic.Answer("BottomBraceType") = "NotApplicable"
                End If
                
            Case Else
                oSelectorLogic.Answer("BottomBraceType") = "NotApplicable"
        End Select
    Else
        oSelectorLogic.Answer("BottomBraceType") = "None"
    End If
    
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub

'*********************************************************************************************
' Method      : CMSetTopInsetBraceAnswer
' Description : TopBraceType is a question to enable creation of TopInsetBrace between two members connected axis-along using a member AC
'When two braces are possible to be placed, we use these questions
'*********************************************************************************************
Public Sub CMSetBottomBraceBoundedEndTypeAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetTopInsetBraceAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
    oSelectorLogic.Answer("BottomBraceBoundedEndType") = "Offset"
              
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub
'*********************************************************************************************
' Method      : CMSetBtmInsetBraceAnswer
' Description : BottomBraceType is a question to enable creation of BottomInsetBrace between two members connected axis-along using a member AC
'When two braces are possible to be placed, we use these questions
'*********************************************************************************************
Public Sub CMSetBottomBraceBoundedEndValueAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetBtmInsetBraceAnswer"
    On Error GoTo ErrorHandler
   
      ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
     
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim sMsg As String
    Dim lStatus As Long

    Set oAppConnection = oSelectorLogic.SmartOccurrence
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    If IsTubularMember(oBoundingData.MemberPart) Then
        oSelectorLogic.Answer("BottomBraceBoundedEndValue") = "0"
        Exit Sub
    End If
    
    Dim oSDO_Bounded As New StructDetailObjects.MemberPart
    Set oSDO_Bounded.object = oBoundedData.MemberPart
    
    Dim bPenetratesWeb As Boolean
    sMsg = "Determining if web or flange is penetrated.."
    bPenetratesWeb = IsWebPenetrated(oBoundingData.AxisPort, oBoundedData.AxisPort)

    If bPenetratesWeb Then
        Dim oTopORWL As ConnectedEdgeInfo
        Dim oBottomOrWR As ConnectedEdgeInfo
        Dim oTFL As ConnectedEdgeInfo
        Dim oTFR As ConnectedEdgeInfo
        Dim oMeasurements As Collection, dDistBtmToBtmOfBounding As Double
        
        GetConnectedEdgeInfo oAppConnection, oBoundedData.AxisPort, oBoundingData.AxisPort, oTopORWL, oBottomOrWR, oTFL, oTFR, oMeasurements
        dDistBtmToBtmOfBounding = oMeasurements.Item("DimPt23ToBottom")
        oSelectorLogic.Answer("BottomBraceBoundedEndValue") = dDistBtmToBtmOfBounding * 4
    Else
        ' From the QA related input get the current Symbol Def.
        oSelectorLogic.Answer("BottomBraceBoundedEndValue") = (((oSDO_Bounded.WebLength - oSDO_Bounded.flangeThickness) / Cos(Atn(0.25))) - 0.1) * 4 'default 1/4 slope maintained to position brace
    End If
   
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub

'*********************************************************************************************
' Method      : CMSetBtmInsetBraceAnswer
' Description : BottomBraceType is a question to enable creation of BottomInsetBrace between two members connected axis-along using a member AC
'When two braces are possible to be placed, we use these questions
'*********************************************************************************************
Public Sub CMSetBottomBraceBoundingEndTypeAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetBtmInsetBraceAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
    oSelectorLogic.Answer("BottomBraceBoundingEndType") = "Offset"
              
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub

'*********************************************************************************************
' Method      : CMSetBtmInsetBraceAnswer
' Description : BottomBraceType is a question to enable creation of BottomInsetBrace between two members connected axis-along using a member AC
'When two braces are possible to be placed, we use these questions
'*********************************************************************************************
Public Sub CMSetBottomBraceBoundingEndValueAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetBtmInsetBraceAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
     
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim sMsg As String
    Dim lStatus As Long

    Set oAppConnection = oSelectorLogic.SmartOccurrence
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    If IsTubularMember(oBoundingData.MemberPart) Then
        oSelectorLogic.Answer("BottomBraceBoundingEndValue") = "0"
        Exit Sub
    End If
    
    Dim oSDO_Bounded As New StructDetailObjects.MemberPart
    Set oSDO_Bounded.object = oBoundedData.MemberPart
    
    Dim bPenetratesWeb As Boolean
    sMsg = "Determining if web or flange is penetrated.."
    bPenetratesWeb = IsWebPenetrated(oBoundingData.AxisPort, oBoundedData.AxisPort)

    If bPenetratesWeb Then
        Dim oTopORWL As ConnectedEdgeInfo
        Dim oBottomOrWR As ConnectedEdgeInfo
        Dim oTFL As ConnectedEdgeInfo
        Dim oTFR As ConnectedEdgeInfo
        Dim oMeasurements As Collection, dDistBtmToBtmOfBounding As Double
        
        GetConnectedEdgeInfo oAppConnection, oBoundedData.AxisPort, oBoundingData.AxisPort, oTopORWL, oBottomOrWR, oTFL, oTFR, oMeasurements
        dDistBtmToBtmOfBounding = oMeasurements.Item("DimPt23ToBottom")
        oSelectorLogic.Answer("BottomBraceBoundingEndValue") = dDistBtmToBtmOfBounding
    Else
        ' From the QA related input get the current Symbol Def.
        oSelectorLogic.Answer("BottomBraceBoundingEndValue") = (((oSDO_Bounded.WebLength - oSDO_Bounded.flangeThickness) / Cos(Atn(0.25))) - 0.1) 'default 1/4 slope maintained to position brace
    End If
    
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub

'*********************************************************************************************
' Method      : CMSetChamferMeasurementAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetChamferMeasurementAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetChamferMeasurementAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)
    oSelectorLogic.Answer("ChamferMeasurement") = gsSlope
    
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub

'*********************************************************************************************
' Method      : CMSetChamferValueAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetChamferValueAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD As String = "CMSetChamferValueAnswer"
    On Error GoTo ErrorHandler
    
    ' From the QA related input get the current Symbol Def.
    Dim oSelectorLogic As IJDSelectorLogic
    Set oSelectorLogic = GetSelectorLogicForCustomMethod(pInput)

    Dim strChamferMeasurement As String

    'Get Chamfer Meaurement Type
    GetSelectorAnswer oSelectorLogic, "ChamferMeasurement", strChamferMeasurement
    
    'Set Chamfer Value based on Measurement Type
    If strChamferMeasurement = gsSlope Then
        oSelectorLogic.Answer("ChamferValue") = 0.333
    ElseIf strChamferMeasurement = gsAngle Then
        oSelectorLogic.Answer("ChamferValue") = 18 'degrees
    End If
    
    Exit Sub
ErrorHandler:
    oSelectorLogic.ReportError "Set Answer", METHOD
End Sub
