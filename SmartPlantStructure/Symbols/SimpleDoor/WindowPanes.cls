VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WindowPanes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*******************************************************************
'  Copyright (C) 1998, Intergraph Corporation.  All rights reserved.
'
'   Definition for Window Panes
'  History:
'******************************************************************

Const definitionProgId As String = "SimpleDoorAsm.WindowPanes"
Private Const MODULE = "WindowPanes:"  'Used for error messages
   
Implements IJDUserSymbolServices



Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
    IJDUserSymbolServices_EditOccurence = False
End Function

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = definitionProgId
End Function
'********************************************************************
' Routine: IJDUserSymbolServices_InstanciateDefinition
'
'
' Description:This instanciates a persistent symbol definition object
' and initializes it for the first time.
'********************************************************************
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParameters As Variant, ByVal ActiveConnection As Object) As Object
    Const METHOD = "IJDUserSymbolServices_InstanciateDefinition:"
    
    On Error GoTo ErrorHandler
    'Lets Create a Symbol Definition Object.
    Dim pSymbolFactory As New DSymbolEntitiesFactory
    Dim pSymbolDefinition As IJDSymbolDefinition
    Set pSymbolDefinition = pSymbolFactory.CreateEntity(definition, ActiveConnection)

    ' Type the symbol definition
    pSymbolDefinition.ProgId = definitionProgId
    pSymbolDefinition.CodeBase = CodeBase
    pSymbolDefinition.Name = IJDUserSymbolServices_GetDefinitionName(Nothing)

    ' Initialize descriptions : Inputs, Representations, Outputs, Libraries, Methods, Commands.
    IJDUserSymbolServices_InitializeSymbolDefinition pSymbolDefinition
    
    'Return symbol definition
    Set IJDUserSymbolServices_InstanciateDefinition = pSymbolDefinition
    Exit Function
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Function

Private Sub DrawWindowPanes(ByVal outputcoll As Object, ByRef arrayOfInputs())
    Const METHOD = "DrawWindowPanes:"
    Dim pPoint1 As New DPosition
    Dim pPoint2 As New DPosition
    Dim pPoint3 As New DPosition
    Dim pStartPos As New DPosition
    Dim pEndPos As New DPosition
    Dim lRowPanes As Long
    Dim lColPanes As Long
    Dim linePane As Line3d
    Dim pFrameLine As LineString3d
    Dim pOffsetVector As New DVector
    Dim dTotalDist As Double
    Dim dOffset As Double
    Dim lIndex As Long
    Dim pPoints(1 To 15) As Double

    On Error GoTo ErrorHandler
        Dim pGeometryFactory As GeometryFactory
        Set pGeometryFactory = New GeometryFactory
        With pGeometryFactory
        
            ' Get inputs
            pPoint1.x = arrayOfInputs(1)
            pPoint1.y = arrayOfInputs(2)
            pPoint1.z = arrayOfInputs(3)
        
            pPoint2.x = arrayOfInputs(4)
            pPoint2.y = arrayOfInputs(5)
            pPoint2.z = arrayOfInputs(6)
        
            pPoint3.x = arrayOfInputs(7)
            pPoint3.y = arrayOfInputs(8)
            pPoint3.z = arrayOfInputs(9)
    
            lRowPanes = arrayOfInputs(10)
            lColPanes = arrayOfInputs(11)
            
            ' create line transoms only if lRowPanes > 1 (1 row pane is no transoms)
            If lRowPanes > 1 Then
                dTotalDist = pPoint3.DistPt(pPoint1)
                Set pOffsetVector = pPoint3.Subtract(pPoint1)
                dOffset = dTotalDist / lRowPanes
                
                'construct (lRowPanes -1) transom lines
                For lIndex = 1 To lRowPanes - 1
                    Set linePane = Nothing
                    pOffsetVector.length = lIndex * dOffset
                    Set pStartPos = pPoint1.offset(pOffsetVector)
                    Set pEndPos = pPoint2.offset(pOffsetVector)
                    
                    Set linePane = .Lines3d.CreateBy2Points(outputcoll.ResourceManager, _
                                pStartPos.x, pStartPos.y, pStartPos.z, _
                                pEndPos.x, pEndPos.y, pEndPos.z)
                    Call outputcoll.AddOutput("RowPanes" & Str(lIndex), linePane)
                Next lIndex
            End If
            
            ' create column transoms only if lColPanes > 1 (1 col pane is no transoms)
            If lColPanes > 1 Then
                dTotalDist = pPoint2.DistPt(pPoint1)
                Set pOffsetVector = pPoint2.Subtract(pPoint1)
                dOffset = dTotalDist / lColPanes
                
                'construct (lRowPanes -1) transom lines
                For lIndex = 1 To lColPanes - 1
                    Set linePane = Nothing
                    pOffsetVector.length = lIndex * dOffset
                    Set pStartPos = pPoint1.offset(pOffsetVector)
                    Set pEndPos = pPoint3.offset(pOffsetVector)
                    
                    Set linePane = .Lines3d.CreateBy2Points(outputcoll.ResourceManager, _
                                pStartPos.x, pStartPos.y, pStartPos.z, _
                                pEndPos.x, pEndPos.y, pEndPos.z)
                    Call outputcoll.AddOutput("ColPanes" & Str(lIndex), linePane)
                Next lIndex
            End If
            
            'Create the frame
            If lColPanes > 1 Or lRowPanes > 1 Then
                pPoints(1) = arrayOfInputs(1)
                pPoints(2) = arrayOfInputs(2)
                pPoints(3) = arrayOfInputs(3)
            
                pPoints(4) = arrayOfInputs(4)
                pPoints(5) = arrayOfInputs(5)
                pPoints(6) = arrayOfInputs(6)
            
                pPoints(7) = arrayOfInputs(4)
                pPoints(8) = arrayOfInputs(5)
                pPoints(9) = arrayOfInputs(9)
        
                pPoints(10) = arrayOfInputs(7)
                pPoints(11) = arrayOfInputs(8)
                pPoints(12) = arrayOfInputs(9)
        
                pPoints(13) = pPoints(1)
                pPoints(14) = pPoints(2)
                pPoints(15) = pPoints(3)
             
                Set pFrameLine = .LineStrings3d.CreateByPoints(outputcoll.ResourceManager, 5, pPoints)
                Call outputcoll.AddOutput("frameLine", pFrameLine)
           End If
            
        End With

    Exit Sub

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Sub

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
    Const METHOD = "IJDUserSymbolServices_InvokeRepresentation:"
    On Error GoTo ErrorHandler
    
    
    If StrComp(repName, "SimplePhysical") = 0 Then
        DrawWindowPanes outputcoll, arrayOfInputs
    End If
    
    Exit Sub

ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Sub

Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pSymbolDefinition As IJDSymbolDefinition)
    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition:"
    
    On Error GoTo ErrorHandler
    ' Feed WindowPanes parameters (pt1, pt2,pt3 gives the rectangular frame + nRowPanes, nColPanes)
    ' pt1 ----- pt2 (pt1/pt2 give row width)
    '  |
    '  |
    ' pt3 (pt1/pt3 give col height)
    '
    ' Inputs:
    '           "pt1X" = 100
    '           "pt1Y" = 100
    '           "pt1Z" = 100
    '           "pt2X" = 200
    '           "pt2Y" = 100
    '           "pt2Z" = 100
    '           "pt3X" = 100
    '           "pt3Y" = 100
    '           "pt3Z" = 700
    '           "nRowPanes = 3
    '           "nColPanes" = 3
    ' Representation : "Physical"
    '       Outputs collection : igCollection_Variable
    '           . ...  }   ellipsis for the horizontal transoms (nRowPanes -1)
    '           . ...  }   ellipsis for the vertical transoms (nColPanes - 1)
    '           .   }
    
    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
    
    
    ' ------------------
    ' Inputs declaration
    ' ------------------
    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition
    
    Dim pInputDesc As IMSSymbolEntities.IJDInput
    Set pInputDesc = New IMSSymbolEntities.DInput
    Dim pPC As IMSSymbolEntities.IJDParameterContent
    Set pPC = New IMSSymbolEntities.DParameterContent
    pPC.Type = igValue

    pInputDesc.Name = "pt1X"
    pInputDesc.Description = "X coord Pt1"
    pInputDesc.Properties = igINPUT_IS_A_PARAMETER
    pPC.uomValue = 100
    pInputDesc.DefaultParameterValue = pPC
    InputsIf.SetInput pInputDesc, 1
    pInputDesc.Reset
  
    pInputDesc.Name = "pt1Y"
    pInputDesc.Description = "Y coord Pt1"
    pInputDesc.Properties = igINPUT_IS_A_PARAMETER
    pPC.uomValue = 100
    pInputDesc.DefaultParameterValue = pPC
    InputsIf.SetInput pInputDesc, 2
    pInputDesc.Reset
   
    pInputDesc.Name = "pt1Z"
    pInputDesc.Description = "Z coord Pt1"
    pInputDesc.Properties = igINPUT_IS_A_PARAMETER
    pPC.uomValue = 100
    pInputDesc.DefaultParameterValue = pPC
    InputsIf.SetInput pInputDesc, 3
    pInputDesc.Reset
  
    pInputDesc.Name = "pt2X"
    pInputDesc.Description = "X coord Pt2"
    pInputDesc.Properties = igINPUT_IS_A_PARAMETER
    pPC.uomValue = 200
    pInputDesc.DefaultParameterValue = pPC
    InputsIf.SetInput pInputDesc, 4
    pInputDesc.Reset
  
    pInputDesc.Name = "pt2Y"
    pInputDesc.Description = "Y coord Pt2"
    pInputDesc.Properties = igINPUT_IS_A_PARAMETER
    pPC.uomValue = 100
    pInputDesc.DefaultParameterValue = pPC
    InputsIf.SetInput pInputDesc, 5
    pInputDesc.Reset
   
    pInputDesc.Name = "pt2Z"
    pInputDesc.Description = "Z coord Pt2"
    pInputDesc.Properties = igINPUT_IS_A_PARAMETER
    pPC.uomValue = 100
    pInputDesc.DefaultParameterValue = pPC
    InputsIf.SetInput pInputDesc, 6
    pInputDesc.Reset
    
    pInputDesc.Name = "pt3X"
    pInputDesc.Description = "X coord Pt3"
    pInputDesc.Properties = igINPUT_IS_A_PARAMETER
    pPC.uomValue = 100
    pInputDesc.DefaultParameterValue = pPC
    InputsIf.SetInput pInputDesc, 7
    pInputDesc.Reset
  
    pInputDesc.Name = "pt3Y"
    pInputDesc.Description = "Y coord Pt3"
    pInputDesc.Properties = igINPUT_IS_A_PARAMETER
    pPC.uomValue = 100
    pInputDesc.DefaultParameterValue = pPC
    InputsIf.SetInput pInputDesc, 8
    pInputDesc.Reset
   
    pInputDesc.Name = "pt3Z"
    pInputDesc.Description = "Z coord Pt3"
    pInputDesc.Properties = igINPUT_IS_A_PARAMETER
    pPC.uomValue = 700
    pInputDesc.DefaultParameterValue = pPC
    InputsIf.SetInput pInputDesc, 9
    pInputDesc.Reset
    
    pInputDesc.Name = "nRowPanes"
    pInputDesc.Description = "Number of Row Panes"
    pInputDesc.Properties = igINPUT_IS_A_PARAMETER
    pPC.uomValue = 3
    pInputDesc.DefaultParameterValue = pPC
    InputsIf.SetInput pInputDesc, 10
    pInputDesc.Reset
    
    pInputDesc.Name = "nColPanes"
    pInputDesc.Description = "Number of Col Panes"
    pInputDesc.Properties = igINPUT_IS_A_PARAMETER
    pPC.uomValue = 3
    pInputDesc.DefaultParameterValue = pPC
    InputsIf.SetInput pInputDesc, 11
    pInputDesc.Reset
    
    ' definition properties
    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE
    'cannot place a custom component it is not an occurence and has no matrix
'    pSymbolDefinition.CacheOption = igSYMBOL_CACHE_OPTION_NOT_SHARED
'    pSymbolDefinition.SupportUpdate = igSYMBOL_SUPPORT_UPDATE
  
    ' ---------------------------------------
    ' "Physical" representation  declaration
    ' ---------------------------------------
    Dim pRepsIf As IMSSymbolEntities.IJDRepresentations
    Set pRepsIf = pSymbolDefinition
    Dim pRepDesc As IMSSymbolEntities.IJDRepresentation
    Set pRepDesc = New IMSSymbolEntities.DRepresentation
    
    pRepDesc.Name = "SimplePhysical"
    pRepDesc.Description = "SimplePhysical"
    pRepDesc.Properties = igREPRESENTATION_ISVBFUNCTION + igCOLLECTION_VARIABLE ' declare that the number of outputs is variable
    pRepDesc.RepresentationId = 1
    
    ' No "Physical" outputs declaration: all are optional
    
    ' Set representation description into the definition.
    pRepsIf.SetRepresentation pRepDesc
    
    Dim oRepEval     As IJDRepresentationEvaluation
    Dim RepsEvalsIf     As IMSSymbolEntities.IJDRepresentationEvaluations
    Set RepsEvalsIf = pSymbolDefinition
    
    Set oRepEval = New DRepresentationEvaluation
    oRepEval.Name = "SimplePhysical"
    oRepEval.Description = "Physical representation"
    oRepEval.Properties = igREPRESENTATION_HIDDEN
    oRepEval.Type = igREPRESENTATION_VBFUNCTION
    oRepEval.ProgId = "SimpleDoorAsm.WindowPanes"
    RepsEvalsIf.AddRepresentationEvaluation oRepEval
    Set oRepEval = Nothing
    Set RepsEvalsIf = Nothing
 
    Exit Sub
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Sub


