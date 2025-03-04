VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ellipse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*******************************************************************
'  Copyright (c) 2003, Intergraph Corporation.  All rights reserved.
'
'  Project: M:\Refdata\Symbols\OutfittingCrosssectionsOutfittingCrossSections.vbp
'  File:  M:\Refdata\Symbols\OutfittingCrosssections\Ellipse.cls
'
'  Description:  For Drawing Ellipse with width & height as specified.
'
'  Author: NVS
'
'  History:
'
'   15th Sep, 1999 : PR
'     Fixed TR# 7632 and TR# 7633 : Unable to Place the Symbol for
'     the OutfittingCrossSections in the Route Environment.
'   27th Sep, 1999: APS
'     Took care of P2R2 symbol impact.
'     For setting outputs on a rep, one
'     needs to query for outputs from rep and set them.
'
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private m_outputColl As IMSSymbolEntities.DOutputCollection

Implements IJDUserSymbolServices

Private Sub Class_Terminate()
     Set m_outputColl = Nothing
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
    IJDUserSymbolServices_EditOccurence = False
End Function


Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = "OutfittingCrossSections.Ellipse"
End Function
'********************************************************************
' Routine: IJDUserSymbolServices_InstanciateDefinition
'
'
' Description:This instanciates a persistent symbol definition object
' and initialize it for the first time.
'********************************************************************
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParameters As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler
  
  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
  
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  
  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "OutfittingCrossSections.Ellipse"
  oSymbolDefinition.CodeBase = CodeBase

  ' Give a unique name to the symbol definition
  oSymbolDefinition.Name = oSymbolDefinition.ProgId

  'returned symbol defintion
  Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
  
  Set oSymbolFactory = Nothing
  Set oSymbolDefinition = Nothing
  Exit Function

ErrorHandler:
    Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    Debug.Assert False
End Function


Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
    Set m_outputColl = outputcoll
    If StrComp(repName, "Ellipse") = 0 Then
        Ellipse arrayOfInputs(1), arrayOfInputs(2)
    End If
End Sub

Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pSymbolDefinition As IJDSymbolDefinition)
   On Error GoTo ErrorHandler
' Feed Ellipse Definition
  '   Width = 0.2
  '   Height = 0.1
  
' Remove all previous Symbol Definition information
   pSymbolDefinition.IJDInputs.RemoveAllInput
   pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
   pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations

   Dim I1 As IMSSymbolEntities.IJDInput
   Set I1 = New IMSSymbolEntities.DInput
   
   Dim I2 As IMSSymbolEntities.IJDInput
   Set I2 = New IMSSymbolEntities.DInput
  
   I1.Name = "Width"
   I1.Description = "Major Axis of the Ellipse"
   I1.Properties = igINPUT_IS_A_PARAMETER
  
 ' Create a default value
     Dim PC As IMSSymbolEntities.DParameterContent
     Set PC = New IMSSymbolEntities.DParameterContent
  
     PC.Type = igValue
     PC.UomValue = 0.2
     I1.DefaultParameterValue = PC
  
     I2.Name = "Height"
     I2.Description = "Minor Axis of the Ellipse"
     I2.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.1
     I2.DefaultParameterValue = PC
   
 ' set the input to the definition
     Dim InputsIf As IMSSymbolEntities.IJDInputs
     Set InputsIf = pSymbolDefinition
   
     InputsIf.SetInput I1, 1
     InputsIf.SetInput I2, 2
   
  ' Define the representation "Ellipse"
     Dim rep1 As IMSSymbolEntities.IJDRepresentation
     Set rep1 = New IMSSymbolEntities.DRepresentation
  
     rep1.Name = "Ellipse"
     rep1.Description = "It's a Ellipse"
     rep1.Properties = igREPRESENTATION_ISVBFUNCTION
     rep1.RepresentationId = 1
    
  ' Create the output
     Dim O1 As IMSSymbolEntities.IJDOutput
     Set O1 = New IMSSymbolEntities.DOutput
 
     O1.Name = "Ellipse"
     O1.Description = "It is a Ellipse"
     O1.Properties = 0
    
    Dim oRepPhysicalOutputs As IMSSymbolEntities.IJDOutputs
    Set oRepPhysicalOutputs = rep1
  
  ' Set the output
     oRepPhysicalOutputs.SetOutput O1
  
  ' Set the representation to definition
      Dim RepsIf As IMSSymbolEntities.IJDRepresentations
      Set RepsIf = pSymbolDefinition
      RepsIf.SetRepresentation rep1
  
  ' Set the evaluation function associated to the Rectangular representation
  
      Dim RepEvalsIf As IMSSymbolEntities.IJDRepresentationEvaluations
      Set RepEvalsIf = pSymbolDefinition
      Dim EvalFuncForEllipse As IJDRepresentationEvaluation
    
      Set EvalFuncForEllipse = New DRepresentationEvaluation
      EvalFuncForEllipse.Name = "Ellipse"
      EvalFuncForEllipse.Description = "evaluation function for the Ellipse representation"
      EvalFuncForEllipse.Properties = igREPRESENTATION_HIDDEN
      EvalFuncForEllipse.Type = igREPRESENTATION_VBFUNCTION
      EvalFuncForEllipse.ProgId = "OutfittingCrossSections.Ellipse"
      RepEvalsIf.AddRepresentationEvaluation EvalFuncForEllipse
  
      Set O1 = Nothing
  
      Set RepEvalsIf = Nothing
      Set rep1 = Nothing
      Set oRepPhysicalOutputs = Nothing
      Set RepsIf = Nothing
      Set EvalFuncForEllipse = Nothing
   
  Exit Sub

ErrorHandler:
    Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    Debug.Assert False


End Sub

' Draw the Ellipse
  Sub Ellipse(ByVal width As Double, ByVal height As Double)
  
    On Error GoTo DrawError
    
    Dim fact As New IngrGeom3D.GeometryFactory
    Dim Geometry As IngrGeom3D.Ellipse3d
    
    Dim CenterX As Double
    Dim CenterY As Double
    Dim CenterZ As Double
    
    Dim NormalX As Double
    Dim NormalY As Double
    Dim NormalZ As Double
    
    Dim MajorX As Double
    Dim MajorY As Double
    Dim MajorZ As Double
    
    Dim MMRatio As Double
    
    CenterX = 0
    CenterY = 0
    CenterZ = 0
    
    NormalX = 0
    NormalY = 0
    NormalZ = 1
   
    MajorX = 0.1
    MajorY = 0
    MajorZ = 0
    
    MMRatio = height / width
   
    Set Geometry = fact.Ellipses3d.CreateByCenterNormMajAxisRatio(m_outputColl.ResourceManager, CenterX, CenterY, CenterZ, NormalX, NormalY, NormalZ, MajorX, MajorY, MajorZ, MMRatio)
    m_outputColl.AddOutput "Ellipse", Geometry
     Set Geometry = Nothing
     Set fact = Nothing
     
 Exit Sub
DrawError:
    Set Geometry = Nothing
    Set fact = Nothing
    
    Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
	Debug.Assert False
End Sub


