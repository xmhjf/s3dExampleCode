VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rectangleCornerRadius"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_SymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory

Private m_SymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
Private m_outputColl As Object

Private Form As Editoval

Implements IJDUserSymbolServices

Private Sub Class_Terminate()
    Set m_SymbolFactory = Nothing
    Set m_SymbolDefinition = Nothing

End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
    
'    Set m_f = New frmEditOccurence1
'    Set m_frmEditOccurence1.SymbolOccurrence = pSymbolOccurence
'    Set m_frmEditOccurence1.TransactionMgr = TransactionMgr
'    m_frmEditOccurence1.Show
    IJDUserSymbolServices_EditOccurence = False
End Function


Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal pDefName As String, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler
  
  Set m_SymbolDefinition = m_SymbolFactory.CreateEntity(Definition, ActiveConnection)
  IJDUserSymbolServices_InitializeSymbolDefinition m_SymbolDefinition
  
  'returned symbol defintion
  Set IJDUserSymbolServices_InstanciateDefinition = m_SymbolDefinition
  
  Exit Function

ErrorHandler:
    MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    Debug.Assert False
End Function


Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
    Set m_outputColl = outputcoll
    If StrComp(repName, "rectangleCornerRadius") = 0 Then
        rectangleCornerRadius arrayOfInputs(1), arrayOfInputs(2), arrayOfInputs(3)
        End If
End Sub

Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pSymbolDefinition As IJDSymbolDefinition)
   On Error GoTo ErrorHandler
' Feed rectangleCornerRadius Definition
' Inputs:
'     Width = 0.2
'     Height =0.1
'     CornerRadius=0.02
' The representation "rectangleCornerRadius" is composed of
'
'           - Frame1
'           - Frame2
'           - Frame3
'           - Frame4
'           - Arc1
'           - Arc2
'           - Arc3
'           - Arc4
  
  'Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations

  'Define a new input by new operator
     Dim I1 As IMSSymbolEntities.IJDInput
     Set I1 = New IMSSymbolEntities.DInput
   
     Dim I2 As IMSSymbolEntities.IJDInput
     Set I2 = New IMSSymbolEntities.DInput
   
     Dim I3 As IMSSymbolEntities.IJDInput
     Set I3 = New IMSSymbolEntities.DInput
  
     I1.Name = "Width"
     I1.Description = "Width of the rectangleCornerRadius"
     I1.Properties = igINPUT_IS_A_PARAMETER
 
  'Create a default value
     Dim PC As IMSSymbolEntities.DParameterContent
     Set PC = New IMSSymbolEntities.DParameterContent
  
     PC.Type = igValue
     PC.UomValue = 0.2
     I1.DefaultParameterValue = PC
  
     I2.Name = "Height"
     I2.Description = "Height of the rectangleCornerRadius"
     I2.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.1
     I2.DefaultParameterValue = PC
  
     I3.Name = "CornerRadius"
     I3.Description = "corner Radius of the rectangleCornerRadius"
     I3.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.02
     I3.DefaultParameterValue = PC
   
   'set the input to the definition
      Dim InputsIf As IMSSymbolEntities.IJDInputs
      Set InputsIf = m_SymbolDefinition
   
      InputsIf.SetInput I1, 1
      InputsIf.SetInput I2, 2
      InputsIf.SetInput I3, 3
  
  ' Define the representation "rectangleCornerRadius"
     Dim rep1 As IMSSymbolEntities.IJDRepresentation
     Set rep1 = New IMSSymbolEntities.DRepresentation
  
     rep1.Name = "rectangleCornerRadius"
     rep1.Description = "It's a rectangleCornerRadius"
     rep1.Properties = igREPRESENTATION_VBFUNCTION
    
  ' Create the output
      Dim O1 As IMSSymbolEntities.IJDOutput
      Set O1 = New IMSSymbolEntities.DOutput
  
      Dim O2 As IMSSymbolEntities.IJDOutput
      Set O2 = New IMSSymbolEntities.DOutput
  
      Dim O3 As IMSSymbolEntities.IJDOutput
      Set O3 = New IMSSymbolEntities.DOutput
  
      Dim O4 As IMSSymbolEntities.IJDOutput
      Set O4 = New IMSSymbolEntities.DOutput
  
      Dim O5 As IMSSymbolEntities.IJDOutput
      Set O5 = New IMSSymbolEntities.DOutput
  
      Dim O6 As IMSSymbolEntities.IJDOutput
      Set O6 = New IMSSymbolEntities.DOutput
  
      Dim O7 As IMSSymbolEntities.IJDOutput
      Set O7 = New IMSSymbolEntities.DOutput
  
      Dim O8 As IMSSymbolEntities.IJDOutput
      Set O8 = New IMSSymbolEntities.DOutput
  
       O1.Name = "Frame1"
       O1.Description = "It is a rectangleCornerRadius"
       O1.Properties = 0
  
       O2.Name = "Frame2"
       O2.Description = "It is a rectangleCornerRadius"
       O2.Properties = 0
  
       O3.Name = "Frame3"
       O3.Description = "It is a rectangleCornerRadius"
       O3.Properties = 0
  
       O4.Name = "Frame4"
       O4.Description = "It is a rectangleCornerRadius"
       O4.Properties = 0
  
       O5.Name = "Arc1"
       O5.Description = "It is a rectangleCornerRadius"
       O5.Properties = 0
  
       O6.Name = "Arc2"
       O6.Description = "It is a rectangleCornerRadius"
       O6.Properties = 0
  
       O7.Name = "Arc3"
       O7.Description = "It is a rectangleCornerRadius"
       O7.Properties = 0
  
       O8.Name = "Arc4"
       O8.Description = "It is a rectangleCornerRadius"
       O8.Properties = 0
  
  ' Set the output
       rep1.SetOutput O1
       rep1.SetOutput O2
       rep1.SetOutput O3
       rep1.SetOutput O4
       rep1.SetOutput O5
       rep1.SetOutput O6
       rep1.SetOutput O7
       rep1.SetOutput O8
  
  ' Set the representation to definition
     Dim RepsIf As IMSSymbolEntities.IJDRepresentations
     Set RepsIf = pSymbolDefinition
     RepsIf.SetRepresentation rep1
  
  ' Set the evaluation function associated to the rectangleCornerRadius representation
  
      Dim RepEvalsIf As IMSSymbolEntities.IJDRepresentationEvaluations
      Set RepEvalsIf = pSymbolDefinition
      Dim EvalFuncForrectangleCornerRadius As IJDRepresentationEvaluation
    
      Set EvalFuncForrectangleCornerRadius = New DRepresentationEvaluation
      EvalFuncForrectangleCornerRadius.Name = "rectangleCornerRadius"
      EvalFuncForrectangleCornerRadius.Description = "evaluation function for the rectangleCornerRadius representation"
      EvalFuncForrectangleCornerRadius.Properties = igREPRESENTATION_HIDDEN
      EvalFuncForrectangleCornerRadius.Type = igREPRESENTATION_VBFUNCTION
      EvalFuncForrectangleCornerRadius.ProgID = "OutfittingCrossSections.rectangleCornerRadius"
      RepEvalsIf.AddRepresentationEvaluation EvalFuncForrectangleCornerRadius
  
      Set O1 = Nothing
      Set O2 = Nothing
      Set O3 = Nothing
      Set O4 = Nothing
      Set O5 = Nothing
      Set O6 = Nothing
      Set O7 = Nothing
      Set O8 = Nothing
      Set RepEvalsIf = Nothing
      Set rep1 = Nothing
      Set RepsIf = Nothing
      Set EvalFuncForrectangleCornerRadius = Nothing
   
  Exit Sub

ErrorHandler:
    MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    Debug.Assert False


End Sub

' Draw the rectangleCornerRadius
  Sub rectangleCornerRadius(ByVal width As Double, ByVal height As Double, ByVal c_radius As Double)
      Dim m_width As Double
      Dim fact As New IMSEntityExamples.DEntityExampleFactory
      Dim Line1 As IJDLine
      Dim Line2 As IJDLine
      Dim line3 As IJDLine
      Dim line4 As IJDLine
      Dim startPt As IJDPosition
      Dim endPt As IJDPosition
    
      Dim coord(8, 3) As Double
    
' Calculate the coordinates of the 8 points of the constructed rectangleCornerRadius
    'First point
       coord(1, 1) = c_radius
       coord(1, 2) = 0#
       coord(1, 3) = 0#
    'Second point
       coord(2, 1) = width - c_radius
       coord(2, 2) = 0#
       coord(2, 3) = 0#
    'Third point
       coord(3, 1) = width
       coord(3, 2) = c_radius
       coord(3, 3) = 0#
    'Fourth point
       coord(4, 1) = width
       coord(4, 2) = height - c_radius
       coord(4, 3) = 0#
    'Fifth point
       coord(5, 1) = width - c_radius
       coord(5, 2) = height
       coord(5, 3) = 0#
    'Sixth point
       coord(6, 1) = c_radius
       coord(6, 2) = height
       coord(6, 3) = 0#
    'Seventh point
       coord(7, 1) = 0
       coord(7, 2) = height - c_radius
       coord(7, 3) = 0#
    'Eighth point
       coord(8, 1) = 0
       coord(8, 2) = c_radius
       coord(8, 3) = 0#
    
    Set startPt = Nothing
    Set endPt = Nothing
    
    ' Line 1
    Set Line1 = fact.CreateEntity(EntityFactoryDLine, m_outputColl.ResourceManager)
    Set startPt = Line1.GetStartPosition
    Set endPt = Line1.GetStartPosition
    
    startPt.x = coord(1, 1)
    startPt.y = coord(1, 2)
    startPt.z = coord(1, 3)
    endPt.x = coord(2, 1)
    endPt.y = coord(2, 2)
    endPt.z = coord(2, 3)
    
    Line1.PutStartPosition startPt
    Line1.PutEndPosition endPt
    
    m_outputColl.AddOutput "Frame1", Line1
    Set startPt = Nothing
    Set endPt = Nothing
    
    ' Line 2
    Set Line2 = fact.CreateEntity(EntityFactoryDLine, m_outputColl.ResourceManager)
    Set startPt = Line2.GetStartPosition
    Set endPt = Line2.GetStartPosition
        
    startPt.x = coord(3, 1)
    startPt.y = coord(3, 2)
    startPt.z = coord(3, 3)
    endPt.x = coord(4, 1)
    endPt.y = coord(4, 2)
    endPt.z = coord(4, 3)
    
    Line2.PutStartPosition startPt
    Line2.PutEndPosition endPt
    
    m_outputColl.AddOutput "Frame2", Line2
    Set startPt = Nothing
    Set endPt = Nothing

    ' Line 3
    Set line3 = fact.CreateEntity(EntityFactoryDLine, m_outputColl.ResourceManager)
    Set startPt = line3.GetStartPosition
    Set endPt = line3.GetStartPosition
        
    startPt.x = coord(5, 1)
    startPt.y = coord(5, 2)
    startPt.z = coord(5, 3)
    endPt.x = coord(6, 1)
    endPt.y = coord(6, 2)
    endPt.z = coord(6, 3)
    
    line3.PutStartPosition startPt
    line3.PutEndPosition endPt
    
    m_outputColl.AddOutput "Frame3", line3
    Set startPt = Nothing
    Set endPt = Nothing

      ' Line 4
    Set line4 = fact.CreateEntity(EntityFactoryDLine, m_outputColl.ResourceManager)
    Set startPt = line4.GetStartPosition
    Set endPt = line4.GetStartPosition
        
    startPt.x = coord(7, 1)
    startPt.y = coord(7, 2)
    startPt.z = coord(7, 3)
    endPt.x = coord(8, 1)
    endPt.y = coord(8, 2)
    endPt.z = coord(8, 3)
    
    line4.PutStartPosition startPt
    line4.PutEndPosition endPt
    
    m_outputColl.AddOutput "Frame4", line4
    Set startPt = Nothing
    Set endPt = Nothing

    Dim fact1 As New IngrGeom3D.GeometryFactory
    Dim Dummy As IngrGeom3D.Arc3d
    
    ' Drawing the Arcs
      Dim StartX As Double
      Dim StartY As Double
      Dim StartZ As Double
      Dim CenterX As Double
      Dim CenterY As Double
      Dim CenterZ As Double
      Dim EndX As Double
      Dim EndY As Double
      Dim EndZ As Double
    
   ' Arc1
      StartX = coord(8, 1)
      StartY = coord(8, 2)
      StartZ = coord(8, 3)
      CenterX = c_radius
      CenterY = c_radius
      CenterZ = 0
      EndX = coord(1, 1)
      EndY = coord(1, 2)
      EndZ = coord(1, 3)

      Set Dummy = fact1.Arcs3d.CreateByCenterStartEnd(m_outputColl.ResourceManager, CenterX, CenterY, CenterZ, StartX, StartY, StartZ, EndX, EndY, EndZ)
      m_outputColl.AddOutput "Arc1", Dummy
      
   ' Arc2
      StartX = coord(2, 1)
      StartY = coord(2, 2)
      StartZ = coord(2, 3)
      CenterX = width - c_radius
      CenterY = c_radius
      CenterZ = 0
      EndX = coord(3, 1)
      EndY = coord(3, 2)
      EndZ = coord(3, 3)
    
      Set Dummy = fact1.Arcs3d.CreateByCenterStartEnd(m_outputColl.ResourceManager, CenterX, CenterY, CenterZ, StartX, StartY, StartZ, EndX, EndY, EndZ)
      m_outputColl.AddOutput "Arc2", Dummy
    
   ' Arc3
      StartX = coord(4, 1)
      StartY = coord(4, 2)
      StartZ = coord(4, 3)
      CenterX = width - c_radius
      CenterY = height - c_radius
      CenterZ = 0
      EndX = coord(5, 1)
      EndY = coord(5, 2)
      EndZ = coord(5, 3)
    
      Set Dummy = fact1.Arcs3d.CreateByCenterStartEnd(m_outputColl.ResourceManager, CenterX, CenterY, CenterZ, StartX, StartY, StartZ, EndX, EndY, EndZ)
      m_outputColl.AddOutput "Arc3", Dummy

   ' Arc4
       StartX = coord(6, 1)
       StartY = coord(6, 2)
       StartZ = coord(6, 3)
       CenterX = c_radius
       CenterY = height - c_radius
       CenterZ = 0
       EndX = coord(7, 1)
       EndY = coord(7, 2)
       EndZ = coord(7, 3)
    
       Set Dummy = fact1.Arcs3d.CreateByCenterStartEnd(m_outputColl.ResourceManager, CenterX, CenterY, CenterZ, StartX, StartY, StartZ, EndX, EndY, EndZ)
       m_outputColl.AddOutput "Arc4", Dummy

End Sub

