VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RectangleCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*******************************************************************
'  Copyright (c) 2003, Intergraph Corporation.  All rights reserved.
'
'  Project: M:\Refdata\Symbols\OutfittingCrosssectionsOutfittingCrossSections.vbp
'  File:  M:\Refdata\Symbols\OutfittingCrosssections\RectangleCR.cls
'
'  Description:  For Drawing Rectangle with Corner Radius as specified.
'
'  Author: NVS
'
'  History:
'    -   Modified on 12/17/98
'     Removed IJDLine Object for Drawing Lines. Line3d Object is used to
'      Draw Lines . It only uses one Output Object ( ComplexString3d Object)
'
'   15th Sep, 1999 : PR
'     Fixed TR# 7632 and TR# 7633 : Unable to Place the Symbol for
'     the OutfittingCrossSections in the Route Environment.
'   27th Sep, 1999: APS [APS]
'     Took care of P2R2 symbol impact.
'     For setting outputs on a rep, one
'     needs to query for outputs from rep and set them.
'
'       SS Mar/08/2000
'           used JObjectCollection instead of UntransactedMiddleElems
'
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added/Updated.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Implements IJDUserSymbolServices

Private m_outputColl As IMSSymbolEntities.DOutputCollection

Private Sub Class_Terminate()
     Set m_outputColl = Nothing
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
    IJDUserSymbolServices_EditOccurence = False
End Function

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = "OutfittingCrossSections.RectangleCR"
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
  oSymbolDefinition.ProgId = "OutfittingCrossSections.RectangleCR"
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
    If StrComp(repName, "RectangleCR") = 0 Then
        RectangleCR arrayOfInputs(1), arrayOfInputs(2), arrayOfInputs(3)
    End If
End Sub

Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pSymbolDefinition As IJDSymbolDefinition)
   On Error GoTo ErrorHandler
' Feed RectangleCR Definition
' Inputs:
'     Width = 0.2
'     Height =0.1
'     CornerRadius=0.02
' The representation "RectangleCR" is composed of
'
'           - Frame1

  
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
     I1.Description = "Width of the RectangleCR"
     I1.Properties = igINPUT_IS_A_PARAMETER
 
  'Create a default value
     Dim PC As IMSSymbolEntities.DParameterContent
     Set PC = New IMSSymbolEntities.DParameterContent
  
     PC.Type = igValue
     PC.UomValue = 0.2
     I1.DefaultParameterValue = PC
  
     I2.Name = "Height"
     I2.Description = "Height of the RectangleCR"
     I2.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.1
     I2.DefaultParameterValue = PC
  
     I3.Name = "CornerRadius"
     I3.Description = "corner Radius of the RectangleCR"
     I3.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.02
     I3.DefaultParameterValue = PC
   
   'set the input to the definition
      Dim InputsIf As IMSSymbolEntities.IJDInputs
      Set InputsIf = pSymbolDefinition
   
      InputsIf.SetInput I1, 1
      InputsIf.SetInput I2, 2
      InputsIf.SetInput I3, 3
  
  ' Define the representation "RectangleCR"
     Dim rep1 As IMSSymbolEntities.IJDRepresentation
     Set rep1 = New IMSSymbolEntities.DRepresentation
  
     rep1.Name = "RectangleCR"
     rep1.Description = "It's a RectangleCR"
     rep1.Properties = igREPRESENTATION_ISVBFUNCTION
     rep1.RepresentationId = 1
    
  ' Create the output
      Dim O1 As IMSSymbolEntities.IJDOutput
      Set O1 = New IMSSymbolEntities.DOutput
  
       O1.Name = "Frame1"
       O1.Description = "It is a RectangleCR"
       O1.Properties = 0
  
     Dim oRepPhysicalOutputs As IMSSymbolEntities.IJDOutputs
     Set oRepPhysicalOutputs = rep1
     
  ' Set the output
       oRepPhysicalOutputs.SetOutput O1
       
  ' Set the representation to definition
     Dim RepsIf As IMSSymbolEntities.IJDRepresentations
     Set RepsIf = pSymbolDefinition
     RepsIf.SetRepresentation rep1
  
  ' Set the evaluation function associated to the RectangleCR representation
  
      Dim RepEvalsIf As IMSSymbolEntities.IJDRepresentationEvaluations
      Set RepEvalsIf = pSymbolDefinition
      Dim EvalFuncForRectangleCR As IJDRepresentationEvaluation
    
      Set EvalFuncForRectangleCR = New DRepresentationEvaluation
      EvalFuncForRectangleCR.Name = "RectangleCR"
      EvalFuncForRectangleCR.Description = "evaluation function for the RectangleCR representation"
      EvalFuncForRectangleCR.Properties = igREPRESENTATION_HIDDEN
      EvalFuncForRectangleCR.Type = igREPRESENTATION_VBFUNCTION
      EvalFuncForRectangleCR.ProgId = "OutfittingCrossSections.RectangleCR"
      RepEvalsIf.AddRepresentationEvaluation EvalFuncForRectangleCR
  
      Set O1 = Nothing
      
      Set RepEvalsIf = Nothing
      Set rep1 = Nothing
      Set oRepPhysicalOutputs = Nothing
      Set RepsIf = Nothing
      Set EvalFuncForRectangleCR = Nothing
   
  Exit Sub

ErrorHandler:
    Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    Debug.Assert False


End Sub

' Draw the RectangleCR
 Sub RectangleCR(ByVal width As Double, ByVal height As Double, ByVal c_radius As Double)
      
   Dim fact As New IngrGeom3D.GeometryFactory
   Dim Geometry As IngrGeom3D.ComplexString3d
   Dim oLine As IngrGeom3D.Line3d
   Dim Arc As IngrGeom3D.Arc3d
   Dim Element As IJElements
   
 '  Use untransacted elements as the elements collection as you are creating
 '  a transient object.You don't have the moniker
   Set Element = New JObjectCollection
   Dim Count As Long
   
   On Error GoTo DrawError

' Set up local temporaries  hh = half height, hw = half width, r = arc radius, ZZ = 0.0
     Dim hh As Double
     Dim hw As Double
     Dim r  As Double
     Dim ZZ As Double
     
     hh = height / 2#
     hw = width / 2#
     
     If (c_radius > 0.00001) Then
         r = c_radius
     Else
         r = 0.00001
     End If
    
     ZZ = 0#
   
 ' Defining Coordinates for Rectangle
    Dim coord(8, 2) As Double
    
    coord(1, 1) = hw
    coord(1, 2) = -hh + r
    
    coord(2, 1) = hw
    coord(2, 2) = hh - r
    
    coord(3, 1) = hw - r
    coord(3, 2) = hh
    
    coord(4, 1) = -hw + r
    coord(4, 2) = hh
    
    coord(5, 1) = -hw
    coord(5, 2) = hh - r
    
    coord(6, 1) = -hw
    coord(6, 2) = -hh + r
    
    coord(7, 1) = -hw + r
    coord(7, 2) = -hh
    
    coord(8, 1) = hw - r
    coord(8, 2) = -hh
    
' Segment 1:  Line
   Set oLine = fact.Lines3d.CreateBy2Points(Nothing, coord(1, 1), coord(1, 2), ZZ, coord(2, 1), coord(2, 2), ZZ)

' Add the Line3d object ( Line 1 )to the Element & Count gives No. of Line3d Objects
   Count = Element.Add(oLine)
   Set oLine = Nothing
   
' Segment 2:  Arc
   Set Arc = fact.Arcs3d.CreateByCenterStartEnd(Nothing, coord(3, 1), coord(2, 2), ZZ, coord(2, 1), coord(2, 2), ZZ, coord(3, 1), coord(3, 2), ZZ)
  
' Add the Arc3d object ( Arc 1 )to the Element & Count gives No. of Arc3d Objects
   Count = Element.Add(Arc)
   Set Arc = Nothing
    
' Segment 3:  Line
   Set oLine = fact.Lines3d.CreateBy2Points(Nothing, coord(3, 1), coord(3, 2), ZZ, coord(4, 1), coord(4, 2), ZZ)

' Add the Line3d object ( Line 2 )to the Element & Count gives No. of Line3d Objects
   Count = Element.Add(oLine)
   Set oLine = Nothing
   
' Segment 4:  Arc
   Set Arc = fact.Arcs3d.CreateByCenterStartEnd(Nothing, coord(4, 1), coord(5, 2), ZZ, coord(4, 1), coord(4, 2), ZZ, coord(5, 1), coord(5, 2), ZZ)
  
' Add the Arc3d object ( Arc 2 )to the Element & Count gives No. of Arc3d Objects
   Count = Element.Add(Arc)
   Set Arc = Nothing

' Segment 5:  Line
   Set oLine = fact.Lines3d.CreateBy2Points(Nothing, coord(5, 1), coord(5, 2), ZZ, coord(6, 1), coord(6, 2), ZZ)

' Add the Line3d object ( Line 3 )to the Element & Count gives No. of Line3d Objects
   Count = Element.Add(oLine)
   Set oLine = Nothing
   
' Segment 6:  Arc
   Set Arc = fact.Arcs3d.CreateByCenterStartEnd(Nothing, coord(7, 1), coord(6, 2), ZZ, coord(6, 1), coord(6, 2), ZZ, coord(7, 1), coord(7, 2), ZZ)
  
' Add the Arc3d object ( Arc 3 )to the Element & Count gives No. of Arc3d Objects
   Count = Element.Add(Arc)
   Set Arc = Nothing

' Segment 7:  Line
   Set oLine = fact.Lines3d.CreateBy2Points(Nothing, coord(7, 1), coord(7, 2), ZZ, coord(8, 1), coord(8, 2), ZZ)
   
' Add the Line3d object ( Line 4 )to the Element & Count gives No. of Line3d Objects
   Count = Element.Add(oLine)
   Set oLine = Nothing
   
' Segment 8:  Arc
   Set Arc = fact.Arcs3d.CreateByCenterStartEnd(Nothing, coord(8, 1), coord(1, 2), ZZ, coord(8, 1), coord(8, 2), ZZ, coord(1, 1), coord(1, 2), ZZ)
   
    
' Add the Arc3d object ( Arc 4 )to the Element & Count gives No. of Arc3d Objects
   Count = Element.Add(Arc)
   Set Arc = Nothing
      
'  Pass the Element ( with Line3d & Arc3d Object collection )as a Parameter to the CreateByCurves Method
     Set Geometry = fact.ComplexStrings3d.CreateByCurves(m_outputColl.ResourceManager, Element)
     m_outputColl.AddOutput "Frame1", Geometry
    
     Set Geometry = Nothing
     Set fact = Nothing
     Set Element = Nothing
     
Exit Sub
DrawError:
    Set Geometry = Nothing
    Set fact = Nothing
    Set oLine = Nothing
    Set Arc = Nothing
    Set Element = Nothing
    
    Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    Debug.Assert False
End Sub
