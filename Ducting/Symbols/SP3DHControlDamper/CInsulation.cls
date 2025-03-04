VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.

'   CInsulation.cls
'   Author:         GL
'   Creation Date:  Friday, Sep 05 2008
'   Description:    "Ingr SmartPlant 3D Symbol SP3DHControlDamper v 1.0"
'                   This symbol is created as per the CR-149510.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Insulation:" 'Used for error messages
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt               As PartFacelets.IJDPart
    Dim parWidth                As Double
    Dim parDepth                As Double
    Dim parDamperLength         As Double
    Dim parLeverLength          As Double
    Dim parCrankLength          As Double
    Dim parHVACShape            As Double
    Dim parInsulationThickness  As Double
    Dim parTransitionLength     As Double
    Dim parDamperWidth          As Double
    Dim parDamperHeight         As Double
    Dim parOffsetDamperHeight   As Double
    Dim parCrankDiameter        As Double
    Dim iCount                  As Integer
    Dim iOutput                 As Double
    Dim CornerRadius            As Double
    Dim oAxis                   As AutoMath.DVector
    Dim oArc                    As IngrGeom3D.Arc3d
    Dim arcs                    As Collection
    Dim oCString                As IngrGeom3D.ComplexString3d
    Dim oGeomFactory            As IngrGeom3D.GeometryFactory
    Dim oLine                   As Object
    Dim objInsDamperBody        As Object
    Dim objInsTransitionBody1   As Object
    Dim objInsTransitionBody2   As Object
    Dim oStPoint                As AutoMath.DPosition
    Dim oEnPoint                As AutoMath.DPosition
    
    CornerRadius = 0
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oAxis = New AutoMath.DVector
    Set arcs = New Collection
    Set oPartFclt = arrayOfInputs(1)
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition

    'Checking for the Part Data Basis Property
    Dim lPartDataBasis As Long
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPartFclt
    lPartDataBasis = oHvacPart.PartDataBasis
    Set oHvacPart = Nothing
    
'        parWidth = arrayOfInputs(2)
'        parDepth = arrayOfInputs(3)
'        parDamperLength = arrayOfInputs(4)
         parLeverLength = arrayOfInputs(5)
         parCrankLength = arrayOfInputs(6)
'        parHVACShape = arrayOfInputs(7)
         parInsulationThickness = arrayOfInputs(8)
'        parTransitionLength = arrayOfInputs(9)
'        parDamperWidth = arrayOfInputs(10)
'        parDamperHeight = arrayOfInputs(11)
'        parOffsetDamperHeight = arrayOfInputs(12)
         parCrankDiameter = arrayOfInputs(13)
'        parActuatorWidth = arrayOfInputs(14)
'        parActuatorDepth = arrayOfInputs(15)
'        parActuatorThickness = arrayOfInputs(16)
         
' Inputs
    Select Case lPartDataBasis
    
    Case Is <= 1, 156
        parWidth = arrayOfInputs(2)
        parDepth = arrayOfInputs(3)
        parDamperLength = arrayOfInputs(4)
        parHVACShape = arrayOfInputs(7)
        
    Case 157
        parWidth = arrayOfInputs(2)
        parDepth = arrayOfInputs(3)
        parDamperLength = arrayOfInputs(4)
        parHVACShape = arrayOfInputs(7)
    
    Case 158
        parWidth = arrayOfInputs(2)
        parDepth = arrayOfInputs(3)
        parDamperLength = arrayOfInputs(4)
        parHVACShape = arrayOfInputs(7)
    
    Case 159
    
        parWidth = arrayOfInputs(2)
        parDepth = arrayOfInputs(3)
        parDamperLength = arrayOfInputs(4)
        parHVACShape = arrayOfInputs(7)
        parTransitionLength = arrayOfInputs(9)
        parDamperWidth = arrayOfInputs(10)
        parDamperHeight = arrayOfInputs(11)
        parOffsetDamperHeight = arrayOfInputs(12)
        
    Case 160
        parWidth = arrayOfInputs(2)
        parDepth = arrayOfInputs(3)
        parDamperLength = arrayOfInputs(4)
        parHVACShape = arrayOfInputs(7)
        parTransitionLength = arrayOfInputs(9)
        parDamperWidth = arrayOfInputs(10)
        parDamperHeight = arrayOfInputs(11)
        parOffsetDamperHeight = arrayOfInputs(12)
        
    End Select
    
    If CmpDblEqual(parDepth, 0) Then
        parDepth = parWidth
    End If
    
    Select Case lPartDataBasis
    
    Case 156, 157, 158
       'Creating Insulation for Damper Body (Output 1)
        If (parHVACShape = 4) Then '(Round)
            oStPoint.Set (-parDamperLength / 2), 0, 0
            oEnPoint.Set (parDamperLength / 2), 0, 0
            parDepth = parWidth
            Set objInsDamperBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parWidth + (2 * parInsulationThickness), True)
        
        ElseIf (parHVACShape = Rectangular) Then
            oStPoint.Set (-parDamperLength / 2), (-parDepth / 2) - parInsulationThickness, _
                         (-parWidth / 2) - parInsulationThickness
            oEnPoint.Set (parDamperLength / 2), (parDepth / 2) + parInsulationThickness, _
                        (parWidth / 2) + parInsulationThickness
            Set objInsDamperBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
        Else  '(Oval)
            If CmpDblEqual(parWidth, parDepth) Then
            
                    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, -parDamperLength / 2, _
                               (parDepth / 2) + parInsulationThickness, ((parWidth - parDepth) / 2), _
                               -parDamperLength / 2, 0, (parWidth / 2) + parInsulationThickness, _
                               -parDamperLength / 2, (-parDepth / 2) - parInsulationThickness, (parWidth - parDepth) / 2)
                    arcs.Add oArc
            
                    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, -parDamperLength / 2, _
                               (-parDepth / 2) - parInsulationThickness, -(parWidth - parDepth) / 2, -parDamperLength / 2, _
                               0, (-parWidth / 2) - parInsulationThickness, -parDamperLength / 2, _
                               parDepth / 2 + parInsulationThickness, -(parWidth - parDepth) / 2)
                    arcs.Add oArc
            Else
                
                Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, -parDamperLength / 2, _
                           (parDepth / 2) + parInsulationThickness, ((parWidth - parDepth) / 2), _
                           -parDamperLength / 2, 0, (parWidth / 2) + parInsulationThickness, _
                           -parDamperLength / 2, (-parDepth / 2) - parInsulationThickness, (parWidth - parDepth) / 2)
                arcs.Add oArc
                
                Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -parDamperLength / 2, _
                             (-parDepth / 2) - parInsulationThickness, (parWidth - parDepth) / 2, -parDamperLength / 2, _
                             (-parDepth / 2) - parInsulationThickness, -(parWidth - parDepth) / 2)
                arcs.Add oLine
                
                Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, -parDamperLength / 2, _
                           (-parDepth / 2) - parInsulationThickness, -(parWidth - parDepth) / 2, -parDamperLength / 2, _
                           0, (-parWidth / 2) - parInsulationThickness, -parDamperLength / 2, _
                           parDepth / 2 + parInsulationThickness, -(parWidth - parDepth) / 2)
                arcs.Add oArc
                
                Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -parDamperLength / 2, _
                            (parDepth / 2) + parInsulationThickness, -(parWidth - parDepth) / 2, -parDamperLength / 2, _
                            (parDepth / 2) + parInsulationThickness, (parWidth - parDepth) / 2)
                arcs.Add oLine
            End If
            oAxis.Set 1, 0, 0
            oStPoint.Set -parDamperLength / 2, parDepth / 2 + parInsulationThickness, (parWidth - parDepth) / 2
            Set oCString = PlaceTrCString(oStPoint, arcs)
            Set objInsDamperBody = PlaceProjection(m_OutputColl, oCString, oAxis, parDamperLength, True)
            
            For iCount = 1 To arcs.Count
                arcs.Remove 1
            Next iCount
            Set oArc = Nothing
            Set arcs = Nothing
            Set oLine = Nothing
       End If
      'Set the Output
      iOutput = iOutput + 1
      m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsDamperBody
      Set objInsDamperBody = Nothing
    
      Set oStPoint = Nothing
      Set oEnPoint = Nothing
      Set oAxis = Nothing
      Set oCString = Nothing
      Set oGeomFactory = Nothing
    
    Case 159, 160
       'Creating Insulation for Transition body1
        If (parHVACShape = 4) Then '(Round)
            oStPoint.Set (-parDamperLength / 2) - parTransitionLength, 0, 0
            oEnPoint.Set -(parDamperLength / 2) - parInsulationThickness, 0, 0
            parDepth = parWidth
            Set objInsTransitionBody1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parWidth + (2 * parInsulationThickness), True)
        
        ElseIf (parHVACShape = Rectangular) Then
            oStPoint.Set (-parDamperLength / 2) - parTransitionLength, (-parDepth / 2) - parInsulationThickness, _
                         (-parWidth / 2) - parInsulationThickness
            oEnPoint.Set -(parDamperLength / 2) - parInsulationThickness, (parDepth / 2) + parInsulationThickness, _
                        (parWidth / 2) + parInsulationThickness
            Set objInsTransitionBody1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
        Else  '(Oval)
            If CmpDblEqual(parWidth, parDepth) Then
                Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, (-parDamperLength / 2) - parTransitionLength, _
                          (parDepth / 2) + parInsulationThickness, ((parWidth - parDepth) / 2), _
                           (-parDamperLength / 2) - parTransitionLength, 0, (parWidth / 2) + parInsulationThickness, _
                           (-parDamperLength / 2) - parTransitionLength, (-parDepth / 2) - parInsulationThickness, (parWidth - parDepth) / 2)
                arcs.Add oArc
 
                Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, (-parDamperLength / 2) - parTransitionLength, _
                           (-parDepth / 2) - parInsulationThickness, -(parWidth - parDepth) / 2, (-parDamperLength / 2) - parTransitionLength, _
                           0, (-parWidth / 2) - parInsulationThickness, (-parDamperLength / 2) - parTransitionLength, _
                           parDepth / 2 + parInsulationThickness, -(parWidth - parDepth) / 2)
                arcs.Add oArc
 
            Else
            
                Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, (-parDamperLength / 2) - parTransitionLength, _
                           (parDepth / 2) + parInsulationThickness, ((parWidth - parDepth) / 2), _
                           (-parDamperLength / 2) - parTransitionLength, 0, (parWidth / 2) + parInsulationThickness, _
                           (-parDamperLength / 2) - parTransitionLength, (-parDepth / 2) - parInsulationThickness, (parWidth - parDepth) / 2)
                arcs.Add oArc
                
                Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, (-parDamperLength / 2) - parTransitionLength, _
                             (-parDepth / 2) - parInsulationThickness, (parWidth - parDepth) / 2, (-parDamperLength / 2) - parTransitionLength, _
                             (-parDepth / 2) - parInsulationThickness, -(parWidth - parDepth) / 2)
                arcs.Add oLine
                
                Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, (-parDamperLength / 2) - parTransitionLength, _
                           (-parDepth / 2) - parInsulationThickness, -(parWidth - parDepth) / 2, (-parDamperLength / 2) - parTransitionLength, _
                           0, (-parWidth / 2) - parInsulationThickness, (-parDamperLength / 2) - parTransitionLength, _
                           parDepth / 2 + parInsulationThickness, -(parWidth - parDepth) / 2)
                arcs.Add oArc
                
                Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, (-parDamperLength / 2) - parTransitionLength, _
                            (parDepth / 2) + parInsulationThickness, -(parWidth - parDepth) / 2, (-parDamperLength / 2) - parTransitionLength, _
                            (parDepth / 2) + parInsulationThickness, (parWidth - parDepth) / 2)
                arcs.Add oLine
            End If
            oAxis.Set 1, 0, 0
            oStPoint.Set (-parDamperLength / 2) - parTransitionLength, parDepth / 2 + parInsulationThickness, (parWidth - parDepth) / 2
            Set oCString = PlaceTrCString(oStPoint, arcs)
            Set objInsTransitionBody1 = PlaceProjection(m_OutputColl, oCString, oAxis, parTransitionLength - parInsulationThickness, True)
       End If
      'Set the Output
       iOutput = iOutput + 1
       m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsTransitionBody1
       Set objInsTransitionBody1 = Nothing
    
      'Creating Insulation for Damper Body
       If CmpDblEqual(parOffsetDamperHeight, 0) Then
            oStPoint.Set -(parDamperLength / 2) - parInsulationThickness, -(parDamperHeight / 2) - parInsulationThickness, _
                         -(parDamperWidth / 2) - parInsulationThickness
            oEnPoint.Set (parDamperLength / 2) + parInsulationThickness, (parDamperHeight / 2) + parInsulationThickness, _
                         (parDamperWidth / 2) + parInsulationThickness
       Else
            oStPoint.Set -(parDamperLength / 2) - parInsulationThickness, -(parDepth / 2) - parOffsetDamperHeight - parInsulationThickness, _
                         -parDamperWidth / 2 - parInsulationThickness
            oEnPoint.Set (parDamperLength / 2) + parInsulationThickness, parDamperHeight - parOffsetDamperHeight - (parDepth / 2) _
                         + parInsulationThickness, parDamperWidth / 2 + parInsulationThickness
       End If
       Set objInsDamperBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
      'Set the output
       m_OutputColl.AddOutput "objInsDamperBody", objInsDamperBody
       Set objInsDamperBody = Nothing
       
      'Creating the Insulation for Transition Body2
       If (parHVACShape = 4) Then '(Round)
            oStPoint.Set (parDamperLength / 2) + parTransitionLength, 0, 0
            oEnPoint.Set (parDamperLength / 2) + parInsulationThickness, 0, 0
            Set objInsTransitionBody2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parWidth + (2 * parInsulationThickness), True)
        
       ElseIf (parHVACShape = Rectangular) Then
            oStPoint.Set (parDamperLength / 2) + parTransitionLength, (-parDepth / 2) - parInsulationThickness, _
                         (-parWidth / 2) - parInsulationThickness
            oEnPoint.Set (parDamperLength / 2) + parInsulationThickness, (parDepth / 2) + parInsulationThickness, _
                        (parWidth / 2) + parInsulationThickness
            Set objInsTransitionBody2 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
       Else  '(Oval)
            Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, (parDamperLength / 2) + parTransitionLength, _
                       (parDepth / 2) + parInsulationThickness, ((parWidth - parDepth) / 2), _
                       (parDamperLength / 2) + parTransitionLength, 0, (parWidth / 2) + parInsulationThickness, _
                       (parDamperLength / 2) + parTransitionLength, (-parDepth / 2) - parInsulationThickness, _
                       (parWidth - parDepth) / 2)
            arcs.Add oArc
            
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, (parDamperLength / 2) + parTransitionLength, _
                         (-parDepth / 2) - parInsulationThickness, (parWidth - parDepth) / 2, (parDamperLength / 2) _
                         + parTransitionLength, (-parDepth / 2) - parInsulationThickness, -(parWidth - parDepth) / 2)
            arcs.Add oLine
            
            Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, (parDamperLength / 2) + parTransitionLength, _
                       (-parDepth / 2) - parInsulationThickness, -(parWidth - parDepth) / 2, (parDamperLength / 2) _
                       + parTransitionLength, 0, (-parWidth / 2) - parInsulationThickness, (parDamperLength / 2) _
                       + parTransitionLength, parDepth / 2 + parInsulationThickness, -(parWidth - parDepth) / 2)
            arcs.Add oArc
            
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, (parDamperLength / 2) + parTransitionLength, _
                        (parDepth / 2) + parInsulationThickness, -(parWidth - parDepth) / 2, (parDamperLength / 2) + _
                        parTransitionLength, (parDepth / 2) + parInsulationThickness, (parWidth - parDepth) / 2)
            arcs.Add oLine
            
            oAxis.Set -1, 0, 0
            oStPoint.Set (parDamperLength / 2) + parTransitionLength, parDepth / 2 + parInsulationThickness, (parWidth - parDepth) / 2
            Set oCString = PlaceTrCString(oStPoint, arcs)
            Set objInsTransitionBody2 = PlaceProjection(m_OutputColl, oCString, oAxis, parTransitionLength - parInsulationThickness, True)
       End If
      'Set the Output
       m_OutputColl.AddOutput "objInsTransitionBody2", objInsTransitionBody2
       Set objInsTransitionBody2 = Nothing
       
       For iCount = 1 To arcs.Count
           arcs.Remove 1
       Next iCount
       Set oArc = Nothing
       Set arcs = Nothing
       Set oLine = Nothing
       Set oStPoint = Nothing
       Set oEnPoint = Nothing
       Set oAxis = Nothing
       Set oCString = Nothing
       Set oGeomFactory = Nothing
End Select
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError2 MODULE, METHOD
End Sub
