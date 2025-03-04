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
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         RRC
'   Creation Date:  Tuesday, 13, nov 2007
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    The symbol is modelled per Model No VC 100 of Hayward Marine Catalog.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ----   ---------------------------------------
'   13.nov.2007     RRC  CR-131696  Created the symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim iOutput     As Double

    Dim parScupperHeight As Double
    Dim parOuterDiameter As Double
    Dim parCylinderHeight As Double
    Dim parInsulationThickness As Double
    Dim parInnerDiameter As Double
    Dim dCylinderHeight As Double
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parScupperHeight = arrayOfInputs(2)
    parOuterDiameter = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    iOutput = 0
'Assumption : Origin is at Port 1.


  Dim oPipeComponent As IJDPipeComponent
  Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
     RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, _
                                                flangeDiam, sptOffset, depth
    Dim oStPoint As AutoMath.DPosition
    Dim oEndPoint As AutoMath.DPosition
    
    Set oStPoint = New AutoMath.DPosition
    
    
    Dim objBottomCylinder As Object
    
    If (lPartDataBasis <> 501) Then
            dCylinderHeight = 0.3 * parScupperHeight
            Set oEndPoint = New AutoMath.DPosition
            oStPoint.Set 0, 0, 0  'origin point
            oEndPoint.Set dCylinderHeight, 0, 0
            Set objBottomCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEndPoint, pipeDiam + 2 * parInsulationThickness, True)
         
            
    ElseIf (lPartDataBasis = 501) Then
            parCylinderHeight = arrayOfInputs(3)
            parInnerDiameter = arrayOfInputs(4)
            dCylinderHeight = parScupperHeight - parCylinderHeight
            oStPoint.Set -parInsulationThickness, 0, 0  'origin point
            Set oEndPoint = New AutoMath.DPosition
            oEndPoint.Set dCylinderHeight, 0, 0
          Dim d2Thk As Double
        d2Thk = (parOuterDiameter - parInnerDiameter)
         Set objBottomCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEndPoint, pipeDiam + d2Thk + 2 * parInsulationThickness, True)
             
        
    End If
     
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBottomCylinder
    Set objBottomCylinder = Nothing
    
                                                 
     Dim objMiddlePortion As Object
     If (lPartDataBasis <> 501) Then
        oStPoint.Set 0.3 * parScupperHeight, 0, 0
        oEndPoint.Set 0.7 * parScupperHeight, 0, 0
         
        Set objMiddlePortion = PlaceCone(m_OutputColl, oStPoint, oEndPoint, 0.5 * pipeDiam + parInsulationThickness, 0.5 * parOuterDiameter + parInsulationThickness, True)
              
     ElseIf (lPartDataBasis = 501) Then
    
        Dim oGeomFactory     As IngrGeom3D.GeometryFactory
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        
        Dim oArc As IngrGeom3D.Arc3d
        Dim oCenter As IJDPosition
        Set oCenter = New DPosition
        Dim dRadius  As Double
         
        dRadius = (parOuterDiameter - pipeDiam) / 4
        oCenter.Set dCylinderHeight, ((0.5 * parOuterDiameter) - dRadius), 0
        Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, oCenter.x, oCenter.y + dRadius + parInsulationThickness, oCenter.z, oCenter.x - dRadius - parInsulationThickness, oCenter.y, oCenter.z, oCenter.x, oCenter.y - dRadius - parInsulationThickness, oCenter.z)
        
        Dim oAxisVec As AutoMath.DVector
        Set oAxisVec = New AutoMath.DVector
        oAxisVec.Set 1, 0, 0
        Dim oCenPoint As AutoMath.DPosition
        Set oCenPoint = New AutoMath.DPosition
        oCenPoint.Set 0, 0, 0
    
        Set objMiddlePortion = PlaceRevolution(m_OutputColl, oArc, oAxisVec, _
                                                    oCenPoint, 8 * Atn(1), True)
        Set oAxisVec = Nothing
        Set oArc = Nothing
           
    End If
                                               
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objMiddlePortion
    Set objMiddlePortion = Nothing
    
    'Insert your code for Top Cylinder
     
    Dim objTopCylinder As Object
    
    If (lPartDataBasis <> 501) Then
        oStPoint.Set 0.7 * parScupperHeight, 0, 0
        oEndPoint.Set parScupperHeight, 0, 0
    
    ElseIf (lPartDataBasis = 501) Then
        oStPoint.Set dCylinderHeight, 0, 0
        oEndPoint.Set parScupperHeight, 0, 0
   
    End If
     Set objTopCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEndPoint, parOuterDiameter + 2 * parInsulationThickness, True)
   
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTopCylinder
    Set objTopCylinder = Nothing
    Set oStPoint = Nothing
    Set oEndPoint = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
