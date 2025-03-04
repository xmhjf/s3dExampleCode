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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         Ramon Him
'   Creation Date:  July, 2007
'   Description:
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI       As Double
Private Const NEGLIGIBLE_THICKNESS = 0.0001

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    
    Dim parFacetoCenter As Double
    Dim parTangentLength As Double
    Dim parBendRadius As Double
    Dim parAngle As Double
    Dim parInsulationThickness As Double
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim ArcStPointToCenter  As Double
    Dim InsDiameterPort  As Double
    Dim InsulationRadius As Double
    
    Dim iOutput     As Double

    Dim objInsulatedTangent1 As Object
    Dim ObjInsulatedBody As Object
    Dim ObjInsulatedTangent2 As Object


' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)
    parTangentLength = arrayOfInputs(3)
    parBendRadius = arrayOfInputs(4)
    parAngle = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)

    pipeDiam = parFacetoCenter
    
    iOutput = 0
    
    Dim dTangentLength1 As Double
    Dim dTangentLength2 As Double
    Dim dElbowRadius As Double
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    Dim dAngle As Double
    dAngle = parAngle
      
   'TangentLength and BendRadius
    dTangentLength1 = parTangentLength
    dTangentLength2 = parTangentLength
    dElbowRadius = parBendRadius
    dFace1toCenter = parBendRadius * Tan(dAngle / 2) + parTangentLength
    dFace2toCenter = dFace1toCenter

    
    pipeDiam = parFacetoCenter
        
    ArcStPointToCenter = dFace1toCenter - dTangentLength1
    InsulationRadius = pipeDiam / 2 + parInsulationThickness
    stPoint.Set -dFace1toCenter, 0, 0
    enPoint.Set -ArcStPointToCenter, 0, 0

    Set objInsulatedTangent1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                                2 * InsulationRadius, True)
                                
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedTangent1
    Set objInsulatedTangent1 = Nothing

 ' Insert your code for output 3(Insulated Body)
    InsulationRadius = pipeDiam / 2 + parInsulationThickness
'   Ensure that the hub radius is not greater than the distance from the center
'   of rotation to the elbow center line
    If CmpDblGreaterthan(InsulationRadius, dElbowRadius) Then
        InsulationRadius = dElbowRadius
    End If

'   Construct a circle that will be used to create the revolution
    Dim objCircleI   As IngrGeom3D.Circle3d

    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set objCircleI = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                         -ArcStPointToCenter, 0, 0, _
                        1, 0, 0, _
                      InsulationRadius)
'   Revolve the circle about the Z axis passing by the origin
    Dim CP As New AutoMath.DPosition 'revolution center point
    Dim CV As New AutoMath.DVector 'rotation vector for rotation
    CP.Set -ArcStPointToCenter, dElbowRadius, 0
    CV.Set 0, 0, 1
    Set ObjInsulatedBody = PlaceRevolution(m_OutputColl, objCircleI, _
                    CV, CP, dAngle, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    Set geomFactory = Nothing
    Set CP = Nothing
    Set CV = Nothing
    Set objCircleI = Nothing

' Insert your code for output 5(Insulated Tangent At Port2)
    ArcStPointToCenter = dFace2toCenter - dTangentLength2
    stPoint.Set (dFace2toCenter) * Cos(dAngle), (dFace2toCenter) * Sin(dAngle), 0
    enPoint.Set ArcStPointToCenter * Cos(dAngle), ArcStPointToCenter * Sin(dAngle), 0
    InsulationRadius = pipeDiam / 2 + parInsulationThickness


    Set ObjInsulatedTangent2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                                2 * InsulationRadius, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedTangent2
    Set ObjInsulatedTangent2 = Nothing

    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

Private Sub Class_Terminate()
    '
End Sub




