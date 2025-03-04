VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         Ramon Him
'   Creation Date:  July, 2007
'   Description:
'   Change History:
'   dd.mmm.yyyy    who       change description
'   -----------    -----     ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private PI       As Double
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
    
    Dim axisVect    As New AutoMath.DVector
    Dim ArcStPointToCenter  As Double
    Dim iCount As Integer
    
    Dim iOutput     As Double

    Dim ObjElbow As Object
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)
    parTangentLength = arrayOfInputs(3)
    parBendRadius = arrayOfInputs(4)
    parAngle = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)

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
    
'----------------------------------------
    
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory

' Insert your code for output 6(Elbow)
'    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, Depth
    pipeDiam = parFacetoCenter


    ArcStPointToCenter = dFace1toCenter - dTangentLength1
'   Construct a circle that will be used to create the revolution
    Dim objCircle   As IngrGeom3D.Circle3d
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                         -ArcStPointToCenter, 0, 0, _
                        1, 0, 0, _
                        pipeDiam / 2)
'   Revolve the circle about the Z axis
    Dim CP As New AutoMath.DPosition 'arc center point
    Dim CV As New AutoMath.DVector 'rotation vector for rotation
    CP.Set -ArcStPointToCenter, dElbowRadius, 0
    CV.Set 0, 0, 1
    Set ObjElbow = PlaceRevolution(m_OutputColl, objCircle, _
                    CV, CP, dAngle, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjElbow
    Set ObjElbow = Nothing
    Set CP = Nothing
    Set CV = Nothing
    Set objCircle = Nothing
    Set geomFactory = Nothing
    

'Place Hvac Nozzle 1
    Dim oHvacNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle

    Dim iDistribPort As GSCADNozzleEntities.IJDistribPort

    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim iPortIndex As Integer
       
    Dim oHvacPort As IJDHvacPort
    Dim oHvacColl As IJDCollection
    
    Dim lEndPrep(1 To 2) As Long
    Dim dThickness(1 To 2) As Double
    Dim dFlangeWidth(1 To 2) As Double
    Dim lFlowDir(1 To 2) As Long
    Dim dPortDepth(1 To 2) As Double
    Dim dCptOffSet(1 To 2) As Double
    Dim dNozzLength(1 To 2) As Double
    Dim CornerRadius As Double
    
    CornerRadius = 0#
    dNozzLength(1) = dTangentLength1
    dNozzLength(2) = dTangentLength2
    
    iPortIndex = 1
    'Set HVAC nozzle parameters
    Set oHvacColl = oPartFclt.GetNozzles()
    For iPortIndex = 1 To oHvacColl.Size
        Set oHvacPort = oHvacColl.Item(iPortIndex)
        lEndPrep(iPortIndex) = oHvacPort.EndPrep
        dThickness(iPortIndex) = oHvacPort.Thickness
        dFlangeWidth(iPortIndex) = oHvacPort.FlangeWidth
        lFlowDir(iPortIndex) = oHvacPort.FlowDirection
        dPortDepth(iPortIndex) = oHvacPort.PortDepth
        dCptOffSet(iPortIndex) = oHvacPort.CptOffset
        'NozzleLength Has to be GREATER than NozzleFlangeThickness
        If CmpDblLessThanOrEqualTo(dThickness(iPortIndex), LINEAR_TOLERANCE) Then
            dThickness(iPortIndex) = 0.0001
        End If
        If CmpDblLessThan(dNozzLength(iPortIndex), dThickness(iPortIndex)) Then
            dNozzLength(iPortIndex) = dThickness(iPortIndex) + 0.001
        End If
    Next iPortIndex
    
    Dim Depth As Double
    Depth = parFacetoCenter
    
    Set oHvacPort = Nothing
    oHvacColl.Clear
    Set oHvacColl = Nothing
    
    iPortIndex = 1
    PortStatus = DistribPortStatus_BASE
    
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, lEndPrep(1), _
                                            dThickness(1), dFlangeWidth(1), lFlowDir(1), parFacetoCenter, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz1", dPortDepth(1), dCptOffSet(1), False, m_OutputColl.ResourceManager)
    Set NozzleFactory = Nothing
    
    CP.Set -dFace1toCenter, 0, 0
    
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation CP
    Dim PortDirection As New AutoMath.DVector
    
    PortDirection.Set -1, 0, 0
    iDistribPort.SetDirectionVector PortDirection

    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(1)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing

' Place Nozzle 2
    PortDirection.Set Cos(dAngle), Sin(dAngle), 0
    CP.Set dFace2toCenter * Cos(dAngle), dFace2toCenter * Sin(dAngle), 0
    iPortIndex = 2
    
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, lEndPrep(2), _
                                            dThickness(2), dFlangeWidth(2), lFlowDir(2), parFacetoCenter, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "HNoz2", dPortDepth(2), dCptOffSet(2), False, m_OutputColl.ResourceManager)
     Set NozzleFactory = Nothing

     Set iDistribPort = oHvacNozzle
     iDistribPort.SetPortLocation CP

    iDistribPort.SetDirectionVector PortDirection
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(2)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    Set CP = Nothing
    Set PortDirection = Nothing
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub


