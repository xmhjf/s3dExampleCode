VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   ProgID:         SP3DHRndReducer.HRndReducer
'   Author:         kkk
'   Creation Date:  Thursday, Feb 09 2006
'   Description:
'   Physical aspect of HVAC Round Reducer
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  23.Jan.2008     VRK  TR-125293 Incorrect Dimensions of various components in HVAC(Lindab Catalog)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:"    'Used for error messages
Private Sub Class_Initialize()
'''
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim iOutput As Double
    Dim parWidth As Double
    Dim parBWidth As Double
    Dim parLength As Double
    Dim parInsulationThickness As Double
    Dim parPlaneOfBranch As Double
    Dim ObjBody As Object
    Dim TakeOff As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parBWidth = arrayOfInputs(3)
    parLength = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)

    iOutput = 0

    Dim dOffsetBetweenCentreLines As Double
    TakeOff = 2 * 25.4 / 1000    'Considering Take-Off as 2 Inches on all the Four sides of the Cross

    'Checking for the Part Data Basis Property
    Dim lPartDataBasis As Long
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPartFclt
    lPartDataBasis = oHvacPart.PartDataBasis
    Set oHvacPart = Nothing

    ' a)Round Reducer, conical-133
    ' b)Round Reducer, eccentric-134
    ' c)Round Reducer, Stepped-135

    Select Case lPartDataBasis

    Case Is <= 1, 133 'Round Reducer, conical

        dOffsetBetweenCentreLines = 0
        'Insert your code for output 1(Reducer)
        Dim stPoint As New AutoMath.DPosition
        Dim enPoint As New AutoMath.DPosition
        stPoint.Set -parLength / 2, 0, 0
        enPoint.Set parLength / 2, 0, 0
        Set ObjBody = PlaceCone(m_OutputColl, stPoint, enPoint, parWidth / 2, parBWidth / 2, False)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
        Set ObjBody = Nothing

    Case 134 'Round Reducer, eccentric

        dOffsetBetweenCentreLines = (parWidth - parBWidth) / 2
        Dim cirNormalX As Double
        Dim cirNormalY As Double
        Dim cirNormalZ As Double

        cirNormalX = 1
        cirNormalY = 0
        cirNormalZ = 0

        Dim objCircle1 As IngrGeom3D.Circle3d
        Dim cir1CenterX As Double
        Dim cir1CenterY As Double
        Dim cir1CenterZ As Double

        cir1CenterX = -parLength / 2
        cir1CenterY = 0
        cir1CenterZ = 0
        Dim geomFactory As New IngrGeom3D.GeometryFactory
        Set objCircle1 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                         Nothing, cir1CenterX, _
                         cir1CenterY, cir1CenterZ, cirNormalX, _
                         cirNormalY, cirNormalZ, parWidth * 0.5)


        'Create Circle2
        Dim objCircle2 As IngrGeom3D.Circle3d
        Dim cir2CenterX As Double
        Dim cir2CenterY As Double
        Dim cir2CenterZ As Double

        cir2CenterX = parLength / 2
        cir2CenterY = -dOffsetBetweenCentreLines
        cir2CenterZ = 0

        Set objCircle2 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                         Nothing, cir2CenterX, _
                         cir2CenterY, cir2CenterZ, cirNormalX, _
                         cirNormalY, cirNormalZ, parBWidth * 0.5)

        'Create Ruled Surface for body
        Dim ObjRuledBody As IngrGeom3D.RuledSurface3d
        Set ObjRuledBody = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                                      objCircle1, objCircle2, True)
        Set ObjBody = ObjRuledBody
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
        Set ObjBody = Nothing
        Set ObjRuledBody = Nothing
        Set geomFactory = Nothing
        ' remove curve
        Dim sp3dElem1 As IJDObject
        Dim sp3dElem2 As IJDObject

        Set sp3dElem1 = objCircle1
        Set sp3dElem2 = objCircle2
        Set objCircle1 = Nothing
        Set objCircle2 = Nothing
        sp3dElem1.Remove
        sp3dElem2.Remove

    Case 135 'Round Reducer, Stepped

        dOffsetBetweenCentreLines = 0
        stPoint.Set -parLength / 2, 0, 0
        enPoint.Set 0, 0, 0
        Set ObjBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parWidth, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
        Set ObjBody = Nothing

        stPoint.Set 0, 0, 0
        enPoint.Set parLength / 2, 0, 0
        Set ObjBody = PlaceCone(m_OutputColl, stPoint, enPoint, parWidth / 2, parBWidth / 2, False)
        ' Set the output
        m_OutputColl.AddOutput "ObjBody", ObjBody
        Set ObjBody = Nothing

    Case Else
        GoTo ErrorLabel:    'Invalid specification.

    End Select

    'Checking for Round recducer case:Ring1 and TakeOff2
    Dim ObjRing1 As Object
    Dim ObjTakeOff1 As Object
    Dim ObjRing2 As Object
    Dim ObjTakeOff2 As Object

    If lPartDataBasis = 135 Then

        '============================================
        'BUILD OuterRing 1 Of Reducer
        '============================================
        stPoint.Set -parLength / 2 - 4 / 1000, 0, 0
        enPoint.Set -parLength / 2 + 4 / 1000, 0, 0
        Set ObjRing1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parWidth * 1.06, True)
        'Set the output
        m_OutputColl.AddOutput "ObjRing1", ObjRing1
        Set ObjRing1 = Nothing

        '============================================
        'BUILD TakeOff 2 Of Reducer
        '============================================
        stPoint.Set parLength / 2, 0, 0
        enPoint.Set parLength / 2 + TakeOff, 0, 0
        Set ObjTakeOff2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parBWidth * 1.01, False)
        'Set the output
        m_OutputColl.AddOutput "ObjTakeOff2", ObjTakeOff2
        Set ObjTakeOff2 = Nothing

    ElseIf lPartDataBasis = 1 Or lPartDataBasis = 133 Or lPartDataBasis = 134 Then

        '============================================
        'BUILD TakeOff 1 Of Reducer
        '============================================
        stPoint.Set -(parLength / 2 + TakeOff), 0, 0
        enPoint.Set -parLength / 2, 0, 0
        Set ObjTakeOff1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parWidth * 1.01, False)
        'Set the output
        m_OutputColl.AddOutput "ObjTakeOff1", ObjTakeOff1
        Set ObjTakeOff1 = Nothing

        '============================================
        'BUILD OuterRing 1 Of Reducer
        '============================================
        stPoint.Set -parLength / 2 - 4 / 1000, 0, 0
        enPoint.Set -parLength / 2 + 4 / 1000, 0, 0
        Set ObjRing1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parWidth * 1.06, True)
        'Set the output
        m_OutputColl.AddOutput "ObjRing1", ObjRing1
        Set ObjRing1 = Nothing

        '============================================
        'BUILD TakeOff 2 Of Reducer
        '============================================
        stPoint.Set parLength / 2, -dOffsetBetweenCentreLines, 0
        enPoint.Set parLength / 2 + TakeOff, -dOffsetBetweenCentreLines, 0
        Set ObjTakeOff2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parBWidth * 1.01, False)
        'Set the output
        m_OutputColl.AddOutput "ObjTakeOff2", ObjTakeOff2
        Set ObjTakeOff2 = Nothing

        '============================================
        'BUILD OuterRing 2 Of Reducer
        '============================================
        stPoint.Set parLength / 2 - 4 / 1000, -dOffsetBetweenCentreLines, 0
        enPoint.Set parLength / 2 + 4 / 1000, -dOffsetBetweenCentreLines, 0
        Set ObjRing2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parBWidth * 1.06, True)
        'Set the output
        m_OutputColl.AddOutput "ObjRing2", ObjRing2
        Set ObjRing2 = Nothing

    End If
    Set stPoint = Nothing
    Set enPoint = Nothing

    'Insert your code for output 1(HvacPort1 )
    
    Dim oNozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim oNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oDistribPort As GSCADNozzleEntities.IJDistribPort
    
    ' HVAC parameters for Nozzle
    Dim bDimBaseOuter As Boolean
    Dim oPos As New AutoMath.DPosition
    Dim oDir As New AutoMath.DVector
    Dim iPortIndex As Integer
    Dim dCornerRadius As Double

    Dim PortStatus As DistribPortStatus
    Dim oHvacPort As IJDHvacPort
    Dim oHvacColl As IJDCollection
    
    Dim lEndPrep(1 To 2) As Long
    Dim dThickness(1 To 2) As Double
    Dim dFlangeWidth(1 To 2) As Double
    Dim lFlowDir(1 To 2) As Long
    Dim dPortDepth(1 To 2) As Double
    Dim dCptOffSet(1 To 2) As Double
    Dim dNozzLength(1 To 2) As Double

    dCornerRadius = 0#
    dNozzLength(1) = 0
    dNozzLength(2) = 0
    
    'Set HVAC nozzle parameters from Catalog
    Set oHvacColl = oPartFclt.GetNozzles()
    For iPortIndex = 1 To oHvacColl.Size
        Set oHvacPort = oHvacColl.Item(iPortIndex)
        lEndPrep(iPortIndex) = oHvacPort.EndPrep
        dThickness(iPortIndex) = oHvacPort.Thickness
        dFlangeWidth(iPortIndex) = oHvacPort.FlangeWidth
        lFlowDir(iPortIndex) = oHvacPort.FlowDirection
        dPortDepth(iPortIndex) = oHvacPort.PortDepth
        dCptOffSet(iPortIndex) = oHvacPort.CptOffset
    Next iPortIndex
        
    Set oHvacPort = Nothing
    oHvacColl.Clear
    Set oHvacColl = Nothing
    
    iPortIndex = 1
    bDimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    Set oHvacNozzle = oNozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, lEndPrep(1), _
                                            dThickness(1), dFlangeWidth(1), lFlowDir(1), parWidth, _
                                            parWidth, dCornerRadius, bDimBaseOuter, PortStatus, _
                                            "HvacNozzle1", dPortDepth(1), dCptOffSet(1), False, m_OutputColl.ResourceManager)

    'Position of the nozzle should be the conenct point of the nozzle
    oPos.Set -parLength / 2, 0, 0
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation oPos

    'Direction specified here of the nozzle should be the direction in which Duct will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

    oDir.Set 0, 0, 1
    oDistribPort.SetRadialOrient oDir

    Set oNozzle = oHvacNozzle
    oNozzle.Length = dNozzLength(1)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set oNozzle = Nothing
    Set oDistribPort = Nothing

    'Insert your code for output 3(HVAC Nozzle - 2)
    iPortIndex = 2
    Set oHvacNozzle = oNozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, lEndPrep(2), _
                                            dThickness(2), dFlangeWidth(2), lFlowDir(2), parBWidth, _
                                            parBWidth, dCornerRadius, bDimBaseOuter, PortStatus, _
                                            "HvacNozzle2", dPortDepth(2), dCptOffSet(2), False, m_OutputColl.ResourceManager)
    
    'Position of the nozzle should be the connect point of the nozzle
    oPos.Set parLength / 2, -dOffsetBetweenCentreLines, 0
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation oPos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.

    oDir.Set 1, 0, 0
    oDistribPort.SetDirectionVector oDir

    oDir.Set 0, 0, 1
    oDistribPort.SetRadialOrient oDir

    Set oNozzle = oHvacNozzle
    oNozzle.Length = dNozzLength(2)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle

    Set oHvacNozzle = Nothing
    Set oNozzle = Nothing
    Set oNozzleFactory = Nothing
    Set oPos = Nothing
    Set oDir = Nothing
    Set oDistribPort = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub

