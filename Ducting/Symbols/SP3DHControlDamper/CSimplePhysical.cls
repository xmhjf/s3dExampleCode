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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.

'   CSimplePhysical.cls
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

Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private m_oEditErrors As IJEditErrors
Private Const NEGLIGIBLE_THICKNESS = 0.0001
Private PI       As Double
Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
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
    Dim parActuatorWidth        As Double
    Dim parActuatorDepth        As Double
    Dim parActuatorThickness    As Double
    Dim iOutput                 As Integer
    Dim iCount                  As Integer
    Dim CornerRadius            As Double
    Dim dKnobLength             As Double
    Dim dFrameWidth             As Double
    Dim dFrameThickness         As Double
    Dim dFrameDepth             As Double
    Dim dLeverWidth             As Double
    Dim dLeverLength            As Double
    Dim dCrankDiameter          As Double
    Dim dCrankLength            As Double
    Dim dGearStemLength         As Double
    Dim dOffsetDamperHeight     As Double

    Dim oStPoint                As AutoMath.DPosition
    Dim oEnPoint                As AutoMath.DPosition
    Dim oAxis                   As AutoMath.DVector
    Dim oPartFclt               As PartFacelets.IJDPart
    Dim oCString                As IngrGeom3D.ComplexString3d
    Dim oGeomFactory            As IngrGeom3D.GeometryFactory
    Dim arcs                    As Collection
    Dim objDamperBody           As Object
    Dim objCrank                As Object
    Dim objLever                As Object
    Dim objTransitionBody       As Object
    Dim oLine                   As Object
    Dim oArc                    As Object
    Dim oKnob1                  As Object
    Dim oKnob2                  As Object
    
    CornerRadius = 0
    Set oAxis = New AutoMath.DVector
    Set oPartFclt = arrayOfInputs(1)
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set arcs = New Collection

   'Checking for the Part Data Basis Property
    Dim lPartDataBasis As Long
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPartFclt
    lPartDataBasis = oHvacPart.PartDataBasis
    Set oHvacPart = Nothing
    
    'Partdatbasis implemented in SP3DHControlDamper
    
        '156 - Control Damper with Lever type Operator
        '157 - Control Damper with gear type operator
        '158 - Control damper with an Actuator
        '159 - Control damper with transition with lever type operator
        '160 - Control damper with transition with an actuator
        
    'Inputs
'        parWidth = arrayOfInputs(2)
'        parDepth = arrayOfInputs(3)
'        parDamperLength = arrayOfInputs(4)
'        parLeverLength = arrayOfInputs(5)
'        parCrankLength = arrayOfInputs(6)
'        parHVACShape = arrayOfInputs(7)
         parInsulationThickness = arrayOfInputs(8)
'        parTransitionLength = arrayOfInputs(9)
'        parDamperWidth = arrayOfInputs(10)
'        parDamperHeight = arrayOfInputs(11)
'        parOffsetDamperHeight = arrayOfInputs(12)
'        parCrankDiameter = arrayOfInputs(13)
'        parActuatorWidth = arrayOfInputs(14)
'        parActuatorDepth = arrayOfInputs(15)
'        parActuatorThickness = arrayOfInputs(16)

    Select Case lPartDataBasis
    
    Case Is <= 1, 156
        parWidth = arrayOfInputs(2)
        parDepth = arrayOfInputs(3)
        parDamperLength = arrayOfInputs(4)
        parLeverLength = arrayOfInputs(5)
        parCrankLength = arrayOfInputs(6)
        parHVACShape = arrayOfInputs(7)
        parCrankDiameter = arrayOfInputs(13)
        
    Case 157
        parWidth = arrayOfInputs(2)
        parDepth = arrayOfInputs(3)
        parDamperLength = arrayOfInputs(4)
        parLeverLength = arrayOfInputs(5)
        parCrankLength = arrayOfInputs(6)
        parHVACShape = arrayOfInputs(7)
    
    Case 158
        parWidth = arrayOfInputs(2)
        parDepth = arrayOfInputs(3)
        parDamperLength = arrayOfInputs(4)
        parLeverLength = arrayOfInputs(5)
        parCrankLength = arrayOfInputs(6)
        parHVACShape = arrayOfInputs(7)
        parActuatorWidth = arrayOfInputs(14)
        parActuatorDepth = arrayOfInputs(15)
        parActuatorThickness = arrayOfInputs(16)
    
    Case 159
    
        parWidth = arrayOfInputs(2)
        parDepth = arrayOfInputs(3)
        parDamperLength = arrayOfInputs(4)
        parLeverLength = arrayOfInputs(5)
        parCrankLength = arrayOfInputs(6)
        parHVACShape = arrayOfInputs(7)
        parTransitionLength = arrayOfInputs(9)
        parDamperWidth = arrayOfInputs(10)
        parDamperHeight = arrayOfInputs(11)
        parOffsetDamperHeight = arrayOfInputs(12)
        parCrankDiameter = arrayOfInputs(13)
        
    Case 160
        parWidth = arrayOfInputs(2)
        parDepth = arrayOfInputs(3)
        parDamperLength = arrayOfInputs(4)
        parLeverLength = arrayOfInputs(5)
        parCrankLength = arrayOfInputs(6)
        parHVACShape = arrayOfInputs(7)
        parTransitionLength = arrayOfInputs(9)
        parDamperWidth = arrayOfInputs(10)
        parDamperHeight = arrayOfInputs(11)
        parOffsetDamperHeight = arrayOfInputs(12)
        parActuatorWidth = arrayOfInputs(14)
        parActuatorDepth = arrayOfInputs(15)
        parActuatorThickness = arrayOfInputs(16)
        
    End Select
    
    If CmpDblEqual(parDepth, 0) Then
        parDepth = parWidth
    End If

    'Assumed values of various parameters for different partdatabasis
    'Crank dimensions for the Lever Type
    If CmpDblEqual(parCrankDiameter, 0) Then
        dCrankDiameter = 0.2 * parDepth
    Else
        dCrankDiameter = parCrankDiameter
    End If
    If CmpDblEqual(parCrankLength, 0) Then
        dCrankLength = 0.25 * parWidth
    Else
        dCrankLength = parCrankLength
    End If
    
    'Assumed Frame and Lever dimensions for the Gear Type and actuator
    If CmpDblEqual(parActuatorWidth, 0) Then
        dFrameWidth = 0.25 * parDamperLength
    Else
        dFrameWidth = parActuatorWidth
    End If
    If CmpDblEqual(parActuatorThickness, 0) Then
        dFrameThickness = 0.2 * parDamperLength
    Else
        dFrameThickness = parActuatorThickness
    End If
    If CmpDblEqual(parActuatorDepth, 0) Then
        dFrameDepth = 0.5 * parDepth
        If lPartDataBasis = 159 Or lPartDataBasis = 160 Then
            dFrameDepth = 0.5 * parDamperHeight
        End If
    Else
        dFrameDepth = parActuatorDepth
    End If
    
    dLeverWidth = 1.25 * dCrankDiameter
    If CmpDblEqual(parLeverLength, 0) Then
        dLeverLength = 0.2 * parDamperLength
    Else
        dLeverLength = parLeverLength
    End If
    dKnobLength = 0.2 * dLeverLength
    dGearStemLength = 0.05 * parDamperLength

    Select Case lPartDataBasis
    
    Case Is <= 1, 156   'Control Damper with Lever type Operator
        
       'Creating Output 1 (Damper Body)
        If (parHVACShape = 4) Then '(Round)
            oStPoint.Set -parDamperLength / 2, 0, 0
            oEnPoint.Set parDamperLength / 2, 0, 0
            parDepth = parWidth
            Set objDamperBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parWidth, True)
        
        ElseIf (parHVACShape = Rectangular) Then
            oStPoint.Set -parDamperLength / 2, -parDepth / 2, -parWidth / 2
            oEnPoint.Set parDamperLength / 2, parDepth / 2, parWidth / 2
            Set objDamperBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
        Else  '(Oval)
            If CmpDblEqual(parWidth, parDepth) Then
                Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, -parDamperLength / 2, _
                           parDepth / 2, (parWidth - parDepth) / 2, -parDamperLength / 2, 0, parWidth / 2, _
                           -parDamperLength / 2, -parDepth / 2, (parWidth - parDepth) / 2)
                arcs.Add oArc
            
                Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, -parDamperLength / 2, _
                           -parDepth / 2, -(parWidth - parDepth) / 2, -parDamperLength / 2, _
                           0, -parWidth / 2, -parDamperLength / 2, parDepth / 2, -(parWidth - parDepth) / 2)
                arcs.Add oArc
            Else
            
                Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, -parDamperLength / 2, _
                           parDepth / 2, (parWidth - parDepth) / 2, -parDamperLength / 2, 0, parWidth / 2, _
                           -parDamperLength / 2, -parDepth / 2, (parWidth - parDepth) / 2)
                arcs.Add oArc
            
                Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -parDamperLength / 2, _
                            -parDepth / 2, (parWidth - parDepth) / 2, -parDamperLength / 2, _
                            -parDepth / 2, -(parWidth - parDepth) / 2)
                arcs.Add oLine
            
                Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, -parDamperLength / 2, _
                          -parDepth / 2, -(parWidth - parDepth) / 2, -parDamperLength / 2, _
                          0, -parWidth / 2, -parDamperLength / 2, parDepth / 2, -(parWidth - parDepth) / 2)
                arcs.Add oArc
            
                Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, -parDamperLength / 2, _
                            parDepth / 2, -(parWidth - parDepth) / 2, -parDamperLength / 2, _
                           parDepth / 2, (parWidth - parDepth) / 2)
                arcs.Add oLine
            
            End If
            oAxis.Set 1, 0, 0
            oStPoint.Set -parDamperLength / 2, parDepth / 2, (parWidth - parDepth) / 2
            Set oCString = PlaceTrCString(oStPoint, arcs)
            Set objDamperBody = PlaceProjection(m_OutputColl, oCString, oAxis, parDamperLength, True)
        End If
       'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objDamperBody
        Set objDamperBody = Nothing
        
       'Creating output 2 (Crank)
        oStPoint.Set 0, 0, 0
        oEnPoint.Set 0, 0, (parWidth / 2) + dCrankLength
        Set objCrank = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCrankDiameter, True)
       'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCrank
        Set objCrank = Nothing
        
       'Creating Output 3(Lever)
        oStPoint.Set -(dLeverWidth / 2), dCrankDiameter / 2, (parWidth / 2) + dCrankLength
        oEnPoint.Set dLeverWidth / 2, -dLeverLength, (parWidth / 2) + dCrankLength - (dLeverWidth / 4)
        Set objLever = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
       'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLever
        Set objLever = Nothing
       
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
    
    Case 157              'Control Damper with gear type operator
       'Creating Output1 (Damper Body)
        oStPoint.Set -parDamperLength / 2, 0, 0
        oEnPoint.Set parDamperLength / 2, 0, 0
        Set objDamperBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parWidth, True)
       'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objDamperBody
        Set objDamperBody = Nothing

       'Creating Output2 (Frame)
        Dim objFrame As Object
        oStPoint.Set -dFrameWidth / 2, -dFrameDepth / 2, 0
        oEnPoint.Set dFrameWidth / 2, dFrameDepth / 2, (parWidth / 2) + dFrameThickness
        Set objFrame = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
       'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objFrame
        Set objFrame = Nothing
        
       'Creating output 3 (Handwheel Operator)
        oStPoint.Set 0, dFrameDepth / 2, (parWidth / 2) + (dFrameThickness / 2)
        oEnPoint.Set 0, (dFrameDepth / 2) + dLeverLength, (parWidth / 2) + (dFrameThickness / 2)
        Set objLever = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dGearStemLength, True)
       'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLever
        Set objLever = Nothing
      
       'Creating Handwheel box
        Dim objHandWheelBox As Object
        oStPoint.Set -dGearStemLength / 2, (dFrameDepth / 2) + dLeverLength, (parWidth / 2) _
                     + (dFrameThickness / 2) - (dGearStemLength / 2)
        oEnPoint.Set dGearStemLength / 2, (dFrameDepth / 2) + dLeverLength + (dGearStemLength), _
                       (parWidth / 2) + (dFrameThickness / 2) + (dGearStemLength / 2)
        Set objHandWheelBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
       'Set the Output
        m_OutputColl.AddOutput "HandWheelBox ", objHandWheelBox
        Set objHandWheelBox = Nothing
        
       'Creating Handwheel
        Dim objHandWheel As Object
        Dim TorOrigin As New AutoMath.DPosition
        Dim TorMajrad As Double
        Dim TorMinRad As Double
        oAxis.Set 0, 1, 0
        TorMajrad = 0.5 * dFrameThickness
        TorMinRad = dGearStemLength / 2
        TorOrigin.Set 0, (dFrameDepth / 2) + dLeverLength + (dGearStemLength / 2), _
                      (parWidth / 2) + (dFrameThickness / 2)
        Set objHandWheel = PlaceTorus(m_OutputColl, TorOrigin, oAxis, TorMajrad, TorMinRad)
       'Set the Output
        m_OutputColl.AddOutput "HandWheel ", objHandWheel
        Set objHandWheel = Nothing
        Set TorOrigin = Nothing
        
        'Creating Spokes for the Handwheel
        'Spoke1
        Dim objHandwheelSpoke As Object
        oStPoint.Set 0, (dFrameDepth / 2) + dLeverLength + (dGearStemLength / 2), (parWidth / 2) + (dFrameThickness / 2)
        oEnPoint.Set -0.5 * dFrameThickness, oStPoint.y, oStPoint.z
        Set objHandwheelSpoke = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dGearStemLength / 2, True)
        'Set the Output
        m_OutputColl.AddOutput "HandWheelSpoke1", objHandwheelSpoke
        Set objHandwheelSpoke = Nothing
        
        'Spoke2
        oEnPoint.Set 0, oStPoint.y, oStPoint.z - 0.5 * dFrameThickness
        Set objHandwheelSpoke = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dGearStemLength / 2, True)
        'Set the Output
        m_OutputColl.AddOutput "HandWheelSpoke2", objHandwheelSpoke
        Set objHandwheelSpoke = Nothing
        
        'Spoke3
        oEnPoint.Set 0.5 * dFrameThickness, oStPoint.y, oStPoint.z
        Set objHandwheelSpoke = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dGearStemLength / 2, True)
        'Set the Output
        m_OutputColl.AddOutput "HandWheelSpoke3", objHandwheelSpoke
        Set objHandwheelSpoke = Nothing
        
        'Spoke4
        oEnPoint.Set 0, oStPoint.y, oStPoint.z + 0.5 * dFrameThickness
        Set objHandwheelSpoke = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dGearStemLength / 2, True)
        'Set the Output
        m_OutputColl.AddOutput "HandWheelSpoke4", objHandwheelSpoke
        Set objHandwheelSpoke = Nothing
        
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        Set oAxis = Nothing
        
    Case 158          'Control damper with an Actuator

       'Creating Output1 (Damper Body)
        oStPoint.Set -parDamperLength / 2, -parDepth / 2, -parWidth / 2
        oEnPoint.Set parDamperLength / 2, parDepth / 2, parWidth / 2
        Set objDamperBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
       'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objDamperBody
        Set objDamperBody = Nothing

       'Creating Output2 (Actuator frame)
        oStPoint.Set -dFrameWidth / 2, -dFrameDepth / 2, (parWidth / 2)
        oEnPoint.Set dFrameWidth / 2, dFrameDepth / 2, (parWidth / 2) + dFrameThickness
        Set objFrame = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
       'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objFrame
        Set objFrame = Nothing

       'Creating Output3 (Actuator Knob1)
        oStPoint.Set 0, 0, (parWidth / 2) + dFrameThickness
        oEnPoint.Set 0, 0, (parWidth / 2) + dFrameThickness + dKnobLength
        Set oKnob1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dGearStemLength, True)
       'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oKnob1
        Set oKnob1 = Nothing

       'Creating Knob2
        oStPoint.Set 0, 0, (parWidth / 2) + dFrameThickness + dKnobLength
        oEnPoint.Set 0, 0, (parWidth / 2) + dFrameThickness + (dKnobLength * 2)
        Set oKnob2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dGearStemLength * 2, True)
       'Set the output
        m_OutputColl.AddOutput "oKnob2", oKnob2
        Set oKnob2 = Nothing
        
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
    
    Case 159, 160    '159 - Control damper with transition with lever type operator
                     '160 - Control damper with transition with an actuator
   
       'Creating output 1(Transition Body)
        If (parHVACShape = 4) Then 'Round
            oStPoint.Set -parTransitionLength - (parDamperLength / 2), 0, 0
            oEnPoint.Set parTransitionLength + (parDamperLength / 2), 0, 0
            parDepth = parWidth
            Set objTransitionBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parWidth, True)
    
        ElseIf (parHVACShape = 1) Then  'Rectangular
            oStPoint.Set -parTransitionLength - (parDamperLength / 2), -parDepth / 2, -parWidth / 2
            oEnPoint.Set parTransitionLength + (parDamperLength / 2), parDepth / 2, parWidth / 2
            Set objTransitionBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

            Else   'Oval
                If CmpDblEqual(parWidth, parDepth) Then
                
                    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, _
                                (-parDamperLength / 2) - parTransitionLength, parDepth / 2, _
                                (parWidth - parDepth) / 2, (-parDamperLength / 2) - parTransitionLength, 0, _
                                parWidth / 2, (-parDamperLength / 2) - parTransitionLength, -parDepth / 2, _
                                (parWidth - parDepth) / 2)
                    arcs.Add oArc
                
                    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, _
                               (-parDamperLength / 2) - parTransitionLength, -parDepth / 2, _
                               -(parWidth - parDepth) / 2, (-parDamperLength / 2) - parTransitionLength, 0, _
                               -parWidth / 2, (-parDamperLength / 2) - parTransitionLength, parDepth / 2, _
                               -(parWidth - parDepth) / 2)
                    arcs.Add oArc
                
                Else
                
                    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, _
                                (-parDamperLength / 2) - parTransitionLength, parDepth / 2, _
                                (parWidth - parDepth) / 2, (-parDamperLength / 2) - parTransitionLength, 0, _
                                parWidth / 2, (-parDamperLength / 2) - parTransitionLength, -parDepth / 2, _
                                (parWidth - parDepth) / 2)
                    arcs.Add oArc
                    
                    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                (-parDamperLength / 2) - parTransitionLength, -parDepth / 2, _
                                (parWidth - parDepth) / 2, (-parDamperLength / 2) - parTransitionLength, _
                                -parDepth / 2, -(parWidth - parDepth) / 2)
                    arcs.Add oLine
                    
                    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, _
                               (-parDamperLength / 2) - parTransitionLength, -parDepth / 2, _
                               -(parWidth - parDepth) / 2, (-parDamperLength / 2) - parTransitionLength, 0, _
                               -parWidth / 2, (-parDamperLength / 2) - parTransitionLength, parDepth / 2, _
                               -(parWidth - parDepth) / 2)
                    arcs.Add oArc
                
                    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                (-parDamperLength / 2) - parTransitionLength, parDepth / 2, _
                                -(parWidth - parDepth) / 2, (-parDamperLength / 2) - parTransitionLength, _
                                parDepth / 2, (parWidth - parDepth) / 2)
                    arcs.Add oLine
                End If
                oAxis.Set 1, 0, 0
                oStPoint.Set (-parDamperLength / 2) - parTransitionLength, parDepth / 2, _
                             (parWidth - parDepth) / 2
                Set oCString = PlaceTrCString(oStPoint, arcs)
                Set objTransitionBody = PlaceProjection(m_OutputColl, oCString, oAxis, _
                                    parDamperLength + (2 * parTransitionLength), True)
        End If
       'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTransitionBody
        Set objTransitionBody = Nothing
    
       'Creating Output 2(Damper Body)
        If CmpDblEqual(parOffsetDamperHeight, 0) Then
            oStPoint.Set -(parDamperLength / 2), -parDamperHeight / 2, -parDamperWidth / 2
            oEnPoint.Set (parDamperLength / 2), parDamperHeight / 2, parDamperWidth / 2
        Else
            oStPoint.Set -(parDamperLength / 2), -(parDepth / 2) - parOffsetDamperHeight, -parDamperWidth / 2
            oEnPoint.Set (parDamperLength / 2), parDamperHeight - parOffsetDamperHeight - (parDepth / 2), parDamperWidth / 2
        End If
        Set objDamperBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
       'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objDamperBody
        Set objDamperBody = Nothing

        If lPartDataBasis = 159 Then
            'Creating output 3 (Crank)
             oStPoint.Set 0, 0, 0
             oEnPoint.Set 0, 0, parDamperWidth / 2 + dCrankLength
             Set objCrank = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCrankDiameter, True)
            'Set the Output
             iOutput = iOutput + 1
             m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCrank
             Set objCrank = Nothing

            'Creating Lever
             oStPoint.Set -(dLeverWidth / 2), dCrankDiameter / 2, dCrankLength + (parDamperWidth / 2)
             oEnPoint.Set dLeverWidth / 2, -dLeverLength, dCrankLength - (dLeverWidth / 4) + (parDamperWidth / 2)
             Set objLever = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            'Set the Output
             m_OutputColl.AddOutput "Lever", objLever
             Set objLever = Nothing

        ElseIf lPartDataBasis = 160 Then
           'Creating Output 3 (Actuator frame)
            oStPoint.Set -dFrameWidth / 2, -dFrameDepth / 2, 0
            oEnPoint.Set dFrameWidth / 2, dFrameDepth / 2, (parDamperWidth / 2) + dFrameThickness
            Set objFrame = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
           'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objFrame
            Set objFrame = Nothing
            
           'Creating Actuator Knob1
            oStPoint.Set 0, 0, dFrameThickness + (parDamperWidth / 2)
            oEnPoint.Set 0, 0, dFrameThickness + (parDamperWidth / 2) + dKnobLength
            Set oKnob1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dGearStemLength, True)
           'Set the output
            m_OutputColl.AddOutput "oKnob1", oKnob1
            Set oKnob1 = Nothing
            
           'Creating Knob2
            oStPoint.Set 0, 0, dFrameThickness + (parDamperWidth / 2) + dKnobLength
            oEnPoint.Set 0, 0, dFrameThickness + (parDamperWidth / 2) + (dKnobLength * 2)
            Set oKnob2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dGearStemLength * 2, True)
           'Set the output
            m_OutputColl.AddOutput "oKnob2", oKnob2
            Set oKnob2 = Nothing
        End If

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
        
    Case Else
        GoTo ErrorLabel:
        
  End Select
    '===================
    'BUILD HVACNOZZLE1
    '===================
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oDistribPort As GSCADNozzleEntities.IJDistribPort
    
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    
    Dim oPos As AutoMath.DPosition
    Set oPos = New AutoMath.DPosition
    Dim oPortDir As AutoMath.DVector
    Set oPortDir = New AutoMath.DVector
    Dim oRadialDir As AutoMath.DVector
    Set oRadialDir = New AutoMath.DVector
    

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
    
    CornerRadius = 0#
    dNozzLength(1) = 0.001
    dNozzLength(2) = 0.001
    
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
            dThickness(iPortIndex) = NEGLIGIBLE_THICKNESS
        End If
        If CmpDblLessThan(dNozzLength(iPortIndex), dThickness(iPortIndex)) Then
            dNozzLength(iPortIndex) = dThickness(iPortIndex) + NEGLIGIBLE_THICKNESS
        End If
    Next iPortIndex
        
    Set oHvacPort = Nothing
    oHvacColl.Clear
    Set oHvacColl = Nothing
    
    iPortIndex = 1
    
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                                parHVACShape, lEndPrep(1), _
                                dThickness(1), dFlangeWidth(1), lFlowDir(1), parWidth, _
                                parDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                "HNoz1", dPortDepth(1), dCptOffSet(1), False, _
                                m_OutputColl.ResourceManager)

    'Position of the nozzle should be the conenct point of the nozzle
    Dim x As Double, y As Double, z As Double
Select Case lPartDataBasis
    Case 156, 157, 158
        x = -parDamperLength / 2
        y = 0#
        z = 0#
   Case 159, 160
        x = (-parDamperLength / 2) - parTransitionLength
        y = 0#
        z = 0#
End Select

    oPos.Set x, y, z
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation oPos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    oPortDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oPortDir
    
    oRadialDir.Set 0, 1, 0
    oDistribPort.SetRadialOrient oRadialDir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(1)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oDistribPort = Nothing

    '===================
    'BUILD HVACNOZZLE2
    '===================
    iPortIndex = 2

    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                    parHVACShape, lEndPrep(2), dThickness(2), dFlangeWidth(2), lFlowDir(2), _
                    parWidth, parDepth, CornerRadius, DimBaseOuter, PortStatus, _
                    "HNoz2", dPortDepth(2), dCptOffSet(2), False, _
                    m_OutputColl.ResourceManager)

    
    'Position of the nozzle should be the conenct point of the nozzle
Select Case lPartDataBasis
    Case 156, 157, 158
        x = parDamperLength / 2
        y = 0
        z = 0
    Case 159, 160
        x = (parDamperLength / 2) + parTransitionLength
        y = 0
        z = 0
End Select

    oPos.Set x, y, z
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation oPos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    oPortDir.Set 1, 0, 0
    oDistribPort.SetDirectionVector oPortDir
    
    oRadialDir.Set 0, 1, 0
    oDistribPort.SetRadialOrient oRadialDir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(2)

    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oDistribPort = Nothing
    
    Set NozzleFactory = Nothing
    Set oPos = Nothing
    Set oPortDir = Nothing
    Set oRadialDir = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError2 MODULE, METHOD
End Sub
