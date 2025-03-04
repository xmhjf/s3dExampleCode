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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   ProgID:         SP3DOPGearTopHandwheel.TopmountedHw
'   Author:         RUK
'   Creation Date:  Wednesday, 2 April 2008
'
'   Description:
'   This Symbol details were taken from Appendix E-107 of the Design Document
'   In this symbol below pVave operator PDBs are implemented
'   Default (1)
'   Gear, top mounted handwheel Type 1 (31)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   04.Apr.2008     RUK     CR-CP-133524  Sample data for Dimensional Basis for valve operators is required
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:"    'Used for error messages
Private PI As Double

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    
    Dim parOperatorHeight As Double
    Dim parOffsetFrmValCen As Double
    Dim parGearDiameter As Double
    Dim parGearRadius As Double
    Dim parGearDepth As Double
    Dim parValCenToHWTop As Double
    Dim parHandwheelDiameter As Double
    Dim parHandwheelRadius As Double
    Dim parHwXOffsetFrmOpStem As Double
    Dim parHwZOffsetOpFrmStem As Double
        
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
'   parOperatorHeight = arrayOfInputs(2)
'   parOffsetFrmValCen = arrayOfInputs(3)
'   parGearDiameter = arrayOfInputs(4)
'   parGearDepth = arrayOfInputs(5)
'   parValCenToHWTop = arrayOfInputs(6)
'   parHWDiameter = arrayOfInputs(7)
'   parHwXOffsetFrmOpStem = arrayOfInputs(8)
'   parYOffset_HW_OPStem = arrayOfInputs(9)
    
    Dim iOutput As Integer
    
    'Checking for the PartDataBasis
    Dim oOperator As IJValveOperator
    Dim lPartDataBasis As Long
    Set oOperator = oPartFclt
    lPartDataBasis = oOperator.DimensionalBasis
    Set oOperator = Nothing
    
    Select Case lPartDataBasis
    Case Is <= 1, 31 'Gear, top mounted handwheel Type 1
        parOperatorHeight = arrayOfInputs(2)
        parOffsetFrmValCen = arrayOfInputs(3)
        parGearDiameter = arrayOfInputs(4)
        parGearDepth = arrayOfInputs(5)
        parHandwheelDiameter = arrayOfInputs(6)
        parValCenToHWTop = arrayOfInputs(7)
        parHwXOffsetFrmOpStem = arrayOfInputs(8)
        parHwZOffsetOpFrmStem = arrayOfInputs(9)
    Case Else
        GoTo ErrorLabel
    End Select
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oCenter As AutoMath.DPosition
    Dim oVector As AutoMath.DVector
    
    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition
    Set oCenter = New DPosition
    Set oVector = New DVector
    
    Dim dOpearatorHeight As Double
    Dim dValCenToHWTop As Double
    Dim dHwXOffsetFrmOpStem As Double
    Dim dHwZOffsetOpFrmStem As Double
    Dim dHandwheelDiameter As Double
    Dim dGearDepth As Double
    Dim dGearDiameter As Double
    
    Select Case lPartDataBasis
    Case Is <= 1, 31
        
        'Assaign Values for optional parameters
        'Calculate Gear Depth
        dGearDepth = IIf(CmpDblGreaterthan(parGearDepth, LINEAR_TOLERANCE), _
                        parGearDepth, _
                        IIf(CmpDblGreaterthan(parOperatorHeight, LINEAR_TOLERANCE), _
                            parOperatorHeight - parOffsetFrmValCen, _
                            2 * (parValCenToHWTop - parOffsetFrmValCen) / 3))
        
        'Spec writer should specify Handwheel X Offset from Operator Stem or Gear Diameter or Gear Radius
        'Spec writer may also to choose to specify value for
        'Calculate Gear Diameter
        dGearDiameter = IIf(CmpDblGreaterthan(parGearDiameter, LINEAR_TOLERANCE), _
                            parGearDiameter, _
                            IIf(CmpDblGreaterthan(parGearRadius, LINEAR_TOLERANCE), _
                            2 * parGearRadius, 1.5 * parHwXOffsetFrmOpStem))
                            
        'Calculate Handwheel X-Offset from Operator Stem
        dHwXOffsetFrmOpStem = IIf(CmpDblGreaterthan(parHwXOffsetFrmOpStem, LINEAR_TOLERANCE), _
                                parHwXOffsetFrmOpStem, dGearDiameter / 1.5)
            
        'Spec writer should specify Handwheel Z Offset from Operator Stem or Handwheel Diameter or Handwheel Radius
        'Spec writer may also to choose to specify value for
        'Handwheel Z Offset from operator Stem and either Handwheel Diameter or Handwheel Radius
        If CmpDblEqual(parHwZOffsetOpFrmStem, 0) And CmpDblEqual(parHandwheelDiameter, 0) And _
                    CmpDblEqual(parHandwheelRadius, 0) Then
            GoTo ErrorLabel
        End If
        
        dHandwheelDiameter = IIf(CmpDblGreaterthan(parHandwheelDiameter, LINEAR_TOLERANCE), _
                                parHandwheelDiameter, _
                                IIf(CmpDblGreaterthan(parHandwheelRadius, LINEAR_TOLERANCE), _
                                    2 * parHandwheelRadius, 2 * parHwZOffsetOpFrmStem))
                                    
        dHwZOffsetOpFrmStem = IIf(CmpDblGreaterthan(parHwZOffsetOpFrmStem, LINEAR_TOLERANCE), _
                                parHwZOffsetOpFrmStem, 0.8 * dGearDiameter / 2)
                                
            
        'The spec writer must specify a dimensional value for either the Valve Operator Height
        'or the Valve Centerline to Top of Handwheel.
        If CmpDblEqual(parOperatorHeight, 0) And CmpDblEqual(parValCenToHWTop, 0) Then
            GoTo ErrorLabel
        End If
        
        'Calculate Opearator Height and Valve Centerline to Top of Handwheel
        dOpearatorHeight = IIf(CmpDblGreaterthan(parOperatorHeight, LINEAR_TOLERANCE), _
                            parOperatorHeight, parValCenToHWTop - 2 * dGearDepth / 3)
        dValCenToHWTop = IIf(CmpDblGreaterthan(parValCenToHWTop, LINEAR_TOLERANCE), _
                            parValCenToHWTop, dOpearatorHeight + 2 * dGearDepth / 3)
        
        'Create the Premitives
        'Create the Gear box (3 Cylinders)
        'Gear box Cylinder 1
        oStPoint.Set 0, parOffsetFrmValCen, 0
        oEnPoint.Set oStPoint.x, oStPoint.y + 0.05 * dGearDepth, oStPoint.z
        iOutput = iOutput + 1
        Call CreateCylinderAddtoOutput(m_OutputColl, oStPoint, oEnPoint, _
                                                dGearDiameter, True, arrayOfOutputs(iOutput))
        'Gear box Cylinder 2
        oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
        oEnPoint.Set oStPoint.x, oEnPoint.y + 0.6 * dGearDepth, oStPoint.z
        iOutput = iOutput + 1
        Call CreateCylinderAddtoOutput(m_OutputColl, oStPoint, oEnPoint, _
                                                dGearDiameter, True, arrayOfOutputs(iOutput))
        'Gear box Cylinder 3
        oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
        oEnPoint.Set oStPoint.x, oEnPoint.y + 0.35 * dGearDepth, oStPoint.z
        iOutput = iOutput + 1
        Call CreateCylinderAddtoOutput(m_OutputColl, oStPoint, oEnPoint, _
                                            0.3 * dGearDiameter, True, arrayOfOutputs(iOutput))
        
        'Create the Connecting Shaft
        Dim objShaft As Object
        Dim dShaftZloc As Double
        dShaftZloc = dHwZOffsetOpFrmStem - 1.6 * dGearDepth / 2 + 0.35 * dGearDepth
        'Shaft should be connected at +ve Z
        dShaftZloc = IIf(CmpDblGreaterthan(dShaftZloc, LINEAR_TOLERANCE), dShaftZloc, 0)
        oStPoint.Set 0, parOffsetFrmValCen + 0.35 * dGearDepth, dShaftZloc
        oEnPoint.Set -0.8 * dHwXOffsetFrmOpStem, oStPoint.y, oStPoint.z
        
        Set objShaft = PlaceCone(m_OutputColl, oStPoint, oEnPoint, _
                                    0.05 * dGearDepth / 2, 0.2 * dGearDepth, True)
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objShaft
        Set objShaft = Nothing
        
        'Create the Gear 2
        oStPoint.Set -1.2 * dHwXOffsetFrmOpStem, parOffsetFrmValCen + dGearDepth / 2, _
                                                                dHwZOffsetOpFrmStem
        oEnPoint.Set -0.8 * dHwXOffsetFrmOpStem, oStPoint.y, oStPoint.z
        iOutput = iOutput + 1
        Call CreateCylinderAddtoOutput(m_OutputColl, oStPoint, oEnPoint, _
            dValCenToHWTop - parOffsetFrmValCen, True, arrayOfOutputs(iOutput))
        
        'Create the Handwheel Stem
        oStPoint.Set -dHwXOffsetFrmOpStem, _
                    parOffsetFrmValCen + dGearDepth / 2 + (dValCenToHWTop - parOffsetFrmValCen) / 2, _
                    dHwZOffsetOpFrmStem
        oEnPoint.Set oStPoint.x, dValCenToHWTop, oStPoint.z
        iOutput = iOutput + 1
        Call CreateCylinderAddtoOutput(m_OutputColl, oStPoint, oEnPoint, _
                                        0.2 * dHwXOffsetFrmOpStem, True, arrayOfOutputs(iOutput))
        'Create the Handwheel
        oCenter.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
        oVector.Set 0, 1, 0
        iOutput = iOutput + 1
        Call CreateHandwheelAddtoOutput(m_OutputColl, oCenter, oVector, _
                    dHandwheelDiameter / 2 - 0.05 * dHandwheelDiameter, _
                                0.05 * dHandwheelDiameter, arrayOfOutputs(iOutput))
                                    
        'Create the Spokes for Handwheel
        oStPoint.Set oCenter.x, oCenter.y, oCenter.z
        
        'Spoke 1
        oEnPoint.Set oCenter.x + (0.95 * dHandwheelDiameter / 2) * Cos(PI / 6), oCenter.y, _
                                    oCenter.z - (0.95 * dHandwheelDiameter / 2) * Sin(PI / 6)
        iOutput = iOutput + 1
        Call CreateCylinderAddtoOutput(m_OutputColl, oStPoint, oEnPoint, _
                                        0.05 * dHandwheelDiameter, True, arrayOfOutputs(iOutput))
        
        'Spoke 2
        oEnPoint.Set oCenter.x - (0.95 * dHandwheelDiameter / 2) * Cos(PI / 6), oCenter.y, _
                                    oCenter.z - (0.95 * dHandwheelDiameter / 2) * Sin(PI / 6)
        iOutput = iOutput + 1
        Call CreateCylinderAddtoOutput(m_OutputColl, oStPoint, oEnPoint, _
                                        0.05 * dHandwheelDiameter, True, arrayOfOutputs(iOutput))
        
        'Spoke 3
        oEnPoint.Set oCenter.x, oCenter.y, oCenter.z + 0.95 * dHandwheelDiameter / 2
        iOutput = iOutput + 1
        Call CreateCylinderAddtoOutput(m_OutputColl, oStPoint, oEnPoint, _
                                        0.05 * dHandwheelDiameter, True, arrayOfOutputs(iOutput))
            
    Case Else
        GoTo ErrorLabel
    End Select
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oCenter = Nothing
    Set oVector = Nothing
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

Private Sub CreateCylinderAddtoOutput(ByVal OutputColl As Object, ByVal oStPoint As IJDPosition, _
                ByVal oEnPoint As IJDPosition, ByVal dDiameter As Double, _
                Optional bIsCapped As Boolean = True, _
                Optional ByVal sOutputName As String = "Cylinder")

    'This Sub creates a Cylinder from oStPoint to oEnPoint having given diameter and adds to the Output
    Dim objCylinder As Object
    Set objCylinder = PlaceCylinder(OutputColl, oStPoint, oEnPoint, dDiameter, bIsCapped)
    
    'Add to the Ouput
    OutputColl.AddOutput sOutputName, objCylinder
    Set objCylinder = Nothing
End Sub

Private Sub CreateHandwheelAddtoOutput(ByVal OutputColl As Object, ByVal oCenter As IJDPosition, _
                    ByVal oAxis As IJDVector, ByVal dMajorRad As Double, _
                    ByVal dMinorRad As Double, _
                    Optional ByVal sOutputName As String = "Handwheel")
    
    Dim objHandwheel As Object
    Set objHandwheel = PlaceTorus(OutputColl, oCenter, oAxis, dMajorRad, dMinorRad)
    
    'Add to the Ouput
    OutputColl.AddOutput sOutputName, objHandwheel
    Set objHandwheel = Nothing
End Sub


