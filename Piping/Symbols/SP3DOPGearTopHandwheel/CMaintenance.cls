VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   ProgID:         SP3DOPGearTopHandwheel.TopmountedHw
'   Author:         RUK
'   Creation Date:  Wednesday, 2 April 2008
'   Description:
'   This Symbol details were taken from Appendix E-107 of the Design Document
'   Default (1)
'   Gear, top mounted handwheel Type 1 (31)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   04.Apr.2008     RUK     CR-CP-133524  Sample data for Dimensional Basis for valve operators is required
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Maintenance:"    'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart

    Dim iOutput As Double
    Dim ObjMaintStem As Object
    Dim ObjMaintOperator As Object
    Dim ObjMaintEccentricShaft As Object
    Dim ObjMainTopmountedHwCylin As Object

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
    'parOperatorHeight = arrayOfInputs(2)
    'parOffsetFrmValCen = arrayOfInputs(3)
    'parGearDiameter = arrayOfInputs(4)
    'parGearDepth = arrayOfInputs(5)
    'parHandwheelDiameter = arrayOfInputs(6)
    'parValCenToHWTop = arrayOfInputs(7)
    'parHwXOffsetFrmOpStem = arrayOfInputs(8)
    'parHwZOffsetOpFrmStem = arrayOfInputs(9)

    iOutput = 0

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
    
    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition
    
    Dim dOpearatorHeight As Double
    Dim dValCenToHWTop As Double
    Dim dHwXOffsetFrmOpStem As Double
    Dim dHwZOffsetOpFrmStem As Double
    Dim dHandwheelDiameter As Double
    Dim dGearDepth As Double
    Dim dGearDiameter As Double
    
    Select Case lPartDataBasis
    Case 31
    
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
        'Create the Handwheel Stem
        oStPoint.Set -dHwXOffsetFrmOpStem, _
                    parOffsetFrmValCen + dGearDepth / 2 + (dValCenToHWTop - parOffsetFrmValCen) / 2, _
                    dHwZOffsetOpFrmStem
        oEnPoint.Set oStPoint.x, dValCenToHWTop, oStPoint.z
        iOutput = iOutput + 1
        Call CreateCylinderAddtoOutput(m_OutputColl, oStPoint, oEnPoint, _
                                        0.2 * dHwXOffsetFrmOpStem, True, arrayOfOutputs(iOutput))
        'Create the Handwheel
        oStPoint.Set oEnPoint.x, oEnPoint.y - 0.05 * dHandwheelDiameter, oEnPoint.z
        oEnPoint.Set oStPoint.x, dValCenToHWTop + 0.05 * dHandwheelDiameter, oStPoint.z
        iOutput = iOutput + 1
        Call CreateCylinderAddtoOutput(m_OutputColl, oStPoint, oEnPoint, _
                                         dHandwheelDiameter, True, arrayOfOutputs(iOutput))
        
    Case Else
        GoTo ErrorLabel
    End Select
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
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


