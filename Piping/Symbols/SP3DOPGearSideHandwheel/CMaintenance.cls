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
'   ProgID:         SP3DGOp_SHandwheel.SHandwheel
'   Author:         RUK
'   Creation Date:  Wednesday, 2 April 2008
'   Description:
'   This Symbol details were taken from Appendix E-107 of the Design Document
'   In this symbol below Valve operator PDBs are implemented
'   Default (1)
'   Gear, side mounted handwheel Type 1 (11)
'   Gear, side mounted handwheel Type 5 (mirror of Type 1) (15)
'   Gear, side mounted handwheel Type 11 (21)
'   Gear, side mounted handwheel Type 12 (mirror of Type 11) (24)
'   Gear, side mounted handwheel Type 11 with stem extension (25)
'   Gear, side mounted handwheel Type 12 with stem extension (mirror of Type 11 with stem extension) (26)
'   Gear, side mounted handwheel Type 11 with shaft (27)
'   Gear, side mounted handwheel Type 12 with shaft (mirror of Type 11 with shaft) (28)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   04.Apr.2008     RUK     CR-CP-133524  Sample data for Dimensional Basis for valve operators is required
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Maintenance:"    'Used for error messages
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
    Dim parGearDepth As Double
    Dim parHandwheelDiameter As Double
    Dim parHwXOffsetFrmOpStem As Double
    Dim parHwZOffsetOpFrmStem As Double
    Dim parHwYOffsetFrmGearBot As Double
    Dim parHwOffsetFrmValCen As Double
    Dim parHwOffsetFrmOpTop As Double
    Dim parValCenToHWTop As Double
    Dim parValCenToHWCen As Double
    Dim parHandwheelRadius As Double
    Dim parExStemLen As Double
    Dim parExStemLenFrmValCen As Double
    Dim parExStemDia As Double
    Dim parGearRadius As Double
    Dim parShaftHeight As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parOperatorHeight = arrayOfInputs(2)
    'parOffsetFrmValCen = arrayOfInputs(3)
    'parGearDiameter = arrayOfInputs(4)
    'parGearRadius = arrayOfInputs(5)
    'parGearDepth = arrayOfInputs(6)
    'parHandwheelDiameter = arrayOfInputs(7)
    'parHandwheelRadius = arrayOfInputs(8)
    'parHwXOffsetFrmOpStem = arrayOfInputs(9)
    'parHwZOffsetOpFrmStem = arrayOfInputs(10)
    'parHwYOffsetFrmGearBot = arrayOfInputs(11)
    'parHwOffsetFrmValCen = arrayOfInputs(12)
    'parHwOffsetFrmOpTop = arrayOfInputs(13)
    'parValCenToHWTop = arrayOfInputs(14)
    'parValCenToHWCen = arrayOfInputs(15)
    'parExStemLen = arrayOfInputs(16)
    'parExStemLenFrmValCen = arrayOfInputs(17)
    'parExStemDia = arrayOfInputs(18)
    'parShaftHeight = arrayOfInputs(19)
    
    Dim iOutput As Integer
    
    'Checking for the PartDataBasis
    Dim oOperator As IJValveOperator
    Dim lPartDataBasis As Long
    Set oOperator = oPartFclt
    lPartDataBasis = oOperator.DimensionalBasis
    Set oOperator = Nothing
    
    Select Case lPartDataBasis
    Case Is <= 1, 11, 15    'Gear, side mounted handwheel Type 1 (11)
                            'Gear, side mounted handwheel Type 5 (mirror of Type 1) (15)
        parOperatorHeight = arrayOfInputs(2)
        parOffsetFrmValCen = arrayOfInputs(3)
        parGearDiameter = arrayOfInputs(4)
        parGearRadius = arrayOfInputs(5)
        parGearDepth = arrayOfInputs(6)
        parHandwheelDiameter = arrayOfInputs(7)
        parHandwheelRadius = arrayOfInputs(8)
        parHwXOffsetFrmOpStem = arrayOfInputs(9)
        parHwZOffsetOpFrmStem = arrayOfInputs(10)
        parHwOffsetFrmValCen = arrayOfInputs(12)
        parHwOffsetFrmOpTop = arrayOfInputs(13)
        
    Case 21, 24 'Gear, side mounted handwheel Type 11
                'Gear, side mounted handwheel Type 12 (mirror of Type 11)
        parOperatorHeight = arrayOfInputs(2)
        parOffsetFrmValCen = arrayOfInputs(3)
        parGearDiameter = arrayOfInputs(4)
        parGearRadius = arrayOfInputs(5)
        parGearDepth = arrayOfInputs(6)
        parHandwheelDiameter = arrayOfInputs(7)
        parHandwheelRadius = arrayOfInputs(8)
        parHwXOffsetFrmOpStem = arrayOfInputs(9)
        parHwZOffsetOpFrmStem = arrayOfInputs(10)
        parHwYOffsetFrmGearBot = arrayOfInputs(11)
        parValCenToHWTop = arrayOfInputs(14)
        parValCenToHWCen = arrayOfInputs(15)
    
    Case 25, 26 'Gear, side mounted handwheel, Type 11 with Extension Stem (25)
                'Gear, side mounted handwheel, Type 12 with Extension Stem (Mirror of Type 11 with Extension Stem) (26)
        parOperatorHeight = arrayOfInputs(2)
        parOffsetFrmValCen = arrayOfInputs(3)
        parGearDiameter = arrayOfInputs(4)
        parGearRadius = arrayOfInputs(5)
        parGearDepth = arrayOfInputs(6)
        parHandwheelDiameter = arrayOfInputs(7)
        parHandwheelRadius = arrayOfInputs(8)
        parHwXOffsetFrmOpStem = arrayOfInputs(9)
        parHwZOffsetOpFrmStem = arrayOfInputs(10)
        parHwYOffsetFrmGearBot = arrayOfInputs(11)
        parValCenToHWTop = arrayOfInputs(14)
        parValCenToHWCen = arrayOfInputs(15)
        parExStemLen = arrayOfInputs(16)
        parExStemLenFrmValCen = arrayOfInputs(17)
        parExStemDia = arrayOfInputs(18)
    
    Case 27, 28 'Gear, side mounted handwheel, Type 11 with Shaft
                'Gear, side mounted handwheel, Type 12 with Shaft (Mirror of Type 11 with Shaft)
        parOperatorHeight = arrayOfInputs(2)
        parOffsetFrmValCen = arrayOfInputs(3)
        parGearDiameter = arrayOfInputs(4)
        parGearRadius = arrayOfInputs(5)
        parGearDepth = arrayOfInputs(6)
        parHandwheelDiameter = arrayOfInputs(7)
        parHandwheelRadius = arrayOfInputs(8)
        parHwXOffsetFrmOpStem = arrayOfInputs(9)
        parHwZOffsetOpFrmStem = arrayOfInputs(10)
        parHwYOffsetFrmGearBot = arrayOfInputs(11)
        parValCenToHWTop = arrayOfInputs(14)
        parValCenToHWCen = arrayOfInputs(15)
        parShaftHeight = arrayOfInputs(19)
    End Select
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    
    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition
    
    Dim dOpearatorHeight As Double
    Dim dHwOffsetFrmValCen As Double
    Dim dHwXOffsetFrmOpStem As Double
    Dim dHwZOffsetOpFrmStem As Double
    Dim dHandwheelDiameter As Double
    Dim dGearDepth As Double
    Dim dGearDiameter As Double
    Dim iHandwheelPos As Integer
    
    Select Case lPartDataBasis
    Case Is <= 1, 11, 15
        'Assaign Values for optional parameters
        'Spec writer should specify Handwheel X Offset from Operator Stem or Handwheel Diameter or Handwheel Radius
        'Spec writer may also to choose to specify value for
        'Handwheel X Offset from operator Stem and either Handwheel Diameter or Handwheel Radius
        
        If CmpDblEqual(parHwZOffsetOpFrmStem, 0) And CmpDblEqual(parHandwheelDiameter, 0) And _
                    CmpDblEqual(parHandwheelRadius, 0) Then
            GoTo ErrorLabel
        End If
            
        dHandwheelDiameter = IIf(CmpDblGreaterthan(parHandwheelDiameter, LINEAR_TOLERANCE), _
                                parHandwheelDiameter, _
                                IIf(CmpDblGreaterthan(parHandwheelRadius, LINEAR_TOLERANCE), _
                                    2 * parHandwheelRadius, 2 * parHwZOffsetOpFrmStem))
                                    
        dHwZOffsetOpFrmStem = IIf(CmpDblGreaterthan(parHwZOffsetOpFrmStem, LINEAR_TOLERANCE), _
                                parHwZOffsetOpFrmStem, dHandwheelDiameter / 2)
        
        'Calculate Gear Depth
        dGearDepth = IIf(CmpDblGreaterthan(parGearDepth, LINEAR_TOLERANCE), _
                                        parGearDepth, 0.95 * dHandwheelDiameter)
        
        'Calculate Gear Diameter
        dGearDiameter = IIf(CmpDblGreaterthan(parGearDiameter, LINEAR_TOLERANCE), _
                            parGearDiameter, 2 * parGearRadius)
        If CmpDblEqual(dGearDiameter, 0) Then GoTo ErrorLabel
        
        'Gear Radius should be greater that Handwheel Z Offset from Operator Stem.
        'Otherwise Handwheel has no contact with Gear box.
        'Here the Handwheel will contact the Gear box at the middle part of the Gear box which has the diameter 0.8 times of the Gear Diameter
        dGearDiameter = IIf(CmpDblGreaterthan(0.8 * dGearDiameter / 2, dHwZOffsetOpFrmStem), _
                                            dGearDiameter, 2 * (dHwZOffsetOpFrmStem / 0.8))
        
        'Calculate Handwheel X-Offset from Operator Stem
        dHwXOffsetFrmOpStem = IIf(CmpDblGreaterthan(parHwXOffsetFrmOpStem, LINEAR_TOLERANCE), _
                                parHwXOffsetFrmOpStem, dGearDiameter)
            
        'The spec writer must specify a dimensional value for either the Valve Operator Height
        'or the Handwheel Offset from Valve Centerline.
        If CmpDblEqual(parOperatorHeight, 0) And CmpDblEqual(parHwOffsetFrmValCen, 0) Then
            GoTo ErrorLabel
        End If
        
        'Calculate Opearator Height and Handwheel Offset from Valve Centerline
        dOpearatorHeight = IIf(CmpDblGreaterthan(parOperatorHeight, LINEAR_TOLERANCE), _
                            parOperatorHeight, _
                            IIf(CmpDblGreaterthan(parHwOffsetFrmOpTop, LINEAR_TOLERANCE), _
                                parHwOffsetFrmValCen + parHwOffsetFrmOpTop, _
                                parHwOffsetFrmValCen + dGearDepth / 2))
        
        dHwOffsetFrmValCen = IIf(CmpDblGreaterthan(parHwOffsetFrmValCen, LINEAR_TOLERANCE), _
                            parHwOffsetFrmValCen, _
                            IIf(CmpDblGreaterthan(parHwOffsetFrmOpTop, LINEAR_TOLERANCE), _
                                parOperatorHeight - parHwOffsetFrmOpTop, _
                                parOperatorHeight - 0.95 * dHandwheelDiameter / 2))
            
        'Create the Premitives
        'Create the Handwheel
        iHandwheelPos = IIf(lPartDataBasis = 11, 1, -1)
        
        'Handwheel Stem
        oStPoint.Set 0, dHwOffsetFrmValCen, _
                                        iHandwheelPos * parHwZOffsetOpFrmStem
        oEnPoint.Set dHwXOffsetFrmOpStem, oStPoint.y, oStPoint.z
        iOutput = iOutput + 1
        Call CreateCylinderAddtoOutput(m_OutputColl, oStPoint, oEnPoint, _
                                    0.05 * dHandwheelDiameter, True, arrayOfOutputs(iOutput))
        
        'Handwheel
        oStPoint.Set oEnPoint.x - 0.15 * dHandwheelDiameter, oEnPoint.y, oEnPoint.z
        oEnPoint.Set oStPoint.x + 0.15 * dHandwheelDiameter, oStPoint.y, oStPoint.z
        
        iOutput = iOutput + 1
        Call CreateCylinderAddtoOutput(m_OutputColl, oStPoint, oEnPoint, dHandwheelDiameter, _
                                                True, arrayOfOutputs(iOutput))
    
    Case 21, 24, 25, 26, 27, 28
        'Assaign Values for optional parameters

        'Spec writer should specify Handwheel X Offset from Operator Stem or Handwheel Diameter or Handwheel Radius
        'Spec writer may also to choose to specify value for
        'Handwheel X Offset from operator Stem and either Handwheel Diameter or Handwheel Radius
        
        If CmpDblLessThanOrEqualTo(parHwXOffsetFrmOpStem, LINEAR_TOLERANCE) And _
                CmpDblLessThanOrEqualTo(parHandwheelDiameter, LINEAR_TOLERANCE) And _
                CmpDblLessThanOrEqualTo(parHandwheelRadius, LINEAR_TOLERANCE) Then
            GoTo ErrorLabel
        End If
        
        dHandwheelDiameter = IIf(CmpDblGreaterthan(parHandwheelDiameter, LINEAR_TOLERANCE), _
                                parHandwheelDiameter, _
                                IIf(CmpDblGreaterthan(parHandwheelRadius, LINEAR_TOLERANCE), _
                                    2 * parHandwheelRadius, _
                                    2 * parHwXOffsetFrmOpStem))
                                    
        dHwXOffsetFrmOpStem = IIf(CmpDblGreaterthan(parHwXOffsetFrmOpStem, LINEAR_TOLERANCE), _
                                parHwXOffsetFrmOpStem, dHandwheelDiameter / 2)
        
        'Calculate Gear Depth
        dGearDepth = IIf(CmpDblGreaterthan(parGearDepth, LINEAR_TOLERANCE), _
                            parGearDepth, 0.65 * dHandwheelDiameter)
        
        'Calculate Gear Diameter and Handwheel Z Offset from Operator Stem
        If CmpDblEqual(parHwZOffsetOpFrmStem, 0) Then
            dGearDiameter = IIf(CmpDblGreaterthan(parGearDiameter, LINEAR_TOLERANCE), _
                            parGearDiameter, 2 * parGearRadius)
            dHwZOffsetOpFrmStem = dGearDiameter
        Else
            dGearDiameter = IIf(CmpDblGreaterthan(parGearDiameter, LINEAR_TOLERANCE), _
                            parGearDiameter, _
                            IIf(CmpDblGreaterthan(parGearRadius, LINEAR_TOLERANCE), _
                                2 * parGearRadius, parHwZOffsetOpFrmStem))
            dHwZOffsetOpFrmStem = parHwZOffsetOpFrmStem
        End If
        
        If CmpDblEqual(dGearDiameter, 0) Then GoTo ErrorLabel
        
        'Gear Radius should be greater that Handwheel X Offset from Operator Stem.
        'Otherwise Handwheel has no contact with Gear box.
        dGearDiameter = IIf(CmpDblGreaterthan(dGearDiameter / 2, dHwXOffsetFrmOpStem), _
                                            dGearDiameter, 2 * dHwXOffsetFrmOpStem)
        
        'Adust the Handwheel Z offset operator Stem such that it was greter than Gear Radius
        If CmpDblLessThan(dHwZOffsetOpFrmStem, dGearDiameter / 2) Then
            dHwZOffsetOpFrmStem = dGearDiameter / 2
        End If
                
        If lPartDataBasis = 25 Or lPartDataBasis = 26 Then
            Dim dExStemDia As Double
            Dim dExStemLen As Double
            'The spec writer must specify either the Extension Stem Length or
            'the Extension Stem Length from Valve Centerline
            If CmpDblEqual(parExStemLen, 0) And CmpDblEqual(parExStemLenFrmValCen, 0) Then
                GoTo ErrorLabel
            End If
            dExStemLen = IIf(CmpDblEqual(parExStemLenFrmValCen, 0), _
                            parExStemLen, parExStemLenFrmValCen - parOffsetFrmValCen)
            
            'Calculate Extension Stem Diameter, if its value is zero
            dExStemDia = IIf(CmpDblEqual(parExStemDia, 0), 0.5 * dGearDiameter, parExStemDia)
        End If
        
        Dim dShaftHeight As Double
        dShaftHeight = IIf(lPartDataBasis = 27 Or lPartDataBasis = 28, parShaftHeight, 0)
        If ((lPartDataBasis = 27 Or lPartDataBasis = 28) And _
                                CmpDblEqual(parShaftHeight, 0)) Then GoTo ErrorLabel
        
        If lPartDataBasis = 21 Or lPartDataBasis = 24 Then
            'The spec writer must specify a dimensional value for either Valve Operator Height,
            'Valve Centerline to Top of Handwheel, or Valve Centerline to Handwheel Centerline.
            If CmpDblEqual(parOperatorHeight, 0) And CmpDblEqual(parValCenToHWTop, 0) And _
                            CmpDblEqual(parValCenToHWCen, 0) Then
                GoTo ErrorLabel
            End If
        End If
        
        Dim dValCenToHWCen As Double
        dValCenToHWCen = parValCenToHWCen
        If CmpDblEqual(parValCenToHWCen, 0) Then
            If CmpDblGreaterthan(parHwYOffsetFrmGearBot, LINEAR_TOLERANCE) Then
                dValCenToHWCen = parOffsetFrmValCen + parHwYOffsetFrmGearBot + dExStemLen + dShaftHeight
                'For partdatabasis = 25, 26 Extension Stem Length has value
                'For partdatabasis = 27, 28 Shaft Height has value
            ElseIf Not CmpDblEqual(parValCenToHWTop, 0) Then
                dValCenToHWCen = parValCenToHWTop - dHandwheelDiameter / 2
            Else
                dValCenToHWCen = IIf(lPartDataBasis = 21 Or lPartDataBasis = 24, _
                                    parOperatorHeight - 0.25 * dHandwheelDiameter, _
                                    IIf(CmpDblGreaterthan(parOperatorHeight, LINEAR_TOLERANCE), _
                                        parOperatorHeight - 0.25 * dHandwheelDiameter, _
                                        parOffsetFrmValCen + dExStemLen + dShaftHeight + dGearDepth / 2))
            End If
        End If
        
        dOpearatorHeight = IIf(CmpDblGreaterthan(parOperatorHeight, LINEAR_TOLERANCE), _
                            parOperatorHeight, _
                            IIf(CmpDblGreaterthan(parGearDepth, LINEAR_TOLERANCE), _
                                parOffsetFrmValCen + dExStemLen + dShaftHeight + dGearDepth, _
                                dValCenToHWCen - 0.4 * dHandwheelDiameter + dGearDepth))
        
        'Create the Premitives
        'Create Handwheel
        iHandwheelPos = IIf(lPartDataBasis = 21 Or lPartDataBasis = 25 Or lPartDataBasis = 27, _
                                                -1, 1)
        
        'Handwheel Stem
        oStPoint.Set iHandwheelPos * dHwXOffsetFrmOpStem, dValCenToHWCen, 0
        'oEnPoint.Set oStPoint.x, oStPoint.y, parHwZOffsetOpFrmStem
        oEnPoint.Set oStPoint.x, oStPoint.y, dHwZOffsetOpFrmStem
        iOutput = iOutput + 1
        Call CreateCylinderAddtoOutput(m_OutputColl, oStPoint, oEnPoint, _
                                    0.05 * dHandwheelDiameter, True, arrayOfOutputs(iOutput))
        
        'Handwheel
        oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z - 0.15 * dHandwheelDiameter
        oEnPoint.Set oStPoint.x, oStPoint.y, oStPoint.z + 0.15 * dHandwheelDiameter
        
        iOutput = iOutput + 1
        Call CreateCylinderAddtoOutput(m_OutputColl, oStPoint, oEnPoint, dHandwheelDiameter, _
                                                True, arrayOfOutputs(iOutput))
        
    Case Else
        GoTo ErrorLabel
    End Select
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
               
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

Private Function ValueForOptionalParameter(OptionalParameter, SourceParameter, Ratio_OptionalToSource) As Double
    If CmpDblEqual(OptionalParameter, 0) Then
        ValueForOptionalParameter = Ratio_OptionalToSource * SourceParameter
    Else
        ValueForOptionalParameter = OptionalParameter
    End If
End Function

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
