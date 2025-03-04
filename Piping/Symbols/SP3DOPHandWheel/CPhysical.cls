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
'   Copyright (c) 2007-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   ProgID:         SP3DOPHandwheel.OPHandWheel
'   Author:         KKC
'   Creation Date:  Tuesday, April 10, 2007
'   Description: This symbol is graphical implementation of the Technegate hand wheel operator
'   This symbol implements following partdatabasis
'   Default
'
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   04.05.2007      KKC     Created: CR-117167  Create valve symbols for use in mining industry
'   22.08.2008      RUK     CR-148066  Provide more realistic handwheel operator symbol
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
    Dim iOutput As Long
    
    Dim parOperatorHeight As Double
    Dim parOperatorDiameter As Double
    Dim parOffsetFrmValCen As Double
    Dim parHandwheelDiameter As Double
    Dim parHandwheelToStemTop As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parOperatorHeight = arrayOfInputs(2)
    'parOperatorDiameter = arrayOfInputs(3)
    'parOffsetFrmValCen = arrayOfInputs(4)
    'parHandwheelDiameter = arrayOfInputs(5)
    'parHandwheelToStemTop = arrayOfInputs(6)

    'Initialize
    iOutput = 0
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oCenter As AutoMath.DPosition
    Dim oVector As AutoMath.DVector
    Dim CenterPos As AutoMath.DPosition
    Dim CylBotOrigin As AutoMath.DPosition
    Dim CylTopOrigin As AutoMath.DPosition
    Dim TorOrigin As AutoMath.DPosition
    Dim TorAxis As AutoMath.DVector
    Dim objHandwheel As Object
    
    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition
    Set oCenter = New DPosition
    Set oVector = New DVector
    Set CenterPos = New AutoMath.DPosition
    Set CylBotOrigin = New AutoMath.DPosition
    Set CylTopOrigin = New AutoMath.DPosition
    Set TorOrigin = New AutoMath.DPosition
    Set TorAxis = New AutoMath.DVector
    
    Dim dStemDia As Double
    Dim dSpokeDia As Double
    Dim dHwTorDia As Double
    
    'Checking for the PartDataBasis
    Dim oOperator As IJValveOperator
    Dim lPartDataBasis As Long
    Set oOperator = oPartFclt
    lPartDataBasis = oOperator.DimensionalBasis
    Set oOperator = Nothing

    Select Case lPartDataBasis
    Case Is <= 1
        parOperatorHeight = arrayOfInputs(2)
        parOperatorDiameter = arrayOfInputs(3)
    
        Dim dTorPriRad As Double
        Dim dTorSecRad As Double
        Dim dActuatorWheelMinimumThickness As Double
        Dim dStemExtension As Double
    
        dStemDia = parOperatorDiameter * 0.1
        '   Check that the stem diameter is within 15mm and 50mm
        If CmpDblLessThan(dStemDia, 0.015) Then dStemDia = 0.015
        If CmpDblGreaterthan(dStemDia, 0.05) Then dStemDia = 0.05
    
        dActuatorWheelMinimumThickness = dStemDia
        dStemExtension = parOperatorHeight * 0.3
    
        'Prepare Torus Radii
        dTorSecRad = dActuatorWheelMinimumThickness / 2
        dTorPriRad = parOperatorDiameter / 2 - dTorSecRad
    
        'Prepare data points
        CenterPos.Set 0, 0, 0
    
        CylBotOrigin.Set CenterPos.x, CenterPos.y, CenterPos.z
        CylTopOrigin.Set CylBotOrigin.x, CylBotOrigin.y + parOperatorHeight, CylBotOrigin.z
    
        TorOrigin.Set CylTopOrigin.x, CylTopOrigin.y - dStemExtension - dTorSecRad, CylTopOrigin.z
    
        'Prepare axis vector for Torus
        TorAxis.Set 0, 1, 0
    
        ' Insert your code for output 1(Stem)
        'Place Cylinder
        Dim ObjStem As IngrGeom3D.Projection3d
    
        Set ObjStem = PlaceCylinder(m_OutputColl, CylBotOrigin, CylTopOrigin, dStemDia, True)
    
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
        Set ObjStem = Nothing
    
        ' Insert your code for output 2(Handwheel)
        'Place Torus
        Set objHandwheel = PlaceTorus(m_OutputColl, TorOrigin, TorAxis, dTorPriRad, dTorSecRad)
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHandwheel
        Set objHandwheel = Nothing
        Set CylBotOrigin = Nothing
        Set CylTopOrigin = Nothing
        Set CenterPos = Nothing
    
    Case 50
        parOperatorHeight = arrayOfInputs(2)
        parOffsetFrmValCen = arrayOfInputs(4)
        parHandwheelDiameter = arrayOfInputs(5)
        parHandwheelToStemTop = arrayOfInputs(6)
        
        dStemDia = 0.05 * parHandwheelDiameter
        dSpokeDia = 0.05 * parHandwheelDiameter
        
        If CmpDblLessThan(dStemDia, 0.015) Then dStemDia = 0.015
        If CmpDblGreaterthan(dStemDia, 0.05) Then dStemDia = 0.05
        
        dSpokeDia = dStemDia
        dHwTorDia = dStemDia
        
        'Handwheel Stem
        Dim objHwStem As Object
        oStPoint.Set 0, parOffsetFrmValCen, 0
        oEnPoint.Set 0, parOperatorHeight, 0
        
        Set objHwStem = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dStemDia, True)
        
        'Add to the Ouput
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHwStem
        Set objHwStem = Nothing
        
        'Handwheel
        oCenter.Set oEnPoint.x, oEnPoint.y - parHandwheelToStemTop, oEnPoint.z
        oVector.Set 0, 1, 0
        
        Set objHandwheel = PlaceTorus(m_OutputColl, oCenter, oVector, _
                                parHandwheelDiameter / 2 - dHwTorDia / 2, _
                                dHwTorDia / 2)
        
        'Add to the Ouput
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHandwheel
        Set objHandwheel = Nothing
        
        'Handwheel Spokes
        Dim objSpoke As Object
        
        'Spoke 1
        oStPoint.Set parHandwheelDiameter / 2 - dHwTorDia / 2, _
                            oEnPoint.y - parHandwheelToStemTop, oEnPoint.z
        oEnPoint.Set -(parHandwheelDiameter / 2 - dHwTorDia / 2), oStPoint.y, oStPoint.z
    
        Set objSpoke = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dSpokeDia, True)
        
        'Add to the Ouput
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "Spoke1", objSpoke
        Set objSpoke = Nothing
        
        'Spoke 2
        oStPoint.Set 0, oEnPoint.y, parHandwheelDiameter / 2 - dHwTorDia / 2
        oEnPoint.Set 0, oStPoint.y, -(parHandwheelDiameter / 2 - dHwTorDia / 2)
    
        Set objSpoke = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dSpokeDia, True)
        
        'Add to the Ouput
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "Spoke2", objSpoke
        Set objSpoke = Nothing
        
    Case Else
        GoTo ErrorLabel
    End Select
    
    'Remove the References
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oCenter = Nothing
    Set oVector = Nothing
    Set CenterPos = Nothing
    Set CylBotOrigin = Nothing
    Set CylTopOrigin = Nothing
    Set TorOrigin = Nothing
    Set TorAxis = Nothing
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

