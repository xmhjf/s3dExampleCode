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
'   Author:         RRK
'   Creation Date:  Monday 29, Jan 2007
'   Description:
'   This implements the physical aspect of Single Basket Strainer symbol to
'   support multiple dimensional bases as per CR-113397.
'   The symbol supports the following dimensional bases
'   1.)Part Data Basis Value 140:-Single Basket Strainer, specified by face-to-face dimension with dimensions from centerline
'   2.)Part Data Basis Value 145:-Single Basket Strainer, specified by face-to-face dimension with overall height
'   3.)Part Data Basis Value 147:-'Single Basket Strainer, Asymmetrical, specified by face-to-face dimension
'                                   with overall height'
'   4.)Part Data Basis Value 1091:-Angle Basket Strainer,Type1
'   5.)Part Data Basis Value 1092:-Angle Basket Strainer,Type2 with a ring at the top
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----    ------------------
'   07.May.2007     RRK      CR-117254.  Modified the code to support new dimensional basis value 147
'   8.Oct.2008     VKC      CR-149709.  Modified the code to support new dimensional basis value 1091 and 1092
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

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
    Dim ObjCylinder As Object
    Dim objCover As Object
    
    Dim parFacetoFace As Double
    Dim parStrainerHeight As Double
    Dim parCentertoBottom As Double
    Dim parCoverDiameter As Double
    Dim parDrainOffset As Double
    Dim parDiameter As Double
    
    Dim parInsulationThickness As Double
    Dim parFacetoCenter As Double

    Dim dCylinderDiam As Double
' Inputs
     Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
'    parStrainerHeight = arrayOfInputs(3)
'    parCentertoBottom = arrayOfInputs(4)
'    parCoverDiameter = arrayOfInputs(5)
'    parInsulationThickness = arrayOfInputs(6)
    
    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    If (lPartDataBasis <= 1 Or lPartDataBasis = 140 Or lPartDataBasis = 145) Then
        parFacetoFace = arrayOfInputs(2)
        parStrainerHeight = arrayOfInputs(3)
        parCentertoBottom = arrayOfInputs(4)
        parCoverDiameter = arrayOfInputs(5)
    End If
    
    If (lPartDataBasis = 147) Then
        parFacetoFace = arrayOfInputs(2)
        parStrainerHeight = arrayOfInputs(3)
        parCentertoBottom = arrayOfInputs(4)
        parCoverDiameter = arrayOfInputs(5)
        parFacetoCenter = arrayOfInputs(7)
    End If
    
    If (lPartDataBasis = 1091 Or lPartDataBasis = 1092) Then
        parStrainerHeight = arrayOfInputs(3)
        parCentertoBottom = arrayOfInputs(4)
        parCoverDiameter = arrayOfInputs(5)
        parFacetoCenter = arrayOfInputs(7)
        parDrainOffset = arrayOfInputs(8)
        parDiameter = arrayOfInputs(9)
    End If
    
    Dim oStPoint   As Object
    Dim oEnPoint   As Object
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition

'   Place Output 1 (Vertical Cylinder)
    Dim dFactor As Double
    dFactor = 0.7
    
    'Assuming diameter of the cylinder as 0.5 times Face to face
    If (lPartDataBasis <= 1 Or lPartDataBasis = 140 Or lPartDataBasis = 145 Or lPartDataBasis = 147) Then
        dCylinderDiam = 0.5 * parFacetoFace
        'Here the Value of Diameter or CoverDiameter may be specified such that the
        'other value is computed as a function of the value that has been specified.
    ElseIf (lPartDataBasis = 1091 Or lPartDataBasis = 1092) Then
        If (parDiameter = 0) Then
            parDiameter = 0.8 * parCoverDiameter
            dCylinderDiam = parDiameter
        Else
            dCylinderDiam = parDiameter
        End If
    End If
    
    'If diameter of the cylinder is less than pipe OD then the cylinder diameter is taken as 1.1 times Pipe OD
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblLessThan(dCylinderDiam, pipeDiam) Then
        dCylinderDiam = 1.1 * pipeDiam
    End If
    
    Dim dStrainerbottomtoInletNozz As Double
    dStrainerbottomtoInletNozz = parFacetoCenter - (dCylinderDiam / 2)

    If (lPartDataBasis <= 1 Or lPartDataBasis = 140) Then
        'Assuming the top position of the vertical cylinder to be dFactor * parStrainerHeight
        'above the origin.
        oStPoint.Set 0, dFactor * parStrainerHeight, 0
        oEnPoint.Set 0, -parCentertoBottom, 0
    ElseIf (lPartDataBasis = 145 Or lPartDataBasis = 147) Then
        oStPoint.Set 0, dFactor * (parStrainerHeight - parCentertoBottom), 0
        oEnPoint.Set 0, -parCentertoBottom, 0
    ElseIf (lPartDataBasis = 1091) Then
        oStPoint.Set 0.4 * (parStrainerHeight - dStrainerbottomtoInletNozz), 0, 0
        oEnPoint.Set -parCentertoBottom, 0, 0
    ElseIf (lPartDataBasis = 1092) Then
        oStPoint.Set 0.3 * (parStrainerHeight - dStrainerbottomtoInletNozz), 0, 0
        oEnPoint.Set -parCentertoBottom, 0, 0
    End If
    
    Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCylinderDiam, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder
    Set ObjCylinder = Nothing

'   Place Output 2(Cover)
    
    If (lPartDataBasis <= 1 Or lPartDataBasis = 140) Then
        oStPoint.Set 0, parStrainerHeight, 0
        oEnPoint.Set 0, dFactor * parStrainerHeight, 0
    ElseIf (lPartDataBasis = 145 Or lPartDataBasis = 147) Then
        oStPoint.Set 0, parStrainerHeight - parCentertoBottom, 0
        oEnPoint.Set 0, dFactor * (parStrainerHeight - parCentertoBottom), 0
    ElseIf (lPartDataBasis = 1091) Then
        oStPoint.Set 0.4 * (parStrainerHeight - dStrainerbottomtoInletNozz), 0, 0
        oEnPoint.Set 0.5 * (parStrainerHeight - dStrainerbottomtoInletNozz), 0, 0
    ElseIf (lPartDataBasis = 1092) Then
        oStPoint.Set 0.3 * (parStrainerHeight - dStrainerbottomtoInletNozz), 0, 0
        oEnPoint.Set 0.4 * (parStrainerHeight - dStrainerbottomtoInletNozz), 0, 0
    End If
    
    If CmpDblEqual(parCoverDiameter, 0) Then
        parCoverDiameter = 1.3 * dCylinderDiam
    End If
    
    Set objCover = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parCoverDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCover
    Set objCover = Nothing
'Place the Handle at the top
    If (lPartDataBasis = 1092) Then
        Dim objHandle As Object
        Dim TorOrigin As New AutoMath.DPosition
        Dim TorMajrad As Double
        Dim TorMinRad As Double
        Dim oAxis     As AutoMath.DVector
        Set oAxis = New AutoMath.DVector
        oAxis.Set 0, 0, 1
        TorMajrad = 0.05 * (parStrainerHeight - dStrainerbottomtoInletNozz)
        TorMinRad = 0.025 * (parStrainerHeight - dStrainerbottomtoInletNozz)
        TorOrigin.Set 0.45 * (parStrainerHeight - dStrainerbottomtoInletNozz), 0, 0
  
        Set objHandle = PlaceTorus(m_OutputColl, TorOrigin, oAxis, TorMajrad, TorMinRad)
'Set the Output
        m_OutputColl.AddOutput "Handle", objHandle
        Set objHandle = Nothing
        Set TorOrigin = Nothing
    End If
    
' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Dim dNozzleLength As Double
    
    Dim dPort1PosX As Double
    Dim dPort2PosX As Double
    Dim dPort2PosY As Double
    
    If (lPartDataBasis <= 1 Or lPartDataBasis = 140 Or lPartDataBasis = 145) Then
        dPort1PosX = -parFacetoFace / 2
        dPort2PosX = parFacetoFace / 2
        dNozzleLength = parFacetoFace
        oDir.Set -1, 0, 0
        oPlacePoint.Set dPort1PosX - sptOffset + depth, 0, 0
    ElseIf lPartDataBasis = 147 Then
        dPort1PosX = -parFacetoFace + parFacetoCenter
        dPort2PosX = parFacetoCenter
        dNozzleLength = parFacetoFace
        oDir.Set -1, 0, 0
        oPlacePoint.Set dPort1PosX - sptOffset + depth, 0, 0
    ElseIf lPartDataBasis = 1091 Or lPartDataBasis = 1092 Then
        dPort1PosX = -parCentertoBottom - dStrainerbottomtoInletNozz
        oPlacePoint.Set dPort1PosX + sptOffset - depth, 0, 0
        dNozzleLength = dStrainerbottomtoInletNozz
        oDir.Set -1, 0, 0
    End If
    
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    If (lPartDataBasis <= 1 Or lPartDataBasis = 140 Or lPartDataBasis = 145 Or lPartDataBasis = 147) Then
         oPlacePoint.Set dPort2PosX + sptOffset - depth, 0, 0
         oDir.Set 1, 0, 0
         Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    ElseIf (lPartDataBasis = 1091 Or lPartDataBasis = 1092) Then
        dPort2PosY = -parFacetoCenter
        dNozzleLength = parFacetoCenter
        oDir.Set 0, -1, 0
        oPlacePoint.Set 0, dPort2PosY - sptOffset + depth, 0
        Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)
    End If
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 3(For the Drain)
    If (lPartDataBasis = 1091 Or lPartDataBasis = 1092) Then
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        oPlacePoint.Set -(parCentertoBottom + dStrainerbottomtoInletNozz) + parDrainOffset, dCylinderDiam / 2, 0
        oDir.Set 0, 1, 0
        Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
        m_OutputColl.AddOutput "Port3", objNozzle
        Set objNozzle = Nothing
    End If
    
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
