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
'   Copyright (c) 2007-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
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
'   5.)Part Data Basis Value 1092:-Angle Basket Strainer,Type2
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----    ------------------
'   07.May.2007     RRK      CR-117254.  Modified the code to support new dimensional basis value 147
'   8.Oct.2008      VKC      CR-149709.  Modified the code to support new dimensional basis value 1091 and 1092.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim iOutput         As Double
    Dim ObjInsVerCylinder As Object
    Dim ObjInsCover     As Object
    Dim ObjInsHorCylinder As Object
    Dim ObjInsPort1     As Object
    Dim ObjInsPort2     As Object
    
    Dim parFacetoFace As Double
    Dim parStrainerHeight As Double 'A
    Dim parCentertoBottom As Double 'B
    Dim parCoverDiameter As Double  'C
    Dim parInsulationThickness As Double
    Dim parFacetoCenter As Double
    Dim parDrainOffset As Double
    Dim parDiameter As Double
    
    Dim dInsulationDia As Double
    Dim dCylinderDiam As Double
' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
'    parStrainerHeight = arrayOfInputs(3) 'A
'    parCentertoBottom = arrayOfInputs(4) 'B
'    parCoverDiameter = arrayOfInputs(5)  'C
'    parInsulationThickness = arrayOfInputs(6)
    
    iOutput = 0
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
        parInsulationThickness = arrayOfInputs(6)
    End If
    
    If (lPartDataBasis = 147) Then
        parFacetoFace = arrayOfInputs(2)
        parStrainerHeight = arrayOfInputs(3)
        parCentertoBottom = arrayOfInputs(4)
        parCoverDiameter = arrayOfInputs(5)
        parFacetoCenter = arrayOfInputs(7)
        parInsulationThickness = arrayOfInputs(6)
    End If
    
    If (lPartDataBasis = 1091 Or lPartDataBasis = 1092) Then
        parStrainerHeight = arrayOfInputs(3)
        parCentertoBottom = arrayOfInputs(4)
        parCoverDiameter = arrayOfInputs(5)
        parFacetoCenter = arrayOfInputs(7)
        parDrainOffset = arrayOfInputs(8)
        parDiameter = arrayOfInputs(9)
        parInsulationThickness = arrayOfInputs(6)
    End If
    
    iOutput = 0
    
    Dim oStPoint   As Object
    Dim oEnPoint   As Object
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
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
    
'   Place Output 5 (InsVerCylinder)
    If (lPartDataBasis <= 1 Or lPartDataBasis = 140) Then
        oStPoint.Set 0, dFactor * parStrainerHeight, 0
        oEnPoint.Set 0, -parCentertoBottom, 0
    ElseIf (lPartDataBasis = 145 Or lPartDataBasis = 147) Then
        oStPoint.Set 0, dFactor * (parStrainerHeight - parCentertoBottom), 0
        oEnPoint.Set 0, -parCentertoBottom, 0
    ElseIf (lPartDataBasis = 1091) Then
        oStPoint.Set 0.4 * (parStrainerHeight - dStrainerbottomtoInletNozz), 0, 0
        oEnPoint.Set -parCentertoBottom - parInsulationThickness, 0, 0
    ElseIf (lPartDataBasis = 1092) Then
        oStPoint.Set 0.3 * (parStrainerHeight - dStrainerbottomtoInletNozz), 0, 0
        oEnPoint.Set -parCentertoBottom - parInsulationThickness, 0, 0
    End If
    
    dInsulationDia = dCylinderDiam + 2 * parInsulationThickness
    Set ObjInsVerCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsVerCylinder
    Set ObjInsVerCylinder = Nothing

'   Place Output 6(InsCover)
    If (lPartDataBasis <= 1 Or lPartDataBasis = 140) Then
        oStPoint.Set 0, parStrainerHeight + parInsulationThickness, 0
        oEnPoint.Set 0, dFactor * parStrainerHeight, 0
    ElseIf (lPartDataBasis = 145 Or lPartDataBasis = 147) Then
        oStPoint.Set 0, parStrainerHeight - parCentertoBottom + parInsulationThickness, 0
        oEnPoint.Set 0, dFactor * (parStrainerHeight - parCentertoBottom), 0
    ElseIf (lPartDataBasis = 1091) Then
        oStPoint.Set 0.4 * (parStrainerHeight - dStrainerbottomtoInletNozz) - parInsulationThickness, 0, 0
        oEnPoint.Set 0.5 * (parStrainerHeight - dStrainerbottomtoInletNozz) + parInsulationThickness, 0, 0
    ElseIf (lPartDataBasis = 1092) Then
        oStPoint.Set 0.3 * (parStrainerHeight - dStrainerbottomtoInletNozz) - parInsulationThickness, 0, 0
        oEnPoint.Set 0.4 * (parStrainerHeight - dStrainerbottomtoInletNozz) + parInsulationThickness, 0, 0
    End If
    
    If CmpDblEqual(parCoverDiameter, 0) Then
        parCoverDiameter = 1.3 * dCylinderDiam
    End If
    
    dInsulationDia = parCoverDiameter + 2 * parInsulationThickness
    Set ObjInsCover = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsCover
    Set ObjInsCover = Nothing

'   Place Output 7(InsHorCylinder)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim dPort1PosX As Double
    Dim dPort2PosX As Double
    Dim dPort2PosY As Double
    
    If (lPartDataBasis <= 1 Or lPartDataBasis = 140 Or lPartDataBasis = 145) Then
        dPort1PosX = -parFacetoFace / 2
        dPort2PosX = parFacetoFace / 2
        oStPoint.Set dPort1PosX, 0, 0
        oEnPoint.Set dPort2PosX, 0, 0
    ElseIf lPartDataBasis = 147 Then
        dPort1PosX = -parFacetoFace + parFacetoCenter
        dPort2PosX = parFacetoCenter
        oStPoint.Set dPort1PosX, 0, 0
        oEnPoint.Set dPort2PosX, 0, 0
    ElseIf (lPartDataBasis = 1091 Or lPartDataBasis = 1092) Then
        dPort2PosY = -parFacetoCenter
        oStPoint.Set 0, dPort2PosY, 0
        oEnPoint.Set 0, -parDiameter / 2, 0
    End If
    
    dInsulationDia = pipeDiam + 2 * parInsulationThickness
    Set ObjInsHorCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsHorCylinder
    Set ObjInsHorCylinder = Nothing
    
' Insert your code for output 8(ObjInsPort1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim dFlangeThickness As Double
    dFlangeThickness = flangeThick
    
    If CmpDblEqual(dFlangeThickness, 0) Then
        dFlangeThickness = 0.0001
    End If
    
    If (lPartDataBasis <= 1 Or lPartDataBasis = 140 Or lPartDataBasis = 145 Or lPartDataBasis = 147) Then
        oStPoint.Set dPort1PosX, 0, 0
        oEnPoint.Set dPort1PosX + dFlangeThickness + parInsulationThickness, 0, 0
    ElseIf (lPartDataBasis = 1091 Or lPartDataBasis = 1092) Then
        dPort1PosX = -parCentertoBottom - dStrainerbottomtoInletNozz
        oStPoint.Set dPort1PosX, 0, 0
        oEnPoint.Set dPort1PosX + dFlangeThickness + parInsulationThickness, 0, 0
    End If
    
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then
        dInsulationDia = pipeDiam + parInsulationThickness * 2
    End If
    dInsulationDia = flangeDiam + parInsulationThickness * 2
    Set ObjInsPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort1
    Set ObjInsPort1 = Nothing

 ' Insert your code for output 9(ObjInsPort2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    dFlangeThickness = flangeThick
    If CmpDblEqual(dFlangeThickness, 0) Then
        dFlangeThickness = 0.0001
    End If
    
    oStPoint.Set 0, dPort2PosY, 0
    oEnPoint.Set 0, dPort2PosY + dFlangeThickness + parInsulationThickness, 0
    
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then
        dInsulationDia = pipeDiam + parInsulationThickness * 2
    End If
    dInsulationDia = flangeDiam + parInsulationThickness * 2
    Set ObjInsPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort2
    Set ObjInsPort2 = Nothing
    
' Insert your code for Drain Port(ObjInsPort3)
    If (lPartDataBasis = 1091 Or lPartDataBasis = 1092) Then
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        
        Dim dPort3PosX As Double
        Dim dPort3PosY As Double
        Dim ObjInsPort3 As Object
        dFlangeThickness = flangeThick
        
        If CmpDblEqual(dFlangeThickness, 0) Then
                dFlangeThickness = 0.0001
        End If
            
            dPort3PosY = dCylinderDiam / 2
            dPort3PosX = -(parCentertoBottom + dStrainerbottomtoInletNozz) + parDrainOffset
            
        If CmpDblEqual(dFlangeThickness, 0) Then
            dFlangeThickness = 0.0001
        End If
            oStPoint.Set dPort3PosX, dPort3PosY, 0
            oEnPoint.Set dPort3PosX, dPort3PosY - sptOffset + depth + parInsulationThickness, 0
            
        If CmpDblGreaterthan(pipeDiam, flangeDiam) Then
           dInsulationDia = pipeDiam + parInsulationThickness * 2
    End If
        dInsulationDia = flangeDiam + parInsulationThickness * 2
        Set ObjInsPort3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)

' Set the output
        m_OutputColl.AddOutput "InsDrainPort", ObjInsPort3
        Set ObjInsPort3 = Nothing
    End If
        
'Insert your code for the Nozzle length at the bottom
    If (lPartDataBasis = 1091 Or lPartDataBasis = 1092) Then
        oStPoint.Set -parCentertoBottom, 0, 0
        oEnPoint.Set -parCentertoBottom - dStrainerbottomtoInletNozz, 0, 0
    
    
        dInsulationDia = flangeDiam + parInsulationThickness * 2
        If CmpDblGreaterthan(pipeDiam, flangeDiam) Then
            dInsulationDia = pipeDiam + parInsulationThickness * 2
        End If
        Set ObjInsHorCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
' Set the output
        m_OutputColl.AddOutput "InsFlange", ObjInsHorCylinder
        Set ObjInsHorCylinder = Nothing
    End If
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
