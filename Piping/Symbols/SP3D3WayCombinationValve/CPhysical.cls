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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          RRK
'   Creation Date:  Friday, Aug 31 2007
'   Description:    This symbol is prepared for 3 way diverter combination valve
'                   of Tyco Flow Control as per CR-126718
'   Source:         http://www.tycoflowcontrol-pc.com/ld/F605_4_07.pdf
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Const E_NOOPERATORDEFINED = &H80040400 'Corresponds to E_PIPECOMP_NOOPERATORDEFINED in RefDataPipingFacelets.idl
Private Const E_NOOPERATORPARTFOUND = &H80040401 'Corresponds to E_PIPECOMP_NOVALVEPARTFOUND in RefDataPipingFacelets.idl

Private Sub Class_Initialize()

'''

End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim pipeDiam1        As Double
    Dim flangeThick1     As Double
    Dim sptOffset1       As Double
    Dim flangeDiam1      As Double
    Dim depth1           As Double
    
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    
    Dim pipeDiam3        As Double
    Dim flangeThick3     As Double
    Dim sptOffset3       As Double
    Dim flangeDiam3      As Double
    Dim depth3           As Double
    
    Dim iOutput     As Double

    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parPorttoPortOffset As Double
    Dim parInsulationThickness As Double
    Dim parHandwheelAngle   As Double

    Dim oStPoint   As AutoMath.DPosition
    Dim oEnPoint   As AutoMath.DPosition
 
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parFace3toCenter = arrayOfInputs(4)
    parPorttoPortOffset = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    parHandwheelAngle = arrayOfInputs(7)
    
    If CmpDblLessThanOrEqualTo(parFace1toCenter, 0) Then
        parFace1toCenter = parFace3toCenter
    End If
    
    iOutput = 0

' Insert your code for output 1(Left Valve Body)
    Dim objLeftValveBody As Object
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam1, sptOffset1, depth1
    
    oStPoint.Set -parFace1toCenter, 0, 0
    oEnPoint.Set 0, 0, 0
    
    Set objLeftValveBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam1, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLeftValveBody
    Set objLeftValveBody = Nothing
    
' Insert your code for output 2(Vertical Valve Body)
    Dim objVertValveBody As Object
    
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam3, sptOffset3, depth3
    'Assumption: Height of the cylinder(valve body 2) is assumed to be twice the Face 2 to center
    oStPoint.Set 0, -parFace2toCenter, 0
    oEnPoint.Set 0, parPorttoPortOffset + (pipeDiam3 / 2), 0
    
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2
    Set objVertValveBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam2, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objVertValveBody
    Set objVertValveBody = Nothing
    
' Insert your code for output 3(Body for two Flanges)
    Dim objFlange As Object
    
    'Assumptions:     1.)The diameter of flanges is assumed to be 1.5 times the smaller of
    '                    parFace1toCenter and parFace3toCenter
    '                 2.)Total Flange Thickness of twp flanges is calculated as:
    '                   parPorttoPortOffset - (pipeDiam1 / 2 + pipeDiam3 / 2 + 0.1 * parPorttoPortOffset)
                      
    Dim dTotalFlangeThk As Double
    Dim dFlangeDia As Double
    
    dTotalFlangeThk = parPorttoPortOffset - (pipeDiam1 / 2 + pipeDiam3 / 2 + 0.1 * parPorttoPortOffset)
    
    'Check in case dTotalFlangeThk becomes zero or negative
    If CmpDblLessThanOrEqualTo(dTotalFlangeThk, 0) Then dTotalFlangeThk = 0.02
    
    dFlangeDia = 1.5 * SmallerDim(parFace3toCenter, parFace1toCenter)
    
    oStPoint.Set 0, (pipeDiam1 / 2) + (0.05 * parPorttoPortOffset), 0
    oEnPoint.Set 0, oStPoint.y + dTotalFlangeThk, 0
    
    Set objFlange = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dFlangeDia, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objFlange
    Set objFlange = Nothing
        
' Insert your code for output 4(Right Valve Body)
    Dim objRightValveBody As Object
    
    oStPoint.Set parFace3toCenter, parPorttoPortOffset, 0
    oEnPoint.Set 0, parPorttoPortOffset, 0
    
    Set objRightValveBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam3, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objRightValveBody
    Set objRightValveBody = Nothing
    
' Place Nozzle 1

    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim faceToFace  As Double

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parFace1toCenter - sptOffset1 + depth1, 0, 0
    oDir.Set -1, 0, 0
    
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set 0, -parFace2toCenter - sptOffset2 + depth2, 0
    oDir.Set 0, -1, 0
    
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 3
    oPlacePoint.Set parFace3toCenter + sptOffset3 - depth3, parPorttoPortOffset, 0
    oDir.Set 1, 0, 0
    
    Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
'   Insert your code for Valve Actuator
'--------------------------------------
    Dim oSymbolHelper As IJSymbolGeometryHelper
    Set oSymbolHelper = New SP3DSymbolHelper.SymbolServices
    oSymbolHelper.OutputCollection = m_OutputColl
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt

    On Error Resume Next
    Dim oDirX As IJDVector
    Dim oDirY As IJDVector
    Dim oDirZ As IJDVector
    Set oDirX = New DVector
    Set oDirY = New DVector
    Set oDirZ = New DVector
    
    oDirX.Set Cos(parHandwheelAngle), 0, Sin(parHandwheelAngle)
    oDirY.Set 0, 1, 0
    oDirZ.Set -Sin(parHandwheelAngle), 0, Cos(parHandwheelAngle)
    
    Dim oOperatorPart As IJDPart
    Dim oOperatorOcc   As IJPartOcc
    
    If Not oPipeComponent Is Nothing Then
'       Note: oPipeComponent.GetValveOperatorPart method call may fail in following two cases:
'        1. If valve operator catalog part number is not defined (to check for E_NOOPERATORDEFINED) or
'        2. If valve operator part is not found (to check for E_NOOPERATORPARTFOUND)
'        symbol should raise error for case 2.
        On Error GoTo HandleOperatorCodeLabel

        Set oOperatorPart = oPipeComponent.GetValveOperatorPart

        If Not oOperatorPart Is Nothing Then
            Dim OpOrigin As IJDPosition
            Set OpOrigin = New DPosition
            
            OpOrigin.Set 0, parPorttoPortOffset, 0
            Set oOperatorOcc = oSymbolHelper.CreateChildPartOcc("ValveOperator", oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)
        End If
    End If
    
HandleOperatorCodeLabel:
'   The symbol code execution should not raise error for case 1 i.e. If Err.Number = E_NOOPERATORDEFINED
    If Err.Number = E_NOOPERATORPARTFOUND Then          'Case 2
        Err.Raise Err.Number, MODULE & METHOD, "No operator part found."
    End If
   
    Set oSymbolHelper = Nothing
    Set oOperatorPart = Nothing
    Set oPipeComponent = Nothing
    Set oOperatorOcc = Nothing
  
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
