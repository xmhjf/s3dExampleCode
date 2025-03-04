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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          BG
'   Creation Date:  Thursday, Jan 31 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   03.Dec.2004     svsmylav        Defined new input parameter HandwheelAngle and modified valve operator direction vectors.
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  27.Dec.2007     VRK  CR-CP·132638 Auto Recirculation Valve without an Operator
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
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjMainCone As Object
    Dim ObjUpperCone As Object
    Dim ObjRightCylinder As Object
    Dim ObjInsulCylinder As Object
    Dim ObjInsulBranch As Object
    Dim ObjNozzle1 As Object
    Dim ObjNozzle2 As Object
    Dim ObjNozzle3 As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parInsulationThickness As Double
    
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    Dim CylinderLen As Double
    Dim parHandwheelAngle   As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parFace3toCenter = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    parHandwheelAngle = arrayOfInputs(6)
    
    iOutput = 0

 ' Insert your code for output 1(Main Cone)
    Dim BaseCenPoint As New AutoMath.DPosition
    Dim TopCenPoint As New AutoMath.DPosition
    Dim grPipeDia As Double
    Dim MiniCylinTk As Double
    
    MiniCylinTk = 0.01
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    BaseCenPoint.Set -parFace1toCenter + flangeThick, 0#, 0#
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2
    If pipeDiam > pipeDiam2 Then                        ' Finding the greater of twi pipe Diameters
        grPipeDia = pipeDiam
    Else
        grPipeDia = pipeDiam2
    End If
    If (0.2 * grPipeDia) > MiniCylinTk Then         ' Determining the minimum thickness of Cylinder(Output 3)
        CylinderLen = MiniCylinTk
    Else
        CylinderLen = 0.2 * grPipeDia
    End If
    If flangeThick2 > 0 Then
        TopCenPoint.Set parFace2toCenter - flangeThick2, 0#, 0
    Else
        TopCenPoint.Set parFace2toCenter - CylinderLen, 0#, 0
    End If
    Set ObjMainCone = PlaceCone(m_OutputColl, BaseCenPoint, TopCenPoint, pipeDiam / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMainCone
    Set ObjMainCone = Nothing
    
 ' Insert your code for output 2(Upper Cone)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim UpperconeBasePt As New AutoMath.DPosition
    Dim UpperconeTopPt As New AutoMath.DPosition
    UpperconeBasePt.Set 0, parFace3toCenter - flangeThick, 0#
    UpperconeTopPt.Set 0#, 0#, 0
    Set ObjUpperCone = PlaceCone(m_OutputColl, UpperconeBasePt, UpperconeTopPt, pipeDiam / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjUpperCone
    Set ObjUpperCone = Nothing
    
 ' Insert your code for output 3(Right Cylinder)
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    stPoint.Set parFace2toCenter - CylinderLen, 0, 0
    enPoint.Set parFace2toCenter, 0, 0
    Set ObjRightCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRightCylinder
    Set ObjRightCylinder = Nothing
    
 ' Insert your code for output 6(Nozzle 1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim faceToFace  As Double

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set oPartFclt = arrayOfInputs(1)
    Set ObjNozzle1 = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle1
    Set ObjNozzle1 = Nothing
    
 ' Insert your code for output 7(Nozzle 2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set parFace2toCenter + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    Set ObjNozzle2 = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle2
    Set ObjNozzle2 = Nothing
    
 ' Insert your code for output 8(Nozzle 3)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set 0, parFace3toCenter + sptOffset - depth, 0
    oDir.Set 0, 1, 0
    Set ObjNozzle3 = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle3
    Set ObjNozzle3 = Nothing
    
' Insert your code for output 9 (Valve Operator)
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
    
    oDirX.Set Cos(parHandwheelAngle), Sin(parHandwheelAngle), 0
    oDirY.Set 0, 0, 1
    oDirZ.Set -Sin(parHandwheelAngle), Cos(parHandwheelAngle), 0
        
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
            OpOrigin.Set 0, 0, 0
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


