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
'   Copyright (c) 2004-06, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         Mallepally Srinivas
'   Creation Date:  Wednesday, Jan 14, 2004
'   Description:
'       This symbol is created from SP3DangleValve.AngleValve, and cache property has been added.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----   ---------------------
'   31.Jul.2006     dkl     CR-28369: Removed hardcoded operator, added code for creating chilpartocc.
'                           Input parameters modified to accept Face1toCenter and Face2toCenter.
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

    Dim oPartFclt As PartFacelets.IJDPart
    Dim PipeDia As Double
    Dim FlangeThick As Double
    Dim CptOffset As Double
    Dim FlangeDia As Double
    Dim Depth As Double
    
    Dim CptOffset2 As Double
    Dim Depth2 As Double
    
    Dim iOutput As Double
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parHandwheelAngle   As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parHandwheelAngle = arrayOfInputs(4)

    RetrieveParameters 1, oPartFclt, m_OutputColl, _
                       PipeDia, FlangeThick, FlangeDia, CptOffset, Depth
    
    iOutput = 0

    Dim CenterPos As New AutoMath.DPosition
    CenterPos.Set 0, 0, 0

 ' Place Inlet
 '---------------------------------------------------------------------------------------
    Dim StartPos As New AutoMath.DPosition
    Dim EndPos As New AutoMath.DPosition
    
    StartPos.Set CenterPos.x - parFace1toCenter + FlangeThick, CenterPos.y, CenterPos.z
    EndPos.Set CenterPos.x, CenterPos.y, CenterPos.z
    
    Dim objInlet As IngrGeom3D.Cone3d
    Set objInlet = PlaceCone(m_OutputColl, StartPos, EndPos, PipeDia / 2, 0)
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInlet
    
    Set objInlet = Nothing
    
 ' Place Outlet
 '---------------------------------------------------------------------------------------
    RetrieveParameters 2, oPartFclt, m_OutputColl, _
                       PipeDia, FlangeThick, FlangeDia, CptOffset2, Depth2
    StartPos.Set CenterPos.x, CenterPos.y + parFace2toCenter - FlangeThick, CenterPos.z
    EndPos.Set CenterPos.x, CenterPos.y, CenterPos.z
    
    Dim objOutlet As IngrGeom3D.Cone3d
    Set objOutlet = PlaceCone(m_OutputColl, StartPos, EndPos, PipeDia / 2, 0)
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objOutlet
    
    Set objOutlet = Nothing
    
    Set StartPos = Nothing
    Set EndPos = Nothing
    
 ' Place Nozzle 1
 ' ------------------------------------------------------------------------------------------
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parFace1toCenter - CptOffset + Depth, 0, 0
    oDir.Set -1, 0, 0

    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

 ' Place Nozzle 2
 ' ------------------------------------------------------------------------------------------
    oPlacePoint.Set 0, parFace2toCenter + CptOffset2 - Depth2, 0
    oDir.Set 0, 1, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Insert your code for output 6 (Valve Operator)
' ------------------------------------------------------------------------------------------
    Dim oSymbolHelper As IJSymbolGeometryHelper
    Set oSymbolHelper = New SP3DSymbolHelper.SymbolServices
    oSymbolHelper.OutputCollection = m_OutputColl

    On Error Resume Next
    Dim oDirX As IJDVector
    Dim oDirY As IJDVector
    Dim oDirZ As IJDVector
    Set oDirX = New DVector
    Set oDirY = New DVector
    Set oDirZ = New DVector
    
    oDirX.Set Cos(parHandwheelAngle), 0, Sin(parHandwheelAngle)
    oDirY.Set 0, -1, 0
    oDirZ.Set Sin(parHandwheelAngle), 0, -Cos(parHandwheelAngle)

    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
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
