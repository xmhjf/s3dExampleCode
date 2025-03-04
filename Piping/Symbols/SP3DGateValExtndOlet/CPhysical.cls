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
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         BG
'   Creation Date:  Monday, Jun 17 2002
'   Description:
'   The Symbol details were taken from PDS PDS Piping Component Data Reference Manual
'   at Page No D - 8 and SN=V2A. The Symbol consist of Physical and Insulation aspects
'   Physical aspect is made up of three cones and two Nozzles.Insulation aspect consist of
'   Simple Cylinder between flange of 2nd Nozzle and to the point where body extension ends
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'   03.Dec.2004     svsmylav        Defined new input parameter HandwheelAngle and modified valve operator direction vectors.
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'   09.Sep.2005     svsmylav        TR-83135: Modified the Y-coordinates of points in the
'                                           symbol by adding 'parValveCentertoHeaderCenter'.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CPhysical:" 'Used for error messages

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
    Dim ObjUpperCone As Object
    Dim ObjLowerCone As Object
    Dim ObjBodyExtension As Object
    
    Dim parValveCentertoHeaderCenter As Double
    Dim parFacetoCenter As Double
    Dim parOletLength As Double
    Dim parInsulationThickness As Double
    Dim sptOffset2 As Double
    Dim depth2 As Double
    Dim parHandwheelAngle   As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parValveCentertoHeaderCenter = arrayOfInputs(2)
    parFacetoCenter = arrayOfInputs(3)
    parOletLength = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    parHandwheelAngle = arrayOfInputs(6)
    
    iOutput = 0
    
 ' Insert your code for output 1(UpperCone)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    'Store the Nozzle 2 values in a variable
    sptOffset2 = sptOffset
    depth2 = depth
    Dim basecenter As AutoMath.DPosition
    Dim topcenter As AutoMath.DPosition
    Set basecenter = New AutoMath.DPosition
    Set topcenter = New AutoMath.DPosition
    
    basecenter.Set 0, parValveCentertoHeaderCenter + parFacetoCenter - flangeThick, 0
    topcenter.Set 0, parValveCentertoHeaderCenter, 0
    
    Set ObjUpperCone = PlaceCone(m_OutputColl, basecenter, topcenter, pipeDiam / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjUpperCone
    Set ObjUpperCone = Nothing
    
 ' Insert your code for output 2(LowerCone)
    basecenter.Set 0, parValveCentertoHeaderCenter - parFacetoCenter, 0
    topcenter.Set 0, parValveCentertoHeaderCenter, 0
    Set ObjLowerCone = PlaceCone(m_OutputColl, basecenter, topcenter, pipeDiam / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLowerCone
    Set ObjLowerCone = Nothing
    
 ' Insert your code for output 3(BodyExtension)
    basecenter.Set 0, parValveCentertoHeaderCenter - parFacetoCenter - parOletLength, 0
    topcenter.Set 0, parValveCentertoHeaderCenter - parFacetoCenter, 0
    Set ObjBodyExtension = PlaceCone(m_OutputColl, basecenter, topcenter, 0.6 * pipeDiam / 2, pipeDiam / 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyExtension
    Set ObjBodyExtension = Nothing
    Set basecenter = Nothing
    Set topcenter = Nothing
    
' Place Nozzle 1

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set 0, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set 0, parValveCentertoHeaderCenter + parFacetoCenter + sptOffset2 - depth2, 0
    oDir.Set 0, 1, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

' Insert your code for output 7 (Valve Operator)
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

    oDirX.Set 0, Cos(parHandwheelAngle), Sin(parHandwheelAngle)
    oDirY.Set 1, 0, 0
    oDirZ.Set 0, Sin(parHandwheelAngle), -Cos(parHandwheelAngle)

    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    On Error GoTo ErrorLabel
    Dim oOperatorPart As IJDPart
    Dim oOperatorOcc   As IJPartOcc
    If Not oPipeComponent Is Nothing Then
        Set oOperatorPart = oPipeComponent.GetValveOperatorPart
        If Not oOperatorPart Is Nothing Then
            Dim OpOrigin As IJDPosition
            Set OpOrigin = New DPosition
            OpOrigin.Set 0, parValveCentertoHeaderCenter, 0
            Set oOperatorOcc = oSymbolHelper.CreateChildPartOcc("ValveOperator", oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)

        End If
    End If
    Set oSymbolHelper = Nothing
    Set oOperatorPart = Nothing
    Set oPipeComponent = Nothing
    Set oOperatorOcc = Nothing
    Set oDirX = Nothing
    Set oDirY = Nothing
    Set oDirZ = Nothing
    Set OpOrigin = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
