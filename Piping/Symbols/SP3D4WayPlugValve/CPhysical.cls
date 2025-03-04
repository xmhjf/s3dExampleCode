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
'   Copyright (c) 2003-04, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          MS
'   Creation Date:  Thursday, Jun 14 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol Model No. is: V19, Page No. D-18 PDS Piping Component Data Reference Guide.
' The Symbol is made of Thirteen output objects. Eleven are physical aspect outputs and Two are insulation aspect outputs.
' The physical outputs are made using 'PlaceCone' and 'CreateNozzle' functions.
'The Insulation outputs are made using 'PlaceCylinder' functions.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'   03.Dec.2004     svsmylav        Defined new input parameter HandwheelAngle and modified valve operator direction vectors.
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was 
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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
    Dim ObjLeftCone As Object
    Dim ObjUpperCone As Object
    Dim ObjLowerCone As Object
    Dim ObjRightCone As Object
    Dim ObjTopCone As Object
    Dim ObjBottomCone As Object
    Dim ObjInsulCylinder1 As Object
    Dim ObjInsulCylinder2 As Object
    
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parFace4toCenter As Double
    Dim parInsulationThickness As Double
    
    Dim parHandwheelAngle   As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)     'A
    parFace2toCenter = arrayOfInputs(3)     'B
    parFace3toCenter = arrayOfInputs(4)     'C
    parFace4toCenter = arrayOfInputs(5)     'C
    parInsulationThickness = arrayOfInputs(6)
    parHandwheelAngle = arrayOfInputs(7)
    
    iOutput = 0

 ' Insert your code for output 1(Left Cone)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim BaseCenPoint As New AutoMath.DPosition
    Dim TopCenPoint As New AutoMath.DPosition
    
    BaseCenPoint.Set -parFace1toCenter + flangeThick, 0#, 0#

'   PlaceCone is giving 'Invalid property value' failure. A very small increment in Y or Z coordinate
'   value removes the 'Invalid property value' failure.
    TopCenPoint.Set 0#, 0#, 0
    
    Set ObjLeftCone = PlaceCone(m_OutputColl, BaseCenPoint, TopCenPoint, pipeDiam / 2, 0, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLeftCone
    Set ObjLeftCone = Nothing
    Set BaseCenPoint = Nothing
    Set TopCenPoint = Nothing
    
 ' Insert your code for output 2(Upper Cone)
    Dim UpperConeBasePt As New AutoMath.DPosition
    Dim UpperConeTopPt As New AutoMath.DPosition
    Dim coneDia As Double
    Dim pipeDiam2 As Double
    Dim coneHt As Double
    
    coneDia = parFace1toCenter * 0.5
    
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam, sptOffset, depth
    If pipeDiam2 > pipeDiam Then
        coneHt = pipeDiam2 * 0.5
    Else
        coneHt = pipeDiam * 0.5
    End If
    
    UpperConeBasePt.Set 0#, 0#, 0#
    UpperConeTopPt.Set 0#, coneHt, 0#
    
    Set ObjUpperCone = PlaceCone(m_OutputColl, UpperConeBasePt, UpperConeTopPt, coneDia / 2, 0, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjUpperCone
    Set ObjUpperCone = Nothing
    Set UpperConeBasePt = Nothing
    Set UpperConeTopPt = Nothing
    
 ' Insert your code for output 3(Lower Cone)
    Dim LowerConeBasePt As New AutoMath.DPosition
    Dim LowerConeTopPt As New AutoMath.DPosition
    
    LowerConeBasePt.Set 0#, 0#, 0#
    LowerConeTopPt.Set 0#, -coneHt, 0#
    
    Set ObjLowerCone = PlaceCone(m_OutputColl, LowerConeBasePt, LowerConeTopPt, coneDia / 2, 0, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLowerCone
    Set ObjLowerCone = Nothing
    Set LowerConeBasePt = Nothing
    Set LowerConeTopPt = Nothing
    
 ' Insert your code for output 4(Right Cone)
    Set oPartFclt = arrayOfInputs(1)
    
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim RightConeBasePt As New AutoMath.DPosition
    Dim RightConeTopPt  As New AutoMath.DPosition
    
    RightConeBasePt.Set parFace2toCenter - flangeThick, 0#, 0#
'   PlaceCone is giving 'Invalid property value' failure. A very small increment in Y or Z coordinate
'   value removes the 'Invalid property value' failure.
    RightConeTopPt.Set 0#, 0#, 0
    
    Set ObjRightCone = PlaceCone(m_OutputColl, RightConeBasePt, RightConeTopPt, pipeDiam / 2, 0, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRightCone
    Set ObjRightCone = Nothing
    Set RightConeBasePt = Nothing
    Set RightConeTopPt = Nothing
    
' Insert your code for output 5(Top Cone)
    Set oPartFclt = arrayOfInputs(1)
    
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim TopConeBasePt As New AutoMath.DPosition
    Dim TopConeTopPt  As New AutoMath.DPosition
    
    TopConeBasePt.Set 0, parFace3toCenter - flangeThick, 0#
'   PlaceCone is giving 'Invalid property value' failure. A very small increment in Y or Z coordinate
'   value removes the 'Invalid property value' failure.
TopConeTopPt.Set 0#, 0#, 0
    
    Set ObjTopCone = PlaceCone(m_OutputColl, TopConeBasePt, TopConeTopPt, pipeDiam / 2, 0, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopCone
    Set ObjTopCone = Nothing
    Set TopConeBasePt = Nothing
    Set TopConeTopPt = Nothing
    
' Insert your code for output 6(Bottom Cone)
    Set oPartFclt = arrayOfInputs(1)
    
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim BottomConeBasePt As New AutoMath.DPosition
    Dim BottomConeTopPt  As New AutoMath.DPosition
    
    BottomConeBasePt.Set 0, -parFace4toCenter + flangeThick, 0#
'   PlaceCone is giving 'Invalid property value' failure. A very small increment in Y or Z coordinate
'   value removes the 'Invalid property value' failure.
    BottomConeTopPt.Set 0#, 0#, 0
    
    Set ObjBottomCone = PlaceCone(m_OutputColl, BottomConeBasePt, BottomConeTopPt, pipeDiam / 2, 0, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBottomCone
    Set ObjBottomCone = Nothing
    Set BottomConeBasePt = Nothing
    Set BottomConeTopPt = Nothing
    
' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    
    Set oPartFclt = arrayOfInputs(1)
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
      
    oPlacePoint.Set parFace2toCenter + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 3
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
     
    oPlacePoint.Set 0, parFace3toCenter + sptOffset - depth, 0
    oDir.Set 0, 1, 0
    
    Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
            
' Place Nozzle 4
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
     
    oPlacePoint.Set 0, -parFace4toCenter - sptOffset + depth, 0
    oDir.Set 0, -1, 0
    
    Set objNozzle = CreateNozzle(4, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
            
' Insert your code for output 13 (Valve Operator)
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
    
    oDirX.Set Cos(parHandwheelAngle), Sin(parHandwheelAngle), 0
    oDirY.Set 0, 0, 1
    oDirZ.Set -Sin(parHandwheelAngle), Cos(parHandwheelAngle), 0

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
            OpOrigin.Set 0, 0, 0
            Set oOperatorOcc = oSymbolHelper.CreateChildPartOcc("ValveOperator", oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)
  
        End If
    End If
    Set oSymbolHelper = Nothing
    Set oOperatorPart = Nothing
    Set oPipeComponent = Nothing
    Set oOperatorOcc = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

