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
'   Author:         RRK
'   Creation Date:  Monday 29, Jan 2007
'   Description:
'   This implements the physical aspect of  Y Strainer symbol which supports multiple dimensional bases
'   as per CR-113397
'   The symbol supports the following dimensional bases
'   1.)Part Data Basis value 95:-'Y Strainer,specified by Face-to-Face Dimension with Offset to Blow Down Connection
'           Face'
'   2.)Part Data Basis value 100:-'Y Strainer,specified by Face-to-Center 1,  Face-to-Center 2, and Face-to-Center 3'
'   3.)Part Data Basis value 105:-Y Strainer,specified by Face-to-Face Dimension with Offset to Blow Down Connection Port

'   Change History:

'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   27.Aug.2007     VRG      TR-124959  Insulation for ports was ensured for different end preparations
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Const NEGLIGIBLE_THICKNESS = 0.0001
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
    
    Dim PI As Double
    PI = 4 * Atn(1)

    Dim iOutput     As Integer
    Dim ObjBranchBody As Object
    Dim ObjStrainerCover As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parFacetoFace As Double
    Dim parOffset As Double
    Dim parPortVOffset As Double
    Dim parInsulationThickness As Double
    
    Dim dBranchDia As Double
    Dim dBranchCenterPosX As Double
    Dim dBranchCenterPosY As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFace1toCenter = arrayOfInputs(2)
'    parFace2toCenter = arrayOfInputs(3)
'    parFace3toCenter = arrayOfInputs(4)
'    parFacetoFace = arrayOfInputs(5)
'    parOffset = arrayOfInputs(6)
'    parPortVOffset = arrayOfInputs(7)
    parInsulationThickness = arrayOfInputs(8)
    
    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    

    iOutput = 0
    
    If lPartDataBasis = 100 Then
        parFace1toCenter = arrayOfInputs(2)
        parFace2toCenter = arrayOfInputs(3)
        parFace3toCenter = arrayOfInputs(4)
    ElseIf lPartDataBasis <= 1 Or lPartDataBasis = 95 Then
        parFacetoFace = arrayOfInputs(5)
        parOffset = arrayOfInputs(6)
    ElseIf lPartDataBasis = 105 Then
        parFacetoFace = arrayOfInputs(5)
        parPortVOffset = arrayOfInputs(7)
    Else: GoTo ErrorLabel
    End If
    
'Assumption for starting point of the branch

    If (lPartDataBasis <= 1 Or lPartDataBasis = 95 Or lPartDataBasis = 105) Then
        parFace1toCenter = 0.25 * parFacetoFace
        parFace2toCenter = parFacetoFace - parFace1toCenter
    End If
    
    Dim oStPoint   As Object
    Dim oEnPoint   As Object
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition

'   Place Output 1(Branch)

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
    End If
    If CmpDblEqual(flangeDiam, 0) Then
            flangeDiam = pipeDiam
    End If
    If (lPartDataBasis <= 1 Or lPartDataBasis = 95) Then
        dBranchCenterPosY = parOffset - (flangeDiam / 2) * Sin(PI / 4)
        dBranchCenterPosX = dBranchCenterPosY / Tan(PI / 4)
                     
    ElseIf (lPartDataBasis = 105) Then

        dBranchCenterPosX = parPortVOffset * Tan(PI / 4)
        dBranchCenterPosY = parPortVOffset
                     
    ElseIf (lPartDataBasis = 100) Then
        dBranchCenterPosX = parFace3toCenter * Cos(PI / 4)
        dBranchCenterPosY = parFace3toCenter * Sin(PI / 4)
        
    End If

    oStPoint.Set 0, 0, 0
    oEnPoint.Set dBranchCenterPosX - (2 * flangeThick * Sin(PI / 4)), _
                     -dBranchCenterPosY + (2 * flangeThick * Cos(PI / 4)), 0
    
'Assumption for branch diameter
    dBranchDia = pipeDiam
    Set ObjBranchBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dBranchDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBranchBody
    Set ObjBranchBody = Nothing

'   Place Output 2(Cover)
    
    oStPoint.Set dBranchCenterPosX - (2 * flangeThick * Sin(PI / 4)), _
                     -dBranchCenterPosY + (2 * flangeThick * Cos(PI / 4)), 0
    oEnPoint.Set dBranchCenterPosX, -dBranchCenterPosY, 0
    
    Dim dCoverDiam As Double
    dCoverDiam = flangeDiam
    
' Check for cover not intersecting Pipe and Flange
    Dim dFlangeEndX As Double
    Dim dFlangeEndYin As Double
    Dim dFlangeEndYout As Double

    dFlangeEndX = dBranchCenterPosX + (dCoverDiam / 2) * Sin(PI / 4)
    dFlangeEndYin = dBranchCenterPosY - _
                    (2 * flangeThick * Sin(PI / 4)) - (dCoverDiam / 2) * Cos(PI / 4)
    dFlangeEndYout = dBranchCenterPosY - (dCoverDiam / 2) * Cos(PI / 4)
    
    If CmpDblGreaterthan(dFlangeEndX, parFace2toCenter - flangeThick) _
                        And CmpDblLessThan(dFlangeEndYout, flangeDiam / 2) Then
        dCoverDiam = 2 * (parFace2toCenter - flangeThick - dBranchCenterPosX) / Sin(PI / 4)
        
        dFlangeEndX = dBranchCenterPosX + (dCoverDiam / 2) * Sin(PI / 4)
        dFlangeEndYin = dBranchCenterPosY - _
                        (2 * flangeThick * Sin(PI / 4)) - (dCoverDiam / 2) * Cos(PI / 4)
        dFlangeEndYout = dBranchCenterPosY - (dCoverDiam / 2) * Cos(PI / 4)
        If CmpDblLessThan(dFlangeEndYin, pipeDiam / 2) Then
            dCoverDiam = 2 * (dBranchCenterPosY - (2 * flangeThick * Sin(PI / 4)) - (pipeDiam / 2)) / Cos(PI / 4)
        End If
    End If
    
    If CmpDblLessThan(dFlangeEndYin, pipeDiam / 2) Then
        dCoverDiam = 2 * (dBranchCenterPosY - (2 * flangeThick * Sin(PI / 4)) - (pipeDiam / 2)) / Cos(PI / 4)
        
        dFlangeEndX = dBranchCenterPosX + (dCoverDiam / 2) * Sin(PI / 4)
        dFlangeEndYin = dBranchCenterPosY - _
                        (2 * flangeThick * Sin(PI / 4)) - (dCoverDiam / 2) * Cos(PI / 4)
        dFlangeEndYout = dBranchCenterPosY - (dCoverDiam / 2) * Cos(PI / 4)
        If CmpDblGreaterthan(dFlangeEndX, parFace2toCenter - flangeThick) _
                        And CmpDblLessThan(dFlangeEndYout, flangeDiam / 2) Then
            dCoverDiam = 2 * (parFace2toCenter - flangeThick - dBranchCenterPosX) / Sin(PI / 4)
        End If
    End If

    Set ObjStrainerCover = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCoverDiam, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStrainerCover
    Set ObjStrainerCover = Nothing
    
' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Dim dNozzleLength As Double
    
    oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    dNozzleLength = parFace1toCenter
        
    If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
    
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    oPlacePoint.Set parFace2toCenter + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    dNozzleLength = parFace2toCenter
    If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
    
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 3
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
      
        oPlacePoint.Set dBranchCenterPosX + (sptOffset - depth) * Sin(PI / 4), _
                        -dBranchCenterPosY - (sptOffset - depth) * Cos(PI / 4), 0
        oDir.Set Cos(PI / 4), Cos(3 * PI / 4), 0
    
    Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

