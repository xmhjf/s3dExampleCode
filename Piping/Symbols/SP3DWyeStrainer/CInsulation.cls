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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'   CInsulation.cls
'   Author:         RRK
'   Creation Date:  Monday 29, Jan 2007
'   Description:
'   This implements the Insulation aspect of  Y Strainer symbol which supports multiple dimensional bases
'   as per CR-113397
'   The symbol supports the following dimensional bases
'   1.)Part Data Basis value 95:-'Y Strainer,specified by Face-to-Face Dimension with Offset to Blow Down Connection
'                                  Face'
'   2.)Part Data Basis value 100:-'Y Strainer,specified by Face-to-Center 1,  Face-to-Center 2, and Face-to-Center 3'
'   3.)Part Data Basis value 105:-Y Strainer,specified by Face-to-Face Dimension with Offset to Blow Down Connection Port

'   Change History:

'   dd.mmm.yyyy     who      change description
'   -----------     -----    ------------------
'   27.Aug.2007     VRG      TR-124959  Insulation for ports was ensured for different end preparations
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Const NEGLIGIBLE_THICKNESS = 0.0001
Private Const MODULE = "Insulation:" 'Used for error messages

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
    Dim ObjInsBranch As Object
    Dim ObjInsStrainer As Object
    Dim ObjInsStrainerCover As Object
    Dim ObjInsPort1 As Object
    Dim ObjInsPort2 As Object
    
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
    
    Dim dInsulationDia As Double

' Insert your code for output 7(Insulation for Strainer Body)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
    End If
    If CmpDblEqual(flangeDiam, 0) Then
            flangeDiam = pipeDiam
    End If
    oStPoint.Set -parFace1toCenter + flangeThick, 0, 0
    oEnPoint.Set parFace2toCenter - flangeThick, 0, 0

    dInsulationDia = pipeDiam + parInsulationThickness * 2
    Set ObjInsStrainer = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsStrainer
    Set ObjInsStrainer = Nothing

 ' Insert your code for output 8(Insulation for Branch)
 
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

    dBranchDia = pipeDiam

    dInsulationDia = dBranchDia + parInsulationThickness * 2
    Set ObjInsBranch = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBranch
    Set ObjInsBranch = Nothing
    
' Insert your code for output 9(Insulation for Strainer Cover)
    
    oStPoint.Set dBranchCenterPosX - (((2 * flangeThick) + parInsulationThickness) * Sin(PI / 4)), _
                 -dBranchCenterPosY + (((2 * flangeThick) + parInsulationThickness) * Cos(PI / 4)), 0

    oEnPoint.Set dBranchCenterPosX + parInsulationThickness * Sin(PI / 4), _
                -dBranchCenterPosY - parInsulationThickness * Cos(PI / 4), 0
        
    
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
    
    dInsulationDia = dCoverDiam + 2 * parInsulationThickness
    
    If CmpDblLessThanOrEqualTo(dInsulationDia, pipeDiam + (2 * parInsulationThickness)) Then
     dInsulationDia = pipeDiam + 2 * parInsulationThickness
    End If
    
    Set ObjInsStrainerCover = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsStrainerCover
    Set ObjInsStrainerCover = Nothing

 ' Insert your code for output 10(Insulation for Port 1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
    End If
    oStPoint.Set -parFace1toCenter, 0, 0
    oEnPoint.Set -parFace1toCenter + flangeThick + parInsulationThickness, 0, 0
    
    dInsulationDia = flangeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then dInsulationDia = pipeDiam + parInsulationThickness * 2
    
    Set ObjInsPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort1
    Set ObjInsPort1 = Nothing

 ' Insert your code for output 11(Insulation for Port 2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
    End If
    oStPoint.Set parFace2toCenter, 0, 0
    oEnPoint.Set parFace2toCenter - flangeThick - parInsulationThickness, 0, 0
    
    dInsulationDia = flangeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then dInsulationDia = pipeDiam + parInsulationThickness * 2
    
    Set ObjInsPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort2
    Set ObjInsPort2 = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
