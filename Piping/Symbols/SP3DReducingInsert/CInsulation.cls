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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          svsmylav
'   Creation Date:  Tuesday, Aug 27 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The Details were taken from PDS Piping Component Data - Reference Guide at Page no D-52
'   SN=F43_2 and MC=BUSH,INSR1,INSR2. Symbol is created with Six Outputs - three are Insulation aspect
'   outputs. Others are physical aspect outputs and corresponding function are as follows:
'   Flange - using 'PlaceCylinder',
'   Nozzle1  - using 'CreateNozzleWithLength' and
'   Nozzle2  - using 'CreateNozzle' function.
'   Symbol is created with three Insulation aspect
'   outputs which are prepared by using 'PlaceCylinder' function.
'
'   Change History:
'   dd.mmm.yyyy     who            change description
'   -----------     -----          ------------------
'   11.Dec.2003   svsmylav TR-52642 - The input parameters are changed
'                                  to match with the prepared PDS data. Output name
'                                  ObjInsulatedFlange is changed to ObjInsulatedPort2.
'   22.Aug.2005   svsmylav TR-83739: Added check to ensure non-zero diameter and non-zero length for
'                                           cylinder to avoid PlaceCylinder function failure.
'   05.Oct.2005   svsmylav TR-85849: Insulation envelope is modified to suit corrected port locations.
'   08.SEP.2006   KKC      DI-95670  Replace names with initials in all revision history sheets and symbols
'   16.JUL.2008   PK       CR-145835 Implemented part data basis options 1020,1021 and 1022
'   16.Sep.2008     MP   CR-149885   Implemented part data basis option 1083
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private Const NEGLIGIBLE_THICKNESS = 0.0001

Private Function ReturnMax4(A#, B#, C#, D#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If MaxValue < B Then MaxValue = B
    If MaxValue < C Then MaxValue = C
    If MaxValue < D Then MaxValue = D
    ReturnMax4 = MaxValue
End Function

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

    Dim iOutput             As Double
    Dim ObjInsulatedBody    As Object
    Dim ObjInsulatedPort1   As Object
    Dim ObjInsulatedPort2   As Object

    Dim parFacetoFace           As Double
    Dim parInsulationThickness  As Double
    Dim parLength1              As Double
    Dim parLength2              As Double
    Dim parWidth                As Double
    
    Dim dFacetoFace             As Double
    Dim dFacetoCenter           As Double

    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
    Dim InsulationDia           As Double
    Dim dInsulationDiameter As Double
'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parFacetoFace = arrayOfInputs(2)  'O
    parInsulationThickness = arrayOfInputs(3)
    'parLength1 = arrayOfInputs(4)
    'parLength2 = arrayOfInputs(5)
    'parWidth = arrayOfInputs(6)
    
    Dim lPartdatabasis As Long
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt

    lPartdatabasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

    If lPartdatabasis <= 1 Or lPartdatabasis = 1023 Or lPartdatabasis = 1083 Then
        parFacetoFace = arrayOfInputs(2)
        dFacetoFace = parFacetoFace
        If lPartdatabasis = 1083 Then
            parLength1 = arrayOfInputs(4)
            parWidth = arrayOfInputs(6)
        End If
    ElseIf lPartdatabasis = 1020 Or lPartdatabasis = 1021 Or lPartdatabasis = 1022 Then
        parLength1 = arrayOfInputs(4)
        parLength2 = arrayOfInputs(5)
        parWidth = arrayOfInputs(6)
        dFacetoFace = parLength1 + parLength2
    End If

    iOutput = 0

    Dim pipeDiam2       As Double
    Dim flangeThick2    As Double
    Dim flangeDiam2     As Double
    Dim depth2          As Double
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset, depth2
            
    
    If lPartdatabasis = 1083 Then
        dFacetoCenter = (dFacetoFace) * 0.5
    Else
        dFacetoCenter = (dFacetoFace - depth2) * 0.5
    End If
    If lPartdatabasis <= 1 Or lPartdatabasis = 1023 Or lPartdatabasis = 1083 Then
        'Insert your code for output 1(InsulatedBody)
        If flangeThick = 0 Then flangeThick = NEGLIGIBLE_THICKNESS
        If flangeThick2 = 0 Then flangeThick2 = NEGLIGIBLE_THICKNESS
        
        stPoint.Set -dFacetoCenter + flangeThick, 0, 0
        If lPartdatabasis = 1083 Then
            enPoint.Set dFacetoCenter - (parFacetoFace - parLength1), 0, 0
        Else
            enPoint.Set stPoint.x + dFacetoFace - flangeThick2, 0, 0
        End If
    
        InsulationDia = ReturnMax4(pipeDiam, flangeThick, pipeDiam2, flangeThick2) _
                            + 2 * parInsulationThickness
        Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
        Set ObjInsulatedBody = Nothing
    
        'Insert your code for output 2(Insulated Port1)
        stPoint.Set -dFacetoCenter, 0, 0
        enPoint.Set stPoint.x + flangeThick + parInsulationThickness, 0, 0
        InsulationDia = pipeDiam + parInsulationThickness * 2
        If flangeDiam > pipeDiam Then InsulationDia = flangeDiam + 2 * parInsulationThickness
        Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
        Set ObjInsulatedPort1 = Nothing
    
        'Insert your code for output 3(Insulated port2)
        If lPartdatabasis = 1083 Then
            stPoint.Set dFacetoCenter - flangeThick2 - parInsulationThickness, 0, 0
'            stPoint.Set dFacetoCenter - (parFacetoFace - parLength1), 0, 0
            enPoint.Set dFacetoCenter, 0, 0
        Else
            stPoint.Set dFacetoFace - dFacetoCenter, 0, 0
            enPoint.Set stPoint.x - flangeThick2 - parInsulationThickness, 0, 0
        End If
        InsulationDia = pipeDiam2 + parInsulationThickness * 2
        If flangeDiam2 > pipeDiam2 Then InsulationDia = flangeDiam2 + 2 * parInsulationThickness
        Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
        Set ObjInsulatedPort2 = Nothing
        
        'Insert code for insulated body2
        Dim oInsbody2 As Object
        If lPartdatabasis = 1083 Then
        stPoint.Set dFacetoCenter - (parFacetoFace - parLength1), 0, 0
        enPoint.Set dFacetoCenter - flangeThick2 - parInsulationThickness, 0, 0
'        If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then
'            dInsulationDiameter = flangeDiam2 + 2 * parInsulationThickness
'        Else
'            dInsulationDiameter = pipeDiam2 + 2 * parInsulationThickness
'        End If
        Set oInsbody2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam2, True)
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "InsBody2", oInsbody2
        
        Set oInsbody2 = Nothing
    'Insert your code for Insulated Head
        Dim dThick As Double
        Dim dHeadDia As Double
        Dim oHead As Object
        dThick = 0.2 * parLength1
        dHeadDia = 1.4 * pipeDiam + (2 * parInsulationThickness)
        
        stPoint.Set dFacetoCenter - (parFacetoFace - parLength1) + parInsulationThickness, 0, 0
        enPoint.Set dFacetoCenter - (parFacetoFace - (0.8 * parLength1)) - parInsulationThickness, 0, 0
        Set oHead = PlaceCylinder(m_OutputColl, stPoint, enPoint, dHeadDia, True)
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "InsHead", oHead
        Set oHead = Nothing
        End If

    ElseIf lPartdatabasis = 1020 Or lPartdatabasis = 1021 Or lPartdatabasis = 1022 Then
        
        'Insert your code for Insulated Body
        Dim oInsBody As Object
        stPoint.Set -dFacetoCenter, 0, 0
        enPoint.Set stPoint.x + parLength1, 0, 0
        
        Set oInsBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam + 2 * parInsulationThickness, True)

        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oInsBody
        Set oInsBody = Nothing

        'Insert your code for Insulated Head for the bushing
        Dim oInsHead    As Object
        Dim lNumSides   As Long
        Dim oTransMat   As AutoMath.DT4x4
        Dim oVec        As AutoMath.DVector

        Set oVec = New AutoMath.DVector
        Set oTransMat = New AutoMath.DT4x4

        If lPartdatabasis = 1020 Then
            lNumSides = 6
        ElseIf lPartdatabasis = 1021 Then
            lNumSides = 8
        ElseIf lPartdatabasis = 1022 Then
            lNumSides = 10
        End If
        
        Dim dSideLength As Double
        If CmpDblEqual(parWidth, 0) Then
            'Assuming that side length of the polygon shaped hub is 0.7 times the Piping Outside Diameter of Port1
            dSideLength = 0.7 * pipeDiam
        Else
            dSideLength = 0.5 * parWidth
        End If
        
        'Check to ensure that the insulation thickness for the polygonal head of the reducing
        'bush does not creep beyond the length1 of the insert
        Dim dInsulationThickness As Double
        
        If CmpDblGreaterthan(parInsulationThickness, parLength1) Then
            dInsulationThickness = parLength1
        Else
            dInsulationThickness = parInsulationThickness
        End If
        
        Set oInsHead = PlaceNnagon(m_OutputColl, lNumSides, dSideLength + parInsulationThickness, parLength2 + dInsulationThickness, True)
        
        oVec.Set parLength1 - dFacetoCenter - dInsulationThickness, 0, 0
        oTransMat.LoadIdentity
        oTransMat.Translate oVec
        oInsHead.Transform oTransMat
                
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oInsHead
        Set oInsHead = Nothing
        Set oTransMat = Nothing
        Set oVec = Nothing
        
        'Insert your code for Insulated Port 1
        Dim oInsPort1           As Object
        stPoint.Set -dFacetoCenter, 0, 0
              
        'Check to ensure that flange thickness is a non-zero value so that symbol does not fail
        'for male ends
        If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
        End If
        
        'Check to ensure that the insulation for Port1 does not cross length1 of the body
        If CmpDblGreaterthan(parInsulationThickness + flangeThick, parLength1) Then
            enPoint.Set -dFacetoCenter + parLength1, 0, 0
        Else
            enPoint.Set -dFacetoCenter + parInsulationThickness + flangeThick, 0, 0
        End If
                        
        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            dInsulationDiameter = flangeDiam + 2 * parInsulationThickness
        Else
            dInsulationDiameter = pipeDiam + 2 * parInsulationThickness
        End If
        
        Set oInsPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)
        
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oInsPort1
        Set oInsPort1 = Nothing
            
    End If
    
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
