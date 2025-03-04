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
'   Copyright (c) 2004-07, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          MS
'   Creation Date:  Thursday, June 2004
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This valve consists of 7 outputs. This symbol has 2 Inputs FacetoFace and InsulationThickness
'   Insulation Diameter is greater of Flange and Pipe Dia of both Pipes mating with the Valve
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   03.Dec.2004     svsmylav        Defined new input parameter HandwheelAngle.
'   02.Jan.2006     KKK       CR-91106  Updated the symbol to be based on either face-to-face
'                                           or face-to-center(Implementation of PartDataBasis Property)
'   31.Mar.2006     SymbolTeam(India)       Modified to accept default value of lPartDataBasis to be less than or equal to 1.
'   25.sep.06       KKC            TR-102869 Modified symbol code to use Face to Center or Face to Face to Dimension.
'                                            Modified symbol code to use Face to Face as a input parameter.
'   22.May.2007     MA             CR-113431 Implemented Part data basis for values 5,10,13,25,35,349,351 and 353.
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
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double

    Dim iOutput     As Double
    Dim ObjValveInsulation As Object
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double
    Dim dInsulationDia As Double
    Dim parHandwheelAngle   As Double
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parSeat1toCenter As Double
    Dim parSeat2toCenter As Double
    Dim parSeattoSeat As Double
    Dim parNippleLength As Double
    Dim parValveFacetoFace As Double
    Dim parValveFacetoCenter As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    parHandwheelAngle = arrayOfInputs(4)
'    parFacetoCenter = arrayOfInputs(5)
'    parFace1toCenter = arrayOfInputs(6)
'    parFace2toCenter = arrayOfInputs(7)
'    parSeat1toCenter = arrayOfInputs(8)
'    parSeat2toCenter = arrayOfInputs(9)
'    parSeattoSeat = arrayOfInputs(10)
'    parNippleLength = arrayOfInputs(11)
'    parValveFacetoFace = arrayOfInputs(12)
'    parValveFacetoCenter = arrayOfInputs(13)

    iOutput = 0
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                            sptOffset, depth
                                                                            
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, _
                                                                sptOffset2, depth2

    Dim dRadiusofCone As Double
    dRadiusofCone = pipeDiam / 2
        
 'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    Dim dNozzleLength As Double 'Used in case where nipples are created
    dNozzleLength = 0
    
    Select Case lPartDataBasis
    
    Case Is <= 1, 5  'FacetoFace
        parFacetoFace = arrayOfInputs(2)
        dFace1toCenter = parFacetoFace / 2
        dFace2toCenter = parFacetoFace / 2
    
    Case 10    'FacetoCenter
        parFacetoCenter = arrayOfInputs(5)
        dFace1toCenter = parFacetoCenter
        dFace2toCenter = parFacetoCenter
    
    Case 13   'Face1toCenter and Face2toCenter
        parFace1toCenter = arrayOfInputs(6)
        parFace2toCenter = arrayOfInputs(7)
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
    
    Case 25   'SeattoSeat
        parSeattoSeat = arrayOfInputs(10)
        dFace1toCenter = (parSeattoSeat / 2) + depth - sptOffset
        'Assumption: If depth is not provided, length of the half of the valve body
        'is 1.2 times half of the seat-to-seat dimension.
        If CmpDblLessThanOrEqualTo(depth, 0) Then _
        dFace1toCenter = (parSeattoSeat / 2) * 1.2
        
        dFace2toCenter = (parSeattoSeat / 2) + depth2 - sptOffset2
        If CmpDblLessThanOrEqualTo(depth2, 0) Then _
        dFace2toCenter = (parSeattoSeat / 2) * 1.2
        
        If CmpDblGreaterthan(flangeDiam, flangeDiam2) Then
            dRadiusofCone = flangeDiam / 2
        Else
            dRadiusofCone = flangeDiam2 / 2
        End If
        If CmpDblLessThanOrEqualTo(dRadiusofCone, 0) Then dRadiusofCone = pipeDiam / 2
    
    Case 35  'Seat1toCenter and Seat2toCenter
        parSeat1toCenter = arrayOfInputs(8)
        parSeat2toCenter = arrayOfInputs(9)
        dFace1toCenter = parSeat1toCenter + depth - sptOffset
        'Assumption: If depth is not provided, length of the half of the valve body
        'is 1.2 times half of the seat-to-seat dimension.
        If CmpDblLessThanOrEqualTo(depth, 0) Then _
        dFace1toCenter = parSeat1toCenter * 1.2
        
        dFace2toCenter = parSeat2toCenter + depth2 - sptOffset2
        If CmpDblLessThanOrEqualTo(depth2, 0) Then _
        dFace2toCenter = parSeat2toCenter * 1.2
        
        If CmpDblGreaterthan(flangeDiam, flangeDiam2) Then
            dRadiusofCone = flangeDiam / 2
        Else
            dRadiusofCone = flangeDiam2 / 2
        End If
        If CmpDblLessThanOrEqualTo(dRadiusofCone, 0) Then dRadiusofCone = pipeDiam / 2
              
    Case 349  'NippleLength and ValveFacetoFace
        parNippleLength = arrayOfInputs(11)
        parValveFacetoFace = arrayOfInputs(12)
        dFace1toCenter = parValveFacetoFace / 2
        dFace2toCenter = parValveFacetoFace / 2
        dNozzleLength = parNippleLength
    
    Case 351  'FacetoFace and ValveFacetoFace
        parFacetoFace = arrayOfInputs(2)
        parValveFacetoFace = arrayOfInputs(12)
        dFace1toCenter = parValveFacetoFace / 2
        dFace2toCenter = parValveFacetoFace / 2
        dNozzleLength = (parFacetoFace - parValveFacetoFace) / 2
    
    Case 353  'NippleLength and ValveFacetoCenter
        parNippleLength = arrayOfInputs(11)
        parValveFacetoCenter = arrayOfInputs(13)
        dFace1toCenter = parValveFacetoCenter
        dFace2toCenter = parValveFacetoCenter
        dNozzleLength = parNippleLength
    
    Case Else
        GoTo ErrorLabel:
        
    End Select
 
 ' Insert your code for output 6(ValveInsulation)
    Dim oStartInsulationPos       As AutoMath.DPosition
    Dim oEndInsulationPos         As AutoMath.DPosition
    Set oStartInsulationPos = New AutoMath.DPosition
    Set oEndInsulationPos = New AutoMath.DPosition
    Dim dInsulationDia1 As Double
    Dim dInsulationDia2 As Double

    oStartInsulationPos.Set -dFace1toCenter - dNozzleLength, 0, 0
    oEndInsulationPos.Set dFace2toCenter + dNozzleLength, 0, 0
    
    'In cases 25 and 35, the radius of the cone is greater than the pipe radius.
    'This radius is considered in the computation of insulation diameter below.
    If CmpDblGreaterthan(flangeDiam, dRadiusofCone * 2) Then
        dInsulationDia1 = flangeDiam + parInsulationThickness * 2
    Else
        dInsulationDia1 = pipeDiam + parInsulationThickness * 2
    End If
       
    If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then
        dInsulationDia2 = flangeDiam2 + parInsulationThickness * 2
    Else
        dInsulationDia2 = pipeDiam2 + parInsulationThickness * 2
    End If
    
    If CmpDblGreaterthan(dInsulationDia1, dInsulationDia2) Then
        dInsulationDia = dInsulationDia1
    Else
        dInsulationDia = dInsulationDia2
    End If
    Set ObjValveInsulation = PlaceCylinder(m_OutputColl, oStartInsulationPos, oEndInsulationPos, _
                                    dInsulationDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveInsulation
    Set ObjValveInsulation = Nothing
    Set oStartInsulationPos = Nothing
    Set oEndInsulationPos = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub


