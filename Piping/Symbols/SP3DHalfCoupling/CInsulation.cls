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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         MS
'   Creation Date:  Tuesday, Aug 27 2002
'   Description:
'   Symbol Model No. is: F31 Page No. D-48 of PDS Piping Component Data Reference Guide.
'   Symbol is created with Three Outputs
'   The Two physical aspect outputs are created as follows:
'   One ObjNozzle object by using 'CreateNozzle' function and another ObjNozzle by using CreateNozzleWithLength
'   The One Insulation aspect output ObjHalfCouplingIns is created 'PlaceCylinder'
'
'   Change History:
'   dd.mmm.yyyy     who         change description
'   -----------     ---         ------------------
'
'  08.SEP.2006     KKC          DI-95670  Replace names with initials in all revision history sheets and symbols
'  22.OCT.2007     svsmylav     CR-128137 Enhance half coupling symbol to support face-to-face or end-to-end
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "CInsulation:" 'Used for error messages

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
    Dim flangeDiam2      As Double
    Dim depth2           As Double
        
    Dim iOutput     As Double
    Dim ObjHalfCouplingIns As Object
    
    Dim parFacetoHeaderCenter As Double
    Dim parFacetoEnd As Double
    Dim parInsulationThickness As Double
    Dim parEndtoEnd As Double
    Dim parFacetoFace As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoHeaderCenter = arrayOfInputs(2)
'    parFacetoEnd = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
'    parEndtoEnd = arrayOfInputs(5)
'    parFacetoFace = arrayOfInputs(6)
    
    iOutput = 0
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    Dim dFacetoHeaderCenter As Double
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, sptOffset, depth2

    Dim dMeetPoint As Double
    dMeetPoint = Sqr((pipeDiam / 2) ^ 2 - (pipeDiam2 / 2) ^ 2)
    
    Select Case lPartDataBasis
        Case Is <= 1, 451 'The default state; Half coupling, specified by end-to-end dimension - 451
            dFacetoHeaderCenter = parFacetoHeaderCenter
            parFacetoEnd = arrayOfInputs(3)
            If UBound(arrayOfInputs) >= 5 And CmpDblLessThanOrEqualTo(parFacetoEnd, 0) Then
                parEndtoEnd = arrayOfInputs(5)
                If CmpDblGreaterthan(parFacetoHeaderCenter, LINEAR_TOLERANCE) Then
                    dFacetoHeaderCenter = parFacetoHeaderCenter
                Else
                    dFacetoHeaderCenter = parEndtoEnd + depth2 + dMeetPoint _
                        - RetrieveWallThickness(1, oPartFclt)
                End If
            End If
        Case Is = 453 'Half coupling, specified by face-to-face dimension
            parFacetoFace = arrayOfInputs(6)
            If CmpDblGreaterthan(parFacetoHeaderCenter, LINEAR_TOLERANCE) Then
                dFacetoHeaderCenter = parFacetoHeaderCenter
            Else
                dFacetoHeaderCenter = parFacetoFace + dMeetPoint _
                    - RetrieveWallThickness(1, oPartFclt)
            End If
        Case Else
            ReportUnanticipatedError MODULE, METHOD
    End Select

 ' Insert your code for output 1(Insulation for Half Coupling)
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim InsulationDia As Double

    stPoint.Set 0, dMeetPoint, 0
    enPoint.Set 0, dFacetoHeaderCenter, 0

    If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then
        InsulationDia = flangeDiam2 + parInsulationThickness * 2
    Else
        InsulationDia = pipeDiam2 + parInsulationThickness * 2
    End If
    
    Set ObjHalfCouplingIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHalfCouplingIns
    Set ObjHalfCouplingIns = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
