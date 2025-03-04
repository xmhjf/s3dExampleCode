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
'   07.May.2005     MS          TR 77169 Modified Port1 & 2 locations
'  08.SEP.2006     KKC          DI-95670  Replace names with initials in all revision history sheets and symbols
'  22.OCT.2007     svsmylav     CR-128137 Enhance half coupling symbol to support face-to-face or end-to-end
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
    Dim pipeDiam2        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
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
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam, sptOffset2, depth2
    Dim NozzleLength As Double

    Dim dMeetPoint As Double
    dMeetPoint = Sqr((pipeDiam / 2) ^ 2 - (pipeDiam2 / 2) ^ 2)
    
    Select Case lPartDataBasis
        Case Is <= 1, 451 'The default state; Half coupling, specified by end-to-end dimension - 451
            dFacetoHeaderCenter = parFacetoHeaderCenter
            parFacetoEnd = arrayOfInputs(3)
            NozzleLength = parFacetoEnd
            If UBound(arrayOfInputs) >= 5 And CmpDblLessThanOrEqualTo(parFacetoEnd, 0) Then
                parEndtoEnd = arrayOfInputs(5)
                NozzleLength = parEndtoEnd + depth2
                If CmpDblGreaterthan(parFacetoHeaderCenter, LINEAR_TOLERANCE) Then
                    dFacetoHeaderCenter = parFacetoHeaderCenter
                Else
                    dFacetoHeaderCenter = parEndtoEnd + depth2 + dMeetPoint _
                        - RetrieveWallThickness(1, oPartFclt)
                End If
            End If
        Case Is = 453 'Half coupling, specified by face-to-face dimension
            parFacetoFace = arrayOfInputs(6)
            NozzleLength = parFacetoFace
            If CmpDblGreaterthan(parFacetoHeaderCenter, LINEAR_TOLERANCE) Then
                dFacetoHeaderCenter = parFacetoHeaderCenter
            Else
                dFacetoHeaderCenter = parFacetoFace + dMeetPoint _
                    - RetrieveWallThickness(1, oPartFclt)
            End If
        Case Else
            ReportUnanticipatedError MODULE, METHOD
    End Select
    
    If CmpDblLessThan(NozzleLength, flangeThick2) Then
        NozzleLength = flangeThick2
    End If
        
' Place Nozzle 1
' The Half Coupling Body is made using createNozzleWithLength

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
    oPlacePoint.Set 0, dFacetoHeaderCenter + sptOffset2 - depth2, 0
    oDir.Set 0, 1, 0
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)

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
