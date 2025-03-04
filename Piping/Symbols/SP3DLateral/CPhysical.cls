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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          NN
'   Creation Date:  Friday, Nov 24 2000
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   07.Jan.2008     RRK                     CR-131505 Enhanced the symbol to support part data basis values of 608, 609 and 610
'   07.JUL.2008     MA                      CR-145695 Implemented part data basis for the value 13
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI As Double
Private Const MODULE = "Physical:"    'Used for error messages

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim pipeDiam As Double
    Dim flangeThick As Double
    Dim cptoffset As Double
    Dim flangeDiam As Double
    Dim depth As Double

    Dim flangeThick2 As Double
    Dim cptoffset2 As Double
    Dim depth2 As Double

    Dim pipeDiam3 As Double
    Dim cptoffset3 As Double
    Dim depth3 As Double

    Dim iOutput As Double

    Dim parFacetoFace As Double
    Dim parFace2toBranch As Double
    Dim parInsulationThickness As Double
    Dim parSeattoSeat As Double
    Dim parSeat1toCenter As Double
    Dim parSeat3toCenter As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parAngle As Double

    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    Dim dFace3toCenter As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parFacetoFace = arrayOfInputs(2)
    'parFace2toBranch = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    'parSeattoSeat = arrayOfInputs(5)
    'parSeat1toCenter = arrayOfInputs(6)
    'parSeat3toCenter = arrayOfInputs(7)
    'parFace1toCenter = arrayOfInputs(8)
    'parFace2toCenter = arrayOfInputs(9)
    'parFace3toCenter = arrayOfInputs(10)
    'parAngle = arrayOfInputs(11)


    parAngle = PI / 4

    'Check to see that old instances of the symbol do not fail
    If UBound(arrayOfInputs) >= 11 Then parAngle = arrayOfInputs(11)

    iOutput = 0

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick2, flangeDiam, cptoffset2, depth2
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick, flangeDiam, cptoffset3, depth3

    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing

    Select Case lPartDataBasis
    Case Is <= 1    'Default
        parFacetoFace = arrayOfInputs(2)
        parFace2toBranch = arrayOfInputs(3)

        dFace1toCenter = parFacetoFace - parFace2toBranch
        dFace2toCenter = parFace2toBranch
        dFace3toCenter = parFace2toBranch

    Case 13    'Asymmetrical Face-to-Center dimension basis
               'Face to center dimension for port 3 is same as that for port 2
        parFace1toCenter = arrayOfInputs(8)
        parFace2toCenter = arrayOfInputs(9)
        
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
        dFace3toCenter = parFace2toCenter
    
    Case 608   'Y branches, single, specified by header face-to-face dimension, face-to-center 1, and face-to-center 3
        parFacetoFace = arrayOfInputs(2)
        parFace1toCenter = arrayOfInputs(8)
        parFace3toCenter = arrayOfInputs(10)

        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFacetoFace - dFace1toCenter
        dFace3toCenter = parFace3toCenter

    Case 609   'Y branches, single, specified by face-to-center 1, face-to-center 2, and face-to-center 3
        parFace1toCenter = arrayOfInputs(8)
        parFace2toCenter = arrayOfInputs(9)
        parFace3toCenter = arrayOfInputs(10)

        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
        dFace3toCenter = parFace3toCenter

    Case 610  'Y branches, single, specified by header seat-to-seat dimension, seat-to-center 1, and seat-to-center 3
        parSeattoSeat = arrayOfInputs(5)
        parSeat1toCenter = arrayOfInputs(6)
        parSeat3toCenter = arrayOfInputs(7)

        dFace1toCenter = parSeat1toCenter + depth - cptoffset
        dFace2toCenter = parSeattoSeat - parSeat1toCenter + depth2 + cptoffset2
        dFace3toCenter = parSeat3toCenter + depth3 - cptoffset3

    Case Else
        GoTo ErrorLabel    'Invalid specification.
    End Select

    'The Body of the Lateral is part of Port1
    'Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Dim objNozzle As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -dFace1toCenter - cptoffset + depth, 0, 0

    Dim Port1Length As Double
    Port1Length = dFace1toCenter + dFace2toCenter - flangeThick2
    oDir.Set -1, 0, 0

    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port1Length)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    'Place Nozzle 2
    oPlacePoint.Set dFace2toCenter + cptoffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    'Place Nozzle 3
    oPlacePoint.Set (dFace3toCenter + cptoffset3 - depth3) * Cos(parAngle), (dFace3toCenter + cptoffset3 - depth3) * Sin(parAngle), 0
    oDir.Set Cos(parAngle), Sin(parAngle), 0

    Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, oPlacePoint, dFace3toCenter)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    Set oPlacePoint = Nothing
    Set oDir = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
