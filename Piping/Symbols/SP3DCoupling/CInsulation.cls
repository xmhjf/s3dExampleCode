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
'   Author:          NN
'   Creation Date:  Friday, Nov 24 2000
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy      who                   change description
'   -----------      -----                 ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   20.Nov.2003     SymbolTeam(India)       TR-52620--Introduced check for output 'Insulated Body' for correct geometry for Butt-welded End Preparations.
'   01.Sep.2006      KKC                    CR-100189  Added Part Data Basis property to specify coupling either in face-to-face
'                                                       or face-to-center or face1-to-center and face2 to center.
'   14-SEP-06        KKC                    DI-95670  Replace names with initials in all revision history sheets and symbols")
'   21-Dec-06       svsmylav                TR-112718: FacetoCenter and Face1toCenter are optional inputs added in V7 -
'                                           these are moved to the end of first two V6.1 existing inputs (otherwise mirror operation failed).
'   06-Nov-07        VRG                    CR-127647 Enhanced the symbol to support Flexible Coupling using PartDataBasis value.
'                                           (Source: http://www.mdcvacuum.com/urd/uniface.urd/ecf0070w.display?1.6.4.1)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:"    'Used for error messages

Private Sub Class_Initialize()


End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim pipeDiam As Double
    Dim flangeThick As Double
    Dim cptOffset As Double
    Dim flangeDiam As Double
    Dim depth As Double
    Dim iOutput As Integer
    Dim ObjInsulatedBody As Object

    Dim parInsulationThickness As Double
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim lPartDataBasis As Long

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInsulationThickness = arrayOfInputs(3)

    iOutput = 0

    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis

    If (lPartDataBasis = 5) Or (lPartDataBasis <= 1) Then   'Applicable for 5 "Face-to-face dimension basis"
        Dim parFacetoFace As Double
        parFacetoFace = arrayOfInputs(2)
        parFacetoCenter = parFacetoFace / 2     'parFace1toCenter and parFace2toCenter: These dimensions are calculated and used in computation.
        parFace1toCenter = parFacetoFace / 2
    ElseIf (lPartDataBasis = 420) Then          'Partdatabasis value 420 is used by flexible coupling for Kwik flange/Large flange
        parFacetoFace = arrayOfInputs(2)
        parFacetoCenter = parFacetoFace / 2
        parFace1toCenter = parFacetoFace / 2
    ElseIf (lPartDataBasis = 10) Then           'Applicable for  1 "Default" & 10 -"Face-to-center dimension basis"
        parFacetoCenter = arrayOfInputs(4)
        parFace1toCenter = arrayOfInputs(5)
        If CmpDblEqual(parFace1toCenter, 0) Then
            parFace1toCenter = parFacetoCenter
        End If
    End If
    'Insert your code for output 2(Insulated Body)
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

    stPoint.Set -parFacetoCenter, 0, 0
    enPoint.Set parFace1toCenter, 0, 0
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    Dim parFlangeDiameter As Double
    parFlangeDiameter = flangeDiam
    If CmpDblLessThanOrEqualTo(flangeDiam, 0) Then parFlangeDiameter = pipeDiam

    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    Dim parFlangeDiameter2 As Double
    parFlangeDiameter2 = flangeDiam
    If CmpDblLessThanOrEqualTo(flangeDiam, 0) Then parFlangeDiameter2 = pipeDiam

    Dim parInsulationDiameter As Double
    If CmpDblGreaterthanOrEqualTo(parFlangeDiameter, parFlangeDiameter2) Then
        parInsulationDiameter = parFlangeDiameter + 2 * parInsulationThickness
    Else
        parInsulationDiameter = parFlangeDiameter2 + 2 * parInsulationThickness
    End If

    If (lPartDataBasis = 420) Then
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
        If CmpDblLessThanOrEqualTo(parInsulationDiameter, 1.2 * pipeDiam) Then parInsulationDiameter = 1.2 * pipeDiam
    End If
    Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
