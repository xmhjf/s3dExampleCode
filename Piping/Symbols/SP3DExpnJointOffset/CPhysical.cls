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
'   Copyright (c) 2003-06, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          BG
'   Creation Date:  Thursday, Jan 3 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)      Copyright Information, Header  is added.
'   16.Feb.2006     svsmylav       RI -28367: Updated the symbol to use Part Data Basis property
'                                          and added new input parameter 'FacetoFace' for Model 'AUD',Metal%20Catalog.pdf,
'                                          of Flexonics, URL http://www.flexonics.ca/products.htm.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

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
    Dim ObjNozzle1withExtn As Object
    Dim ObjJunction1 As Object
    Dim ObjSkewPipe As Object
    Dim ObjJunction2 As Object
    Dim objNozzle As Object
    Dim ObjNozzle2withExtn As Object
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parOffset As Double
    Dim parInsulationThickness As Double
    Dim parFacetoFace As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFace1toCenter = arrayOfInputs(2)
'    parFace2toCenter = arrayOfInputs(3)
    parOffset = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
'    parFacetoFace = arrayOfInputs(6)

    iOutput = 0
    
 'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis

    If (lPartDataBasis <> 5) Then
        'Consider Face-to-center dimension basis
        parFace1toCenter = arrayOfInputs(2)
        parFace2toCenter = arrayOfInputs(3)
    
    ElseIf (lPartDataBasis = 5) Then
        'Face-to-face dimension basis
        parFacetoFace = arrayOfInputs(6)
        'Consider parFace1toCenter and parFace2toCenter to be half of parFacetoFace
        parFace1toCenter = parFacetoFace / 2
        parFace2toCenter = parFacetoFace / 2
    End If

' Place Nozzle 1 (Nozzle with Extension)
    Dim Port1Length As Double
    Dim oPlacePoint As New AutoMath.DPosition
    Dim oDir As New AutoMath.DVector
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Port1Length = parFace1toCenter - parFace1toCenter / 3 - parFace1toCenter / 3
    oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set ObjNozzle1withExtn = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port1Length)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle1withExtn
    Set ObjNozzle1withExtn = Nothing
 
 ' Insert your code for output 2(Junction1 (Cylinder))
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    stPoint.Set -parFace1toCenter / 3 - parFace1toCenter / 3, 0, 0
    enPoint.Set -parFace1toCenter / 3, 0, 0
    Set ObjJunction1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam * 1.5, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjJunction1
    Set ObjJunction1 = Nothing
    
 ' Insert your code for output 3(Skewed Pipe)
    Dim Angle As Double
    Dim Hypotenous As Double
    Dim skewpipeStPt As New AutoMath.DPosition
    Dim skewpipeEdPt As New AutoMath.DPosition
    Angle = Atn(parOffset / (parFace1toCenter / 3))
    Hypotenous = 0.5 * pipeDiam * Tan(Angle)
    skewpipeStPt.x = -(Hypotenous * Cos(Angle) + parFace1toCenter / 3)
    skewpipeStPt.y = -Hypotenous * Sin(Angle)
    skewpipeStPt.z = 0
    skewpipeEdPt.x = (Hypotenous * Cos(Angle) + parFace2toCenter / 3)
    skewpipeEdPt.y = (parOffset + (parFace2toCenter / parFace1toCenter) * parOffset) + _
                        Hypotenous * Sin(Angle)
    skewpipeEdPt.z = 0
    Set ObjSkewPipe = PlaceCylinder(m_OutputColl, skewpipeStPt, skewpipeEdPt, pipeDiam, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSkewPipe
    Set ObjSkewPipe = Nothing
    
 ' Insert your code for output 4(Junction2 (Cylinder))
    Dim JunpipeStPt As New AutoMath.DPosition
    Dim JunpipeEdPt As New AutoMath.DPosition
    Dim PipeDia As Double
    RetrieveParameters 2, oPartFclt, m_OutputColl, PipeDia, flangeThick, flangeDiam, sptOffset, depth
    JunpipeStPt.Set parFace2toCenter / 3, parOffset + (parFace2toCenter / parFace1toCenter) * parOffset, 0
    JunpipeEdPt.Set (parFace2toCenter / 3 + parFace2toCenter / 3), parOffset + (parFace2toCenter / parFace1toCenter) * parOffset, 0
    Set ObjJunction2 = PlaceCylinder(m_OutputColl, JunpipeStPt, JunpipeEdPt, PipeDia * 1.5, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjJunction2
    Set ObjJunction2 = Nothing
    
 ' Insert your code for output 5(Nozzle with Extension) Place Nozzle 2
    Dim Port2Length As Double
    Dim oPlacePointA As New AutoMath.DPosition
    Dim oDirA As New AutoMath.DVector
    oPlacePointA.Set parFace2toCenter + sptOffset - depth, parOffset + (parFace2toCenter / parFace1toCenter) * parOffset, 0
    oDirA.Set 1, 0, 0
    Port2Length = parFace2toCenter - parFace2toCenter / 3 - parFace2toCenter / 3
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDirA, oPlacePointA, Port2Length)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
