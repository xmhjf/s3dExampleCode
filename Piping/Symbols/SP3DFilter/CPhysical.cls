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
'   Author:          svsmylav
'   Creation Date:  Wednesday, May 29 2002
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This symbol constructed as per the catalog available at URL http://www.yardneyfilters.com/basket_specs.htm
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     -----               ------------------
'   09.Jul.2003     SymbolTeam(India)   Copyright Information, Header  is added.
'   27.Jan.2006     svsmylav    RI-28367: Deleted reference to 'S3AC or S3AZ' of PDS in the
'                                       revision history and added hyper link to Yardney's site.
'                                       Also, modified symbol input parameters to match that of the
'                                       manufacturer's skecth.
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
    Dim ObjFilterBody  As Object
    
    Dim parFilterBodyHeight1 As Double
    Dim parNozzleOffset As Double
    Dim parOffset As Double
    Dim parFacetoFace As Double
    Dim parInsulationThickness  As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFilterBodyHeight1 = arrayOfInputs(2) 'A
    parNozzleOffset = arrayOfInputs(3)      'B
    parOffset = arrayOfInputs(4)            'C
    parFacetoFace = arrayOfInputs(5)        'D
    parInsulationThickness = arrayOfInputs(6)

    iOutput = 0
    
'   Assumption: Origin of the symbol is assumed to be at half of the parFilterBodyHeight1

' Insert your code for output 4(Body of Filter)
'   Assumption: Diameter of the body of the filter is taken 80% of parFacetoFace
    Dim dFilterBodyDiameter   As Double
    dFilterBodyDiameter = 0.8 * parFacetoFace
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

    stPoint.Set 0, -parFilterBodyHeight1 / 2, 0
    enPoint.Set 0, parFilterBodyHeight1 / 2, 0
    Set ObjFilterBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, dFilterBodyDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFilterBody
    Set ObjFilterBody = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

' Place Nozzle 1
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Dim dNozzleLength As Double

    oPlacePoint.Set -parFacetoFace / 2 - sptOffset + depth, _
                (parOffset + parNozzleOffset) - parFilterBodyHeight1 / 2, 0
    oDir.Set -1, 0, 0
    dNozzleLength = parFacetoFace / 2
    If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, _
                    parOffset - parFilterBodyHeight1 / 2, 0
    oDir.Set 1, 0, 0
    dNozzleLength = parFacetoFace / 2
    If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
