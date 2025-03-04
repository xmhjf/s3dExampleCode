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
'   Copyright (c) 2005, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         svsmylav
'   Creation Date:  Monday, Nov 21, 2005
'   Description:
'     This class module is the place for user to implement graphical part of VBSymbol for this aspect
'     This is Trimmable Generic Cross symbol. (PDS Reducing run and branches cross (MC=XRRB)
'     (SN=F163)Symbol provides a close example to this symbol).
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
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
    Dim ObjMainNozwithLen As Object
    Dim ObjTaperCone As Object
    Dim ObjBranchNozLen As Object
    Dim ObjBranchNozLen1 As Object
    Dim ObjNozzle As Object
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parFace4toCenter As Double
    Dim parAngle As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2) 'A (of PDS SN=F163)
    parFace2toCenter = arrayOfInputs(3) 'B
    parFace3toCenter = arrayOfInputs(4) 'C
    parFace4toCenter = arrayOfInputs(5) 'D
    parAngle = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)
    
    iOutput = 0
    
 ' Insert your code for output 1 (Main run Nozzle with Length)  Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Dim Port1Length As Double
    Port1Length = parFace1toCenter
    If CmpDblLessThan(Port1Length, flangeThick) Then Port1Length = flangeThick
    oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set ObjMainNozwithLen = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port1Length)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMainNozwithLen
    Set ObjMainNozwithLen = Nothing
    
' Place Nozzle 2  Insert your code for output 2(Nozzle in Main Line)
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2
    oPlacePoint.Set parFace2toCenter + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0
    Set ObjNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle
    Set ObjNozzle = Nothing

' Insert your code for output 2(Main run Reducer)
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition

    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

    stPoint.Set 0, 0, 0
    enPoint.Set parFace2toCenter - flangeThick2, 0, 0
    
    Set ObjTaperCone = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam * 0.5, pipeDiam2 * 0.5, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTaperCone
    Set ObjTaperCone = Nothing

' Insert your code for output 3(Branch run Nozzle with Length)
    Dim Port3Length As Double
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    oPlacePoint.Set (parFace3toCenter + sptOffset - depth) * Cos(parAngle), _
        (parFace3toCenter + sptOffset - depth) * Sin(parAngle), 0
    oDir.Set Cos(parAngle), Sin(parAngle), 0
    Port3Length = parFace3toCenter
    If CmpDblLessThan(Port3Length, flangeThick) Then Port3Length = flangeThick
    Set ObjBranchNozLen = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port3Length)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBranchNozLen
    Set ObjBranchNozLen = Nothing
    
' Insert your code for output 4(Branch run Nozzle with Length)
    Dim Port4Length As Double
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    oPlacePoint.Set (parFace4toCenter + sptOffset - depth) * Cos(parAngle + PI), _
            (parFace4toCenter + sptOffset - depth) * Sin(parAngle + PI), 0

    Port4Length = parFace4toCenter
    If CmpDblLessThan(Port4Length, flangeThick) Then Port4Length = flangeThick

    oDir.Set Cos(parAngle + PI), Sin(parAngle + PI), 0
    Set ObjBranchNozLen1 = CreateNozzleWithLength(4, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port4Length)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBranchNozLen1
    Set ObjBranchNozLen1 = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
