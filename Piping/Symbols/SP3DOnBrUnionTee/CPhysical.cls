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
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         MS
'   Creation Date:  Monday, Aug 26 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol Model No. is: F128 Page No. D-66 of PDS Piping Component Data Reference Guide.
'  Symbol is created with Ten Outputs
'   The Four physical aspect outputs are created as follows:
'   ObjUnionBody- Using 'PlaceProjection' function,
'   One ObjNozzle object by using 'CreateNozzle' function and another ObjNozzle by using CreateNozzleWithLength
' The Six Insulation aspect outputs are created as follows:
' ObjInsulatedBody ,ObjInsulatedPort1, ObjInsulatedPort2 , ObjInsulatedBranch, ObjInsulatedPort3 and
' ObjInsulatedUnion using PlaceCylinder.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CPhysical:" 'Used for error messages
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
    Dim ObjUnionBody As Object
    Dim ObjInsulatedBody As Object
    Dim ObjInsulatedPort1 As Object
    Dim ObjInsulatedPort2 As Object
    Dim ObjInsulatedBranch As Object
    Dim ObjInsulatedUnion As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parUnionDiameter As Double
    Dim parUnionOffset As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parFace3toCenter = arrayOfInputs(4)
    parUnionDiameter = arrayOfInputs(5)
    parUnionOffset = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)
    
    iOutput = 0

 ' Insert your code for output 1(Body of Union)
     Dim LineStrPts(0 To 20)  As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim dAngle As Double
    dAngle = PI / 6
'Point 1
    LineStrPts(0) = 0
    LineStrPts(1) = parFace3toCenter - parUnionOffset
    LineStrPts(2) = parUnionDiameter / 2
'Point 2
    LineStrPts(3) = parUnionDiameter / 2 * Cos(dAngle)
    LineStrPts(4) = parFace3toCenter - parUnionOffset
    LineStrPts(5) = parUnionDiameter / 2 * Sin(dAngle)
'Point 3
    LineStrPts(6) = parUnionDiameter / 2 * Cos(dAngle)
    LineStrPts(7) = parFace3toCenter - parUnionOffset
    LineStrPts(8) = -parUnionDiameter / 2 * Sin(dAngle)
'Point 4
    LineStrPts(9) = 0
    LineStrPts(10) = parFace3toCenter - parUnionOffset
    LineStrPts(11) = -parUnionDiameter / 2
'Point 5
    LineStrPts(12) = -parUnionDiameter / 2 * Cos(dAngle)
    LineStrPts(13) = parFace3toCenter - parUnionOffset
    LineStrPts(14) = -parUnionDiameter / 2 * Sin(dAngle)
'Point 6
    LineStrPts(15) = -parUnionDiameter / 2 * Cos(dAngle)
    LineStrPts(16) = parFace3toCenter - parUnionOffset
    LineStrPts(17) = parUnionDiameter / 2 * Sin(dAngle)
'Point 7
    LineStrPts(18) = 0
    LineStrPts(19) = parFace3toCenter - parUnionOffset
    LineStrPts(20) = parUnionDiameter / 2

'   Refer to Note 21 - assumption: Union Thickness is 0.2 times parUnionOffset
    Dim UnionThickness As Double
    Dim axisVect    As New AutoMath.DVector
    UnionThickness = 0.2 * parUnionOffset
    
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, LineStrPts)
     
    axisVect.Set 0, 1, 0
    Set ObjUnionBody = PlaceProjection(m_OutputColl, oLineString, axisVect, UnionThickness, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjUnionBody
    Set ObjUnionBody = Nothing
    Set axisVect = Nothing
    Set geomFactory = Nothing
    Dim sp3dElem1 As IJDObject
    Set sp3dElem1 = oLineString
    sp3dElem1.Remove
    Set oLineString = Nothing
    
'THE Body of the Tee is part of the Port1

' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim Port1Length As Double
    Port1Length = parFace1toCenter + parFace2toCenter - flangeThick
    If Port1Length < flangeThick Then
        Port1Length = flangeThick
    End If
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port1Length)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2
    oPlacePoint.Set parFace2toCenter + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 3
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set 0, parFace3toCenter + sptOffset - depth, 0
    oDir.Set 0, 1, 0
    If parFace3toCenter < flangeThick Then
        parFace3toCenter = flangeThick
    End If
    Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, oPlacePoint, parFace3toCenter)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
 
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

