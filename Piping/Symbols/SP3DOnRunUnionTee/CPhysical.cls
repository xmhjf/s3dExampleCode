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
'   Author:         BG
'   Creation Date:  Monday, Aug 27 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Symbol Model No. is: F127 Page No. D-66 of PDS Piping Component Data Reference Guide.
'   Symbol is created with Ten Outputs
'   The Four physical aspect outputs are created as follows: 1.Union of shape Hexagon made up of
'   Line String and projecting it, 2.Nozzle-1 with length towards -ive X-axis, 3. Nozzle-2 along +ive X-axis
'   and 4. Nozzle-3 with length along +ive Y-axis.
'   Insulation aspect consist of 1. Insulation for body, 2. Nozzle-1, 3. Nozzle-2, 4. Insulation for Branch,
'   5. Insulation for Nozzle-3 and 6. Insulation for Union.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
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
    dAngle = PI / 3
'Point 1
    LineStrPts(0) = parFace2toCenter - parUnionOffset
    LineStrPts(1) = 0
    LineStrPts(2) = parUnionDiameter / 2
'Point 2
    LineStrPts(3) = parFace2toCenter - parUnionOffset
    LineStrPts(4) = Sin(dAngle) * parUnionDiameter / 2
    LineStrPts(5) = Cos(dAngle) * parUnionDiameter / 2
'Point 3
    LineStrPts(6) = parFace2toCenter - parUnionOffset
    LineStrPts(7) = Sin(dAngle) * parUnionDiameter / 2
    LineStrPts(8) = -Cos(dAngle) * parUnionDiameter / 2
'Point 4
    LineStrPts(9) = parFace2toCenter - parUnionOffset
    LineStrPts(10) = 0
    LineStrPts(11) = -parUnionDiameter / 2
'Point 5
    LineStrPts(12) = parFace2toCenter - parUnionOffset
    LineStrPts(13) = -Sin(dAngle) * parUnionDiameter / 2
    LineStrPts(14) = -Cos(dAngle) * parUnionDiameter / 2
'Point 6
    LineStrPts(15) = parFace2toCenter - parUnionOffset
    LineStrPts(16) = -Sin(dAngle) * parUnionDiameter / 2
    LineStrPts(17) = Cos(dAngle) * parUnionDiameter / 2
'Point 7
    LineStrPts(18) = parFace2toCenter - parUnionOffset
    LineStrPts(19) = 0
    LineStrPts(20) = parUnionDiameter / 2
    
'   Refer to Note 21 - assumption: Union Thickness is 0.2 times parUnionOffset
    Dim UnionThickness As Double
    Dim axisVect    As New AutoMath.DVector
    UnionThickness = 0.2 * parUnionOffset
    
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, LineStrPts)
     
    axisVect.Set 1, 0, 0
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
    
'The Body of the Tee is part of the Port1

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

