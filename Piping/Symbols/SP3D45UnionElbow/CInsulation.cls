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
'   Author:         SSP
'   Creation Date:  Wednesday, Aug 27 2003
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    The Details were taken from PDS Piping Component Data - Reference Guide at Page no D-55
'    SN=F63 and MC=E45U. Symbol is created with 12 Outputs - Seven are Insulation aspect
'    outputs. Other are physical aspect outputs and corresponding function are as follows:
'      ObjUnion1 - 'PlaceProjection'
'      ObjElbow - 'PlaceRevolution'
'      ObjUnion2 - 'PlaceProjection' and
'      two objNozzle - 'CreateNozzleWithLength'.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   19.Aug.2005     svsmylav        TR-83739: Added check to ensure non-zero diameter and non-zero length for
'                                           cylinder to avoid PlaceCylinder function failure.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  07.Nov.2007     MA   TR-128456: Provided a check on end points of non flanged insulation port cylinders
'                                in which case small cylinders of negligible thickness will be created
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private m_oGeomHelper As IJSymbolGeometryHelper
Private PI       As Double
Private Const NEGLIGIBLE_THICKNESS = 0.0001

Private Sub Class_Initialize()
    
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
    PI = 4 * Atn(1)
    Set m_oGeomHelper = New SymbolServices
    
    Exit Sub
    
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo Errx
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim flangeDiam2      As Double
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parUnionDiameter As Double
    Dim parTangentLength As Double
    Dim parInsulationThickness As Double
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim ArcStPointToCenter  As Double
    Dim InsDiameterPort  As Double
    Dim InsulationRadius As Double
    Dim UnionInsRadius   As Double
    Dim ElbowRadius      As Double

    Dim iOutput     As Double
    Dim ObjInsulatedBody As Object

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parUnionDiameter = arrayOfInputs(4)
    parTangentLength = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    
    m_oGeomHelper.OutputCollection = m_OutputColl
    
    iOutput = 0
 ' Insert your code for output 1(Insulated Tangent At Port1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset, depth
'   If tangent length is zero and flange thickness is zero then insulation for tangent will fail (Zero length
'   cylinder). We use negligible thickness to protect symbol.
    Dim neglThk As Double
    If CmpDblLessThanOrEqualTo(flangeThick2, NEGLIGIBLE_THICKNESS) Or _
                    CmpDblLessThanOrEqualTo(parTangentLength, NEGLIGIBLE_THICKNESS) Then
        neglThk = NEGLIGIBLE_THICKNESS
    End If
    
    ArcStPointToCenter = parFace1toCenter - parTangentLength
    InsulationRadius = pipeDiam / 2 + parInsulationThickness
    stPoint.Set -parFace1toCenter + flangeThick - neglThk, 0, 0
    enPoint.Set -ArcStPointToCenter - parFace1toCenter / 5, 0, 0

    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, 2 * InsulationRadius

 ' Insert your code for output 2(Insulated Union At Port1)
    stPoint.Set -parFace1toCenter + parTangentLength + parInsulationThickness, 0, 0
    enPoint.Set stPoint.x - parFace1toCenter / 5 - 2 * parInsulationThickness, 0, 0
    UnionInsRadius = parUnionDiameter / 2 + parInsulationThickness
     
     iOutput = iOutput + 1
     m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, 2 * UnionInsRadius

 ' Insert your code for output 3(Insulated Body)
    Dim Angle As Double
    Angle = PI / 4
    ElbowRadius = ArcStPointToCenter / Tan(Angle / 2)

    InsulationRadius = pipeDiam / 2 + parInsulationThickness
'   Ensure that the hub radius is not greater than the distance from the center
'   of rotation to the elbow center line
    If CmpDblGreaterthan(InsulationRadius, ElbowRadius) Then
        InsulationRadius = ElbowRadius
    End If

'   Construct a circle that will be used to create the revolution
    Dim objCircleI   As IngrGeom3D.Circle3d

    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set objCircleI = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                         -ArcStPointToCenter, 0, 0, _
                        1, 0, 0, _
                      InsulationRadius)
'   Revolve the circle about the Z axis passing by the origin
    Dim CP As New AutoMath.DPosition 'revolution center point
    Dim CV As New AutoMath.DVector 'rotation vector for rotation
    CP.Set -ArcStPointToCenter, ElbowRadius, 0
    CV.Set 0, 0, 1
    Set ObjInsulatedBody = PlaceRevolution(m_OutputColl, objCircleI, _
                    CV, CP, Angle, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    Set geomFactory = Nothing
    Set CP = Nothing
    Set CV = Nothing
    Dim sp3dElem1 As IJDObject
    Set sp3dElem1 = objCircleI
    Set objCircleI = Nothing
    sp3dElem1.Remove
    Set sp3dElem1 = Nothing
    
' Insert your code for output 4(Insulated Union At Port2)
    stPoint.Set (ArcStPointToCenter - parInsulationThickness) * Cos(Angle), _
                (ArcStPointToCenter - parInsulationThickness) * Sin(Angle), 0
    enPoint.Set (ArcStPointToCenter + parFace1toCenter / 5 + parInsulationThickness) * Cos(Angle), _
                (ArcStPointToCenter + parFace1toCenter / 5 + parInsulationThickness) * Sin(Angle), 0
    
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, 2 * UnionInsRadius

' Insert your code for output 5(Insulated Tangent At Port2)
    stPoint.Set (parFace2toCenter - flangeThick2 + neglThk) * Cos(Angle), _
                (parFace2toCenter - flangeThick2 + neglThk) * Sin(Angle), 0
    enPoint.Set (ArcStPointToCenter + parFace1toCenter / 5) * Cos(Angle), _
                 (ArcStPointToCenter + parFace1toCenter / 5) * Sin(Angle), 0
    InsulationRadius = pipeDiam2 / 2 + parInsulationThickness

    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, 2 * InsulationRadius


' Insert your code for output 6(Insulated Port1)
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set -parFace1toCenter + flangeThick + parInsulationThickness, 0, 0
    InsDiameterPort = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then _
            InsDiameterPort = flangeDiam + 2 * parInsulationThickness
    
    If CmpDblEqual(flangeThick, 0) Then
        enPoint.Set -parFace1toCenter + NEGLIGIBLE_THICKNESS, 0, 0
    End If
    
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, InsDiameterPort

' Insert your code for output 7(Insulated Port2)
    stPoint.Set parFace2toCenter * Cos(Angle), parFace2toCenter * Sin(Angle), 0
    enPoint.Set (parFace2toCenter - flangeThick2 - parInsulationThickness) * Cos(Angle), _
                (parFace2toCenter - flangeThick2 - parInsulationThickness) * Sin(Angle), 0
    InsDiameterPort = pipeDiam2 + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then _
        InsDiameterPort = flangeDiam2 + 2 * parInsulationThickness
    
    If CmpDblEqual(flangeThick2, 0) Then
        enPoint.Set (parFace2toCenter - NEGLIGIBLE_THICKNESS) * Cos(Angle), _
                        (parFace2toCenter - NEGLIGIBLE_THICKNESS) * Sin(Angle), 0
    End If
    
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, InsDiameterPort
    
    Set stPoint = Nothing
    Set enPoint = Nothing
   
    Exit Sub
    
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
End Sub
