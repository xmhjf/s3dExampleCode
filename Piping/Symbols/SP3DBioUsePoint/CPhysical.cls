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
'   Author:          BG
'   Creation Date:  Wednesday, March 27 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'This Symbol is Constructed using 11 number of Outputs (Physical and Insulation)
'Physical Aspect: 1.port-1 with length, 2.port-2 with length,3.port-3 with length and 4.Return body
'Insulation Aspect: Closed Insulation is provided.
'1.Port -1 Insulation, 2.Port-1 Tangent Insulation,3.Port -2 Insulation, 4.Port-2 Tangent Insulation,
'5.Port -3 Insulation, 6.Port-3 Tangent Insulation, and 7.Insulation for return
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   01.Aug.2006     KKK                        TR: 80254 Replaced 'm_OutputColl.ResourceManager' with "Nothing'.

'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI           As Double


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
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjReturn As Object
    Dim ObjInsulatedBody As Object
    
    Dim parFacetoFace As Double
    Dim parFacetoCenter As Double
    Dim parFace1toFace3 As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parFacetoCenter = arrayOfInputs(3)
    parFace1toFace3 = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    

    iOutput = 0

' Insert your code for output 1(Fitting - Return )
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim objCirc1 As IngrGeom3D.Circle3d
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim cirNormalX As Double, cirNormalY As Double, cirNormalZ As Double
    Dim cir1CenterX As Double, cir1CenterY As Double, cir1CenterZ As Double
    Dim centPoint   As New AutoMath.DPosition
    Dim axis        As New AutoMath.DVector
    cirNormalX = -1
    cirNormalY = 0
    cirNormalZ = 0
    cir1CenterX = CenterPos.x - parFacetoFace / 2
    cir1CenterY = CenterPos.y
    cir1CenterZ = CenterPos.z
    Set objCirc1 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                                                Nothing, cir1CenterX, _
                                                cir1CenterY, cir1CenterZ, cirNormalX, _
                                                cirNormalY, cirNormalZ, pipeDiam / 2)

    centPoint.Set CenterPos.x - parFacetoFace / 2, _
                  CenterPos.y + parFacetoFace / 2, _
                  CenterPos.z
    axis.Set 0, 0, 1
    Set ObjReturn = PlaceRevolution(m_OutputColl, objCirc1, _
                                    axis, centPoint, PI, _
                                    True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjReturn
    Set ObjReturn = Nothing
    Set centPoint = Nothing
    Set axis = Nothing
    Set geomFactory = Nothing
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim PortLength As Double
    
    'Dim faceToFace  As Double
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    PortLength = parFacetoCenter - parFacetoFace / 2
    'faceToFace = arrayOfInputs(2)
    oPlacePoint.Set CenterPos.x - parFacetoCenter - sptOffset + depth, _
                     CenterPos.y, _
                     CenterPos.z
    oDir.Set -1, 0, 0
    'Set oPartFclt = arrayOfInputs(1)
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, PortLength)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set CenterPos.x - parFacetoCenter - sptOffset + depth, _
                    CenterPos.y + parFacetoFace, _
                    CenterPos.z
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, PortLength)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 3
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set CenterPos.x + (parFace1toFace3 - parFacetoCenter) + sptOffset - depth, _
                    CenterPos.y + parFacetoFace / 2, _
                    CenterPos.z
    PortLength = parFace1toFace3 - parFacetoCenter
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, oPlacePoint, PortLength)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
' remove curve
    Dim sp3dElem1 As IJDObject
    Set sp3dElem1 = objCirc1
    Set objCirc1 = Nothing
    sp3dElem1.Remove
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
