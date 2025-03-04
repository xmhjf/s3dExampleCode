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
'    Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'    CPhysical.Cls
'Author:                RH
'    Creation Date:  01-May-08
'Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'    Change History:
'    dd.mmm.yyyy     who     change description
'    -----------     ---     ------------------
'    13.June.2008     VRK     CR-134560:Provide pull-pit/manhole equipment symbol for use with duct banks
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:"    'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim iOutput As Integer

    Dim parPullBoxHeight As Double
    Dim parPullBoxWidth As Double
    Dim parPullBoxLength As Double
    Dim parPullBoxThickness As Double

    Dim PullBoxHole1Height As Double
    Dim PullBoxHole1Width As Double
    Dim PullBoxHole1Offset As Double
    Dim PullBoxHole2Height As Double
    Dim PullBoxHole2Width As Double
    Dim PullBoxHole2Offset As Double
    Dim PullBoxHole1ThruStart As Boolean
    Dim PullBoxHole1ThruEnd As Boolean
    Dim PullBoxHole2ThruStart As Boolean
    Dim PullBoxHole2ThruEnd As Boolean
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parPullBoxHeight = arrayOfInputs(2)         'A
    parPullBoxWidth = arrayOfInputs(3)          'B
    parPullBoxLength = arrayOfInputs(4)         'C
    parPullBoxThickness = arrayOfInputs(5)      'D
    PullBoxHole1Height = arrayOfInputs(6)       'E
    PullBoxHole1Width = arrayOfInputs(7)        'F
    PullBoxHole1Offset = arrayOfInputs(8)       'G
    PullBoxHole1ThruStart = arrayOfInputs(9)
    PullBoxHole1ThruEnd = arrayOfInputs(10)
    PullBoxHole2Height = arrayOfInputs(11)      'H
    PullBoxHole2Width = arrayOfInputs(12)       'I
    PullBoxHole2Offset = arrayOfInputs(13)      'J
    PullBoxHole2ThruStart = arrayOfInputs(14)
    PullBoxHole2ThruEnd = arrayOfInputs(15)
    
    iOutput = 0

    'Insert your code for output (Default surface)
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New GeometryFactory
    Dim objPlane As IngrGeom3D.Plane3d
    Dim dLineStrPoints(0 To 14) As Double
    Dim oLineString As IngrGeom3D.LineString3d

    dLineStrPoints(0) = 0.25 * parPullBoxLength
    dLineStrPoints(1) = 0.25 * parPullBoxWidth
    dLineStrPoints(2) = -0.5 * parPullBoxHeight

    dLineStrPoints(3) = 0.75 * parPullBoxLength
    dLineStrPoints(4) = 0.25 * parPullBoxWidth
    dLineStrPoints(5) = dLineStrPoints(2)

    dLineStrPoints(6) = 0.75 * parPullBoxLength
    dLineStrPoints(7) = -0.25 * parPullBoxWidth
    dLineStrPoints(8) = dLineStrPoints(2)

    dLineStrPoints(9) = 0.25 * parPullBoxLength
    dLineStrPoints(10) = -0.25 * parPullBoxWidth
    dLineStrPoints(11) = dLineStrPoints(2)

    dLineStrPoints(12) = dLineStrPoints(0)
    dLineStrPoints(13) = dLineStrPoints(1)
    dLineStrPoints(14) = dLineStrPoints(2)

    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)

    'Create persistent default surface plane - the plane can mate ---
    Set objPlane = oGeomFactory.Planes3d.CreateByOuterBdry _
                   (m_OutputColl.ResourceManager, oLineString)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPlane
    Set objPlane = Nothing
    Set oLineString = Nothing
    Set oGeomFactory = Nothing

    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
