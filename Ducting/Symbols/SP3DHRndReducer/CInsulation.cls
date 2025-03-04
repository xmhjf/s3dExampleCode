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
'   Copyright (c) 2006-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   ProgID:         SP3DHRndReducer.HRndReducer
'   Author:         kkk
'   Creation Date:  Thursday, Feb 09 2006
'   Description:
'    Insulation aspect of HVAC Round Reducer
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  23.Jan.2008     VRK  TR-125293 Incorrect Dimensions of various components in HVAC(Lindab Catalog)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = Insulation:    'Used for error messages
Private Sub Class_Initialize()
'''
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim iOutput As Double

    Dim ObjBodyIns As Object

    Dim parWidth As Double
    Dim parBWidth As Double
    Dim parLength As Double
    Dim parInsulationThickness As Double

    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parBWidth = arrayOfInputs(3)
    parLength = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)

    iOutput = 0

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

    'Checking for the Part Data Basis Property
    Dim lPartDataBasis As Long
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPartFclt
    lPartDataBasis = oHvacPart.PartDataBasis
    Set oHvacPart = Nothing

    ' a)Round Reducer, conical-133
    ' b)Round Reducer, eccentric-134
    ' c)Round Reducer, Stepped-135

    Select Case lPartDataBasis

    Case Is <= 1, 133 'Round Reducer, conical

        'Insert your code for output 1(Reducer)
        stPoint.Set -parLength / 2, 0, 0
        enPoint.Set parLength / 2, 0, 0
        Set ObjBodyIns = PlaceCone(m_OutputColl, stPoint, enPoint, (parWidth / 2) + parInsulationThickness, (parBWidth / 2) + parInsulationThickness, False)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyIns
        Set ObjBodyIns = Nothing

    Case 134 'Round Reducer, eccentric

        Dim dOffsetBetweenCentreLines As Double
        dOffsetBetweenCentreLines = (parWidth - parBWidth) / 2
        Dim cirNormalX As Double
        Dim cirNormalY As Double
        Dim cirNormalZ As Double

        cirNormalX = 1
        cirNormalY = 0
        cirNormalZ = 0

        Dim objCircle1 As IngrGeom3D.Circle3d
        Dim cir1CenterX As Double
        Dim cir1CenterY As Double
        Dim cir1CenterZ As Double

        cir1CenterX = -parLength / 2
        cir1CenterY = 0
        cir1CenterZ = 0
        Dim geomFactory As New IngrGeom3D.GeometryFactory
        Set objCircle1 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                         Nothing, cir1CenterX, _
                         cir1CenterY, cir1CenterZ, cirNormalX, _
                         cirNormalY, cirNormalZ, (parWidth / 2) + parInsulationThickness)


        'Create Circle2
        Dim objCircle2 As IngrGeom3D.Circle3d
        Dim cir2CenterX As Double
        Dim cir2CenterY As Double
        Dim cir2CenterZ As Double

        cir2CenterX = parLength / 2
        cir2CenterY = -dOffsetBetweenCentreLines
        cir2CenterZ = 0
        Set objCircle2 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                         Nothing, cir2CenterX, _
                         cir2CenterY, cir2CenterZ, cirNormalX, _
                         cirNormalY, cirNormalZ, (parBWidth / 2) + parInsulationThickness)

        'Create Ruled Surface for body
        Dim ObjRuledBody As IngrGeom3D.RuledSurface3d
        Set ObjRuledBody = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                                      objCircle1, objCircle2, True)
        Set ObjBodyIns = ObjRuledBody
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyIns
        Set ObjBodyIns = Nothing
        Set ObjRuledBody = Nothing
        Set geomFactory = Nothing
        ' remove curve
        Dim sp3dElem1 As IJDObject
        Dim sp3dElem2 As IJDObject

        Set sp3dElem1 = objCircle1
        Set sp3dElem2 = objCircle2
        Set objCircle1 = Nothing
        Set objCircle2 = Nothing
        sp3dElem1.Remove
        sp3dElem2.Remove

    Case 135 'Round Reducer, Stepped

        stPoint.Set -parLength / 2, 0, 0
        enPoint.Set 0, 0, 0
        Set ObjBodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, parWidth + 2 * parInsulationThickness, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyIns
        Set ObjBodyIns = Nothing

        stPoint.Set 0, 0, 0
        enPoint.Set parLength / 2, 0, 0
        Set ObjBodyIns = PlaceCone(m_OutputColl, stPoint, enPoint, (parWidth / 2) + parInsulationThickness, (parBWidth / 2) + parInsulationThickness, False)
        'Set the output
        m_OutputColl.AddOutput "ObjInsBody", ObjBodyIns
        Set ObjBodyIns = Nothing

    Case Else
        GoTo ErrorLabel:    'Invalid specification.

    End Select
    Set stPoint = Nothing
    Set enPoint = Nothing
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
