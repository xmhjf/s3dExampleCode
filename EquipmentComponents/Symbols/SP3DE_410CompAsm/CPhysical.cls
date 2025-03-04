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
'   Copyright (c) 2004-06 Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, Apr 1 2004
'   Description:
'   This Symbol detail is taken from PDS Equipment Modeling User's Guide,
'   E410 Symbol in Page no 325.
'   Symbol is created using the following Outputs:
'       a) One Insulation aspect outputs,
'       b) One maintenance aspect output and
'       c) Two Physical aspect outputs: , top surface
'           ObjEquipment uses 'PlaceCylinder' and
'       d) Two ReferenceGeometry aspect outputs: A default Surface and a control point.
'       Other Physical and Insulation aspect outputs are variable outputs.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   06.Dec.2004     SymbolTeam(India)      Removed the small value (0.0000001) that was
'                                          added to the end point for cones in symbol code (Which was a temporary fix).
'   11.Jul.2006     kkc           DI 95670-Replaced names with initials in the revision history.
'   25.Sep.2006     dkl           TR-84724 If variables of type double are compared, appropriate CmpDblXXX functions are used.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private Const TOP_BODY = 1
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput     As Double
    Dim ObjEquipment As Object

    Dim parBottomBodyHeight1 As Double
    Dim parBottomBodyHeight2 As Double
    Dim parBottomBodyDiameter As Double
    Dim parShaftLength As Double
    Dim parShaftDiameter As Double
    Dim parTaperBodyHeight As Double
    Dim parTopBodyHeight As Double
    Dim parTopBodyDiameter As Double
    Dim parBundlePullingLength As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBottomBodyHeight1 = arrayOfInputs(2) 'P1
    parBottomBodyHeight2 = arrayOfInputs(3) 'P2
    parBottomBodyDiameter = arrayOfInputs(4) 'P3
    parShaftLength = arrayOfInputs(5)       'P4
    parShaftDiameter = arrayOfInputs(6)     'P5
    parTaperBodyHeight = arrayOfInputs(7)   'P6
    parTopBodyHeight = arrayOfInputs(8)     'P7
    parTopBodyDiameter = arrayOfInputs(9)   'P8
    parBundlePullingLength = arrayOfInputs(10) 'P9
    parInsulationThickness = arrayOfInputs(11)

    iOutput = 0

    ''Origin is at DP/PP1

'   Input validation
    If CmpDblLessThanOrEqualTo(parBottomBodyHeight1, LINEAR_TOLERANCE) Or CmpDblLessThanOrEqualTo(parBottomBodyHeight2, LINEAR_TOLERANCE) Or _
            CmpDblLessThanOrEqualTo(parBottomBodyDiameter, LINEAR_TOLERANCE) Or CmpDblLessThanOrEqualTo(parShaftLength, LINEAR_TOLERANCE) Then GoTo ErrorLabel

    If CmpDblGreaterthan(parShaftDiameter, LINEAR_TOLERANCE) And CmpDblGreaterthan(parShaftDiameter, parBottomBodyDiameter) Then GoTo ErrorLabel
    If CmpDblGreaterthan(parTopBodyDiameter, LINEAR_TOLERANCE) And CmpDblLessThan(parTopBodyDiameter, parBottomBodyDiameter) Then GoTo ErrorLabel
    
'   Insert your code for output 3(TopSurface)
    Dim dEqpTopPointZ As Double
    dEqpTopPointZ = parBottomBodyHeight1
    Dim oCircle As IngrGeom3D.Circle3d
    Dim baseCenter As New AutoMath.DPosition
    Dim normalVect As New AutoMath.DVector
    Dim circRadius  As Double
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
  
  
    If CmpDblLessThanOrEqualTo(parTopBodyDiameter, LINEAR_TOLERANCE) Then
            circRadius = parBottomBodyDiameter / 2
     Else
            circRadius = parTopBodyDiameter / 2
            dEqpTopPointZ = parBottomBodyHeight1 + parTaperBodyHeight + parTopBodyHeight
     End If

    Dim ObjSurface As Object
    baseCenter.Set 0, 0, dEqpTopPointZ
    
    normalVect.Set 0, 0, 1
    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                    baseCenter.x, baseCenter.y, baseCenter.z, _
                                                    normalVect.x, normalVect.y, normalVect.z, _
                                                    circRadius)
                                                        
'   create persistent default surface plane - the plane can mate ---
    Set ObjSurface = geomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, oCircle)

'   Set  the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSurface
    Set ObjSurface = Nothing
    
    Set baseCenter = Nothing
    Set normalVect = Nothing

    Dim ObjTmpcurves As IJDObject
    Set ObjTmpcurves = oCircle
    ObjTmpcurves.Remove
    Set ObjTmpcurves = Nothing
    Set oCircle = Nothing

    Set geomFactory = Nothing

' Insert your code for output 4(Equipment)
'   Notes:
'       The following variables need to be positive values greater than zero for
'       placing i) Top Body ii) Taper Body and iii) Shaft respectively:
'           - For the Top Body placement, parTopBodyHeight and parTopBodyDiameter,
'           - For the Taper Body placement, parTopBodyDiameter and parTaperBodyHeight,
'           - For the Shaft placement, parShaftLength and parShaftDiameter.
    Dim flag      As Integer
    flag = 0

    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

'   Place Top Body

    If CmpDblGreaterthan(parTopBodyHeight, LINEAR_TOLERANCE) And CmpDblGreaterthan(parTopBodyDiameter, LINEAR_TOLERANCE) Then
        Dim ObjTopBody As Object

        flag = TOP_BODY
        stPoint.Set 0, 0, _
                        parTopBodyHeight + parTaperBodyHeight + parBottomBodyHeight1

        enPoint.Set 0, 0, parTaperBodyHeight + parBottomBodyHeight1
        Set ObjTopBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parTopBodyDiameter, False)

        'Set the output
        m_OutputColl.AddOutput "TopBody", ObjTopBody
        Set ObjTopBody = Nothing
    End If

'   Place Taper Body
    If CmpDblGreaterthan(parTopBodyDiameter, LINEAR_TOLERANCE) And CmpDblGreaterthan(parTaperBodyHeight, LINEAR_TOLERANCE) Then
        Dim ObjTaperBody As Object

        stPoint.Set 0, 0, parTaperBodyHeight + parBottomBodyHeight1
        enPoint.Set 0, 0, parBottomBodyHeight1
        Set ObjTaperBody = PlaceCone(m_OutputColl, stPoint, enPoint, parTopBodyDiameter / 2, parBottomBodyDiameter / 2, False)

        'Set the output
        m_OutputColl.AddOutput "TaperBody", ObjTaperBody
        Set ObjTaperBody = Nothing
    End If

'   Place Bottom Body
    stPoint.Set 0, 0, parBottomBodyHeight1
    enPoint.Set 0, 0, -parBottomBodyHeight2
    Set ObjEquipment = PlaceCylinder(m_OutputColl, stPoint, enPoint, parBottomBodyDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEquipment
    Set ObjEquipment = Nothing

'   Place shaft
    If CmpDblGreaterthan(parShaftDiameter, LINEAR_TOLERANCE) Then
        Dim ObjShaft As Object

        stPoint.Set 0, 0, -parBottomBodyHeight2

        enPoint.Set 0, 0, -parBottomBodyHeight2 - parShaftLength
        Set ObjShaft = PlaceCylinder(m_OutputColl, stPoint, enPoint, parShaftDiameter, True)

        'Set the output
        m_OutputColl.AddOutput "Shaft", ObjShaft
        Set ObjShaft = Nothing
    End If
    Set stPoint = Nothing
    Set enPoint = Nothing


    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext

End Sub

