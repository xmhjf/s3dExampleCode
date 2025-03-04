VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReferenceGeometry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CReferenceGeometry.cls
'   Author:         KKC
'   Creation Date:  Tuesday, Aug 28 2006
'   Description:
'
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'    06.Nov.2006         KKC          TR-106572  All of the new Mitered components should be capped.
'    06.Nov.2006         KKC          TR-106569  New Mitered components do not modify correctly
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


Option Explicit
Private Const MODULE = "ReferenceGeometry:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim iOutput     As Double
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim parwidth As Double
    Dim pardepth As Double
    Dim parBendRadius As Double
    Dim lNoOfMiterCuts As Long
    Dim parBendAngle As Double
    Dim parInsulationThickness As Double
    Dim parNoOfSegments As Long
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBendRadius = arrayOfInputs(2)
    parNoOfSegments = arrayOfInputs(3)
    parBendAngle = arrayOfInputs(4)
    parwidth = arrayOfInputs(5)
    pardepth = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)
    lNoOfMiterCuts = parNoOfSegments - 1
    iOutput = 0
    
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    If lNoOfMiterCuts = 0 Then
        stPoint.Set -parBendRadius, 0, 0
        enPoint.Set 0, 0, 0
    Else
        stPoint.Set -parBendRadius * Tan(parBendAngle / 2), 0, 0 ' The starting point
        enPoint.Set parBendRadius * Tan(parBendAngle / 2) * Cos(parBendAngle), _
                    0, parBendRadius * Tan(parBendAngle / 2) * Sin(parBendAngle)
    End If
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory

'Place a Point3D at one of the surfaces
    Dim ObjPoint As IngrGeom3D.Point3d
    
    Set ObjPoint = geomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, _
                                        stPoint.x, stPoint.y, stPoint.z)

    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPoint
    Set ObjPoint = Nothing
'Place a Horizontal Line along symbol X - axis
    Dim ObjHoriLine As IngrGeom3D.Line3d
    Dim startPoint(0 To 2) As Double
    Dim endPoint(0 To 2) As Double
    
    startPoint(0) = stPoint.x + 0.1 * parBendRadius
    startPoint(1) = 0
    startPoint(2) = 0
    
    endPoint(0) = -0.1 * parBendRadius
    endPoint(1) = 0
    endPoint(2) = 0

    Set ObjHoriLine = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, startPoint(0), startPoint(1), startPoint(2), _
                                                                                                    endPoint(0), endPoint(1), endPoint(2))
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriLine
    Set ObjHoriLine = Nothing

'Place a vertical Line
    Dim ObjVertLine As IngrGeom3D.Line3d


    startPoint(0) = 0.1 * enPoint.x
    startPoint(1) = 0
    startPoint(2) = 0.1 * enPoint.z

    endPoint(0) = 0.9 * enPoint.x
    endPoint(1) = 0
    endPoint(2) = 0.9 * enPoint.z
    Set ObjVertLine = geomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, startPoint(0), startPoint(1), startPoint(2), _
                                                                                                    endPoint(0), endPoint(1), endPoint(2))
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVertLine
    Set ObjVertLine = Nothing

'Place a Default Surface
    Dim oCircle As IngrGeom3D.Circle3d
    Dim ObjDefSurface As IngrGeom3D.Plane3d
    Dim normalX As Double
    Dim normalY As Double
    Dim normalZ As Double

    normalX = 0
    normalY = 1
    normalZ = 0
    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                               stPoint.x, stPoint.y, stPoint.z, _
                                                normalX, normalY, normalZ, _
                                                (parwidth) / 2)

    Set ObjDefSurface = geomFactory.Planes3d.CreateByOuterBdry(m_OutputColl.ResourceManager, oCircle)

    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefSurface
    Set ObjDefSurface = Nothing
    Set oCircle = Nothing
    
'Place a Reference Surface
    Dim ObjRefSurface As Object
    
    normalX = Sin(parBendAngle)
    normalY = 0
    normalZ = -Cos(parBendAngle)
    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                enPoint.x, enPoint.y, enPoint.z, _
                                                normalX, normalY, normalZ, _
                                                (parwidth) / 2)
    Set ObjRefSurface = geomFactory.Planes3d.CreateByOuterBdry(m_OutputColl.ResourceManager, oCircle)

    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRefSurface
    Set ObjRefSurface = Nothing
    Set geomFactory = Nothing

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
End Sub


