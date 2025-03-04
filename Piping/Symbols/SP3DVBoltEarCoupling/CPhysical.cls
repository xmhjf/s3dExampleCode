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
'   Author:         kkk
'   Creation Date:  Monday, Nov 21, 2005

'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mm.yyyy          who          change description
'   -----------         -----        ------------------
'    22.06.2006         KKC    TR-97303. Removed nozzle entities from the code as clamps
'                                      donot require nozzles entities to be placed.
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
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim iOutput     As Double
    Dim ObjBody As Object
    Dim ObjInsulatedBody As Object
    
    Dim parInsulationThickness As Double
    Dim parClampWidth As Double
    Dim parClampOuterDiameter As Double
    Dim parClampEarOuterDiameter As Double
    Dim parNumberofBoltEars As Long

' Inputs

    Set oPartFclt = arrayOfInputs(1)
    parClampWidth = arrayOfInputs(2)
    parClampOuterDiameter = arrayOfInputs(3)
    parClampEarOuterDiameter = arrayOfInputs(4)
    parNumberofBoltEars = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    
    iOutput = 0


    Dim ObjClampBody As Object
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
     
' Insert your code for output 1(Clamp Body)
    
     Dim stPoint   As New AutoMath.DPosition
     Dim enPoint   As New AutoMath.DPosition
    
     stPoint.Set -parClampWidth / 2, 0, 0
     enPoint.Set parClampWidth / 2, 0, 0
    
     Set ObjClampBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parClampOuterDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjClampBody
    Set ObjClampBody = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    

'' Place Nozzle 1
'
'    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
'
'    Dim oPlacePoint As AutoMath.DPosition
'    Dim oDir        As AutoMath.DVector
'    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
'    Set oPlacePoint = New AutoMath.DPosition
'    Set oDir = New AutoMath.DVector
'    oPlacePoint.Set -parClampWidth / 2 - cptOffset + depth, 0, 0
'    oDir.Set -1, 0, 0
'
'    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'' Set the output
'    iOutput = iOutput + 1
'    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
'    Set objNozzle = Nothing
'
'' Place Nozzle 2
'    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
'    oPlacePoint.Set parClampWidth / 2 + cptOffset - depth, 0, 0
'    oDir.Set 1, 0, 0
'
'    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'
'' Set the output
'    iOutput = iOutput + 1
'    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
'    Set objNozzle = Nothing
'    Set oPlacePoint = Nothing
'    Set oDir = Nothing
'End If
    '-------------------------
    ' Placement of Ears
    '-------------------------
    
    Dim dLugWidth As Double
    Dim dLugHeight As Double
    Dim objLugbody   As Object

    Dim dRadius  As Double
    Dim dTheta As Double
    Dim dAngle As Double
    
    Dim dPlacePoints(0 To 14) As Double
    Dim axisVect As New AutoMath.DVector
    'Assuming that the LugHeight = 20% of Clamp Outer Diameter
    
    dLugHeight = 0.2 * parClampOuterDiameter / 2   'H
    dLugWidth = (parClampEarOuterDiameter - parClampOuterDiameter) / 2  'W

    If CmpDblLessThanOrEqualTo(parNumberofBoltEars, 2) Then parNumberofBoltEars = 2

    dAngle = 2 * PI / parNumberofBoltEars
    dRadius = (parClampEarOuterDiameter / 2 - dLugWidth / 2) * 0.98 'R
    
    For dTheta = 0 To 2 * PI Step dAngle

        dPlacePoints(0) = -parClampWidth / 2
        dPlacePoints(1) = (dRadius - dLugWidth / 2) * Cos(dTheta) + (dLugHeight / 2) * Sin(dTheta)
        dPlacePoints(2) = (dRadius - dLugWidth / 2) * Sin(dTheta) - (dLugHeight / 2) * Cos(dTheta)
        
        dPlacePoints(3) = -parClampWidth / 2
        dPlacePoints(4) = (dRadius + dLugWidth / 2) * Cos(dTheta) + (dLugHeight / 2) * Sin(dTheta)
        dPlacePoints(5) = (dRadius + dLugWidth / 2) * Sin(dTheta) - (dLugHeight / 2) * Cos(dTheta)
        
        dPlacePoints(6) = -parClampWidth / 2
        dPlacePoints(7) = (dRadius + dLugWidth / 2) * Cos(dTheta) - (dLugHeight / 2) * Sin(dTheta)
        dPlacePoints(8) = (dRadius + dLugWidth / 2) * Sin(dTheta) + (dLugHeight / 2) * Cos(dTheta)
        
        dPlacePoints(9) = -parClampWidth / 2
        dPlacePoints(10) = (dRadius - dLugWidth / 2) * Cos(dTheta) - (dLugHeight / 2) * Sin(dTheta)
        dPlacePoints(11) = (dRadius - dLugWidth / 2) * Sin(dTheta) + (dLugHeight / 2) * Cos(dTheta)
        
        dPlacePoints(12) = -parClampWidth / 2
        dPlacePoints(13) = (dRadius - dLugWidth / 2) * Cos(dTheta) + (dLugHeight / 2) * Sin(dTheta)
        dPlacePoints(14) = (dRadius - dLugWidth / 2) * Sin(dTheta) - (dLugHeight / 2) * Cos(dTheta)
        
        axisVect.Set 1, 0, 0
        
        Set objLugbody = PlaceProjectedBox(m_OutputColl, dPlacePoints, axisVect, parClampWidth, True)
        m_OutputColl.AddOutput "BoltEars_", objLugbody
        Set objLugbody = Nothing
        Set axisVect = Nothing

    Next dTheta

Exit Sub
    
ErrorLabel:
ReportUnanticipatedError MODULE, METHOD
 
End Sub
Private Function PlaceProjectedBox(ByVal objOutputColl As Object, _
                        dPoints() As Double, _
                        ByVal axisVector As AutoMath.DVector, _
                        height As Double, _
                        isCapped As Boolean) As Object

    Const METHOD = "PlaceProjectedBox:"
    On Error GoTo ErrorHandler
      
    Dim objProjection   As IngrGeom3D.Projection3d
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim oLineStr As IngrGeom3D.LineString3d
        
    Set oLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dPoints)
    Set objProjection = oGeomFactory.Projections3d.CreateByCurve( _
                                                    objOutputColl.ResourceManager, _
                                                    oLineStr, _
                                                    axisVector.x, axisVector.y, axisVector.z, _
                                                    height, isCapped)
        
    Set PlaceProjectedBox = objProjection
    Set objProjection = Nothing
    Set oLineStr = Nothing
    Set oGeomFactory = Nothing

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
        
End Function
