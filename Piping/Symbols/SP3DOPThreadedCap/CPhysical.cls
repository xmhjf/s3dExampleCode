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
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         RRK
'   Creation Date:  Thursday, Jul 24 2006
'   Description:
'       Threaded Cap Operator (for Relief Valve, Source: KNUCKLE, at www.thevalveshop.com).
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "CPhysical:" 'Used for error messages
Private m_oGeomHelper As IJSymbolGeometryHelper

Private PI As Double
Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
  
    PI = Atn(1) * 4
    Set m_oGeomHelper = New SymbolServices
    
    Exit Sub
    
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
        Err.HelpFile, Err.HelpContext
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
    iOutput = 0
    
    Dim parOperatorHeight As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parOperatorHeight = arrayOfInputs(2)
   
    Dim stPoint  As Object
    Dim enPoint  As Object
    
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
     
     m_oGeomHelper.OutputCollection = m_OutputColl

' Place Output 1(Cylinder)
'------------------------------------------------------------------------------------------------------------------------------------------
    Dim dCylDia As Double
    Dim dCylHt As Double
    
'Assumption:
    dCylDia = 0.0254
    dCylHt = 0.75 * parOperatorHeight
    
    stPoint.Set 0, 0, 0
    enPoint.Set 0, dCylHt, 0
    
    ' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dCylDia
    
    Set enPoint = Nothing
    
     
' Place Output 2(Hexagonal Nut)
'------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim ObjNut      As Object
    Dim geomFactory As Object
    Dim iCount      As Integer
    Dim dAngle      As Double
    Dim dNutWidth    As Double
    Dim dNutHeight   As Double
    Dim dNominalDia As Double

    Dim axisVect   As Object
    Dim LineStrPoints(0 To 20)  As Double
    Dim oLineString As IngrGeom3D.LineString3d
    
'Assumption
    dNominalDia = 0.0127
    dNutWidth = 1.5 * dNominalDia
    dNutHeight = parOperatorHeight - dCylHt
    dAngle = PI / 3

    
    stPoint.Set 0, dCylHt, 0
 
    For iCount = 1 To 7
        LineStrPoints(3 * iCount - 3) = stPoint.x + (dNutWidth / 2) * Cos(PI / 6 + iCount * dAngle)
        LineStrPoints(3 * iCount - 2) = stPoint.y
        LineStrPoints(3 * iCount - 1) = stPoint.z + (dNutWidth / 2) * Sin(PI / 6 + iCount * dAngle)
    Next iCount
    
    Set ObjNut = New IngrGeom3D.Projection3d
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, LineStrPoints)
    Set axisVect = New AutoMath.DVector
    axisVect.Set 0, 1, 0
  
    Set ObjNut = PlaceProjection(m_OutputColl, oLineString, axisVect, dNutHeight, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNut
    Set ObjNut = Nothing
    Set axisVect = Nothing
    Set geomFactory = Nothing
    Set oLineString = Nothing
    Set stPoint = Nothing

    Exit Sub
 
ErrorLabel:
 ReportUnanticipatedError MODULE, METHOD
 
End Sub
