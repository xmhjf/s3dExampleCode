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
' Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          PS
'   Creation Date:   Tuesday, NOV 13 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol
''''''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private PI           As Double
Private Const MODULE = "Insulation:"
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
    Dim objcyl      As Object
    Dim ocbody      As Object
    
    
    Dim oCircle     As IngrGeom3D.Circle3d
    Dim oline       As IngrGeom3D.Line3d
    
    
    Dim BaseCenPoint As New AutoMath.DPosition
    Dim TopCenPoint  As New AutoMath.DPosition
    Dim startPoint   As New AutoMath.DPosition
    Dim endPoint     As New AutoMath.DPosition
    Dim axisVect     As New AutoMath.DVector
    Dim RevCenPt     As New AutoMath.DPosition
    Dim geomFactory  As New IngrGeom3D.GeometryFactory
    
    Dim pardimen1 As Double
    Dim pardimen2 As Double
    Dim pardimen3 As Double
    Dim pardimen4 As Double
    Dim pardimen5 As Double
    Dim parInsulationThickness As Double
        
    
    Set oPartFclt = arrayOfInputs(1)
    pardimen1 = arrayOfInputs(2)       ' Length of the part'
    pardimen2 = arrayOfInputs(3)       'Diameter of the part'
    pardimen3 = arrayOfInputs(4)       'height1 of the part'
    pardimen4 = arrayOfInputs(5)       'offset distance of the part'
    pardimen5 = arrayOfInputs(6)       'height2 of the part'
    parInsulationThickness = arrayOfInputs(7)
    iOutput = 0
    
    
    ' OutPut 11:Cylinder 1
    
    BaseCenPoint.Set 0, 0, 0
    TopCenPoint.Set pardimen1, 0, 0
    Set objcyl = PlaceCylinder(m_OutputColl, BaseCenPoint, TopCenPoint, pardimen2 + 2 * parInsulationThickness, True)
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objcyl
    Set objcyl = Nothing
      
    ' OutPut 12: Placing the neck
    startPoint.Set pardimen3, 0, -pardimen4 / 2
    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, startPoint.x, startPoint.y, startPoint.z, 0, 0, -1, parInsulationThickness + pardimen2 / 2)
    RevCenPt.Set pardimen5, 0, -pardimen4 / 2
    axisVect.Set 0, 1, 0
    Set ocbody = PlaceRevolution(m_OutputColl, oCircle, axisVect, RevCenPt, 7 * PI / 18, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ocbody
    Set ocbody = Nothing
    Set oCircle = Nothing
    Set axisVect = Nothing
    Set RevCenPt = Nothing
 
    'OutPut 13: Placing the conections i.e the neck to the cylider body
    startPoint.Set pardimen3, 0, -pardimen4 / 2
    endPoint.Set pardimen3, 0, 0
    Set objcyl = PlaceCylinder(m_OutputColl, startPoint, endPoint, pardimen2 + 2 * parInsulationThickness, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objcyl
    Set objcyl = Nothing
    Set geomFactory = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    
    ' OutPut 14:Place Cylinder2
    Dim radius As Double
    radius = pardimen3 - pardimen5
    BaseCenPoint.Set pardimen5 + radius * Cos(7 * PI / 18), 0, -(pardimen4 / 2) - radius * Sin(7 * PI / 18)
    TopCenPoint.Set BaseCenPoint.x - 0.15 * pardimen1 * Sin(7 * PI / 18), 0, BaseCenPoint.z - 0.15 * pardimen1 * Cos(7 * PI / 18)
    Set objcyl = PlaceCylinder(m_OutputColl, BaseCenPoint, TopCenPoint, pardimen2 + 2 * parInsulationThickness, True)
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objcyl
    Set objcyl = Nothing
    Set BaseCenPoint = Nothing
    Set TopCenPoint = Nothing
        
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
