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
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author: Veena
'   Creation Date:  Friday, Nov 25 2006
'
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'******************************************************************************


Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private PI       As Double

Private Const CYLINDRICAL_SUPPORTS = 1
Private Const CUBOID_SUPPORTS = 2

Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
     PI = 4 * Atn(1)
     Set m_oSymGeomHelper = New SymbolServices
     Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
End Sub
Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim CenterPos       As New AutoMath.DPosition
   
 
    Dim parSupportAngularLocation As Double
    Dim parNumberOfSupports As Long
    Dim parSupportHeight As Double
    Dim parSupportLength As Double
    Dim parSupportThickness As Double
    Dim parSupportRadialLocation As Double
    Dim parStartHeight As Double
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parSupportAngularLocation = arrayOfInputs(2)
    parNumberOfSupports = arrayOfInputs(3)
    parSupportHeight = arrayOfInputs(4)
    parSupportLength = arrayOfInputs(5)
    parSupportThickness = arrayOfInputs(6)
    parSupportRadialLocation = arrayOfInputs(7)
    parStartHeight = arrayOfInputs(8)
    
    
  
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
     
    CenterPos.Set 0, 0, 0
    
   If parNumberOfSupports > 0 And parSupportLength > 0 And parSupportHeight > 0 Then
        ' Insert your code for output  (Supports/DefaultSurfaces)
        Dim supportType As Integer
        Dim normalVect As New AutoMath.DVector
        Dim oCircle As Object
        Dim oPoint As Object
        Dim parAngle As Double
        Dim i As Integer
    
        'Set the number of Supports/DefaultSurfaces and their type
        Dim ObjSupport As Object
        Dim ObjDefaultSurface As Object

        If parSupportThickness = 0 Then
            supportType = CYLINDRICAL_SUPPORTS
        Else
            supportType = CUBOID_SUPPORTS
        End If

        normalVect.Set 0, 0, 1
        
        Dim iOutput As Double
        iOutput = 0
        'Place supports
        Select Case supportType
           Case CYLINDRICAL_SUPPORTS
               For i = 1 To parNumberOfSupports
                    'Center point position in horizontal plane is needed
                    parAngle = PI / 2 - parSupportAngularLocation + _
                                    (i - 1) * (2 * PI / parNumberOfSupports)
                    stPoint.Set CenterPos.x + (parSupportRadialLocation + parSupportLength / 2) * Cos(parAngle), _
                                CenterPos.y + (parSupportRadialLocation + parSupportLength / 2) * Sin(parAngle), _
                                CenterPos.z + (parSupportHeight - parStartHeight)
                    enPoint.Set stPoint.x, _
                                stPoint.y, _
                                stPoint.z - parSupportHeight
                    Set ObjSupport = PlaceCylinder(m_outputColl, stPoint, enPoint, parSupportLength, True)
                    'Set the output
                    m_outputColl.AddOutput "Supports_", ObjSupport
                    Set ObjSupport = Nothing
                    
                    'Place Point at both Ends
                    Set oPoint = geomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, enPoint.x, enPoint.y, enPoint.z)
                    m_outputColl.AddOutput "Point1_", oPoint
                    Set oPoint = Nothing
                    
                    Set oPoint = geomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, stPoint.x, stPoint.y, stPoint.z)
                    m_outputColl.AddOutput "Point2_", oPoint
                    Set oPoint = Nothing

                    'Place Circle3d at both Ends
                    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(m_outputColl.ResourceManager, _
                                                    enPoint.x, enPoint.y, enPoint.z, _
                                                    normalVect.x, normalVect.y, normalVect.z, _
                                                    parSupportLength / 2)
                    'Set the output
                    m_outputColl.AddOutput "Edge1_", oCircle
                    Set oCircle = Nothing

                    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(m_outputColl.ResourceManager, _
                                                    stPoint.x, stPoint.y, stPoint.z, _
                                                    normalVect.x, normalVect.y, normalVect.z, _
                                                    parSupportLength / 2)
                    'Set the output
                    m_outputColl.AddOutput "Edge2_", oCircle
                    Set oCircle = Nothing

                    
                Next i
                Set normalVect = Nothing
    
           Case CUBOID_SUPPORTS
               'Consider an initial support cross section to be on the X axis, in horizontal plane
                Dim initSecPoints(0 To 14)  As Double
    
                'Left hand side Bottom point
                initSecPoints(0) = CenterPos.x + parSupportRadialLocation
                initSecPoints(1) = CenterPos.y - parSupportLength / 2
                initSecPoints(2) = CenterPos.z - parStartHeight
    
                'Right hand side Bottom point
                initSecPoints(3) = CenterPos.x + parSupportRadialLocation + parSupportThickness
                initSecPoints(4) = CenterPos.y - parSupportLength / 2
                initSecPoints(5) = CenterPos.z - parStartHeight
                
                'Right hand side Top point
                initSecPoints(6) = CenterPos.x + parSupportRadialLocation + parSupportThickness
                initSecPoints(7) = CenterPos.y + parSupportLength / 2
                initSecPoints(8) = CenterPos.z - parStartHeight
            
                'Left hand side Top point
                initSecPoints(9) = CenterPos.x + parSupportRadialLocation
                initSecPoints(10) = CenterPos.y + parSupportLength / 2
                initSecPoints(11) = CenterPos.z - parStartHeight
                
                ' Left hand side Bottom point
                initSecPoints(12) = initSecPoints(0)
                initSecPoints(13) = initSecPoints(1)
                initSecPoints(14) = initSecPoints(2)
                         
            
                'Prepare profile points
                Dim oLineString As IngrGeom3D.LineString3d
                Dim objLineString As IJDObject
                Dim lineStrPoints(0 To 14)  As Double
                Dim J As Integer
                
                Dim ProjVector As New AutoMath.DVector
                ProjVector.Set 0, 0, 1
    
                For i = 1 To parNumberOfSupports
                    Set oLineString = New IngrGeom3D.LineString3d
                    'Center point position in horizontal plane is needed
                    parAngle = PI / 2 - parSupportAngularLocation + _
                                    (i - 1) * (2 * PI / parNumberOfSupports)
    
                    'Rotate initial support cross section to the required location
                    For J = 0 To 14 Step 3
                        lineStrPoints(J) = initSecPoints(J) * Cos(parAngle) - initSecPoints(J + 1) * Sin(parAngle)
                        lineStrPoints(J + 1) = initSecPoints(J) * Sin(parAngle) + initSecPoints(J + 1) * Cos(parAngle)
                        lineStrPoints(J + 2) = initSecPoints(J + 2)
                    Next J
                    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, lineStrPoints)
                    Set ObjSupport = PlaceProjection(m_outputColl, oLineString, ProjVector, parSupportHeight, True)
                   
                    'Set the output
                    m_outputColl.AddOutput "Supports_", ObjSupport
                    Set ObjSupport = Nothing
                            
                         Dim oPoint1 As Point3d
                         Dim oPoint2 As Point3d
                        
                        
                        Set oPoint1 = geomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, lineStrPoints(0), lineStrPoints(1), lineStrPoints(2))
                        Set oPoint2 = geomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, lineStrPoints(6), lineStrPoints(7), lineStrPoints(8))
                       
                        oPoint1.GetPoint lineStrPoints(0), lineStrPoints(1), lineStrPoints(2)
                        oPoint2.GetPoint lineStrPoints(6), lineStrPoints(7), lineStrPoints(8)

                        Dim x, y, z As Double
                        x = (lineStrPoints(0) + lineStrPoints(6)) / 2
                        y = (lineStrPoints(1) + lineStrPoints(7)) / 2
                        z = (lineStrPoints(2) + lineStrPoints(8)) / 2
                         
                        Set oPoint = geomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, x, y, z)
                        m_outputColl.AddOutput "Point1_", oPoint
                        Set oPoint = Nothing
                        
                        Set oPoint = geomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, x, y, (z + parSupportHeight))

                        m_outputColl.AddOutput "Point2_", oPoint
                        Set oPoint = Nothing
                        
                        Dim linepoints(0 To 11) As Double
                        Dim line As Line3d

                        Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, lineStrPoints(0), lineStrPoints(1), lineStrPoints(2), lineStrPoints(3), lineStrPoints(4), lineStrPoints(5))
                        m_outputColl.AddOutput "Edge1_", line
                        Set line = Nothing

                        Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, lineStrPoints(3), lineStrPoints(4), lineStrPoints(5), lineStrPoints(6), lineStrPoints(7), lineStrPoints(8))
                        m_outputColl.AddOutput "Edge2_", line
                        Set line = Nothing

                        Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, lineStrPoints(6), lineStrPoints(7), lineStrPoints(8), lineStrPoints(9), lineStrPoints(10), lineStrPoints(11))
                        m_outputColl.AddOutput "Edge3_", line
                        Set line = Nothing

                        Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, lineStrPoints(0), lineStrPoints(1), lineStrPoints(2), lineStrPoints(9), lineStrPoints(10), lineStrPoints(11))
                        m_outputColl.AddOutput "Edge4_", line
                        Set line = Nothing
                 
                      
                   'Remove  linestring
                    Set objLineString = oLineString
                    objLineString.Remove
                    Set oLineString = Nothing
                Next i
                Set ProjVector = Nothing
                Set geomFactory = Nothing
        End Select
    End If
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set CenterPos = Nothing
    Set objLineString = Nothing
         
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub


