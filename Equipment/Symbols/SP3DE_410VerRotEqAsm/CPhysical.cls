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
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         RUK
'   Creation Date:  Monday, Feb 19 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Thirteen Outputs:
'      In which One Default Surface is created at the bottom end of the Shaft.
'      The Circular Foundation Port is created at the bottom end of the Bottom Cylinder
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
    Set m_oSymGeomHelper = New SymbolServices
     
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double
    
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

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBottomBodyHeight1 = arrayOfInputs(2)     'P1
    parBottomBodyHeight2 = arrayOfInputs(3)     'P2
    parBottomBodyDiameter = arrayOfInputs(4)    'P3
    parShaftLength = arrayOfInputs(5)           'P4
    parShaftDiameter = arrayOfInputs(6)         'P5
    parTaperBodyHeight = arrayOfInputs(7)       'P6
    parTopBodyHeight = arrayOfInputs(8)         'P7
    parTopBodyDiameter = arrayOfInputs(9)       'P8
    parBundlePullingLength = arrayOfInputs(10)  'P9
    parInsulationThickness = arrayOfInputs(11)

    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_OutputColl

    'Origin is at DP/PP1
    'Input validation
    If CmpDblLessThanOrEqualTo(parBottomBodyHeight1, LINEAR_TOLERANCE) Or CmpDblLessThanOrEqualTo(parBottomBodyHeight2, LINEAR_TOLERANCE) Or _
            CmpDblLessThanOrEqualTo(parBottomBodyDiameter, LINEAR_TOLERANCE) Or CmpDblLessThanOrEqualTo(parShaftLength, LINEAR_TOLERANCE) Then GoTo ErrorLabel

    If CmpDblGreaterthan(parShaftDiameter, LINEAR_TOLERANCE) And CmpDblGreaterthan(parShaftDiameter, parBottomBodyDiameter) Then GoTo ErrorLabel
    If CmpDblGreaterthan(parTopBodyDiameter, LINEAR_TOLERANCE) And CmpDblLessThan(parTopBodyDiameter, parBottomBodyDiameter) Then GoTo ErrorLabel
    
    'Inser your code for creating the Surfaces and Edges
            '--Default Surface
            '--Surface at the Top
            '--Edges for the TopBody
            '--Edges for the BottomBody
            '--Edges for the Shaft
    Dim oCircle     As IngrGeom3D.Circle3d
    Dim baseCenter  As New AutoMath.DPosition
    Dim normalVect  As New AutoMath.DVector
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim circRadius  As Double
    
    'Create the Default Surface at the bottom end of the Shaft
    Dim ObjDefSurface As IngrGeom3D.Plane3d
    
    If CmpDblGreaterthan(parShaftLength, LINEAR_TOLERANCE) And CmpDblGreaterthan(parShaftDiameter, LINEAR_TOLERANCE) Then
        baseCenter.Set 0, 0, -parBottomBodyHeight2 - parShaftLength
        circRadius = parShaftDiameter / 2
    Else
        baseCenter.Set 0, 0, -parBottomBodyHeight2
        circRadius = parBottomBodyDiameter / 2
    End If
    normalVect.Set 0, 0, -1
    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                    baseCenter.x, baseCenter.y, baseCenter.z, _
                                                    normalVect.x, normalVect.y, normalVect.z, _
                                                    circRadius)
                                                            
    'Create persistent default surface plane - the plane can mate ---
    Set ObjDefSurface = geomFactory.Planes3d.CreateByOuterBdry _
                                        (m_OutputColl.ResourceManager, oCircle)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefSurface
    Set ObjDefSurface = Nothing
    
    'Insert your code for TopSurface
    Dim dEqpTopPointZ As Double
    dEqpTopPointZ = parBottomBodyHeight1
          
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
                                                        
    'Create persistent default surface plane - the plane can mate ---
    Set ObjSurface = geomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, oCircle)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSurface
    Set ObjSurface = Nothing
    
    'insert your code for edges
    'Edges for Top Body
     If CmpDblGreaterthan(parTopBodyHeight, LINEAR_TOLERANCE) And CmpDblGreaterthan(parTopBodyDiameter, LINEAR_TOLERANCE) Then
        'Top Edge
        Dim ObjEdge1 As Object
        Dim ObjEdge2 As Object
        Set ObjEdge1 = geomFactory.Circles3d.CreateByCenterNormalRadius(m_OutputColl.ResourceManager, _
                                                        baseCenter.x, baseCenter.y, baseCenter.z, _
                                                        normalVect.x, normalVect.y, normalVect.z, _
                                                        circRadius)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEdge1
        Set ObjEdge1 = Nothing
        
        'Bottom Edge
        baseCenter.z = parBottomBodyHeight1 + parTaperBodyHeight
        Set ObjEdge2 = geomFactory.Circles3d.CreateByCenterNormalRadius(m_OutputColl.ResourceManager, _
                                                        baseCenter.x, baseCenter.y, baseCenter.z, _
                                                        normalVect.x, normalVect.y, normalVect.z, _
                                                        circRadius)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEdge2
        Set ObjEdge2 = Nothing
    End If
    
    'Edges for "Bottom Body"
    'Top Edge
    Dim ObjEdge3 As Object
    Dim ObjEdge4 As Object
    baseCenter.z = parBottomBodyHeight1
    Set ObjEdge3 = geomFactory.Circles3d.CreateByCenterNormalRadius(m_OutputColl.ResourceManager, _
                                                        baseCenter.x, baseCenter.y, baseCenter.z, _
                                                        normalVect.x, normalVect.y, normalVect.z, _
                                                        parBottomBodyDiameter / 2)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEdge3
    Set ObjEdge3 = Nothing
    
    'Bootom Edge
    baseCenter.z = -parBottomBodyHeight2
    Set ObjEdge4 = geomFactory.Circles3d.CreateByCenterNormalRadius(m_OutputColl.ResourceManager, _
                                                    baseCenter.x, baseCenter.y, baseCenter.z, _
                                                    normalVect.x, normalVect.y, normalVect.z, _
                                                    parBottomBodyDiameter / 2)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEdge4
    Set ObjEdge4 = Nothing
    
    'Edges for "Shaft"
    If CmpDblGreaterthan(parShaftDiameter, LINEAR_TOLERANCE) Then
        'Top Edge
        Dim ObjEdge5 As Object
        Dim ObjEdge6 As Object
        baseCenter.z = -parBottomBodyHeight2
        Set ObjEdge5 = geomFactory.Circles3d.CreateByCenterNormalRadius(m_OutputColl.ResourceManager, _
                                                        baseCenter.x, baseCenter.y, baseCenter.z, _
                                                        normalVect.x, normalVect.y, normalVect.z, _
                                                        parShaftDiameter / 2)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEdge5
        Set ObjEdge5 = Nothing
        
        'Bottom Edge
        baseCenter.z = -(parBottomBodyHeight2 + parShaftLength)
        Set ObjEdge6 = geomFactory.Circles3d.CreateByCenterNormalRadius(m_OutputColl.ResourceManager, _
                                                        baseCenter.x, baseCenter.y, baseCenter.z, _
                                                        normalVect.x, normalVect.y, normalVect.z, _
                                                        parShaftDiameter / 2)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEdge6
        Set ObjEdge6 = Nothing
    End If
    'Removing the References
    Set baseCenter = Nothing
    Set normalVect = Nothing
    Dim ObjTmpcurves As IJDObject
    Set ObjTmpcurves = oCircle
    ObjTmpcurves.Remove
    Set ObjTmpcurves = Nothing
    Set oCircle = Nothing
    Set geomFactory = Nothing

    'Insert your code for Equipment
    'Notes:
        'The following variables need to be positive values greater than zero for
        'placing i) Top Body ii) Taper Body and iii) Shaft respectively:
        '    - For the Top Body placement, parTopBodyHeight and parTopBodyDiameter,
        '    - For the Taper Body placement, parTopBodyDiameter and parTaperBodyHeight,
        '    - For the Shaft placement, parShaftLength and parShaftDiameter.
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

    'Place Top Body
    If CmpDblGreaterthan(parTopBodyHeight, LINEAR_TOLERANCE) And CmpDblGreaterthan(parTopBodyDiameter, LINEAR_TOLERANCE) Then
        Dim ObjTopBody As Object
        
        stPoint.Set 0, 0, parTopBodyHeight + parTaperBodyHeight + parBottomBodyHeight1
        enPoint.Set 0, 0, parTaperBodyHeight + parBottomBodyHeight1
        Set ObjTopBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parTopBodyDiameter, False)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopBody
        Set ObjTopBody = Nothing
    End If

    'Place Taper Body
    If CmpDblGreaterthan(parTopBodyDiameter, LINEAR_TOLERANCE) And CmpDblGreaterthan(parTaperBodyHeight, LINEAR_TOLERANCE) Then
        Dim ObjTaperBody As Object
        
        stPoint.Set 0, 0, parTaperBodyHeight + parBottomBodyHeight1
        enPoint.Set 0, 0, parBottomBodyHeight1
        Set ObjTaperBody = PlaceCone(m_OutputColl, stPoint, enPoint, parTopBodyDiameter / 2, parBottomBodyDiameter / 2, False)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTaperBody
        Set ObjTaperBody = Nothing
    End If

    'Place Bottom Body
    Dim ObjBottomBody As Object
    
    stPoint.Set 0, 0, parBottomBodyHeight1
    enPoint.Set 0, 0, -parBottomBodyHeight2
    Set ObjBottomBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parBottomBodyDiameter, True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBottomBody
    Set ObjBottomBody = Nothing
    
    'Place Circular FoundationPort
    Dim ObjFoundationPort As IJEqpFoundationPort
    Dim NozzlePHFactory As NozzlePHFactory
    Set NozzlePHFactory = New NozzlePHFactory
    Dim dOrigin(0 To 2) As Double
    Dim dXaxis(0 To 2) As Double
    Dim dZaxis(0 To 2) As Double
    
    'The origin of the port is taken to be at the centre point of the support base.
    dOrigin(0) = 0
    dOrigin(1) = 0
    dOrigin(2) = -parBottomBodyHeight2
    
    dXaxis(0) = 0
    dXaxis(1) = 1
    dXaxis(2) = 0#
    
    dZaxis(0) = 0#
    dZaxis(1) = 0#
    dZaxis(2) = -1#

    Set ObjFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "STFndPort1", _
                                                    False, m_OutputColl.ResourceManager)

    Call ObjFoundationPort.PutCS(dOrigin(0), dOrigin(1), dOrigin(2), _
                            dXaxis(0), dXaxis(1), dXaxis(2), _
                            dZaxis(0), dZaxis(1), dZaxis(2))
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFoundationPort
    Set ObjFoundationPort = Nothing
    Set NozzlePHFactory = Nothing
    
    'Place shaft
    If CmpDblGreaterthan(parShaftDiameter, LINEAR_TOLERANCE) Then
        Dim ObjShaft As Object
        
        stPoint.Set 0, 0, -parBottomBodyHeight2
        enPoint.Set 0, 0, -parBottomBodyHeight2 - parShaftLength
        Set ObjShaft = PlaceCylinder(m_OutputColl, stPoint, enPoint, parShaftDiameter, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjShaft
        Set ObjShaft = Nothing
    End If
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

