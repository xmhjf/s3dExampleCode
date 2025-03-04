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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:19:41 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          svsmylav
'   Creation Date:  Thursday, Apr 18 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol is created using the following Outputs:
'       a) One Insulation aspect output,
'       b) One maintenance aspect output and
'       c) Two Physical aspect outputs: A default Surface,
'           ObjEquipment uses 'PlaceRevolution.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   April 1 2003    svsmylav        Single complex string revolution is replaced with
'                                                                  cylinder/cone primitives.
'   22.Jul.2003     ACM                              TR 44842: Added New Plane on the top of the Equipment.
'   29.Nov.2004     V6UpgradeSO             Made compatible with Smart Occurrence based Equipments
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'   06.May.2005     svsmylav        CR-76070: Removed the Nozzle related parameters and code to create nozzles.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  26.Sep.2006     RRK  TR-84724  If variables of type double are compared, appropriate CmpDblXXX functions are used.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private Const TOP_BODY = 1
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint
    
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
 
'   Input validation
    If CmpDblLessThanOrEqualTo(parBottomBodyHeight1, LINEAR_TOLERANCE) Or CmpDblLessThanOrEqualTo(parBottomBodyHeight2, LINEAR_TOLERANCE) Or CmpDblLessThanOrEqualTo(parBottomBodyDiameter, LINEAR_TOLERANCE) Then GoTo ErrorLabel
    If CmpDblGreaterthan(parShaftDiameter, LINEAR_TOLERANCE) And CmpDblGreaterthan(parShaftDiameter, parBottomBodyDiameter) Then GoTo ErrorLabel
    If CmpDblGreaterthan(parTopBodyDiameter, LINEAR_TOLERANCE) And CmpDblLessThan(parTopBodyDiameter, parBottomBodyDiameter) Then GoTo ErrorLabel

'   Insert your code for output 6(ObjDefSurface)
    Dim ObjDefSurface As IngrGeom3D.Plane3d
    Dim oCircle As IngrGeom3D.Circle3d
    Dim oComplexStr As New IngrGeom3D.ComplexString3d
    Dim baseCenter As New AutoMath.DPosition
    Dim normalVect As New AutoMath.DVector
    Dim circRadius  As Double
    
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory

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
    Call oComplexStr.AddCurve(oCircle, False)
    Set ObjDefSurface = geomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, _
                                                baseCenter.x, baseCenter.y, baseCenter.z, _
                                                normalVect.x, normalVect.y, normalVect.z)
    Call ObjDefSurface.AddBoundary(oComplexStr)
    
'   Set  the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefSurface
    Set ObjDefSurface = Nothing
   
'   Insert your code for output 7(TOP ObjDefSurface)
    Set oComplexStr = New ComplexString3d
            
    baseCenter.Set 0, 0, parBottomBodyHeight1 + parTaperBodyHeight + parTopBodyHeight
    circRadius = parTopBodyDiameter / 2
    
    normalVect.Set 0, 0, 1
    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                    baseCenter.x, baseCenter.y, baseCenter.z, _
                                                    normalVect.x, normalVect.y, normalVect.z, _
                                                    circRadius)
    Call oComplexStr.AddCurve(oCircle, False)
    Set ObjDefSurface = geomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, _
                                                        baseCenter.x, baseCenter.y, baseCenter.z, _
                                                        normalVect.x, normalVect.y, normalVect.z)
    Call ObjDefSurface.AddBoundary(oComplexStr)
    
'   Set  the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefSurface
    Set ObjDefSurface = Nothing
    Set baseCenter = Nothing
    Set normalVect = Nothing
    
    oComplexStr.RemoveCurve True
    Set oComplexStr = Nothing
    
    Dim ObjTmpcurves As IJDObject
    Set ObjTmpcurves = oCircle
    ObjTmpcurves.Remove
    Set ObjTmpcurves = Nothing
    Set oCircle = Nothing
    
    Set geomFactory = Nothing
    
' Insert your code for output 8(Equipment)
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
        Set ObjTopBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parTopBodyDiameter, True)
        
        'Set the output
        m_OutputColl.AddOutput "TopBody", ObjTopBody
        Set ObjTopBody = Nothing
    End If
    
'   Place Taper Body
    If CmpDblGreaterthan(parTopBodyDiameter, parBottomBodyDiameter) And CmpDblGreaterthan(parTaperBodyHeight, LINEAR_TOLERANCE) Then
        Dim ObjTaperBody As Object
        
        stPoint.Set 0, 0, parTaperBodyHeight + parBottomBodyHeight1

        enPoint.Set 0, 0, parBottomBodyHeight1
        Set ObjTaperBody = PlaceCone(m_OutputColl, stPoint, enPoint, parTopBodyDiameter / 2, parBottomBodyDiameter / 2, True)
        
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
    If CmpDblGreaterthan(parShaftLength, LINEAR_TOLERANCE) And CmpDblGreaterthan(parShaftDiameter, LINEAR_TOLERANCE) Then
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
    
'==========================================
'Construction of  Control Point
'==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, 0, 0, 0, 0.05)

    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    m_OutputColl.AddOutput "VerRotatingEquipmentControlPoint", m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

