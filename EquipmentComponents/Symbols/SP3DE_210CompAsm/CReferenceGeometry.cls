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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:17:29 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004-06, Intergraph Corporation. All rights reserved.
'
'   CReferenceGeometry.cls
'   Author:          MS
'   Creation Date: Wednesday, Feb 18 2004
'   Description:
'    The symbol is prepared from PDS Equipment guide E210. The symbol consists of 3 Physical outputs. The skirt is variable output.
'
'   Change History:
'   dd.mmm.yyyy     who                  change description
'   -----------    -----                 ------------------
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   11.Jul.2006     kkc       DI 95670-Replaced names with initials in the revision history.
'   26.Sep.2006     dkl       TR-84724 If variables of type double are compared, appropriate CmpDblXXX functions are used.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "ReferenceGeometry:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint
  
    Dim iOutput     As Double
    
    Dim parVesselHeight As Double
    Dim parVesselDiameter As Double
    Dim parSkirtTopDiameter As Double
    Dim parSkirtBottomDiameter As Double
    Dim parSkirtHeight As Double
    Dim parVesselStartHeight As Double
    Dim parInsulationThickness  As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselHeight = arrayOfInputs(2)
    parVesselDiameter = arrayOfInputs(3)
    parSkirtTopDiameter = arrayOfInputs(4)
    parSkirtBottomDiameter = arrayOfInputs(5)
    parSkirtHeight = arrayOfInputs(6)
    parVesselStartHeight = arrayOfInputs(7)
    parInsulationThickness = arrayOfInputs(8)
   
    iOutput = 0
        
' Origin is at Mid of Bottom plane of Skirt.

'Place Bottom Default Surface
    Dim geomFactory As GeometryFactory
    Set geomFactory = New GeometryFactory
    Dim ObjBotDefaultSurface As IngrGeom3D.Plane3d
    Dim oCircle As IngrGeom3D.Circle3d
    Dim xValue As Double
    Dim yValue As Double
    Dim zValue As Double
    Dim normalX As Double
    Dim normalY As Double
    Dim normalZ As Double
    xValue = 0
    yValue = 0
    zValue = 0

    normalX = 0
    normalY = 0
    normalZ = -1
    Dim dDefaultSurfaceDiameter As Double
    
    If (CmpDblGreaterthan(parVesselStartHeight, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSkirtTopDiameter, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSkirtBottomDiameter, LINEAR_TOLERANCE)) Then
        dDefaultSurfaceDiameter = parSkirtBottomDiameter
    Else
        dDefaultSurfaceDiameter = 0.25 * parVesselDiameter
   End If
    
    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                xValue, yValue, zValue, _
                                                normalX, normalY, normalZ, _
                                                dDefaultSurfaceDiameter / 2)
    Set ObjBotDefaultSurface = geomFactory.Planes3d.CreateByOuterBdry(m_OutputColl.ResourceManager, oCircle)

    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBotDefaultSurface
    Set ObjBotDefaultSurface = Nothing
    Set oCircle = Nothing
    Set geomFactory = Nothing

'==========================================
'Construction of  Control Point
'==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, 0, 0, 0, 0.02, , , False)
    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
    
End Sub
