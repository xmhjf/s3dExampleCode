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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:08:46 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CReferenceGeometry.cls
'   Author:         svsmylav
'   Creation Date:  Tuesday, May 13, 2003
'   Description:
'   This is a Clamped Trunnion Pipe Guide for small Lines through 6 inch pipe.
'   This symbol consists of 4 inputs and 21 outputs,It has two aspects one is Simple Physical
'   another is ReferenceGeometry.There are 15 Physical and 5 Reference Outputs..
'   This is a Dow symbol Equipment
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "ReferenceGeometry:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double
    
    Dim parVesselDiameter As Double
    Dim parPipeCLtoTowerCL As Double
    Dim parTowertoGuideBase As Double
    Dim parPipeOD As Double
    
    Dim ObjDefaultsurface As Object
    Dim ObjPipeAxisLine As IngrGeom3D.Line3d
    Dim ObjPipeAxisPoint As IngrGeom3D.Point3d
    Dim ObjVesselAxisLine As IngrGeom3D.Line3d
    Dim ObjVesselAxisPoint As IngrGeom3D.Point3d
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parPipeCLtoTowerCL = arrayOfInputs(3)
    parTowertoGuideBase = arrayOfInputs(4)
    parPipeOD = arrayOfInputs(5)
    
    iOutput = 0
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
   
' Insert your code for output 16(Reference plane on Tower)
    ' create non-persistent circle to use for creating default surface ---
    Dim objCircle As IngrGeom3D.Circle3d
    Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            0, 0, 0, _
                            0, 0, -1, _
                            0.1 * parVesselDiameter / 2)

'   create persistent default surface plane - the plane can mate ---
    Set ObjDefaultsurface = oGeomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, objCircle)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefaultsurface
    Set ObjDefaultsurface = Nothing
    Set objCircle = Nothing
    
' Insert your code for output 17( Reference line on Pipe)
'The line is on vertical centreline of the pipe
    Dim oStPoint As New AutoMath.DPosition
    Dim oEnPoint As New AutoMath.DPosition
    Dim dGuidePlateHeight As Double
    
    dGuidePlateHeight = 0.204 ''8 inch  (trapezoidal plate)
    
    oStPoint.Set 0, parPipeCLtoTowerCL, dGuidePlateHeight / 2
    oEnPoint.Set 0, parPipeCLtoTowerCL, -dGuidePlateHeight / 2
    
    Set ObjPipeAxisLine = oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                            oStPoint.x, oStPoint.y, oStPoint.z, _
                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPipeAxisLine
    Set ObjPipeAxisLine = Nothing

' Insert your code for output 18( Reference line on Tower)
    'The line is a radial line on the radius of the tower along the line connecting tower centreline to pipe centreline.
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0, parVesselDiameter / 2, 0
    
    Set ObjVesselAxisLine = oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                            oStPoint.x, oStPoint.y, oStPoint.z, _
                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVesselAxisLine
    Set ObjVesselAxisLine = Nothing
        
' Insert your code for output 19( Point at origin )
    Set ObjVesselAxisPoint = oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, 0, 0, 0)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVesselAxisPoint
    Set ObjVesselAxisPoint = Nothing
    
' Insert your code for output 20( Point at Pipe Axis )
    Set ObjPipeAxisPoint = oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, 0, _
                                                        parPipeCLtoTowerCL, 0)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPipeAxisPoint
    Set ObjPipeAxisPoint = Nothing
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oGeomFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
        Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

