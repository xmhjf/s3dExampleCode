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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:10:17 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CReferenceGeometry.cls
'   Author:         svsmylav
'   Creation Date:  Wednesday, 7th May 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on symbol of The DOW Chemical company.
'   "Dow Welded Trunnion Type Pipe Support Detail For Medium Diameter Lines 6 inch Through 14 inch Pipe"
'   The symbol consists of 17 Physical outputs. Of Which 10 plates, 2 saddles, One Reference plane, 2 lines and 2 points.
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
        
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parPipeCLtoTowerCL = arrayOfInputs(3)
    parTowertoGuideBase = arrayOfInputs(4)
    parPipeOD = arrayOfInputs(5)
    
    iOutput = 0
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
' Insert your code for output 13(Reference plane on Tower)
    ' create non-persistent circle to use for creating default surface ---
    Dim objPlane As IngrGeom3D.Plane3d
    Dim objCircle As IngrGeom3D.Circle3d
    Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            0, 0, 0, _
                            0, 0, -1, _
                            0.1 * parVesselDiameter / 2)

'   create persistent default surface plane - the plane can mate ---
    Set objPlane = oGeomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, objCircle)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPlane
    Set objPlane = Nothing
    Set objCircle = Nothing
        
    ' Insert your code for output 14(Vertical Line on Centreline of Pipe)
'The line is on vertical centreline of the pipe
    Dim dSupportHalfHeight As Double
    dSupportHalfHeight = 0.229 ' 9 inch
    Dim oVLine As IngrGeom3D.Line3d
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    stPoint.Set 0, parPipeCLtoTowerCL, dSupportHalfHeight
    enPoint.Set 0, parPipeCLtoTowerCL, -dSupportHalfHeight

    Set oVLine = oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                            stPoint.x, stPoint.y, stPoint.z, _
                            enPoint.x, enPoint.y, enPoint.z)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oVLine
    Set oVLine = Nothing

' Insert your code for output 15(Radial line on Tower)
    'The line is a radial line on the radius of the tower along the line connecting tower centreline to pipe centreline.
    Dim oRLine As IngrGeom3D.Line3d
    stPoint.Set 0, 0, 0
    enPoint.Set 0, parVesselDiameter / 2, 0

    Set oRLine = oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                            stPoint.x, stPoint.y, stPoint.z, _
                            enPoint.x, enPoint.y, enPoint.z)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oRLine
    Set oRLine = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

' Insert your code for output 16(Point on Tower)
    Dim oTowerPoint As IngrGeom3D.Point3d
    Set oTowerPoint = oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, 0, 0, 0)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oTowerPoint
    Set oTowerPoint = Nothing

' Insert your code for output 17(Point on Pipe CentreLine)
    Dim oPipePoint As IngrGeom3D.Point3d
    Set oPipePoint = oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, 0, parPipeCLtoTowerCL, 0)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oPipePoint
    Set oPipePoint = Nothing
    Set oGeomFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub

