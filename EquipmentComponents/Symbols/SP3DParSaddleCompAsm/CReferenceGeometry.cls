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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:09:46 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CReferenceGeometry.cls
'   Author:         svsmylav
'   Creation Date:  Friday, May 9, 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'   This is a Saddle Support for 1.2mts to 3 mts Dia Vessels.
'   This symbol consists of 13 inputs and 12 outputs  and It has two aspects one is Simple Physical
'   another is ReferenceGeometry. Among them 9 are physical and 3 are ReferenceGeometry.
'   This is a Dow Equipment symbol

Option Explicit


Private Const MODULE = "ReferenceGeometry:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim parVesselDiameter As Double
    Dim parSupportLength As Double
    Dim parSupportHeight As Double
    Dim parSupportWidth As Double
    Dim parCPtoSaddle As Double
    Dim parBasePlateThickness As Double
    Dim parBackingPlateThickness As Double
    Dim parBackingPlateWidth As Double
    Dim parStiffnerThickness As Double
    Dim parStiffnerThickness1 As Double
    Dim parStiffnerSpacing As Double
    Dim parXboltHole As Double
    Dim parYboltHole As Double
    
    Dim iOutput     As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parSupportLength = arrayOfInputs(3)
    parSupportHeight = arrayOfInputs(4)
    parSupportWidth = arrayOfInputs(5)
    parCPtoSaddle = arrayOfInputs(6)
    parBasePlateThickness = arrayOfInputs(7)
    parBackingPlateThickness = arrayOfInputs(8)
    parBackingPlateWidth = arrayOfInputs(9)
    parStiffnerThickness = arrayOfInputs(10)
    parStiffnerThickness1 = arrayOfInputs(11)
    parStiffnerSpacing = arrayOfInputs(12)
    parXboltHole = arrayOfInputs(13)
    parYboltHole = arrayOfInputs(14)
    
    iOutput = 0
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
   
' Insert your code for output 10(Reference plane on at Control Point of vessel)
    ' create non-persistent circle to use for creating default surface ---
    
    Dim ObjDefaultSurface As Object
    Dim objCircle As IngrGeom3D.Circle3d
    Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            0, 0, 0, _
                            1, 0, 0, _
                            0.1 * parVesselDiameter / 2)

'   create persistent default surface plane - the plane can mate ---
    Set ObjDefaultSurface = oGeomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, objCircle)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefaultSurface
    Set ObjDefaultSurface = Nothing
    Set objCircle = Nothing
    
    
' Insert your code for output 11( Reference line on the Vessel axis)
' The line passes through Vessel axis.

    Dim ObjLine As IngrGeom3D.Line3d
    Dim oStPoint As New AutoMath.DPosition
    Dim oEnPoint As New AutoMath.DPosition
    
    oStPoint.Set parCPtoSaddle - parSupportWidth / 2, 0, 0
    oEnPoint.Set parCPtoSaddle + parSupportWidth / 2, 0, 0
    
    Set ObjLine = oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, _
                            oStPoint.x, oStPoint.y, oStPoint.z, _
                            oEnPoint.x, oEnPoint.y, oEnPoint.z)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLine
    Set ObjLine = Nothing

' Insert your code for output 12( Point on Vessel axis )
    Dim ObjPoint As IngrGeom3D.Point3d

    Set ObjPoint = oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, 0, 0, 0)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPoint
    Set ObjPoint = Nothing
        
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oGeomFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub



