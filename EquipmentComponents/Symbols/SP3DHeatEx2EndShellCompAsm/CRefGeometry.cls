VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRefGeometry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CRefGeometry.cls
'   Author:          ACM
'   Creation Date:  Tuesday, May 16, 2003
'   Description:
'   This is a Shell and Tube Exchanger shell Body.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   23.Aug.2003  SymbolTeam(India)       Updated Error Handling Code
'   20.May.2005    MS                    CR-76070: Modified the nozzle creation code(as members),added datum points.
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

''  This is an Exchanger Shell Body Symbol.
'   This symbol is Fully parametric symbol, consists of 77 inputs and 10 outputs.It has 2 aspects one is Simple Physical
'   another is ReferenceGeometry. Among them 8 are physical and 2 is ReferenceGeometry. It has Five nozzles and all
'   are fully parametric.
'   This is a Dow Equipment E-161 Exchanger Shell symbol.

Option Explicit

Private Const MODULE = "RefGeometry:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim parVesselDiameter As Double
    Dim parVesselHeight As Double
    Dim parCPtoFace As Double
    Dim parFaceDiameter As Double
    Dim parFaceThickness As Double
    Dim parInsulationThickness As Double
    
    Dim iOutput     As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parVesselHeight = arrayOfInputs(3)
    parCPtoFace = arrayOfInputs(4)
    parFaceDiameter = arrayOfInputs(5)
    parFaceThickness = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)
     
    iOutput = 0
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

' Insert your code for output 9( Point at origin )
    Dim ObjExchangerPoint As IngrGeom3D.Point3d
    
    Set ObjExchangerPoint = oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, 0, 0, 0)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerPoint
    Set ObjExchangerPoint = Nothing
    
'   Create DefaultSurface - for output 10(a Circular shape is chosen)
    Dim ObjDefaultSurface As Object
    Dim objCircle As IngrGeom3D.Circle3d
    Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            0, 0, 0, _
                            0, 0, 1, _
                            0.1 * parVesselDiameter / 2)

'   create persistent default surface plane - the plane can mate ---
    Set ObjDefaultSurface = oGeomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, objCircle)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefaultSurface
    Set ObjDefaultSurface = Nothing
    Set objCircle = Nothing
    Set oGeomFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
        Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
