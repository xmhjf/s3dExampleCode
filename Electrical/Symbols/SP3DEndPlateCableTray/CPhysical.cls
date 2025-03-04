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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          BG
'   Creation Date:  Friday, Mar 15 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'Tray End Plate (functions as a plug, a blind flange) is drawn as a plane at 0,0,0
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   07.May.2008     dkl         CR-141967 Updated the symbol to address insertion depth.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

'''

End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim ActualWidth As Double
    Dim ActualDepth As Double
    Dim oPort1 As New AutoMath.DPosition 'Port 1 center point
    Dim HD              As Double
    Dim HW              As Double
    Dim iOutput     As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    
    iOutput = 0
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Dim oTrayPart As IJCableTrayPart
    Set oTrayPart = oPartFclt
    
    'Insertion depth is not applicable now to this symbol. It is being implemented for consistency/future part data basis cases.
    Dim dInsertionDepth As Double
    'Resuming to next line on error to ensure functioning in V7 Service packs.
    On Error Resume Next
    dInsertionDepth = oTrayPart.InsertionDepth
    On Error GoTo ErrorLabel
    'variable for relocating the port considering insertion depth.
    Dim oPortLocation As AutoMath.DPosition
    Set oPortLocation = New AutoMath.DPosition
    
    Call RetrieveCableTrayPortProperties(1, oPartFclt, ActualWidth, ActualDepth)
    HW = ActualWidth / 2
    HD = ActualDepth / 2
    oPort1.Set 0, 0, 0
    
    'Points  of Cable Tray End Plate
    Dim VertiPlanePts(0 To 11) As Double
    Dim ObjEndPlate As IngrGeom3D.Plane3d
    
    VertiPlanePts(0) = oPort1.x
    VertiPlanePts(1) = oPort1.y - HW
    VertiPlanePts(2) = oPort1.z - HD
    
    VertiPlanePts(3) = oPort1.x
    VertiPlanePts(4) = oPort1.y - HW
    VertiPlanePts(5) = oPort1.z + HD
    
    VertiPlanePts(6) = oPort1.x
    VertiPlanePts(7) = oPort1.y + HW
    VertiPlanePts(8) = oPort1.z + HD
    
    VertiPlanePts(9) = oPort1.x
    VertiPlanePts(10) = oPort1.y + HW
    VertiPlanePts(11) = oPort1.z - HD
 ' Insert your code for output 1(Tray End Plate)
     Set ObjEndPlate = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, VertiPlanePts)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEndPlate
    Set ObjEndPlate = Nothing
    
' Place Nozzle 1
    'Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim oRadialOrient As AutoMath.DVector
    Dim objCableTrayPort   As GSCADNozzleEntities.IJCableTrayPortOcc

    'Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    
    oDir.Set -1, 0, 0
    oRadialOrient.Set 0, 0, 1
    oPortLocation.Set oPort1.x - dInsertionDepth * oDir.x, oPort1.y - dInsertionDepth * oDir.y, oPort1.z - dInsertionDepth * oDir.z

    Set objCableTrayPort = CreateCableTrayPort(oPartFclt, 1, oPortLocation, oDir, oRadialOrient, m_OutputColl)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCableTrayPort
    Set objCableTrayPort = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPort1 = Nothing
    Set oGeomFactory = Nothing
    Set oPortLocation = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
    

