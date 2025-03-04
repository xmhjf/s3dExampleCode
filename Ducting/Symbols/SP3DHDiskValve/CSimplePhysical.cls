VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright(c)2005-06 Intergraph GmbH
'   All Rights Reserved
'
'  "SimplePhysical:".cls
'  Author:          GW
'   Creation Date:  Dienstag, Sep 20 2005
'  Description:
'      This class module is the place for user to implement graphical part of VBSymbol for this aspect
'      TODO - fill in header description information
'
'  Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     ---                 ------------------
'   03.Nov.2005     kkk      CR-87366  Create TROX HVAC parts and symbols
'   08.SEP.2006     KKC      DI-95670  Replace names with initials in all revision history sheets and symbols
'   11.Sept.2006    RRK      TR-104419 Changed the number of outputs defined
'                                                 Removed "Depth" as an input parameter
'                                                 Removed the Variables and functions that
'                                                 were not used
'                                                 Removed the modDiskValve.bas module
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "CSimplePhysical" 'Used for error messages
Private m_oGeomHelper As IJSymbolGeometryHelper

Private m_GeomFactory As IngrGeom3D.GeometryFactory

Private Const E_FAIL = &H80004005
Const NEGLIGIBLE_THICKNESS = 0.0001
Private PI       As Double

Private Const clngNPoints As Long = 10

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
      PI = 4 * Atn(1)
    Set m_oGeomHelper = New SymbolServices
    
    Set m_GeomFactory = New IngrGeom3D.GeometryFactory
    
    Exit Sub
    
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext

End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim CP As New AutoMath.DPosition
    Dim PortDirection As New AutoMath.DVector
    Dim RadialDirection As New AutoMath.DVector
    
    Dim iOutput     As Double
        
'    {<(InputParamDec)>}
    Dim Width As Double
    Dim blnOpposite As Boolean
    Dim CentralDiskDia As Double
    Dim ValveLength As Double
    Dim DiskPosition As Double
    Dim ValveDia As Double
    Dim SealingRingLength As Double
    Dim dblToInner As Double
    Dim dblToDown As Double
    Dim lngIndex As Long
    
    Dim Depth As Double
       
    lngIndex = 0
    
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    Width = arrayOfInputs(2)
    CentralDiskDia = arrayOfInputs(3)
    ValveLength = arrayOfInputs(4)
    DiskPosition = arrayOfInputs(5)
    ValveDia = arrayOfInputs(6)
    SealingRingLength = arrayOfInputs(7)

    Depth = Width
    dblToInner = 0.005
    dblToDown = 0.01
    
    iOutput = 0
        
' Place Nozzle 1
    Dim FlangeWidth As Double
    Dim NozzleFlangeThickness As Double
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim iDistribPort As GSCADNozzleEntities.IJDistribPort
    
    Dim EndPrep As Long
    Dim FlowDir As DistribFlow
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim PortDepth As Double
    Dim CptOffset As Double
    Dim pos As New AutoMath.DPosition
    Dim Dir As AutoMath.DVector
    Set Dir = New AutoMath.DVector
    Dim RadialDir As New AutoMath.DVector
    Dim CornerRadius As Double
        
'   Set HVAC nozzle parameters
    PortDepth = 0#

'   FlangeWidth and Flange thickness assumed to be negigible thickness, Nozzle length to be
'   greater than flange thickness in general.
    FlangeWidth = NEGLIGIBLE_THICKNESS
    NozzleFlangeThickness = NEGLIGIBLE_THICKNESS
    
    FlowDir = DistribFlow_UNDEFINED
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    EndPrep = 301
    CptOffset = 0
    CornerRadius = 0

    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(1, "SymbDefn", GSCADNozzleEntities.Round, EndPrep, _
                                            NozzleFlangeThickness, FlangeWidth, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "hvac1", PortDepth, CptOffset, False, m_OutputColl.ResourceManager)
    pos.Set -ValveLength / 2# - 0.035, 0, 0
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    Dir.Set -1, 0, 0
    iDistribPort.SetDirectionVector Dir
    
    Dir.Set 0, 1, 0
    iDistribPort.SetRadialOrient Dir
    
    Set iNozzle = oHvacNozzle
    iNozzle.Length = 0.035
                                                             
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    
'Place Output 2 (Luft)
    
    Dim objP(clngNPoints) As AutoMath.DPosition
    Dim i As Long
    
    For i = 0 To clngNPoints
       Set objP(i) = New AutoMath.DPosition
    Next i
     
    ' Outer bound
    Dim dbl8mm As Double
    dbl8mm = 0.008
    
    objP(0).Set -SealingRingLength, ValveDia * 0.5, 0#
    objP(1).Set 0, ValveDia * 0.5, 0#
    objP(2).Set 0, CentralDiskDia * 0.5, 0#
    objP(3).Set dbl8mm, CentralDiskDia * 0.5, 0#
    objP(4).Set dbl8mm, ValveDia * 0.5, 0#
    objP(5).Set 0, CentralDiskDia * 0.5 * 0.75, 0#
    objP(6).Set -DiskPosition, CentralDiskDia * 0.5 * 0.75, 0#
    
    ' build curve
    Dim complex As IngrGeom3D.ComplexString3d
    Set complex = buildCurve(objP, 7)
    
    ' rotate curve
    Dim objAxis As New AutoMath.DPosition
    Dim objCenter As New AutoMath.DPosition
    Dim dblAngle As Double
    
    dblAngle = 2 * PI
    objCenter.Set 0, 0, 0
    objAxis.Set 1#, 0, 0
    
    Call rotateCurve(m_OutputColl, _
                        arrayOfOutputs(), _
                        iOutput, _
                        complex, _
                        objAxis, _
                        objCenter, _
                        dblAngle, _
                        True)
    
    
    
    Set iDistribPort = Nothing

    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing

    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
    Exit Sub
    Resume
    
End Sub
Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
    Set m_GeomFactory = Nothing
End Sub
                    
Private Function buildCurve(objP() As AutoMath.DPosition, lngMax As Long) As IngrGeom3D.ComplexString3d

    Dim i As Long
    Dim oLine As IngrGeom3D.Line3d
    Dim iElements As IJElements
    
    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP(0).x, objP(0).y, objP(0).z, objP(1).x, objP(1).y, objP(1).z)
    Set iElements = New JObjectCollection ' IMSElements.DynElements
    iElements.Add oLine
    Set buildCurve = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
    Set iElements = Nothing
    
    For i = 2 To lngMax - 1
        Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, _
                objP(i - 1).x, objP(i - 1).y, objP(i - 1).z, objP(i).x, objP(i).y, objP(i).z)
        buildCurve.AddCurve oLine, True
    Next i

    Set oLine = Nothing
    Set iElements = Nothing
    
End Function

Private Sub rotateCurve(ByVal objOutputColl As Object, _
                        strName() As String, _
                        lngIndex As Double, _
                        complex As IngrGeom3D.ComplexString3d, _
                        objAxis As AutoMath.DPosition, _
                        objCenter As AutoMath.DPosition, _
                        dblAngle As Double, _
                        blnCapped As Boolean)
                    
    Dim objRevolution           As IngrGeom3D.Revolution3d
    
    Set objRevolution = m_GeomFactory.Revolutions3d.CreateByCurve(objOutputColl.ResourceManager, _
                                                    complex, _
                                                    objAxis.x, objAxis.y, objAxis.z, _
                                                    objCenter.x, objCenter.y, objCenter.z, _
                                                    dblAngle, blnCapped)
                                                    
    lngIndex = lngIndex + 1
    objOutputColl.AddOutput strName(lngIndex), objRevolution
    
    Set objRevolution = Nothing
    
End Sub



