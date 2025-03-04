VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          BG
'   Creation Date:  Monday, Dec 31 2001
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------    -----                ------------------
'   09.Jul.2003     SymbolTeam(India) Copyright Information, Header  is added.
'   19.Aug.2005     svsmylav          TR-83739: Added check to ensure non-zero diameter and non-zero length for
'                                          cylinder to avoid PlaceCylinder function failure.
'   08.SEP.2006     KKC               DI-95670 Replace names with initials in all revision history sheets and symbols.
'   04.SEP.2007     dkl               TR-122010 Added the parameter 'Face4toCenter' used for Nozzle 4 creation.
'                                     Used CmpDblXXX functions where appropriate.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private Const NEGLIGIBLE_THICKNESS = 0.0001

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Integer
    Dim ObjPort1Insulation As Object
    Dim ObjPort2Insulation As Object
    Dim ObjPort3Insulation As Object
    Dim ObjPort4Insulation As Object
    Dim ObjMainBodyInsul As Object
    Dim ObjTaperConeInsul As Object
    Dim ObjBranchInsulation As Object
    Dim ObjBranchInsulation1 As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parInsulationThickness As Double
    Dim parFace4toCenter As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parFace3toCenter = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    
    'To ensure compatibility with previous versions of the symbol.
    Dim dNozzle4Location As Double
    dNozzle4Location = parFace3toCenter
    If UBound(arrayOfInputs) >= 6 Then
        parFace4toCenter = arrayOfInputs(6)
        dNozzle4Location = parFace4toCenter
    End If
    
    iOutput = 0

 ' Insert your code for output 6(Insulation for Port 1)
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim parInsulationDiameter As Double
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set -(parFace1toCenter - flangeThick - parInsulationThickness), 0, 0
    If CmpDblLessThanOrEqualTo(flangeThick, 0) Then _
        enPoint.Set -(parFace1toCenter - NEGLIGIBLE_THICKNESS), 0, 0
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    Set ObjPort1Insulation = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPort1Insulation
    Set ObjPort1Insulation = Nothing

' Insert your code for output 10(Insulation for the Main Body)
    stPoint.Set -(parFace1toCenter - flangeThick), 0, 0
    enPoint.Set 0, 0, 0
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    Set ObjMainBodyInsul = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMainBodyInsul
    Set ObjMainBodyInsul = Nothing

' Insert your code for output 7(Insulation for Port 2)
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2
    stPoint.Set parFace2toCenter, 0, 0
    enPoint.Set (parFace2toCenter - flangeThick2 - parInsulationThickness), 0, 0
    If CmpDblLessThanOrEqualTo(flangeThick2, 0) Then _
        enPoint.Set parFace2toCenter - NEGLIGIBLE_THICKNESS, 0, 0
    parInsulationDiameter = pipeDiam2 + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then parInsulationDiameter = flangeDiam2 + parInsulationThickness * 2
    Set ObjPort2Insulation = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPort2Insulation
    Set ObjPort2Insulation = Nothing
    
' Insert your code for output 11(Insulation for Taper Cone)
    Dim cirNormalX As Double, cirNormalY As Double, cirNormalZ As Double
    Dim objCircle1 As IngrGeom3D.Circle3d, objCircle2 As IngrGeom3D.Circle3d
    Dim cir2CenterX As Double, cir2CenterY As Double, cir2CenterZ As Double
    Dim cir1CenterX As Double, cir1CenterY As Double, cir1CenterZ As Double
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    pipeDiam = pipeDiam + parInsulationThickness * 2
    pipeDiam2 = pipeDiam2 + parInsulationThickness * 2
    cirNormalX = 1
    cirNormalY = 0
    cirNormalZ = 0
    cir1CenterX = 0
    cir1CenterY = 0
    cir1CenterZ = 0
    Set objCircle1 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                                                m_OutputColl.ResourceManager, cir1CenterX, _
                                                cir1CenterY, cir1CenterZ, cirNormalX, _
                                                cirNormalY, cirNormalZ, pipeDiam * 0.5)
'Create Circle2
    cir2CenterX = parFace2toCenter - flangeThick2
    cir2CenterY = 0
    cir2CenterZ = 0
    Set objCircle2 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                                                m_OutputColl.ResourceManager, cir2CenterX, _
                                                cir2CenterY, cir2CenterZ, cirNormalX, _
                                                cirNormalY, cirNormalZ, pipeDiam2 * 0.5)
'Create Ruled Surface for body
    Set ObjTaperConeInsul = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                            objCircle1, objCircle2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTaperConeInsul
    Set ObjTaperConeInsul = Nothing
    Set objCircle2 = Nothing
    Set objCircle1 = Nothing
    
 ' Insert your code for output 8(Insulation for Port 3)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    stPoint.Set 0, parFace3toCenter, 0
    enPoint.Set 0, (parFace3toCenter - flangeThick - parInsulationThickness), 0
    If CmpDblLessThanOrEqualTo(flangeThick, 0) Then _
        enPoint.Set 0, parFace3toCenter - NEGLIGIBLE_THICKNESS, 0
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    Set ObjPort3Insulation = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPort3Insulation
    Set ObjPort3Insulation = Nothing

' Insert your code for output 12(Insulation for Branch)
    stPoint.Set (parFace3toCenter - flangeThick), 0, 0
    enPoint.Set 0, 0, 0
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    Set ObjBranchInsulation = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBranchInsulation
    Set ObjBranchInsulation = Nothing
    
' Insert your code for output 9(Insulation for Port 4)
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    stPoint.Set 0, -dNozzle4Location, 0
    enPoint.Set 0, -(dNozzle4Location - flangeThick - parInsulationThickness), 0
    If CmpDblLessThanOrEqualTo(flangeThick, 0) Then _
        enPoint.Set 0, -(dNozzle4Location - NEGLIGIBLE_THICKNESS), 0
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    Set ObjPort4Insulation = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPort4Insulation
    Set ObjPort4Insulation = Nothing
        
 ' Insert your code for output 13(Insulation for Branch 4)
    stPoint.Set -(dNozzle4Location - flangeThick), 0, 0
    enPoint.Set 0, 0, 0
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    Set ObjBranchInsulation1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBranchInsulation1
    Set ObjBranchInsulation1 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
