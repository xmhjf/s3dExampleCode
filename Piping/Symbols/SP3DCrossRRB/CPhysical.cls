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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          BG
'   Creation Date:  Monday, Dec 31 2001
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------    -----                    ------------------
'   09.Jul.2003     SymbolTeam(India)  Copyright Information, Header  is added.
'   01.Aug.2006     KKK                TR: 80254 Replaced 'm_OutputColl.ResourceManager' with "Nothing'.
'   08.SEP.2006     KKC                DI-95670 Replace names with initials in all revision history sheets and symbols.
'   04.SEP.2007     dkl                TR-122010 Added the parameter 'Face4toCenter' used for Nozzle 4 creation.
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
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Integer
    Dim ObjMainNozwithLen As Object
    Dim ObjTaperCone As Object
    Dim ObjBranchNozLen As Object
    Dim ObjBranchNozLen1 As Object
    Dim ObjNozzle As Object
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
    
 ' Insert your code for output 1 (Main run Nozzle with Length)  Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Dim Port1Length As Double
    Port1Length = parFace1toCenter
    oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set ObjMainNozwithLen = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port1Length)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMainNozwithLen
    Set ObjMainNozwithLen = Nothing
    
' Place Nozzle 2  Insert your code for output 2(Nozzle in Main Line)
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2
    oPlacePoint.Set parFace2toCenter + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0
    Set ObjNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle
    Set ObjNozzle = Nothing
    
' Insert your code for output 2(Main run Reducer)
    Dim cirNormalX As Double, cirNormalY As Double, cirNormalZ As Double
    Dim objCircle1 As IngrGeom3D.Circle3d, objCircle2 As IngrGeom3D.Circle3d
    Dim cir2CenterX As Double, cir2CenterY As Double, cir2CenterZ As Double
    Dim cir1CenterX As Double, cir1CenterY As Double, cir1CenterZ As Double
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    cirNormalX = 1
    cirNormalY = 0
    cirNormalZ = 0
    cir1CenterX = 0
    cir1CenterY = 0
    cir1CenterZ = 0
    Set objCircle1 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                                                Nothing, cir1CenterX, _
                                                cir1CenterY, cir1CenterZ, cirNormalX, _
                                                cirNormalY, cirNormalZ, pipeDiam * 0.5)
'Create Circle2
    cir2CenterX = parFace2toCenter - flangeThick2
    cir2CenterY = 0
    cir2CenterZ = 0
    Set objCircle2 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                                                Nothing, cir2CenterX, _
                                                cir2CenterY, cir2CenterZ, cirNormalX, _
                                                cirNormalY, cirNormalZ, pipeDiam2 * 0.5)
'Create Ruled Surface for body
    Set ObjTaperCone = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                            objCircle1, objCircle2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTaperCone
    Set ObjTaperCone = Nothing
    Set objCircle1 = Nothing
    Set objCircle2 = Nothing

' Insert your code for output 3(Branch run Nozzle with Length)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set 0, parFace3toCenter + sptOffset - depth, 0
    oDir.Set 0, 1, 0
    Dim PortLength As Double
    PortLength = parFace3toCenter
    If CmpDblLessThan(PortLength, flangeThick) Then PortLength = flangeThick
    Set ObjBranchNozLen = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, oPlacePoint, PortLength)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBranchNozLen
    Set ObjBranchNozLen = Nothing
    
' Insert your code for output 4(Branch run Nozzle with Length)
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set 0, -dNozzle4Location - sptOffset + depth, 0
    oDir.Set 0, -1, 0
    PortLength = dNozzle4Location
    If CmpDblLessThan(PortLength, flangeThick) Then PortLength = flangeThick
    Set ObjBranchNozLen1 = CreateNozzleWithLength(4, oPartFclt, m_OutputColl, oDir, oPlacePoint, PortLength)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBranchNozLen1
    Set ObjBranchNozLen1 = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
