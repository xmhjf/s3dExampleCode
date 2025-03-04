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
'   Copyright (c) 2004-07, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          MS
'   Creation Date:  Tuesday, Sept 14 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    Symbol Model No.: V58, Page No. D-28 of PDS Piping Component Data Reference Guide.
'    This Symbol is created with Four Outputs, all of which are physical aspect outputs.
'       BodyCone1, BodyCone2 using 'PlaceCone' function,
'       Strainer body using 'PlaceCylinder',
'       One ObjNozzle object by using 'CreateNozzle' function
'   The following Part data basis cases are addressed for the parameters specified:
'   Case A (Part data Basis value -13): Asymmetrical Face-to-Center dimension basis
'   Case B (Part data Basis value -379): Foot Valve, specified by overall length and strainer length

'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  13.AUG.2007     MA   CR-125108 Implemented part data basis for values 13 and 379.
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
        
    Dim iOutput     As Double
    Dim ObjBodyCone1  As Object
    Dim ObjBodyCone2 As Object
    Dim ObjStrainerBody As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parOverallLength As Double
    Dim parStrainerLength As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFace1toCenter = arrayOfInputs(2)
'    parFace2toCenter = arrayOfInputs(3)
'    parOverallLength = arrayOfInputs(4)
'    parStrainerLength = arrayOfInputs(5)

    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim oPlacePoint As AutoMath.DPosition
    Set oPlacePoint = New AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Set oDir = New AutoMath.DVector
    
    iOutput = 0
 
' This symbol is based on the following Part data basis values that govern its geometry
'   (Part data Basis value -13): Asymmetrical Face-to-Center dimension basis
'   (Part data Basis value -379): Foot Valve, specified by overall length and strainer length
 
 'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    Select Case lPartDataBasis
    
    Case Is <= 1, 13
        
        parFace1toCenter = arrayOfInputs(2)
        parFace2toCenter = arrayOfInputs(3)
     
     ' Insert your code for output 2(Cone - Port1 Side)
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
        stPoint.Set -parFace1toCenter + flangeThick, 0, 0
        enPoint.Set 0, 0, 0.000001
        
        Set ObjBodyCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
    
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1
        Set ObjBodyCone1 = Nothing
        
     ' Insert your code for output 3(Cone - Port2 Side)
        stPoint.Set parFace1toCenter, 0, 0
        enPoint.Set 0, 0, 0.000001
        Set ObjBodyCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
    
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
        Set ObjBodyCone2 = Nothing
        
     ' Insert your code for output 4(Hose Connector Body)
        stPoint.Set parFace2toCenter, 0, 0
        enPoint.Set parFace1toCenter, 0, 0
        Set ObjStrainerBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam, True)
            
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStrainerBody
        Set ObjStrainerBody = Nothing
        Set stPoint = Nothing
        Set enPoint = Nothing
        
        oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
        oDir.Set -1, 0, 0
    
    Case 379
        
        parOverallLength = arrayOfInputs(4)
        parStrainerLength = arrayOfInputs(5)
        
     ' Insert your code for output 2(Cone - Port1 Side)
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
        stPoint.Set -(parOverallLength - parStrainerLength) / 2 + flangeThick, 0, 0
        enPoint.Set 0, 0, 0
        
        Set ObjBodyCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
       
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1
        Set ObjBodyCone1 = Nothing
    
     ' Insert your code for output 3(Cone - Port2 Side)
        stPoint.Set (parOverallLength - parStrainerLength) / 2, 0, 0
        enPoint.Set 0, 0, 0
    
        Set ObjBodyCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
        
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
        Set ObjBodyCone2 = Nothing
    
     ' Insert your code for output 4(Hose Connector Body)
        stPoint.Set (parOverallLength - parStrainerLength) / 2 + parStrainerLength, 0, 0
        enPoint.Set (parOverallLength - parStrainerLength) / 2, 0, 0
        
        Set ObjStrainerBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam, True)
        
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStrainerBody
        Set ObjStrainerBody = Nothing
        Set stPoint = Nothing
        Set enPoint = Nothing
    
        oPlacePoint.Set -(parOverallLength - parStrainerLength) / 2 - sptOffset + depth, 0, 0
        oDir.Set -1, 0, 0
    
    Case Else
        GoTo ErrorLabel:
        
    End Select
    

' Place Nozzle 1
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
        
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
