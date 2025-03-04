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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author: VRK
'   Creation Date:  Friday, Oct 25 2006
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has One Output:
'           Insulated Flanged head
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'******************************************************************************
Option Explicit

Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Insulation:" 'Used for error messages
Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
    Set m_oSymGeomHelper = New SymbolServices
    
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub
    Private Sub Class_Terminate()
        Set m_oSymGeomHelper = Nothing
    End Sub
Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim ObjInsFLGDHeadbody1 As Object
    Dim ObjInsFLGDHeadbody2 As Object
    Dim iOutput     As Double
    Dim parVesselDiameter As Double
    Dim parEndHeadFlangedThickness1 As Double
    Dim parEndHeadFlangedThickness2 As Double
    Dim stPoint  As New AutoMath.DPosition
    Dim enPoint  As New AutoMath.DPosition
    Dim parInsulationThickness As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)                'P52
    parEndHeadFlangedThickness1 = arrayOfInputs(3)      'P50
    parEndHeadFlangedThickness2 = arrayOfInputs(4)      'P51
    parInsulationThickness = arrayOfInputs(5)
    
    m_oSymGeomHelper.OutputCollection = m_outputColl
    
    iOutput = 0
    parVesselDiameter = parVesselDiameter + 2 * parInsulationThickness

    'ObjInsFLGDHeadbody
    stPoint.Set 0, 0, 0
    enPoint.Set 0, 0, parEndHeadFlangedThickness1

    Set ObjInsFLGDHeadbody1 = PlaceCylinder(m_outputColl, stPoint, enPoint, parVesselDiameter, True)
    
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjInsFLGDHeadbody1
    Set ObjInsFLGDHeadbody1 = Nothing

    stPoint.Set 0, 0, 0
    enPoint.Set 0, 0, -parEndHeadFlangedThickness2

    Set ObjInsFLGDHeadbody2 = PlaceCylinder(m_outputColl, stPoint, enPoint, parVesselDiameter, True)
    
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjInsFLGDHeadbody2
    Set ObjInsFLGDHeadbody2 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

    




