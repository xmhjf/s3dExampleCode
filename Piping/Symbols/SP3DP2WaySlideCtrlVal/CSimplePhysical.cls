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
'**********************************************************************************
'   Copyright (c) 2004-08, Intergraph Corporation. All rights reserved.
'
'   File:   SimplePhysical.cls
'   Author: svsmylav
'   Date:   Tuesday, Sep 21 2004
'
'   Description:
'      This class module is the place for user to implement graphical part of VBSymbol for this aspect
'       Symbol Model No. is:  I20AZ Page No. 360 of PDS Piping Component Data Reference Guide pcd.pdf.
'
'  Change History:
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  12.May.2008     KKC  TR-139585  Placing tag inst using "SP3DP2WaySlideCtrlVal" throws “Error updating symbol”
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI As Double
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
  PI = 4 * Atn(1)
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorHandler
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double

    Dim parFacetoFace As Double
    Dim parFace2toCenter As Double
    Dim parBranchLength As Double
    Dim parProjectionLength As Double
    Dim parBranchHeight As Double
    Dim parProjectionHeight As Double
    Dim parProjectionWidth As Double
    Dim parBranchWidth As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)  'P2 + P3 as per EDEN
    parBranchLength = arrayOfInputs(3)   'P4
    parProjectionLength = arrayOfInputs(4)   'P5
    parBranchHeight = arrayOfInputs(5)   'P6
    parProjectionHeight = arrayOfInputs(6)   'P7
    parProjectionWidth = arrayOfInputs(7)    'P8
    parBranchWidth = arrayOfInputs(8)    'P9
    
    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
   
    ' Place the geometry for the Branch
    oStPoint.Set -parBranchLength / 2, 0, -parBranchWidth / 2
    oEnPoint.Set parBranchLength / 2, parBranchHeight - parProjectionHeight, parBranchWidth / 2
     
    Dim ObjBranch As Object
    Set ObjBranch = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    ' Set the output
    '   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBranch
    Set ObjBranch = Nothing
    
    ' Place the geometry for the Branch Projection
    oStPoint.Set -parProjectionLength / 2, parBranchHeight - parProjectionHeight, -parProjectionWidth / 2
    oEnPoint.Set parProjectionLength / 2, parBranchHeight, parProjectionWidth / 2
    Dim ObjBranchProjection As Object
    Set ObjBranchProjection = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    ' Set the output
    '   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBranchProjection
    Set ObjBranchProjection = Nothing
    
    
    ' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parFacetoFace / 2 - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0

    Dim dNozzleLength As Double

    dNozzleLength = parFacetoFace - flangeThick
    If dNozzleLength < flangeThick Then dNozzleLength = flangeThick
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub

