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
'
'   Copyright (c) 2004-08, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         BBA
'   Creation Date:  Tuesday, October 26
'   Description:
'   PDS PullBox is taken as reference for preparing this symbol.
'   The symbol is created using Two outputs:
'   Output 1 is prepared using 'PlaceBox' function,
'   Outputs 2 to 7 are created using 'CreateConduitNozzle' function.
'   This symbol is based on the following Part data basis value that governs its geometry,
'   1)Single gang box, dead end, specified by length, height, width and extension
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   13.Jan.2005     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   19.Jan.2005     svsmylav   CR-69202     Pullbox is removed from the delivered electrical equipments and added it to conduit fittings,
'                                           So undone the code changes needed for Equipment based on Smart-Occurrence.
'   08.SEP.2006     KKC        DI-95670     Replace names with initials in all revision history sheets and symbols
'   22.MAY.2008     PK         CR-93287     Enhanced the pull box symbol implementing partdatabasis 9034
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' This class module is the place for user to implement graphical part of VBSymbol for this aspect

Option Explicit
Private Const COVERHEIGHT = 0.002    'Assumed height of cover
Private Const MODULE = "Physical:"   'Used for error messages

Private Sub Class_Initialize()

'''

End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim ConduitOD       As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJConduitPortOcc

    Dim iOutput         As Double
    Dim ObjPullBoxBody  As Object
    
    Dim parPullBoxLength    As Double
    Dim parPullBoxWidth     As Double
    Dim parPullBoxHeight    As Double

    Dim parLength       As Double
    Dim parWidth        As Double
    Dim parHeight       As Double
    Dim parExtension    As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parPullBoxLength = arrayOfInputs(2)
    'parPullBoxWidth = arrayOfInputs(3)
    'parPullBoxHeight = arrayOfInputs(4)
    'parLength = arrayOfInputs(5)
    'parWidth = arrayOfInputs(6)
    'parHeight = arrayOfInputs(7)
    'parExtension = arrayOfInputs(8)
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    If lPartDataBasis <= 1 Then
         parPullBoxLength = arrayOfInputs(2)
         parPullBoxWidth = arrayOfInputs(3)
         parPullBoxHeight = arrayOfInputs(4)
    ElseIf lPartDataBasis = 9034 Then
         parLength = arrayOfInputs(5)
         parWidth = arrayOfInputs(6)
         parHeight = arrayOfInputs(7)
         parExtension = arrayOfInputs(8)
    End If

    iOutput = 0
    Dim stPoint   As AutoMath.DPosition
    Dim enPoint   As AutoMath.DPosition
    
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
    If lPartDataBasis <= 1 Then
    
        'Insert your code for output 1(PullBoxBody)
        stPoint.Set -parPullBoxLength / 2, _
                    -parPullBoxWidth / 2, _
                    -parPullBoxHeight / 2
        enPoint.Set parPullBoxLength / 2, _
                    parPullBoxWidth / 2, _
                    parPullBoxHeight / 2
        Set ObjPullBoxBody = PlaceBox(m_OutputColl, stPoint, enPoint)
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPullBoxBody
        Set ObjPullBoxBody = Nothing
        Set stPoint = Nothing
        Set enPoint = Nothing
         
        'Place Nozzle 1
        RetrieveParameters 1, oPartFclt, m_OutputColl, ConduitOD, flangeThick, flangeDiam, sptOffset, depth
    
        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector
        oPlacePoint.Set -parPullBoxLength / 2 - sptOffset + depth, 0, 0
        oDir.Set -1, 0, 0
        Set objNozzle = CreateConduitNozzle(oPlacePoint, oDir, m_OutputColl, oPartFclt, 1)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        'Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, ConduitOD, flangeThick, flangeDiam, sptOffset, depth
        oPlacePoint.Set parPullBoxLength / 2 + sptOffset - depth, 0, 0
        oDir.Set 1, 0, 0
        Set objNozzle = CreateConduitNozzle(oPlacePoint, oDir, m_OutputColl, oPartFclt, 2)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
    
        'Place Nozzle 3
        RetrieveParameters 3, oPartFclt, m_OutputColl, ConduitOD, flangeThick, flangeDiam, sptOffset, depth
        oPlacePoint.Set 0, -parPullBoxWidth / 2 - sptOffset + depth, 0
        oDir.Set 0, -1, 0
        Set objNozzle = CreateConduitNozzle(oPlacePoint, oDir, m_OutputColl, oPartFclt, 3)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        'Place Nozzle 4
        RetrieveParameters 4, oPartFclt, m_OutputColl, ConduitOD, flangeThick, flangeDiam, sptOffset, depth
        oPlacePoint.Set 0, parPullBoxWidth / 2 + sptOffset - depth, 0
        oDir.Set 0, 1, 0
        Set objNozzle = CreateConduitNozzle(oPlacePoint, oDir, m_OutputColl, oPartFclt, 4)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        'Place Nozzle 5
        RetrieveParameters 5, oPartFclt, m_OutputColl, ConduitOD, flangeThick, flangeDiam, sptOffset, depth
        oPlacePoint.Set 0, 0, -parPullBoxHeight / 2 - sptOffset + depth
        oDir.Set 0, 0, -1
        Set objNozzle = CreateConduitNozzle(oPlacePoint, oDir, m_OutputColl, oPartFclt, 5)
        
        'Set the output
        m_OutputColl.AddOutput "Nozzle5", objNozzle
        Set objNozzle = Nothing
        
        'Place Nozzle 6
        RetrieveParameters 6, oPartFclt, m_OutputColl, ConduitOD, flangeThick, flangeDiam, sptOffset, depth
        oPlacePoint.Set 0, 0, parPullBoxHeight / 2 + sptOffset - depth
        oDir.Set 0, 0, 1
        Set objNozzle = CreateConduitNozzle(oPlacePoint, oDir, m_OutputColl, oPartFclt, 6)
        
        'Set the output
        m_OutputColl.AddOutput "Nozzle6", objNozzle
        Set objNozzle = Nothing
        Set oPlacePoint = Nothing
        Set oDir = Nothing
        
    ElseIf lPartDataBasis = 9034 Then
    
        Dim dCylDiam As Double
        dCylDiam = 0.7 * parHeight
        
        If CmpDblLessThanOrEqualTo(dCylDiam, ConduitOD) Then
            dCylDiam = 1.1 * ConduitOD
        End If
 
        'Create the Gang Box
        Dim oBox As Object
        stPoint.Set parExtension, -(0.5 * dCylDiam - 0.2 * parHeight), -0.5 * parWidth
        enPoint.Set parExtension + parLength, parHeight - 0.5 * dCylDiam, 0.5 * parWidth
        
        iOutput = iOutput + 1
        Set oBox = PlaceBox(m_OutputColl, stPoint, enPoint)
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
        Set oBox = Nothing
        
        'Create the Base
        Dim oBase As Object
        stPoint.Set 0, -0.5 * dCylDiam, -0.5 * parWidth
        enPoint.Set parExtension + parLength, -(0.5 * dCylDiam - 0.2 * parHeight), 0.5 * parWidth
        
        iOutput = iOutput + 1
        Set oBase = PlaceBox(m_OutputColl, stPoint, enPoint)
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBase
        Set oBase = Nothing

        'Create the Nozzle Shape
        Dim oNCylinder  As Object
        stPoint.Set 0, 0, 0
        enPoint.Set parExtension, 0, 0
        
        iOutput = iOutput + 1
        Set oNCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, dCylDiam, True)
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNCylinder
        Set oNCylinder = Nothing

        'Create the cover
        Dim oCover  As Object
        stPoint.Set parExtension + 0.1 * parLength, parHeight - 0.5 * dCylDiam, -0.4 * parWidth
        enPoint.Set parExtension + 0.9 * parLength, parHeight - 0.5 * dCylDiam + COVERHEIGHT, 0.4 * parWidth
        
        iOutput = iOutput + 1
        Set oCover = PlaceBox(m_OutputColl, stPoint, enPoint)
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oCover
        Set oCover = Nothing

        'Place Nozzle 1
        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector
        
        RetrieveParameters 1, oPartFclt, m_OutputColl, ConduitOD, flangeThick, flangeDiam, sptOffset, depth
        
        oPlacePoint.Set -sptOffset + depth, 0, 0
        oDir.Set -1, 0, 0
        
        Set objNozzle = CreateConduitNozzle(oPlacePoint, oDir, m_OutputColl, oPartFclt, 1)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        Set oDir = Nothing
        Set oPlacePoint = Nothing
    End If
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub


