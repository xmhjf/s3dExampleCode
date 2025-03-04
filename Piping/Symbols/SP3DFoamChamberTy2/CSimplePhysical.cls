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
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, Sep 16 2004
'   Description:
'       This is Foam Chamber Type 2 symbol, Symbol Model No. is: FS26B Page No. 336 of PDS Piping
'       Component Data Reference Guide pcd.pdf. Symbol consists of six physical outputs and one of the outputs
'       is a nozzle.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CSimplePhysical" 'Used for error messages
Private m_oGeomHelper As IJSymbolGeometryHelper


Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo ErrorHandler
    
    Set m_oGeomHelper = New SymbolServices
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs() As Variant, ByRef arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorHandler
    
    Dim oPart As IJDPart
    Dim oPipeComponent As IJDPipeComponent
     
    Dim parFace1toCenter As Double
    Dim parFoamChamberWidth As Double
    Dim parFoamChamberLength1 As Double
    Dim parFoamChamberHeight1 As Double
    Dim parFoamChamberHeight2 As Double
    Dim parFoamChamberHeight3 As Double
    Dim parFoamChamberCyltoCenter As Double
    Dim parFoamChamberInletWidth As Double
    Dim parFoamChamberInletHeight1 As Double
    Dim parFoamChamberInletHeight2 As Double
    Dim parFoamChamberCylDia1 As Double
    Dim parFoamChamberCylThickness As Double
    Dim parFoamChamberCylDia2 As Double

' Inputs
    Set oPart = arrayOfInputs(1)
    Set oPipeComponent = oPart
    parFace1toCenter = arrayOfInputs(2)          'D1
    parFoamChamberWidth = arrayOfInputs(3)       'D2
    parFoamChamberLength1 = arrayOfInputs(4)     'D3
    parFoamChamberHeight1 = arrayOfInputs(5)     'D4
    parFoamChamberHeight2 = arrayOfInputs(6)     'D5
    parFoamChamberHeight3 = arrayOfInputs(7)     'D6
    parFoamChamberCyltoCenter = arrayOfInputs(8) 'D7
    parFoamChamberInletWidth = arrayOfInputs(9)  'D8
    parFoamChamberInletHeight1 = arrayOfInputs(10)   'D9
    parFoamChamberInletHeight2 = arrayOfInputs(11)   'D10
    parFoamChamberCylDia1 = arrayOfInputs(12)        'D11
    parFoamChamberCylThickness = arrayOfInputs(13)   'D12
    parFoamChamberCylDia2 = arrayOfInputs(14)    'D13

    Dim oNozzleData() As NozzleData

    NozzleInitialize oPart, oNozzleData
   
    Dim oNozzleCol As Collection
    Set oNozzleCol = New Collection
    
    Dim oOutputCol As Collection
    Set oOutputCol = New Collection
    
    Dim oT4x4Current As IJDT4x4
    Set oT4x4Current = New DT4x4
    
    Dim oT4x4Temp As IJDT4x4
    Set oT4x4Temp = New DT4x4
    
    Dim oGeometryFactory As GeometryFactory
    Set oGeometryFactory = New GeometryFactory

    Dim oOriginPos As DPosition
    Set oOriginPos = New DPosition
    oOriginPos.Set 0, 0, 0
    
'**********************************************************************************
'**********************************************************************************
'   BEGIN
'**********************************************************************************
'**********************************************************************************

        NozzleDefineDirectionVector oNozzleData, NozzleGeometryType_SingleCP
        

        Call NozzlePlace(m_OutputColl, oPart, oNozzleData, NozzleId_CP1, oOriginPos, oT4x4Current, oNozzleCol, oOutputCol)

'        MoveAlongAxis oT4x4Current, AxisDirection_Primary, oNozzleData(1).dCptOffset
        
        Dim len_1 As Double
        len_1 = parFace1toCenter - oNozzleData(1).dFlangeThickness - parFoamChamberLength1

        Dim oTrapezoid1 As Object
        Dim oTrapezoidBaseCenterPos1 As DPosition
        Set oTrapezoidBaseCenterPos1 = New DPosition
        oTrapezoidBaseCenterPos1.Set oOriginPos.x + oNozzleData(1).dFlangeThickness / 2, oOriginPos.y, oOriginPos.z - parFoamChamberInletWidth / 2
        Set oTrapezoid1 = PlaceTrapezoid(m_OutputColl, oTrapezoidBaseCenterPos1, oNozzleData(1).dFlangeThickness, parFoamChamberInletHeight2, oNozzleData(1).dFlangeThickness, parFoamChamberInletHeight2, parFoamChamberInletWidth, True)

'        oTrapezoid1.Transform oT4x4Current
        oOutputCol.Add oTrapezoid1

        Dim oTrapezoid2 As Object
        Dim oTrapezoidBaseCenterPos2 As DPosition
        Set oTrapezoidBaseCenterPos2 = New DPosition
        oTrapezoidBaseCenterPos2.Set oOriginPos.x + oNozzleData(1).dFlangeThickness + len_1 / 2, oOriginPos.y, oOriginPos.z - parFoamChamberWidth / 2
        Set oTrapezoid2 = PlaceTrapezoid(m_OutputColl, oTrapezoidBaseCenterPos2, len_1, parFoamChamberInletHeight1, len_1, parFoamChamberInletHeight1, parFoamChamberWidth, True)
        
'        oTrapezoid2.Transform oT4x4Current
        oOutputCol.Add oTrapezoid2

'        MoveAlongAxis oT4x4Current, AxisDirection_Primary, parFoamChamberLength1
'
'        Call NozzlePlace(m_OutputColl, oPart, oNozzleData, NozzleId_CP0, oOriginPos, oT4x4Current, oNozzleCol, oOutputCol)

'        Dim angle As Double
'        angle = 90#

'        RotateOrientation oT4x4Current, AxisDirection_Secondary, angle
'
'
'        MoveAlongAxis oT4x4Current, AxisDirection_Primary, -parFoamChamberHeight1
        

        Dim proj As Double
        proj = parFoamChamberHeight1 + parFoamChamberHeight2

        Dim height As Double
        height = parFoamChamberWidth

        Dim width As Double
        width = 2# * parFoamChamberLength1

        Dim oTrapezoid3 As Object
        Dim oTrapezoidBaseCenterPos3 As DPosition
        Set oTrapezoidBaseCenterPos3 = New DPosition
        oTrapezoidBaseCenterPos3.Set oOriginPos.x + parFace1toCenter, oOriginPos.y + _
            parFoamChamberHeight2 / 2 - parFoamChamberHeight1 / 2, oOriginPos.z - height / 2
        Set oTrapezoid3 = PlaceTrapezoid(m_OutputColl, oTrapezoidBaseCenterPos3, width, proj, width, proj, height, True)

'        oTrapezoid3.Transform oT4x4Current
        oOutputCol.Add oTrapezoid3

        Dim offset_1 As Double
        offset_1 = parFoamChamberHeight2 - parFoamChamberHeight3

        MoveAlongAxis oT4x4Current, AxisDirection_Primary, -offset_1
        

'        RotateOrientation oT4x4Current, AxisDirection_Normal, angle
        

'        Dim half_ht As Double
'        half_ht = 0.5 * parFoamChamberWidth

'        Dim len_2 As Double
'        len_2 = parFoamChamberCyltoCenter - parFoamChamberCylThickness - half_ht
'
'        MoveAlongAxis oT4x4Current, AxisDirection_Primary, half_ht
'
'
'        oT4x4Temp.LoadIdentity
'        oT4x4Temp.IndexValue(12) = len_2
        
        Dim oCylinderCapped1 As Object
        Dim stPoint As AutoMath.DPosition
        Dim enPoint As AutoMath.DPosition
        Set stPoint = New AutoMath.DPosition
        Set enPoint = New AutoMath.DPosition

        stPoint.Set oOriginPos.x + parFace1toCenter, oOriginPos.y + parFoamChamberHeight3, oOriginPos.z + parFoamChamberWidth / 2
        enPoint.Set oOriginPos.x + parFace1toCenter, oOriginPos.y + parFoamChamberHeight3, oOriginPos.z + parFoamChamberCyltoCenter - parFoamChamberCylThickness
                
        Set oCylinderCapped1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, CDbl(parFoamChamberCylDia1), True)
'       oCylinderCapped1.Transform oT4x4Current
        
        oOutputCol.Add oCylinderCapped1
        
'        oT4x4Current.MultMatrix oT4x4Temp
        

'        oT4x4Temp.LoadIdentity
'        oT4x4Temp.IndexValue(12) = parFoamChamberCylThickness
        
        Dim oCylinderCapped2 As Object

        stPoint.Set oOriginPos.x + parFace1toCenter, oOriginPos.y + parFoamChamberHeight3, oOriginPos.z + parFoamChamberCyltoCenter - parFoamChamberCylThickness
        enPoint.Set oOriginPos.x + parFace1toCenter, oOriginPos.y + parFoamChamberHeight3, oOriginPos.z + parFoamChamberCyltoCenter
        Set oCylinderCapped2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, CDbl(parFoamChamberCylDia2), True)
        
'        oCylinderCapped2.Transform oT4x4Current
        
        oOutputCol.Add oCylinderCapped2
        
'        oT4x4Current.MultMatrix oT4x4Temp
     
'**********************************************************************************
'**********************************************************************************
'   END
'**********************************************************************************
'**********************************************************************************
    
    Dim oNozzle As IJDNozzle
    Dim lNozzleId As Long
    Dim lNozzleIndex As Long
    
    Dim oNozzlePos As DPosition
    Dim oPos As AutoMath.DPosition
    Set oPos = New AutoMath.DPosition
    
    Dim oNozzleVec As DVector
    Set oNozzleVec = New DVector

    'place the nozzles & add to the output collection
    For lNozzleIndex = 1 To oNozzleCol.Count
        lNozzleId = oNozzleCol.Item(lNozzleIndex)
        
        'move the nozzle in the opposite direction by the flange thickness
        oNozzleVec.Set -oNozzleData(lNozzleId).oDirectionVector.x, -oNozzleData(lNozzleId).oDirectionVector.y, -oNozzleData(lNozzleId).oDirectionVector.z
        oNozzleVec.Length = oNozzleData(lNozzleId).dFlangeThickness
        
'        Set oNozzlePos = oNozzleData(lNozzleId).oPlacementPoint.Clone.Offset(oNozzleVec)
        oPos.Set oOriginPos.x - oNozzleData(lNozzleIndex).dCptOffset + oNozzleData(lNozzleIndex).dDepth, oOriginPos.y, oOriginPos.z
        Set oNozzle = CreateNozzleJustaCircle(CInt(lNozzleId), oPart, m_OutputColl, oNozzleData(lNozzleId).oDirectionVector, oPos)
        
        oOutputCol.Add oNozzle
    Next lNozzleIndex
    
    Dim lOutputIndex As Long
    'add all outputs
    For lOutputIndex = 1 To oOutputCol.Count
        m_OutputColl.AddOutput arrayOfOutputs(lOutputIndex), oOutputCol.Item(lOutputIndex)
    Next lOutputIndex
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub
