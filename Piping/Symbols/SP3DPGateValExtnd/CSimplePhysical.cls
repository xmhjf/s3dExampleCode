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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   File:   SimplePhysical.cls
'   Author: MS
'   Date:   Tuesday, 21st Sept 2004
'
'   Description:
'      This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'  Change History:
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
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
'    Dim oPipeComponent As IJDPipeComponent
    
    Set oPart = arrayOfInputs(1)

'    Set oPipeComponent = oPart
    
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
    
        Dim FacetoCenter As Variant
        Dim Face1toCenter As Variant

        FacetoCenter = arrayOfInputs(2)
        Face1toCenter = arrayOfInputs(3)


        NozzleDefineDirectionVector oNozzleData, NozzleGeometryType_Olet
        

        Call NozzlePlace(m_OutputColl, oPart, oNozzleData, NozzleId_CP1, oOriginPos, oT4x4Current, oNozzleCol, oOutputCol)
        

'        If (gen_type_2 = TerminationClass_Bolted) Then
'
'                Dim extra2 As Variant
'                extra2 = oNozzleData(2).dCptOffset
'
'        Else
'
'                extra2 = 0#
'
'        End If

        Dim length_1 As Variant
        length_1 = FacetoCenter - Face1toCenter '+ extra2

        Dim length_2 As Variant
        length_2 = Face1toCenter - oNozzleData(2).dFlangeThickness

        Dim diameter As Variant
        diameter = 0#

        Dim dia_end As Variant
        dia_end = 0.6 * oNozzleData(2).dPipeDiameter

        oT4x4Temp.LoadIdentity
        oT4x4Temp.IndexValue(12) = length_1
        
        Dim oConeCapped1 As Object
        Dim stpoint As IJDPosition
        Dim enpoint As IJDPosition
        Set stpoint = oOriginPos
        Set enpoint = oT4x4Temp.TransformPosition(oOriginPos)
        
        Set oConeCapped1 = PlaceCone(m_OutputColl, stpoint, enpoint, CDbl(dia_end) / 2, CDbl(oNozzleData(2).dPipeDiameter) / 2, True)
        
        oConeCapped1.Transform oT4x4Current
        
        oOutputCol.Add oConeCapped1
        
        oT4x4Current.MultMatrix oT4x4Temp
        

        oT4x4Temp.LoadIdentity
'        oT4x4Temp.IndexValue(12) = oNozzleData(2).dDepth
         
'        Dim oCylinderCapped1 As Object
'        Set oCylinderCapped1 = PlaceCylinder(m_OutputColl, oOriginPos, oT4x4Temp.TransformPosition(oOriginPos), CDbl(oNozzleData(2).dFlangeDiameter), True)
'
'        oCylinderCapped1.Transform oT4x4Current
'
'        oOutputCol.Add oCylinderCapped1
        
        oT4x4Current.MultMatrix oT4x4Temp
        

        oT4x4Temp.LoadIdentity
        oT4x4Temp.IndexValue(12) = length_2
        
        Dim oConeCapped2 As Object
        Set stpoint = oOriginPos
        Set enpoint = oT4x4Temp.TransformPosition(oOriginPos)
        
        Set oConeCapped2 = PlaceCone(m_OutputColl, stpoint, enpoint, CDbl(oNozzleData(2).dPipeDiameter) / 2, CDbl(diameter) / 2, True)
        
        oConeCapped2.Transform oT4x4Current
        
        oOutputCol.Add oConeCapped2
        
        oT4x4Current.MultMatrix oT4x4Temp
        

        Call NozzlePlace(m_OutputColl, oPart, oNozzleData, NozzleId_CP0, oOriginPos, oT4x4Current, oNozzleCol, oOutputCol)
        

        oT4x4Temp.LoadIdentity
        oT4x4Temp.IndexValue(12) = length_2
        
        Dim oConeCapped3 As Object
        Set stpoint = oOriginPos
        Set enpoint = oT4x4Temp.TransformPosition(oOriginPos)
        enpoint.z = enpoint.z
        Set oConeCapped3 = PlaceCone(m_OutputColl, enpoint, stpoint, CDbl(oNozzleData(2).dPipeDiameter) / 2, CDbl(diameter) / 2, True)
        
        oConeCapped3.Transform oT4x4Current
        
        oOutputCol.Add oConeCapped3
        
        oT4x4Current.MultMatrix oT4x4Temp
        

        oT4x4Temp.LoadIdentity
'        oT4x4Temp.IndexValue(12) = oNozzleData(2).dDepth
        
'        Dim oCylinderCapped2 As Object
'        Set oCylinderCapped2 = PlaceCylinder(m_OutputColl, oOriginPos, oT4x4Temp.TransformPosition(oOriginPos), CDbl(oNozzleData(2).dFlangeDiameter), True)
'
'        oCylinderCapped2.Transform oT4x4Current
'
'        oOutputCol.Add oCylinderCapped2
        
        oT4x4Current.MultMatrix oT4x4Temp
        

        MoveAlongAxis oT4x4Current, AxisDirection_Primary, oNozzleData(2).dCptOffset
        

        Call NozzlePlace(m_OutputColl, oPart, oNozzleData, NozzleId_CP2, oOriginPos, oT4x4Current, oNozzleCol, oOutputCol)
        



    
'**********************************************************************************
'**********************************************************************************
'   END
'**********************************************************************************
'**********************************************************************************
    
    Dim oNozzle As IJDNozzle
    Dim lNozzleId As Long
    Dim lNozzleIndex As Long
    
    Dim oNozzlePos As DPosition
    
    Dim oNozzleVec As DVector
    Set oNozzleVec = New DVector
    
    'place the nozzles & add to the output collection
    For lNozzleIndex = 1 To oNozzleCol.Count
        lNozzleId = oNozzleCol.Item(lNozzleIndex)
        
        'move the nozzle in the opposite direction by the flange thickness
'        oNozzleVec.Set -oNozzleData(lNozzleId).oDirectionVector.x, -oNozzleData(lNozzleId).oDirectionVector.y, -oNozzleData(lNozzleId).oDirectionVector.z
'        oNozzleVec.Length = oNozzleData(lNozzleId).dFlangeThickness
        
'        Set oNozzlePos = oNozzleData(lNozzleId).oPlacementPoint.Clone.Offset(oNozzleVec)
        Set oNozzlePos = oNozzleData(lNozzleId).oPlacementPoint.Clone
        ' The nozzle needs to be positioned away by (cptoffset-depth) value in the nozzle direction.
        oNozzlePos.x = oNozzlePos.x + (oNozzleData(lNozzleId).oDirectionVector.x) * (oNozzleData(lNozzleId).dCptOffset - oNozzleData(lNozzleId).dDepth)
        oNozzlePos.y = oNozzlePos.y + (oNozzleData(lNozzleId).oDirectionVector.y) * (oNozzleData(lNozzleId).dCptOffset - oNozzleData(lNozzleId).dDepth)
        oNozzlePos.z = oNozzlePos.z + (oNozzleData(lNozzleId).oDirectionVector.z) * (oNozzleData(lNozzleId).dCptOffset - oNozzleData(lNozzleId).dDepth)
        
        Set oNozzle = CreateNozzle(CInt(lNozzleId), oPart, m_OutputColl, oNozzleData(lNozzleId).oDirectionVector, oNozzlePos)
        
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
'    Resume Next
End Sub

