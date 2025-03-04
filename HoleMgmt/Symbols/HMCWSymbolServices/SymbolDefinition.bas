Attribute VB_Name = "SymbolDefinition"
'--------------------------------------------------------------------------------------------'
'    Copyright (C) 1998, 1999 Intergraph Corporation. All rights reserved.
'
'
'Abstract
'    Subs to define the symbol definitions for each particular input
'    Intended to be more generic later
'
'
'Notes
'
'
'History
'
'    sypark@ship.samsung.co.kr    30/01/02                Creation.
'                                                         Originally created for StructDetailing
'                                                         brackets and features
'--------------------------------------------------------------------------------------------'

Option Explicit

' The actual Library name of the Commonn custom methods
Public Const LIBRARYNAME_OF_COMMONCUSTOMMETHODS = "CommonSymbolUtils.CommonCM"
Public Const COMMONCM_LIB = "CommonCustomMethodsLib" ' Library name for the proxy to the Common custom methods

Public Const LIBRARYNAME_OF_CUSTOMMETHODS = "HMCWSymbolServices.CableHoleCM"  ' The actual Library name of the custom methods
Public Const CABLEHOLECM_LIB = "CableWayCustomMethodsLib" ' Library name for the proxy to the custom methods

Private Const MODULE = "HMCWSymbolServices.SymbolDefinition"

Public Sub DefineRepresentationForCableWayHole(ByVal ServerDoc As RAD2D.Document, GSCADRepsColl As IMSSymbolEntities.IJDRepresentations, pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    
    Const MT = "DefineRepresentationForCableWayHole"
    On Error GoTo ErrorHandler
    
    Dim RAD2DSymbolDef As RAD2D.SymbolProperties
    Dim RAD2DGroup As RAD2D.group
    Dim oOutputGraphic As IMSSymbolEntities.IJDOutput
    Dim oIJDOutputs As IJDOutputs
    Dim oGroups As RAD2D.Groups
    Dim type2D As RAD2D.ObjectType
    Dim RAD2DObj As Object
    Dim EndPointsConnectedElems As Collection
    Dim ASCableWayHole2DRep As IMSSymbolEntities.IJDRepresentation
    Dim TheGraphicAttributes As RAD2D.AttributeSets
    Dim oAttributeSet As RAD2D.AttributeSet
    Dim oAttribute As RAD2D.Attribute
    
   Set oOutputGraphic = New DOutput
   Set ASCableWayHole2DRep = New IMSSymbolEntities.DRepresentation

   ASCableWayHole2DRep.RepresentationId = 1 'define a aspect 0 (Simple_physical)
   ASCableWayHole2DRep.Name = SYMBOL_MACRO_REPRESENTATION_NAME
   ASCableWayHole2DRep.Properties = igCOLLECTION_VARIABLE

   'Only one representation and one RAD2D group gives one output
   'Set CM on Representation to evaluate GSCAD output of symbol
   'SetCMOnRepresentation ASCableWayHole2DRep, CABLEHOLECM_LIB, "EvaluateCableHole", pSymbolDefinition
   SetCMOnRepresentation ASCableWayHole2DRep, "SketchingAndMacroCustomMethodsLib", "StdEvaluationForMacro", pSymbolDefinition
 
   'Look for the group which is the output of the symbol: this group will be mapped in GSCAD as
   'a 3D complexstring3D. Set this output in the representation:
   Set oGroups = ServerDoc.ActiveSheet.Groups
   If oGroups.Count = 0 Then
       MsgBox "No group was defined on the .sym file as output of the symbol"
   Else
         'Query IJDOutputs interface
         Set oIJDOutputs = ASCableWayHole2DRep
         For Each RAD2DGroup In oGroups
           Set TheGraphicAttributes = RAD2DGroup.AttributeSets
           If Not TheGraphicAttributes Is Nothing Then
               If TheGraphicAttributes.Count <> 0 Then
                   For Each oAttributeSet In TheGraphicAttributes
                       If oAttributeSet.SetName <> "Input" Then
                           oOutputGraphic.Key = RAD2DGroup.Key
                           oOutputGraphic.Name = "Output_" & oOutputGraphic.Key
                           oIJDOutputs.SetOutput oOutputGraphic
                           oOutputGraphic.Reset
                       End If
                   Next
               Else
                   oOutputGraphic.Key = RAD2DGroup.Key
                   oOutputGraphic.Name = "Output_" & oOutputGraphic.Key
                   oIJDOutputs.SetOutput oOutputGraphic
                   oOutputGraphic.Reset
               End If
           Else
               oOutputGraphic.Key = RAD2DGroup.Key
               oOutputGraphic.Name = "Output_" & oOutputGraphic.Key
               oIJDOutputs.SetOutput oOutputGraphic
               oOutputGraphic.Reset
           End If
         Next
   End If
   
   'Set the representation to definition
   GSCADRepsColl.SetRepresentation ASCableWayHole2DRep
   ASCableWayHole2DRep.Reset

    Exit Sub
ErrorHandler:
        HandleError MODULE, MT
End Sub
 