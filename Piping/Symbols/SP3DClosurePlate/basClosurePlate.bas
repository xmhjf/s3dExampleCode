Attribute VB_Name = "basClosurePlate"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   basClosurePlate.bas
'   Author:          RUK
'   Creation Date:  Fiday, Feb 15 2008
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     -----               ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Public Type InputType
    name        As String
    description As String
    properties  As IMSDescriptionProperties
    uomValue    As Double
End Type

Public Type OutputType
    name            As String
    description     As String
    properties      As IMSDescriptionProperties
    Aspect          As SymbolRepIds
End Type

Public Type AspectType
    name                As String
    description         As String
    properties          As IMSDescriptionProperties
    AspectId            As SymbolRepIds
End Type

Public Sub ReportUnanticipatedError(InModule As String, InMethod As String)
    Const E_FAIL = -2147467259
    Err.Raise E_FAIL
End Sub

'CR-33401 - Function retrieves the wall thickness of the pipe.
Public Function RetrievePipeWallThick(index As Long, _
                                ByRef partInput As PartFacelets.IJDPart, _
                                ByRef dPipeThick As Double)

    Dim oPipePort  As IJCatalogPipePort
    Dim oCollection As IJDCollection
    Dim pPortIndex As Integer
    
    Set oCollection = partInput.GetNozzles()
    For pPortIndex = 1 To oCollection.Size
        Set oPipePort = oCollection.Item(pPortIndex)
            If oPipePort.PortIndex = index Then
                dPipeThick = oPipePort.WallThicknessOrGrooveSetback
            End If
    Next pPortIndex
    
    Set oPipePort = Nothing
    
End Function


