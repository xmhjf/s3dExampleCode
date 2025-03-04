Attribute VB_Name = "basCLatrolet"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   basCLatrolet.bas
'   Author:          NN
'   Creation Date:  Sunday, Feb 4 2001
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     -----               ------------------
'   09.Jul.2003     SymbolTeam(India)   Copyright Information, Header  is added.
'   08.SEP.2006     KKC  DI-95670       Replace names with initials in all revision history sheets and symbols
'   12.Feb.2008     RUK                 CR-136268  Enhance the latrolet symbol to be more realistic per Bonney Forge catalog
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

'CR-136268 - Function retrieves the wall thickness of the pipe.
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

