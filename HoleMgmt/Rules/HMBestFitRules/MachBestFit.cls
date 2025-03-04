VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MachBestFit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
' Copyright (C) 1998-2000 Intergraph Corporation.  All Rights Reserved.
'
' File: MachBestFit.cls
'
' Author: Prayaga Avinash
'
' Abstract: Hole Clearance Rules for Mach BestFit
'
' Description:
' This is used for Hole Clearance Rules in the creation of a Hole Trace.
'********************************************************************

Option Explicit

Implements IJHMPipeClearance
Implements IJHMDuctClearance
Implements IJHMCablewayClearance
Implements IJHMConduitClearance

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' IJHMCablewayClearance implementation
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Joint Type   Tightness  Insulation  Round Value  Hole Size
' flange         Yes        -          Yes         Flange O-dim + 10 mm (clearance)
'                No         Yes        Yes         Flange O-dim + 25 mm (clearance)
'                                                    or Cableway O-dim + I.T * 2 + 25 mm (clearance) (large one)
'                No         No         Yes         Flange O-dim + 10 mm (clearance)
' sleeve         Yes        -          No          Sleeve O-dim + 2 mm (clearance)
'                No         Yes        No          Sleeve O-dim + I.T * 2  + 2 mm (clearance)
'                No         No         Yes         Sleeve O-dim + 25 mm (clearance)
' none/other     Yes        -          No          Cableway O-dim + 2 mm (clearance)
'                No         Yes        Yes         Cableway O-dim + I.T * 2 + 25 mm (clearance)
'                No         No         Yes         Cableway O-dim + 25 mm (clearance)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' cableway hole corner radius
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJHMCablewayClearance_CornerRadius() As Double
    'return the radius corner value
    IJHMCablewayClearance_CornerRadius = 0.025
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' cableway has flange fitting
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJHMCablewayClearance_FlangeClr(ByVal bTight As Boolean, ByVal bInsulation As Boolean) As Double
    'return the required clearance
    If bTight Then
        'structure is watertight
        IJHMCablewayClearance_FlangeClr = 0.04
    Else
        If bInsulation Then
            'structure is not watertight and cableway has insulation
            IJHMCablewayClearance_FlangeClr = 0.1
        Else
            'structure is not watertight and cableway has no insulation
            IJHMCablewayClearance_FlangeClr = 0.04
        End If
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' cableway has flange fitting
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJHMCablewayClearance_FlangeRound(ByVal bTight As Boolean, ByVal bInsulation As Boolean) As Boolean
    'return if hole size can be rounded
    If bTight Then
        'structure is watertight
        IJHMCablewayClearance_FlangeRound = True
    Else
        If bInsulation Then
            'structure is not watertight and cableway has insulation
            IJHMCablewayClearance_FlangeRound = True
        Else
            'structure is not watertight and cableway has no insulation
            IJHMCablewayClearance_FlangeRound = True
        End If
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' cableway has fitting other than flange or sleeve
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' to provide for future upgrades the lType parameter is added but is currently not used
Private Function IJHMCablewayClearance_OtherClr(ByVal bTight As Boolean, ByVal bInsulation As Boolean, _
                                                ByVal lType As Long) As Double
    'return the required clearance
    If bTight Then
        'structure is watertight
        IJHMCablewayClearance_OtherClr = 0.008
    Else
        If bInsulation Then
            'structure is not watertight and cableway has insulation
            IJHMCablewayClearance_OtherClr = 0.1
        Else
            'structure is not watertight and cableway has no insulation
            IJHMCablewayClearance_OtherClr = 0.1
        End If
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' cableway has fitting other than flange or sleeve
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' to provide for future upgrades the lType parameter is added but is currently not used
Private Function IJHMCablewayClearance_OtherRound(ByVal bTight As Boolean, ByVal bInsulation As Boolean, _
                                                  ByVal lType As Long) As Boolean
    'return if hole size can be rounded
    If bTight Then
        'structure is watertight
        IJHMCablewayClearance_OtherRound = False
    Else
        If bInsulation Then
            'structure is not watertight and cableway has insulation
            IJHMCablewayClearance_OtherRound = True
        Else
            'structure is not watertight and cableway has no insulation
            IJHMCablewayClearance_OtherRound = True
        End If
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' cableway has sleeve fitting
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJHMCablewayClearance_SleeveClr(ByVal bTight As Boolean, ByVal bInsulation As Boolean) As Double
    'return the required clearance
    If bTight Then
        'structure is watertight
        IJHMCablewayClearance_SleeveClr = 0.008
    Else
        If bInsulation Then
            'structure is not watertight and cableway has insulation
            IJHMCablewayClearance_SleeveClr = 0.008
        Else
            'structure is not watertight and cableway has no insulation
            IJHMCablewayClearance_SleeveClr = 0.1
        End If
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' cableway has sleeve fitting
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJHMCablewayClearance_SleeveRound(ByVal bTight As Boolean, ByVal bInsulation As Boolean) As Boolean
    'return if hole size can be rounded
    If bTight Then
        'structure is watertight
        IJHMCablewayClearance_SleeveRound = False
    Else
        If bInsulation Then
            'structure is not watertight and cableway has insulation
            IJHMCablewayClearance_SleeveRound = False
        Else
            'structure is not watertight and cableway has no insulation
            IJHMCablewayClearance_SleeveRound = True
        End If
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' IJHMConduitClearance implementation
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' note that conduits do not have insulation

'Joint Type      Tightness  Round Value  Hole Size
' flange            -         Yes         Flange O.D + 10mm (clearance)
' square flange     -         Yes         Flange diagonal + 10 mm (clearance)
' sleeve            Yes       No          Sleeve O.D + 2mm (clearance)
'                   No        Yes         Conduit O.D + 25 mm (clearance)
' none/other        Yes       No          Conduit O.D + 2 mm (clearance)
'                   No        Yes         Conduit O.D + 25 mm (clearance)

'If the nearest distance between two holes is less than 50 mm, then two hole size is same

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' conduit has flange fitting
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJHMConduitClearance_FlangeClr() As Double
    'return the required clearance
    IJHMConduitClearance_FlangeClr = 0.04
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' conduit has flange fitting
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJHMConduitClearance_FlangeRound() As Boolean
    'return if hole size can be rounded
    IJHMConduitClearance_FlangeRound = True
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' minimum distance for hole simplification
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJHMConduitClearance_HoleMinDist() As Double
    'return minimum distance between holes for simplification
    IJHMConduitClearance_HoleMinDist = 0.05
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' conduit has fitting other than flange, square flange, or sleeve
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' to provide for future upgrades the lType parameter is added but is currently not used
Private Function IJHMConduitClearance_OtherClr(ByVal bTight As Boolean, ByVal lType As Long) As Double
    'return the required clearance
    If bTight Then
        'structure is watertight
        IJHMConduitClearance_OtherClr = 0.008
    Else
        'structure is not watertight
        IJHMConduitClearance_OtherClr = 0.1
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' conduit has fitting other than flange, square flange, or sleeve
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' to provide for future upgrades the lType parameter is added but is currently not used
Private Function IJHMConduitClearance_OtherRound(ByVal bTight As Boolean, ByVal lType As Long) As Boolean
    'return if hole size can be rounded
    If bTight Then
        'structure is watertight
        IJHMConduitClearance_OtherRound = False
    Else
        'structure is not watertight
        IJHMConduitClearance_OtherRound = True
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' conduit has sleeve fitting
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJHMConduitClearance_SleeveClr(ByVal bTight As Boolean) As Double
    'return the required clearance
    If bTight Then
        'structure is watertight
        IJHMConduitClearance_SleeveClr = 0.008
    Else
        'structure is not watertight
        IJHMConduitClearance_SleeveClr = 0.1
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' conduit has sleeve fitting
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJHMConduitClearance_SleeveRound(ByVal bTight As Boolean) As Boolean
    'return if hole size can be rounded
    If bTight Then
        'structure is watertight
        IJHMConduitClearance_SleeveRound = False
    Else
        'structure is not watertight
        IJHMConduitClearance_SleeveRound = True
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' conduit has square flange fitting
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJHMConduitClearance_SquareFlangeClr() As Double
    'return the required clearance
    IJHMConduitClearance_SquareFlangeClr = 0.04
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' conduit has square flange fitting
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJHMConduitClearance_SquareFlangeRound() As Boolean
    'return if hole size can be rounded
    IJHMConduitClearance_SquareFlangeRound = True
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' IJHMDuctClearance implementation
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Joint Type   Tightness  Insulation  Round Value  Hole Size
' flange         Yes        -          Yes         Flange O-dim + 10 mm (clearance)
'                No         Yes        Yes         Flange O-dim + 25 mm (clearance)
'                                                    or Duct O-dim + I.T * 2 + 25 mm (clearance) (large one)
'                No         No         Yes         Flange O-dim + 10 mm (clearance)
' sleeve         Yes        -          No          Sleeve O-dim + 2 mm (clearance)
'                No         Yes        No          Sleeve O-dim + I.T * 2  + 2 mm (clearance)
'                No         No         Yes         Sleeve O-dim + 25 mm (clearance)
' none/other     Yes        -          No          Duct O-dim + 2 mm (clearance)
'                No         Yes        Yes         Duct O-dim + I.T * 2 + 25 mm (clearance)
'                No         No         Yes         Duct O-dim + 25 mm (clearance)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' duct hole corner radius
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJHMDuctClearance_CornerRadius() As Double
    'return the radius corner value
    IJHMDuctClearance_CornerRadius = 0.025
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' duct has flange fitting
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJHMDuctClearance_FlangeClr(ByVal bTight As Boolean, ByVal bInsulation As Boolean) As Double
    'return the required clearance
    If bTight Then
        'structure is watertight
        IJHMDuctClearance_FlangeClr = 0.04
    Else
        If bInsulation Then
            'structure is not watertight and duct has insulation
            IJHMDuctClearance_FlangeClr = 0.1
        Else
            'structure is not watertight and duct has no insulation
            IJHMDuctClearance_FlangeClr = 0.04
        End If
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' duct has flange fitting
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJHMDuctClearance_FlangeRound(ByVal bTight As Boolean, ByVal bInsulation As Boolean) As Boolean
    'return if hole size can be rounded
    If bTight Then
        'structure is watertight
        IJHMDuctClearance_FlangeRound = True
    Else
        If bInsulation Then
            'structure is not watertight and duct has insulation
            IJHMDuctClearance_FlangeRound = True
        Else
            'structure is not watertight and duct has no insulation
            IJHMDuctClearance_FlangeRound = True
        End If
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' duct has fitting other than flange or sleeve
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' to provide for future upgrades the lType parameter is added but is currently not used
Private Function IJHMDuctClearance_OtherClr(ByVal bTight As Boolean, ByVal bInsulation As Boolean, _
                                            ByVal lType As Long) As Double
    'return the required clearance
    If bTight Then
        'structure is watertight
        IJHMDuctClearance_OtherClr = 0.008
    Else
        If bInsulation Then
            'structure is not watertight and duct has insulation
            IJHMDuctClearance_OtherClr = 0.1
        Else
            'structure is not watertight and duct has no insulation
            IJHMDuctClearance_OtherClr = 0.1
        End If
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' duct has fitting other than flange or sleeve
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' to provide for future upgrades the lType parameter is added but is currently not used
Private Function IJHMDuctClearance_OtherRound(ByVal bTight As Boolean, ByVal bInsulation As Boolean, _
                                              ByVal lType As Long) As Boolean
    'return if hole size can be rounded
    If bTight Then
        'structure is watertight
        IJHMDuctClearance_OtherRound = False
    Else
        If bInsulation Then
            'structure is not watertight and duct has insulation
            IJHMDuctClearance_OtherRound = True
        Else
            'structure is not watertight and duct has no insulation
            IJHMDuctClearance_OtherRound = True
        End If
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' duct has sleeve fitting
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJHMDuctClearance_SleeveClr(ByVal bTight As Boolean, ByVal bInsulation As Boolean) As Double
    'return the required clearance
    If bTight Then
        'structure is watertight
        IJHMDuctClearance_SleeveClr = 0.008
    Else
        If bInsulation Then
            'structure is not watertight and duct has insulation
            IJHMDuctClearance_SleeveClr = 0.008
        Else
            'structure is not watertight and duct has no insulation
            IJHMDuctClearance_SleeveClr = 0.1
        End If
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' duct has sleeve fitting
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJHMDuctClearance_SleeveRound(ByVal bTight As Boolean, ByVal bInsulation As Boolean) As Boolean
    'return if hole size can be rounded
    If bTight Then
        'structure is watertight
        IJHMDuctClearance_SleeveRound = False
    Else
        If bInsulation Then
            'structure is not watertight and cableway has insulation
            IJHMDuctClearance_SleeveRound = False
        Else
            'structure is not watertight and cableway has no insulation
            IJHMDuctClearance_SleeveRound = True
        End If
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' IJHMPipeClearance implementation
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Joint Type      Tightness  Insulation  Round Value  Hole Size
' flange            -          -          Yes         Flange O.D + 10mm (clearance)
' square flange     -          -          Yes         Flange diagonal + 10 mm (clearance)
' sleeve            Yes        -          No          Sleeve O.D + 2mm (clearance)
'                   No         Yes        Yes         Pipe O.D + I.T * 2 + 25 mm (clearance)
'                   No         No         Yes         Pipe O.D + 25 mm (clearance)
' none/other        Yes        -          No          Pipe O.D + 2 mm (clearance)
'                   No         Yes        Yes         Pipe O.D + I.T * 2 + 25 mm (clearance)
'                   No         No         Yes         Pipe O.D + 25 mm (clearance)

'If the nearest distance between two holes is less than 50 mm, then two hole size is same

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' pipe has flange fitting
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJHMPipeClearance_FlangeClr() As Double
    'return the required clearance
    IJHMPipeClearance_FlangeClr = 0.04
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' pipe has flange fitting
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJHMPipeClearance_FlangeRound() As Boolean
    'return if hole size can be rounded
    IJHMPipeClearance_FlangeRound = True
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' minimum distance for hole simplification
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJHMPipeClearance_HoleMinDist() As Double
    'return minimum distance between holes for simplification
    
    IJHMPipeClearance_HoleMinDist = 0.05
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' pipe has fitting other than flange, square flange, or sleeve
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' to provide for future upgrades the lType parameter is added but is currently not used
Private Function IJHMPipeClearance_OtherClr(ByVal bTight As Boolean, ByVal bInsulation As Boolean, _
                                            ByVal lType As Long) As Double
    'return the required clearance
    If bTight Then
        'structure is watertight
        IJHMPipeClearance_OtherClr = 0.008
    Else
        If bInsulation Then
            'structure is not watertight and pipe has insulation
            IJHMPipeClearance_OtherClr = 0.1
        Else
            'structure is not watertight and pipe has no insulation
            IJHMPipeClearance_OtherClr = 0.1
        End If
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' pipe has fitting other than flange, square flange, or sleeve
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' to provide for future upgrades the lType parameter is added but is currently not used
Private Function IJHMPipeClearance_OtherRound(ByVal bTight As Boolean, ByVal bInsulation As Boolean, _
                                              ByVal lType As Long) As Boolean
    'return if hole size can be rounded
    If bTight Then
        'structure is watertight
        IJHMPipeClearance_OtherRound = False
    Else
        If bInsulation Then
            'structure is not watertight and pipe has insulation
            IJHMPipeClearance_OtherRound = True
        Else
            'structure is not watertight and pipe has no insulation
            IJHMPipeClearance_OtherRound = True
        End If
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' pipe has sleeve fitting
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJHMPipeClearance_SleeveClr(ByVal bTight As Boolean, ByVal bInsulation As Boolean) As Double
    'return the required clearance
    If bTight Then
        'structure is watertight
        IJHMPipeClearance_SleeveClr = 0.008
    Else
        If bInsulation Then
            'structure is not watertight and pipe has insulation
            IJHMPipeClearance_SleeveClr = 0.1
        Else
            'structure is not watertight and pipe has no insulation
            IJHMPipeClearance_SleeveClr = 0.1
        End If
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' pipe has sleeve fitting
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJHMPipeClearance_SleeveRound(ByVal bTight As Boolean, ByVal bInsulation As Boolean) As Boolean
    'return if hole size can be rounded
    If bTight Then
        'structure is watertight
        IJHMPipeClearance_SleeveRound = False
    Else
        If bInsulation Then
            'structure is not watertight and pipe has insulation
            IJHMPipeClearance_SleeveRound = True
        Else
            'structure is not watertight and pipe has no insulation
            IJHMPipeClearance_SleeveRound = True
        End If
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' pipe has square flange fitting
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJHMPipeClearance_SquareFlangeClr() As Double
    'return the required clearance
    IJHMPipeClearance_SquareFlangeClr = 0.04
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' pipe has square flange fitting
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJHMPipeClearance_SquareFlangeRound() As Boolean
    'return if hole size can be rounded
    IJHMPipeClearance_SquareFlangeRound = True
End Function


 