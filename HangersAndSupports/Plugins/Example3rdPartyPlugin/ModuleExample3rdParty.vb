Option Explicit On
'*******************************************************************
'Copyright (C) 2009 Intergraph Corporation. All rights reserved.
'
'Abstract:
'    ModuleExample3rdParty.vb - utility routines
'
'Description:
'
'
'Notes:
'
'
'History
'   GWE                 Sep/01/2007      created
'   Chandra Sekhar R    APR/01/2009      Converted to .Net
'   IRK                 Aug/19/2010      TR-CP-186151  Example 3rdParty Plugin Should Set the Name attribute to DesignSupport  
'******************************************************************

Imports Ingr.SP3D.Common.Client.Services
Imports Ingr.SP3D.Common.Middle

Imports System.Math

Module ModuleExample3rdParty

    Const MODULENAME = "modExp3dParty"
    Public Const gc_HS_PLUGIN_Example = "HS_PLUGIN_Example"
    Public gDefaultPartsList As String = "HangerAttribute,Name,SupportContaining3rdPartyParts" & vbCrLf & _
                        "!PartNumber,         E, N,  EL,  RotE,  RotN,  RotEL, Len,  Attr=Val" & vbCrLf & _
                        "Anvil_Fig212_14,  0,  0,   0,     0,        0,         0,        0" & vbCrLf & _
                        "Anvil_FIG80C_16, 0,  0,   1,     0,        0,         0,        0,     TOTAL_TRAV=0.02"
    Public gsExchangeFilePath As String
    Public gs3dPartyParts As String
    Public gsLogFilePath As String
    Public gLogLevel As Integer
    Public gbShowLogfile As Boolean
    Public gsIniPath As String

    Public Function Vector(ByVal x As Double, ByVal y As Double, ByVal z As Double) As String
        Vector = "( " & sDouble(x) & ", " & sDouble(y) & ", " & sDouble(z) & " )"
    End Function

    Public Function sDouble(ByVal d As Double) As String
        If ((Abs(d) < 0.0001 And Abs(d) > 0.000000000001) Or Abs(d) > 1000000.0#) Then
            sDouble = d
        Else
            sDouble = Format(d, "####0.000")
        End If
    End Function
    Public Function VectorAsString(ByVal x As Double, ByVal y As Double, ByVal z As Double) As String
        VectorAsString = x & "," & y & "," & z
    End Function

    Public Function VectorP(ByVal oPos As Position) As String
        VectorP = Vector(oPos.X, oPos.Y, oPos.Z)
    End Function
    Public Function VectorV(ByVal oPos As Vector) As String
        VectorV = Vector(oPos.X, oPos.Y, oPos.Z)
    End Function

    '   Const HgrPipingDisciplineType = 1
    '   Const HgrHVACDisciplineType = 2
    '   Const HgrCableWayDisciplineType = 4
    '       Const HgrCableWayDuctDisciplineType = 6
    '   Const HgrConduitDisciplineType = 8
    Public Function DisciplineAsString(ByVal iDisp As Integer) As String
        DisciplineAsString = New String("")

        If ((iDisp And 1) <> 0) Then
            DisciplineAsString = DisciplineAsString & "," & "Piping"
        End If
        If ((iDisp And 2) <> 0) Then
            DisciplineAsString = DisciplineAsString & "," & "HVAC"
        End If
        If ((iDisp And 4) <> 0) Then
            DisciplineAsString = DisciplineAsString & "," & "Cableway"
        End If
        If ((iDisp And 8) <> 0) Then
            DisciplineAsString = DisciplineAsString & "," & "Conduit"
        End If
        If ((iDisp And 10) <> 0) Then
            DisciplineAsString = DisciplineAsString & "," & "Equipment"
        End If

        DisciplineAsString = Mid(DisciplineAsString, 2)
    End Function
    '    1 Along
    '    2 Branch
    '    3 End
    '    4 Transition
    '    5 Straight
    '    6 Turn
    '    7 Surface
    '    8 EntryExit
    '    9 Segment
    '    0 undefined

    Public Function FeatureAsString(ByVal iFeat As Integer) As String

        Select Case iFeat
            Case 1
                FeatureAsString = "Along"
            Case 2
                FeatureAsString = "Branch"
            Case 3
                FeatureAsString = "End"
            Case 4
                FeatureAsString = "Transition"
            Case 5
                FeatureAsString = "Straight"
            Case 6
                FeatureAsString = "Turn"
            Case 7
                FeatureAsString = "Surface"
            Case 8
                FeatureAsString = "EntryExit"
            Case 9
                FeatureAsString = "Segment"
            Case Else
                FeatureAsString = "Undefined"
        End Select
    End Function
    Public Sub ReadFromPreference()
        With ClientServiceProvider.Preferences
            gsExchangeFilePath = .GetStringValue(gc_HS_PLUGIN_Example & "1", Environ("TEMP") & "\test.txt")
            gs3dPartyParts = .GetStringValue(gc_HS_PLUGIN_Example & "2", "")
            gsIniPath = .GetStringValue("IMPORT_ASSEMBLIES_CFG", System.Environment.CurrentDirectory & "\HangerProg.ini")
            gLogLevel = .GetIntegerValue("IMPORT_ASSEMBLIES_LOGLEVEL", 0)
            gsLogFilePath = .GetStringValue("IMPORT_ASSEMBLIES_LOGFILE", "")
            gbShowLogfile = .GetBooleanValue("IMPORT_ASSEMBLIES_SHOWLOG", False)
        End With

    End Sub

    Public Sub WriteToPreference()
        With ClientServiceProvider.Preferences
            .SetValue(gc_HS_PLUGIN_Example & "1", gsExchangeFilePath)
            .SetValue(gc_HS_PLUGIN_Example & "2", gs3dPartyParts)
            .SetValue("IMPORT_ASSEMBLIES_CFG", gsIniPath)
            .SetValue("IMPORT_ASSEMBLIES_LOGLEVEL", gLogLevel)
            .SetValue("IMPORT_ASSEMBLIES_LOGFILE", gsLogFilePath)
            .SetValue("IMPORT_ASSEMBLIES_SHOWLOG", gbShowLogfile)
        End With
    End Sub
End Module
