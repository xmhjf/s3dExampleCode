'*******************************************************************
'Copyright (C) 2009, Intergraph Corporation. All rights reserved.
'
'Abstract:
'    PlugIn3rdParty - class implementing HangerProg
'    This is an Example of a PlugIn Which demonstrates the way to develop 
'    customized Third Party Plugin'S which places assemblies in SmartPlant3D
'
'Description:
'
'
'Notes:
'
'
'History
'GWE                Sep/01/2007      created
'Chandra Sekhar R   Apr/01/2009      Converted to .net
'******************************************************************

Imports Ingr.SP3D.Support.Client
Imports Ingr.SP3D.Common.Client
Imports Ingr.SP3D.Common.Middle
Imports System.Math
Imports System.Reflection.Assembly
Imports System.Diagnostics.FileVersionInfo
Imports Ingr.SP3D.Common.Client.Services

Public Class PlugIn3rdParty

    Inherits HangerProg
    Const MODULENAME = "PlugIn3rdParty"
    Private m_oParams As HgrParameters
    Private m_sPartArr() As String              ' description of parts to place
    Private m_lIndex As Long                    ' index of current part
    Private m_sCurrentPart() As String          ' information about current part
    Private Const cPartName = 0                 ' m_sCurrentPart(0) = partname to place
    Private Const cPartX = 1                    ' m_sCurrentPart(1) = x position (East Direction) in m
    Private Const cPartY = 2                    ' m_sCurrentPart(2) = y position (North Direction) in m
    Private Const cPartZ = 3                    ' m_sCurrentPart(3) = z position in m
    Private Const cPartRotX = 4                 ' m_sCurrentPart(4) = rotation about x-axis in degree
    Private Const cPartRotY = 5                 ' m_sCurrentPart(5) = rotation about y-axis in degree
    Private Const cPartRotZ = 6                 ' m_sCurrentPart(6) = rotation about z-axis in degree
    Private Const cPartLength = 7               ' m_sCurrentPart(7) = length of part (if part is stretchable)
    Private Const cPartParameterSetting = 8     ' m_sCurrentPart(8) = "name=value" to set a part attribute
    Private Const cPartMax = 8

    '*********************************************************************************************
    '  ComputePositions
    '
    '  This is the first call from the "Start 3rd Party App" command [H&S].
    '  Call the 3dParty application and setup a list of parts to place
    '
    '  m_oParams    = object containing all information about the Design Support
    '                 
    '
    '***************************************************************************************
    Public Overrides Function ComputePositions(ByRef oPars As HgrParameters) As Long
        Try
            Dim lFnr As Long
            Dim lSts As Long
            Dim b As Boolean
            Dim i As Long = 0
            Dim j As Long = 0
            Dim sSupportInfo As String = Nothing

            ComputePositions = -1  ' indicates abort
            m_oParams = oPars

            If (gsExchangeFilePath = "") Then
                ' Assume we need a configured exchangefile path,
                ' so we call the editconfiguration utility, if there is none
                lSts = EditConfiguration()
                If (lSts < 0) Then
                    Exit Function ' user pressed cancel
                End If
                If (gsExchangeFilePath = "") Then
                    MsgBox("Please specify an exchange file path.")
                    Exit Function
                End If
            End If
            '  Store all information about the Support in a string sSupportInfo
            '
            '  retrieve  attributes, and add to String
            '
            Dim oPropName As PropertyValueString = Nothing
            b = oPars.GetSupportAttribute("Name", oPropName)
            If (b) Then
                sSupportInfo = sSupportInfo & vbCrLf & "Support_Name=" & oPropName.PropValue
            End If

            Dim oPropLoad As PropertyValueDouble = Nothing
            b = oPars.GetSupportAttribute("MaxLoad", oPropLoad)

            If (b) Then
                sSupportInfo = sSupportInfo & vbCrLf & "Support_MaxLoad=" & oPropLoad.PropValue
            End If

            sSupportInfo = sSupportInfo & vbCrLf & "SupportPosition=" & VectorP(oPars.SupportPosition)

            If (oPars.NumberRoutes > 0) Then
                For i = 0 To oPars.NumberRoutes - 1

                    sSupportInfo = sSupportInfo & vbCrLf & "[Route_" & i & "]"
                    b = oPars.GetRouteAttribute(i, "Name", oPropName)
                    If (b) Then
                        sSupportInfo = sSupportInfo & vbCrLf & "Route_Name=" & oPropName.PropValue
                    End If
                    sSupportInfo = sSupportInfo & vbCrLf & "Discipline=" & oPars.RouteDiscipline(i) & " " & DisciplineAsString(oPars.RouteDiscipline(i))
                    sSupportInfo = sSupportInfo & vbCrLf & "FeatureType=" & oPars.RouteFeatureType(i) & " " & FeatureAsString(oPars.RouteFeatureType(i))
                    sSupportInfo = sSupportInfo & vbCrLf & "NominalDiameter=" & oPars.RteNominalDiaEntered(i)
                    sSupportInfo = sSupportInfo & vbCrLf & "NominalDiameterUnit=" & oPars.RteNominalDiaUnit(i)
                    sSupportInfo = sSupportInfo & vbCrLf & "NominalDiameterMeter=" & oPars.RteNominalDiaInMeter(i)
                    sSupportInfo = sSupportInfo & vbCrLf & "RouteDir=" & VectorV(oPars.RouteDir(i, 0)) & "/" & VectorV(oPars.RouteDir(i, 1)) & "/" & VectorV(oPars.RouteDir(i, 2))
                    sSupportInfo = sSupportInfo & vbCrLf & "Pressure=" & oPars.RoutePressure(i)
                    sSupportInfo = sSupportInfo & vbCrLf & "Temperature=" & oPars.RouteTemperature(i)
                    sSupportInfo = sSupportInfo & vbCrLf & "InsulationPurpose=" & oPars.RouteInsulationPurpose(i)
                    sSupportInfo = sSupportInfo & vbCrLf & "Insulation=" & oPars.RouteInsulation(i)
                    sSupportInfo = sSupportInfo & vbCrLf & "Height=" & oPars.RouteHeight(i)
                    sSupportInfo = sSupportInfo & vbCrLf & "Width=" & oPars.RouteWidth(i)
                    sSupportInfo = sSupportInfo & vbCrLf & "TurnRadius=" & oPars.RouteTurnRadius(i)

                Next i
            End If

            If (oPars.NumberStructs > 0) Then
                For i = 0 To oPars.NumberStructs - 1
                    sSupportInfo = sSupportInfo & vbCrLf & "[Struct_" & i & "]"
                    b = oPars.GetStructAttribute(i, "Name", oPropName)

                    If (b) Then
                        sSupportInfo = sSupportInfo & vbCrLf & "Struct_Name=" & oPropName.PropValue
                    End If

                    sSupportInfo = sSupportInfo & vbCrLf & "SectionName=" & oPars.StructSectionName(i)
                    sSupportInfo = sSupportInfo & vbCrLf & "SectionType=" & oPars.StructSectionType(i)
                    sSupportInfo = sSupportInfo & vbCrLf & "NearPos=" & VectorP(oPars.StructPosition(i))
                    sSupportInfo = sSupportInfo & vbCrLf & "FarPos=" & VectorP(oPars.StructPositionFar(i))
                    sSupportInfo = sSupportInfo & vbCrLf & "Direction=" & VectorV(oPars.StructDirection(i, 0))
                    sSupportInfo = sSupportInfo & vbCrLf & "UpVector=" & VectorV(oPars.StructDirection(i, 2))
                    sSupportInfo = sSupportInfo & vbCrLf & "ConnectionSide=" & oPars.StructConnect(i)

                    For j = 0 To 8
                        sSupportInfo = sSupportInfo & vbCrLf & "Cp" & j + 1 & "=" & VectorP(oPars.StructStart(i, j)) & " -> " & VectorP(oPars.StructEnd(i, j))
                    Next j
                Next i
            End If
            If (oPars.NumberWalls > 0) Then
                For i = 0 To oPars.NumberWalls - 1
                    sSupportInfo = sSupportInfo & vbCrLf & "[Wall_" & i & "]"
                    sSupportInfo = sSupportInfo & vbCrLf & "NearPoint=" & VectorP(oPars.WallNearPosition(i))
                    sSupportInfo = sSupportInfo & vbCrLf & "FarPoint=" & VectorP(oPars.WallFarPosition(i))
                    sSupportInfo = sSupportInfo & vbCrLf & "FarPoint=" & VectorV(oPars.WallNormal(i))
                Next i
            End If
            '
            ' Write the information also onto the exchange file
            '
            lFnr = FreeFile()

            Try
                FileOpen(lFnr, gsExchangeFilePath, OpenMode.Binary, OpenAccess.Write)
            Catch ex As Exception
                MsgBox("Cannot open file: " & gsExchangeFilePath)
            End Try

            Try
                FilePutObject(lFnr, sSupportInfo)
            Catch ex As Exception
                MsgBox("Cannot Write to file: " & gsExchangeFilePath)
            End Try

            Try
                FileClose(lFnr)
            Catch ex As Exception
            End Try

            ' Now an external application could be called.
            ' This application would read the exchangefile,
            ' may ask the user for further input and can
            ' finally write a file, which this program could read in
            ' to know, which parts are to place

            '        sCmd = "My3dpartyApplication.exe" & " " & gsExchangeFilePath
            '        lSts = ExecCmd(sCmd, False)

            ' For now, we will display a form, showing the  input from Sp3d
            ' and the parts, we want to place.
            ' The user may edit the list of parts to place.

            Dim oFrmExample As New frm3rdPartyExpl()
            oFrmExample.txtInput.Text = sSupportInfo
            If (String.IsNullOrEmpty(gs3dPartyParts)) Then
                gs3dPartyParts = gDefaultPartsList
            End If

            oFrmExample.txtParts.Text = gs3dPartyParts
            oFrmExample.ShowDialog()
            If (oFrmExample.lStatus < 0) Then
                Exit Function ' do not process, cancel button pressed
            End If
            '   if the user has changed the list of parts, write this as preference to the session file
            '
            If (oFrmExample.txtParts.Text <> gs3dPartyParts) Then
                gs3dPartyParts = oFrmExample.txtParts.Text
                WriteToPreference()
            End If

            '   gs3dPartyParts contains the information about all parts,
            '   one part per line
            '   split this information to get an array of part information:

            m_sPartArr = Split(gs3dPartyParts, vbCrLf)

            '
            '   check if there are commands within this list to proceed immediately
            '   (setting attributes of the hanger)
            '
            For i = LBound(m_sPartArr) To UBound(m_sPartArr)
                m_sCurrentPart = Split(m_sPartArr(i), ",")
                ReDim Preserve m_sCurrentPart(2)
                '
                '  check for "HangerAttribute".
                '  This indicates, that column 1 contains an attribute name and
                '  column 3 the corresponding value
                '
                If (StrComp(m_sCurrentPart(0), "HangerAttribute", vbTextCompare) = 0) Then
                    m_sPartArr(i) = "" ' avoid trying to place a part with the name HangerAttribute later
                    If (m_sCurrentPart(1) <> "") Then
                        b = oPars.SetSupAttribute(m_sCurrentPart(1), m_sCurrentPart(2))
                        If (Not b) Then
                            MsgBox("Cannot set support attribute " & m_sCurrentPart(1) & "=" & m_sCurrentPart(2))
                        End If
                    End If
                End If
            Next i

            '
            '   Tell calling application to proceed
            '
            ComputePositions = 0
            Exit Function
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Critical, MODULENAME)
            Debug.Assert(False)
            Exit Function
            Err.Clear()
        End Try
    End Function
    '**********************************************************************************
    '
    '
    '
    '**********************************************************************************
    Public Overrides Function getPart(ByVal lIndex As Long, ByRef sName As String, ByRef sPath As String, ByRef sDBName As String, ByRef oPos As Ingr.SP3D.Common.Middle.Position, ByRef xDir As Ingr.SP3D.Common.Middle.Vector, ByRef yDir As Ingr.SP3D.Common.Middle.Vector) As Long
        m_lIndex = lIndex
        m_sCurrentPart = Split(m_sPartArr(lIndex - 1), ",")
        ReDim Preserve m_sCurrentPart(cPartMax)
        sName = m_sCurrentPart(cPartName)
        sDBName = m_sCurrentPart(cPartName)
        'remove the space from PartNumber string at the start/end positions
        sName = sName.TrimStart()
        sDBName = sDBName.TrimStart()
        sName = sName.TrimEnd()
        sDBName = sDBName.TrimEnd()

        If (sName = "") Then
            getPart = -1 ' do not place this part (was an empty line), but continue with next line
        End If
        If (sName.Contains("!")) Then
            getPart = -1 ' do not place this part (was a comment line), but continue with next line
        End If

        oPos = New Position(m_oParams.SupportPosition)

        '
        '   add the z value
        '
        On Error Resume Next ' if there is no z coordinate value
        oPos.X = oPos.X + m_sCurrentPart(cPartX)
        oPos.Y = oPos.Y + m_sCurrentPart(cPartY)
        oPos.Z = oPos.Z + m_sCurrentPart(cPartZ)
    End Function
    Public Overrides Function SpecialBehaviour(ByRef oPars As HgrParameters) As Long

        Dim sP, sPx, sPy, sPz As String
        Dim sAttribute() As String
        Dim oMatrix As Matrix4X4
        Dim vx, vy, vz As Vector
        Dim dx, dy, dz As Double
        '
        '  check, if there is an attribute  to set for the part
        '
        sP = m_sCurrentPart(cPartParameterSetting)
        If (sP <> "") Then
            sAttribute = Split(sP, "=")
            ReDim Preserve sAttribute(1)
            If (sAttribute(0) <> "") Then
                If (Not oPars.SetPartAttribute(sAttribute(0), sAttribute(1))) Then
                    MsgBox("Error setting attribute: " & sAttribute(0) & "=" & sAttribute(1))
                End If
            End If
        End If
        '
        ' Check if there is a rotation necessary for the part
        sPx = m_sCurrentPart(cPartRotX)
        sPy = m_sCurrentPart(cPartRotY)
        sPz = m_sCurrentPart(cPartRotZ)
        If (sPx <> "") Or (sPy <> "") Or (sPz <> "") Then
            If (IsNumeric(sPx)) And (IsNumeric(sPy)) And (IsNumeric(sPz)) Then

                On Error Resume Next
                dx = sPx ' convert to double, if failed, we get 0, no rotation
                dy = sPy ' convert to double, if failed, we get 0, no rotation
                dz = sPz ' convert to double, if failed, we get 0, no rotation
                On Error GoTo 0
                ' convert angle into radians
                dx = dx * PI / 180.0#
                dy = dy * PI / 180.0#
                dz = dz * PI / 180.0#
                oMatrix = oPars.PartTransformationMatrix
                vx = New Vector()
                vy = New Vector()
                vz = New Vector()
                vx.Set(1.0#, 0.0#, 0.0#)
                vy.Set(0.0#, 1.0#, 0.0#)
                vz.Set(0.0#, 0.0#, 1.0#)

                If oMatrix IsNot Nothing Then

                    Call oMatrix.Rotate(dx, vx)
                    Call oMatrix.Rotate(dy, vy)
                    Call oMatrix.Rotate(dz, vz)

                End If


                oPars.PartTransformationMatrix = oMatrix
            End If
        End If
CleanUp:
    End Function

    Public Overrides ReadOnly Property Length() As Double
        Get
            Length = 1.0# ' 1 meter
            On Error Resume Next ' if there is not length given
            Length = m_sCurrentPart(cPartLength)
        End Get
    End Property

    Public Overrides ReadOnly Property NumberParts() As Long
        Get
            NumberParts = UBound(m_sPartArr) + 1
        End Get
    End Property

    Public Overrides ReadOnly Property SystemName() As String
        Get
            SystemName = "Third Party Example"
        End Get
    End Property

    Public Overrides ReadOnly Property Version() As String
        Get
            Dim a, b, c As Integer
            a = GetVersionInfo(GetExecutingAssembly.Location).ProductMajorPart
            b = GetVersionInfo(GetExecutingAssembly.Location).ProductMinorPart
            c = GetVersionInfo(GetExecutingAssembly.Location).ProductBuildPart
            Version = a & "." & b & "." & c
        End Get
    End Property

    Public Overrides Function EditConfiguration() As Long
        Dim ofrmConfig As New frm3rdPartyConfig()
        EditConfiguration = ofrmConfig.ShowModal
    End Function

    Public Sub New()
        ReadFromPreference()
    End Sub
End Class
