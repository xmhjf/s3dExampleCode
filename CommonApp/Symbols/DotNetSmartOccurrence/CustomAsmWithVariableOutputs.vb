Option Strict On
Imports System
Imports System.Math
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services

Namespace Symbols
    ' Declare:
    '   IJDAttributes interface declared as output
    '   IJDGeometry interface declared as output
    <OutputNotification("IJDAttributes")> _
    <OutputNotification("IJDGeometry")> _
    Public MustInherit Class CustomAsmVariableOutputsBaseClass : Inherits CustomAssemblyDefinition
    End Class

    ' Declare:
    '   non-cached symbol
    '   symbol version
    <CacheOption(CacheOptionType.NonCached)> _
    <SymbolVersion("1.0.0.0")> _
    Public Class CustomAsmWithVariableOutputs : Inherits CustomAsmVariableOutputsBaseClass
        '   1. "Part"  ( Catalog part )
        '   2. "Diameter"
        <InputCatalogPart(1)> _
            Public m_catalogPart As InputCatalogPart
        <InputDouble(2, "BaseWidth", "Base Width", 4)> _
            Public m_dblBaseWidth As InputDouble
        <InputDouble(3, "BaseLength", "Base Length", 4)> _
            Public m_dblBaseLength As InputDouble
        <InputDouble(4, "BaseHeight", "Base Height", 0.5)> _
            Public m_dblBaseHeight As InputDouble
        <InputDouble(5, "SnowmenCount", "Snowmen Count", 1.0)> _
            Public m_dblSnowmenCount As InputDouble

        Private Const CONST_Base As String = "Base"
        Private Const CONST_Bottoms As String = "Bottoms"
        Private Const CONST_Bodies As String = "Bodies"
        Private Const CONST_Heads As String = "Heads"
        Private Const CONST_PropertyBaseWidth As String = "BaseWidth"
        Private Const CONST_PropertyBaseHeight As String = "BaseHeight"
        Private Const CONST_PropertyBaseLength As String = "BaseLength"
        Private Const CONST_PropertySnowmenCount As String = "SnowmenCount"
        Private Const CONST_IJUATestSnowmen As String = "IJUATestSnowmen"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        Public m_physicalAspect As AspectDefinition

        <AssemblyOutput(1, CONST_Base)> _
        Public m_objBase As AssemblyOutput

        <AssemblyOutput(2, CONST_Bottoms)> _
        Public m_objBottoms As AssemblyOutputs

        <AssemblyOutput(3, CONST_Bodies)> _
        Public m_objBodies As AssemblyOutputs

        <AssemblyOutput(4, CONST_Heads)> _
        Public m_objHeads As AssemblyOutputs

        ''' <summary>
        ''' Construct symbol outputs/aspects
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overrides Sub ConstructOutputs()
        End Sub

        ''' <summary>
        ''' Evaluate assembly outputs
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub EvaluateAssembly()
            Dim oBusinessObject As BusinessObject
            oBusinessObject = Me.Occurrence
            If oBusinessObject Is Nothing Then
                Throw New CmnException("Occurrence property is null!")
            End If

            Dim oConnection As SP3DConnection
            oConnection = oBusinessObject.DBConnection

            Dim snowmenCount As PropertyValueInt = DirectCast(oBusinessObject.GetPropertyValue(CONST_IJUATestSnowmen, CONST_PropertySnowmenCount), PropertyValueInt)
            Dim baseWidth As PropertyValueDouble = DirectCast(oBusinessObject.GetPropertyValue(CONST_IJUATestSnowmen, CONST_PropertyBaseWidth), PropertyValueDouble)
            Dim baseLength As PropertyValueDouble = DirectCast(oBusinessObject.GetPropertyValue(CONST_IJUATestSnowmen, CONST_PropertyBaseLength), PropertyValueDouble)
            Dim baseHeight As PropertyValueDouble = DirectCast(oBusinessObject.GetPropertyValue(CONST_IJUATestSnowmen, CONST_PropertyBaseHeight), PropertyValueDouble)


            Dim origin As New Position(baseWidth.PropValue.Value / 2.0, baseLength.PropValue.Value / 2.0, 0.0)

            Dim dblBottomRadius As Double
            If (snowmenCount.PropValue.Value > 0) Then
                dblBottomRadius = (baseWidth.PropValue.Value / snowmenCount.PropValue.Value) / 2.0
            Else
                dblBottomRadius = baseWidth.PropValue.Value
            End If
            Dim dblBodyRadius As Double
            dblBodyRadius = dblBottomRadius * 2.0 / 3.0
            Dim dblHeadRadius As Double
            dblHeadRadius = dblBodyRadius / 2.0

            ' Define the boundary for the base projection
            Dim posColl As New Collection(Of Position)
            posColl.Add(New Position(0.0, 0.0, 0.0))
            posColl.Add(New Position(baseWidth.PropValue.Value, 0.0, 0.0))
            posColl.Add(New Position(baseWidth.PropValue.Value, baseLength.PropValue.Value, 0.0))
            posColl.Add(New Position(0.0, baseLength.PropValue.Value, 0.0))
            posColl.Add(New Position(0.0, 0.0, 0.0))

            If m_objBase.Output Is Nothing Then
                '********************************
                ' Construct base
                '********************************
                m_objBase.Output = New Projection3d(oConnection, New LineString3d(posColl), _
                                                    New Vector(0.0, 0.0, 1.0), baseHeight.PropValue.Value, True)
            Else
                '********************************
                ' Update body
                '********************************
                Dim oBase As Projection3d
                oBase = DirectCast(m_objBase.Output, Projection3d)
                oBase.DefineByCurve(New LineString3d(posColl), _
                                    New Vector(0.0, 0.0, 1.0), baseHeight.PropValue.Value, True)
            End If

            ' Set the position of the bottom relative to the base
            origin.Z = origin.Z + baseHeight.PropValue.Value + dblBottomRadius
            origin.X = baseWidth.PropValue.Value

            Dim iCnt As Integer

            '********************************************************
            ' Construct / Update the variable output bottom spheres
            '********************************************************
            For iCnt = 1 To Convert.ToInt32(snowmenCount.PropValue.Value)
                If m_objBottoms.Count < iCnt Then
                    '********************************
                    ' Construct bottom
                    '********************************
                    m_objBottoms.Add(New Torus3d(oConnection, origin, New Vector(1.0, 0.0, 0.0), _
                                                 New Vector(0.0, 0.0, 1.0), dblBottomRadius, dblBottomRadius / 2, True))
                Else
                    '********************************
                    ' Update bottom
                    '********************************
                    Dim oBottom As Torus3d
                    oBottom = DirectCast(m_objBottoms(iCnt - 1), Torus3d)
                    oBottom.DefineByAxisCenterRadius(origin, New Vector(1.0, 0.0, 0.0), _
                                                     New Vector(0.0, 0.0, 1.0), dblBottomRadius, dblBottomRadius / 2, True)
                End If
                origin.X = origin.X + dblBottomRadius * 2.0
            Next iCnt

            ' Set the position of the body relative to the bottom
            origin.Z = origin.Z + dblBottomRadius + dblBodyRadius
            origin.X = baseWidth.PropValue.Value

            ' Remove extra bottoms
            For iCnt = m_objBottoms.Count - 1 To Convert.ToInt32(snowmenCount.PropValue.Value) Step -1
                m_objBottoms.RemoveAt(iCnt)
            Next

            '********************************************************
            ' Construct / Update the variable output body spheres
            '********************************************************
            For iCnt = 1 To Convert.ToInt32(snowmenCount.PropValue.Value)
                If m_objBodies.Count < iCnt Then
                    '********************************
                    ' Construct body
                    '********************************
                    m_objBodies.Add(New Sphere3d(oConnection, origin, dblBodyRadius, True))
                Else
                    '********************************
                    ' Update body
                    '********************************
                    Dim oBody As Sphere3d
                    oBody = DirectCast(m_objBodies(iCnt - 1), Sphere3d)
                    oBody.Radius = dblBodyRadius
                    oBody.Center() = origin
                End If
                origin.X = origin.X + dblBottomRadius * 2.0
            Next iCnt

            ' Remove extra bodies
            For iCnt = m_objBodies.Count - 1 To Convert.ToInt32(snowmenCount.PropValue.Value) Step -1
                m_objBodies.RemoveAt(iCnt)
            Next

            ' Set the position of the cone head relative to the body
            origin.Z = origin.Z + dblBodyRadius
            origin.X = baseWidth.PropValue.Value

            '********************************************************
            ' Construct / Update the variable output cone heads
            '********************************************************
            For iCnt = 1 To Convert.ToInt32(snowmenCount.PropValue.Value)
                If m_objHeads.Count < iCnt Then
                    '********************************
                    ' Construct cone head
                    '********************************
                    m_objHeads.Add(New Cone3d(oConnection, origin, _
                                              New Position(origin.X, origin.Y, origin.Z + dblHeadRadius * 2.0), _
                                              New Position(origin.X + dblHeadRadius, origin.Y, origin.Z), _
                                              New Position(origin.X, origin.Y, origin.Z + dblHeadRadius * 2.0), True))
                Else
                    '********************************
                    ' Update cone head
                    '********************************
                    Dim oHead As Cone3d
                    oHead = DirectCast(m_objHeads(iCnt - 1), Cone3d)
                    oHead.DefineBy4Pts(origin, _
                                       New Position(origin.X, origin.Y, origin.Z + dblHeadRadius * 2.0), _
                                       New Position(origin.X + dblHeadRadius, origin.Y, origin.Z), _
                                       New Position(origin.X, origin.Y, origin.Z + dblHeadRadius * 2.0), True)
                End If
                origin.X = origin.X + dblBottomRadius * 2.0
            Next iCnt

            ' Remove extra heads
            For iCnt = m_objHeads.Count - 1 To Convert.ToInt32(snowmenCount.PropValue.Value) Step -1
                m_objHeads.RemoveAt(iCnt)
            Next

        End Sub
    End Class
End Namespace
