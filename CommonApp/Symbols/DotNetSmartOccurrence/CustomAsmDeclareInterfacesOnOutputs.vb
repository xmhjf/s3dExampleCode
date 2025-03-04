Option Strict On
Imports System
Imports System.Math
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.Common.Middle.Services.Hidden

Namespace Symbols
    ' Declare:
    '   IJGenericVolume interface declared as output
    '   IJSurfaceArea interface declared as output
    '   IJUATestDotNetShere interface declared as output
    '   IJDGeometry interface declared as output
    '   IJDGeometry on assembly outputs
    <OutputNotification("{E897A3AF-E949-457C-968D-67E5DFDAD154}", DotNetSmartOccurrencesResIds.ResourceId_Volume)> _
    <OutputNotification("{2A7D37F3-0440-4ACF-A731-90CE0613ABA0}", DotNetSmartOccurrencesResIds.ResourceId_SurfaceArea)> _
    <OutputNotification("{5201D970-8E75-4245-BB3B-D932FA9CD534}", DotNetSmartOccurrencesResIds.ResourceId_SphereInfo)> _
    <OutputNotification("IJDGeometry")> _
    <OutputNotification("IJDGeometry", True)> _
    <AssemblyOutputsCopyBehavior(AssemblyOutputsCopyOptions.CopyWithAssembly)> _
    <DefaultLocalizer(DotNetSmartOccurrencesResIds.DotNetSmartOccurrencesResources, DotNetSmartOccurrencesResIds.DotNetSmartOccurrencesAssembly)>
    Public MustInherit Class CustomAsmDeclareInterfacesOnOutputsBaseClass : Inherits CustomAssemblyDefinition
    End Class

    ' Declare:
    '   non-cached symbol
    '   symbol version
    '   DoNotCopyIndependently property which should do nothing because we do not allow this on the concrete class.
    '   
    <CacheOption(CacheOptionType.NonCached)> _
    <SymbolVersion("1.0.0.0")> _
    <AssemblyOutputsCopyBehavior(AssemblyOutputsCopyOptions.DoNotCopyIndependently)> _
    Public Class CustomAsmDeclareInterfacesOnOutputs : Inherits CustomAsmDeclareInterfacesOnOutputsBaseClass
        '   1. "Part"  ( Catalog part )
        '   2. "Diameter"
        <InputCatalogPart(1)> _
        Public m_catalogPart As InputCatalogPart
        <InputDouble(2, "Diameter", "Diameter", 0.3)> _
        Public m_dblDiameter As InputDouble

        Private Const CONST_Base As String = "Base"
        Private Const CONST_Body As String = "Body"
        Private Const CONST_Head As String = "Head"

        Private Const CONST_IJUATestDotNetSphere As String = "IJUATestDotNetSphere"
        Private Const CONST_IJGenericVolume As String = "IJGenericVolume"
        Private Const IID_IJGenericVolume As String = "{E897A3AF-E949-457C-968D-67E5DFDAD154}"
        Private Const CONST_PropertyVolume As String = "Volume"
        Private Const CONST_PropertyDiameter As String = "Diameter"
        Private Const CONST_PropertyOriginX As String = "OriginX"
        Private Const CONST_PropertyOriginY As String = "OriginY"
        Private Const CONST_PropertyOriginZ As String = "OriginZ"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Base, "Base of snowman")> _
        Public m_physicalAspect As AspectDefinition

        <AssemblyOutput(1, CONST_Body)> _
        Public m_objBody As AssemblyOutput

        <AssemblyOutput(2, CONST_Head)> _
        Public m_objHead As AssemblyOutput

        ''' <summary>
        ''' Construct symbol outputs/aspects
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overrides Sub ConstructOutputs()
            ' Get the connection
            Dim oSP3DConnection As SP3DConnection
            oSP3DConnection = Me.OccurrenceConnection

            Dim dblBaseRadius As Double
            dblBaseRadius = m_dblDiameter.Value / 2.0

            '=================================================
            ' Construct base
            '=================================================
            Dim objBase As Sphere3d

            objBase = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, dblBaseRadius), dblBaseRadius, True)

            m_physicalAspect.Outputs(CONST_Base) = objBase
        End Sub

        ''' <summary>
        ''' Evaluate assembly outputs (construct them if they
        '''  do not exist / update them if they do exist)
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

            Dim oPropVal As PropertyValueDouble
            Dim origin As New Position(0.0, 0.0, 0.0)
            oPropVal = DirectCast(oBusinessObject.GetPropertyValue(CONST_IJUATestDotNetSphere, CONST_PropertyOriginX), PropertyValueDouble)
            If oPropVal.PropValue().HasValue Then
                origin.X = oPropVal.PropValue().Value
            End If
            oPropVal = DirectCast(oBusinessObject.GetPropertyValue(CONST_IJUATestDotNetSphere, CONST_PropertyOriginY), PropertyValueDouble)
            If oPropVal.PropValue().HasValue Then
                origin.Y = oPropVal.PropValue().Value
            End If
            oPropVal = DirectCast(oBusinessObject.GetPropertyValue(CONST_IJUATestDotNetSphere, CONST_PropertyOriginZ), PropertyValueDouble)
            If oPropVal.PropValue().HasValue Then
                origin.Z = oPropVal.PropValue().Value
            End If

            ' Test updating the diameter property in the Evaluate
            Dim dblBaseRadius As Double
            dblBaseRadius = m_dblDiameter.Value / 2.0
            Dim dblBodyRadius As Double
            dblBodyRadius = dblBaseRadius * 2.0 / 3.0
            Dim dblHeadRadius As Double
            dblHeadRadius = dblBaseRadius / 2.0

            ' Set where the origin of the body should be
            origin.Z = origin.Z + dblBaseRadius * 2.0 + dblBodyRadius

            '*********************************************
            ' Construct the body (if it does not exist)
            '*********************************************
            If m_objBody.Output Is Nothing Then
                '********************************
                ' Construct body
                '********************************
                m_objBody.Output = New Sphere3d(oConnection, origin, dblBodyRadius, True)
            Else
                '********************************
                ' Update body
                '********************************
                Dim oBody As Sphere3d
                oBody = DirectCast(m_objBody.Output, Sphere3d)
                oBody.Radius = dblBodyRadius
                oBody.Center() = origin
            End If

            ' Modify where the origin of the head should be
            origin.Z = origin.Z + dblBodyRadius + dblHeadRadius

            '*********************************************
            ' Construct the body (if it does not exist)
            '*********************************************
            If m_objHead.Output Is Nothing Then
                '********************************
                ' Construct body
                '********************************
                m_objHead.Output = New Sphere3d(oConnection, origin, dblHeadRadius, True)
            Else
                '********************************
                ' Update head
                '********************************
                Dim oHead As Sphere3d
                oHead = DirectCast(m_objHead.Output, Sphere3d)
                oHead.Radius = dblHeadRadius
                oHead.Center() = origin
            End If

            '*********************************
            ' Compute/Set Volume
            '*********************************
            Dim dRadius As Double
            dRadius = m_dblDiameter.Value / 2.0
            Dim dVolume As Double = PI * dRadius * dRadius * dRadius

            dRadius = dblBodyRadius
            dVolume = dVolume + PI * dRadius * dRadius * dRadius

            dRadius = dblHeadRadius
            dVolume = dVolume + PI * dRadius * dRadius * dRadius

            ' The IsInterfaceUserControlled property indicates if the property is controlled
            Dim bVolumeIsUserControlled As Boolean = IsInterfaceUserControlled(CONST_IJGenericVolume)
            Dim bVolumeIsUserControlledToo As Boolean = IsInterfaceUserControlled(IID_IJGenericVolume)
            If (bVolumeIsUserControlled <> bVolumeIsUserControlledToo) Then
                ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, "Controlled interface check return something different for Name versus IID")
            End If

            ' Set volume only if it is not user controlled
            If Not bVolumeIsUserControlled Then
                Dim oVolume As PropertyValueDouble = New PropertyValueDouble(CONST_IJGenericVolume, CONST_PropertyVolume, dVolume)
                oBusinessObject.SetPropertyValue(oVolume)
            End If

        End Sub
    End Class
End Namespace


