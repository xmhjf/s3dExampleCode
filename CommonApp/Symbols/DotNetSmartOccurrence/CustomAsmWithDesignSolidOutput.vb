Option Strict On
Imports System
Imports System.Math
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.Common.Exceptions
Imports Ingr.SP3D.Equipment.Middle
Imports Ingr.SP3D.Equipment.Middle.Services
Imports Ingr.SP3D.ReferenceData.Middle.Services

Namespace Symbols
    ' Currently, the EquipmentAssemblyDefinition is not CLS-compliant so this symbol will not
    ' be either so until this is fixed, this is set to False.
    <CLSCompliant(False)> _
    <SymbolVersion("1.0.0.0")> _
    <CacheOption(CacheOptionType.NonCached)> _
    Public Class CustomAsmWithDesignSolidOutput : Inherits EquipmentAssemblyDefinition
#Region "Declaration of inputs"
        '   1. "Part"  ( Catalog part )
        '   2. "Base Height"
        '   3. "Height"
        '   4, "Diameter"
        '   5, "HoleType" - either round or rectagular
        <InputCatalogPart(1)> _
        Public m_catalogPart As InputCatalogPart
        <InputDouble(2, "BaseHeight", "Base Height", 1.0)> _
        Public m_BaseHeight As InputDouble
        <InputDouble(3, "Height", "Height", 5.0)> _
        Public m_Height As InputDouble
        <InputDouble(4, "Diameter", "Diameter", 2.0)> _
        Public m_Diameter As InputDouble
        <InputDouble(5, "HoleType", "HoleType", 1.0)> _
        Public m_HoleType As InputDouble
        <InputDouble(6, "CapType", "CapType", 1.0)> _
        Public m_CapType As InputDouble
#End Region

#Region "Private constants"
        Private Const CONST_Base As String = "Base"
        Private Const CONST_Pole As String = "Pole"
        Private Const CONST_Cylinder As String = "Cylinder"
        Private Const CONST_Rectangle As String = "Rectangle"
        Private Const CONST_Holes As String = "Holes"
        Private Const CONST_Cap As String = "Cap"
        Private Const INTERFACE_IJUARectSolid As String = "IJUARectSolid"
        Private Const RECTANGLE_PROPERTY_Height As String = "A"
        Private Const RECTANGLE_PROPERTY_Length As String = "B"
        Private Const RECTANGLE_PROPERTY_Width As String = "C"
        Private Const INTERFACE_IJUACylinder As String = "IJUACylinder"
        Private Const CYLINDER_PROPERTY_Height As String = "A"
        Private Const CYLINDER_PROPERTY_Diameter As String = "B"
        Private Const INTERFACE_IJUASemiElliptical As String = "IJUASemiElliptical"
        Private Const ELLIPTICALHEAD_PROPERTY_Diameter As String = "A"
        Private Const ELLIPTICALHEAD_PROPERTY_Height As String = "B"
        Private Const INTERFACE_IJUACone As String = "IJUACone"
        Private Const CONEHEAD_PROPERTY_Height As String = "A"
        Private Const CONEHEAD_PROPERTY_BottomDiameter As String = "B"
        Private Const CONEHEAD_PROPERTY_TopDiameter As String = "C"
#End Region

#Region "Declaration of Symbol outputs"
        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Base, "Base of pole")> _
        Public m_physicalAspect As AspectDefinition
#End Region

#Region "Declaration of Assembly Outputs"
        <AssemblyOutput(1, CONST_Pole)> _
        Public m_objColumn As AssemblyOutput
        <AssemblyOutput(2, CONST_Cylinder)> _
        Public m_objCylinder As AssemblyOutput
        <AssemblyOutput(3, CONST_Rectangle)> _
        Public m_objRectangle As AssemblyOutput
        ' Declare a variable output for the hole shapes
        <AssemblyOutput(4, CONST_Holes)> _
        Public m_objHoles As AssemblyOutputs
        ' Cap the pole
        <AssemblyOutput(5, CONST_Cap)> _
        Public m_objCap As AssemblyOutput
#End Region

#Region "Symbol construction"
        ''' <summary>
        ''' Construct the symbol outputs
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overrides Sub ConstructOutputs()
            Dim dBaseHeight As Double = m_BaseHeight.Value
            Dim dDiameter As Double = m_Diameter.Value
            Dim oCone As Cone3d = New Cone3d(Me.OccurrenceConnection, New Position(0, 0, 0), _
                                             New Position(0.0, 0.0, dBaseHeight), _
                                             New Position(dDiameter / 2.0 * 1.25, 0.0, 0.0), _
                                             New Position(dDiameter / 2.0, 0.0, dBaseHeight), True)
            m_physicalAspect.Outputs(CONST_Base) = oCone
        End Sub
#End Region

#Region "Construct / Evaluate assembly outputs"
        ''' <summary>
        ''' Evaluate assembly outputs
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub EvaluateAssembly()
            Dim oSystemParent As ISystem
            oSystemParent = DirectCast(Me.Occurrence, ISystem)

            ' Create the solid (if required) and return the solid object
            Dim oSolid As DesignSolid = CreateAndRetrieveSolid(oSystemParent)

            ' We still have our solid output?
            If oSolid Is Nothing Then
                ' Define a ToDo record
                Throw New SymbolErrorException("DotNetCustomAsmErrorCodeList", 2, Nothing, "Solid output is missing!", Nothing)
            End If

            ' Compute the origin from the symbol base height 
            Dim oOrigin As Position = New Position(0.0, 0.0, m_BaseHeight.Value)

            ' Create the cylinder (or just retrieve it if already created)
            Dim oCylinder As GenericShape = CreateAndRetrieveCylinderShape(DirectCast(oSolid, ISystem))

            ' Create the rectangle (or just retrieve it if already created)
            Dim oRectangle As GenericShape = CreateAndRetrieveRectangleShape(DirectCast(oSolid, ISystem))

            ' Set dimensions of cylinder
            Dim dDiameter As Double = m_Diameter.Value
            Dim dHeight As Double = m_Height.Value
            oCylinder.SetPropertyValue(dDiameter, INTERFACE_IJUACylinder, CYLINDER_PROPERTY_Diameter)
            oCylinder.SetPropertyValue(dHeight, INTERFACE_IJUACylinder, CYLINDER_PROPERTY_Height)

            ' Set dimensions of rectangle
            Dim holeDiameter As Double = dDiameter * 0.6
            oRectangle.SetPropertyValue(holeDiameter, INTERFACE_IJUARectSolid, RECTANGLE_PROPERTY_Length)
            oRectangle.SetPropertyValue(holeDiameter, INTERFACE_IJUARectSolid, RECTANGLE_PROPERTY_Width)
            oRectangle.SetPropertyValue(dHeight, INTERFACE_IJUARectSolid, RECTANGLE_PROPERTY_Height)

            ' Make sure we have at least two shapes
            Dim oChildren As List(Of DesignSolidChild) = oSolid.DesignSolidChildren
            If oChildren.Count < 2 Then
                ' Define a ToDo record
                Throw New SymbolErrorException("DotNetCustomAsmErrorCodeList", 2, Nothing, "Solid is missing its Shapes", Nothing)
            End If

            ' Don't care about the order of the shapes other than the first must be the cylinder and the
            ' second the rectangle
            Dim oNewChildrenList As List(Of DesignSolidChild)
            oNewChildrenList = New List(Of DesignSolidChild)(oChildren.Count)
            oNewChildrenList.Add(New DesignSolidChild(oCylinder, DesignSolidOperationType.Add))
            oNewChildrenList.Add(New DesignSolidChild(oRectangle, DesignSolidOperationType.Subtract))

            ' Retrieve the cap if not removed by the user
            Dim bIsEllipticalHead As Boolean = True
            Dim oCap As GenericShape = CreateAndRetrieveCap(DirectCast(oSolid, ISystem), bIsEllipticalHead)
            If Not oCap Is Nothing Then
                If bIsEllipticalHead Then
                    ' Size the head
                    oCap.SetPropertyValue(dDiameter, INTERFACE_IJUASemiElliptical, ELLIPTICALHEAD_PROPERTY_Diameter)
                    ' 4:1 elliptical head
                    oCap.SetPropertyValue(dDiameter / 4.0, INTERFACE_IJUASemiElliptical, ELLIPTICALHEAD_PROPERTY_Height)
                Else
                    ' Size the head
                    oCap.SetPropertyValue(dDiameter, INTERFACE_IJUACone, CONEHEAD_PROPERTY_BottomDiameter)
                    oCap.SetPropertyValue(dDiameter * 0.75, INTERFACE_IJUACone, CONEHEAD_PROPERTY_TopDiameter)
                    oCap.SetPropertyValue(dDiameter / 4.0, INTERFACE_IJUACone, CONEHEAD_PROPERTY_Height)
                End If
                Dim oCylinderOrigin As Position = oCylinder.Origin
                oCylinderOrigin.Z = oCylinderOrigin.Z + m_Height.Value
                oCap.Origin = oCylinderOrigin
                ' Add the elliptical head to the top
                oNewChildrenList.Add(New DesignSolidChild(oCap, DesignSolidOperationType.Add))
            End If

            ' Compute the number of holes needed
            Dim neededHoles As Integer = Convert.ToInt32((m_Height.Value - 0.4 - holeDiameter) / (holeDiameter + 0.2)) * 2

            'Anything change?
            Dim bHasAnythingChanged As Boolean = True

            ' Do we have any holes?
            Dim holeCount As Integer = m_objHoles.Count
            If holeCount <> 0 Then
                Dim oShape As GenericShape = DirectCast(m_objHoles(0), GenericShape)
                Dim bIsRound As Boolean
                If oShape.SupportsInterface(INTERFACE_IJUACylinder) Then
                    bIsRound = True
                End If

                ' Check if the type of hole changed - if yes then we need to empty the entire collection
                If (bIsRound And m_HoleType.Value = 2.0) Or (Not bIsRound And m_HoleType.Value = 1.0) Then
                    m_objHoles.Clear() ' Remove all objects
                    holeCount = 0
                End If

                If holeCount = neededHoles Then
                    bHasAnythingChanged = False
                End If

                ' Remove the unnecessary holes
                For iCnt As Integer = 0 To holeCount - neededHoles - 1 Step 1
                    m_objHoles.RemoveAt(neededHoles)
                Next

            End If

            If bHasAnythingChanged Then
                Dim sPartNumber As String
                If m_HoleType.Value = 1.0 Then
                    sPartNumber = "RtCircularCylinder 001"
                Else
                    sPartNumber = "RectangularSolid 001"
                End If

                ' Add new holes
                For iCnt As Integer = 0 To neededHoles - holeCount - 1 Step 1
                    m_objHoles.Add(New GenericShape(sPartNumber, oSolid))
                Next

                Dim radius As Double = m_Diameter.Value / 2.0
                Dim dStartPos As Double = radius + 0.2

                ' Get orientation of equipment
                Dim oMatrix As Matrix4X4 = New Matrix4X4()
                Dim oEqpMatrix As Matrix4X4 = DirectCast(Me.Occurrence, ILocalCoordinateSystem).Matrix

                ' Position the holes
                For iCnt As Integer = 0 To m_objHoles.Count - 2 Step 2
                    Dim oShape As GenericShape = DirectCast(m_objHoles(iCnt), GenericShape)
                    oShape.Origin = New Position(0.0, 0.0, 0.0)
                    oShape.SetOrientation(New Vector(1.0, 0.0, 0.0), New Vector(0.0, 1.0, 0.0))
                    oShape.Origin = New Position(oOrigin.X - radius - 0.05, oOrigin.Y, oOrigin.Z + dStartPos + (0.2 + holeDiameter) * iCnt / 2.0)
                    oShape.Transform(oMatrix)
                    oShape.Transform(oEqpMatrix)

                    If m_HoleType.Value = 1.0 Then
                        ' Set dimensions of cylinder
                        oShape.SetPropertyValue(m_Diameter.Value * 0.6, INTERFACE_IJUACylinder, CYLINDER_PROPERTY_Diameter)
                        oShape.SetPropertyValue(m_Diameter.Value + 0.1, INTERFACE_IJUACylinder, CYLINDER_PROPERTY_Height)
                    Else
                        ' Set dimensions of rectangle
                        oShape.SetPropertyValue(m_Diameter.Value * 0.6, INTERFACE_IJUARectSolid, RECTANGLE_PROPERTY_Length)
                        oShape.SetPropertyValue(m_Diameter.Value * 0.6, INTERFACE_IJUARectSolid, RECTANGLE_PROPERTY_Width)
                        oShape.SetPropertyValue(m_Diameter.Value + 0.1, INTERFACE_IJUARectSolid, RECTANGLE_PROPERTY_Height)
                    End If

                    oShape = DirectCast(m_objHoles(iCnt + 1), GenericShape)
                    oShape.Origin = New Position(0.0, 0.0, 0.0)
                    oShape.SetOrientation(New Vector(1.0, 0.0, 0.0), New Vector(0.0, 1.0, 0.0))
                    oShape.Origin = New Position(oOrigin.X - radius - 0.05, oOrigin.Y, oOrigin.Z + dStartPos + (0.2 + holeDiameter) * iCnt / 2.0)
                    oMatrix.Rotate(Math.PI / 2.0, New Vector(0.0, 0.0, 1.0))
                    oShape.Transform(oMatrix)
                    oShape.Transform(oEqpMatrix)

                    If m_HoleType.Value = 1.0 Then
                        ' Set dimensions of cylinder
                        oShape.SetPropertyValue(m_Diameter.Value * 0.6, INTERFACE_IJUACylinder, CYLINDER_PROPERTY_Diameter)
                        oShape.SetPropertyValue(m_Diameter.Value + 0.1, INTERFACE_IJUACylinder, CYLINDER_PROPERTY_Height)
                    Else
                        ' Set dimensions of rectangle
                        oShape.SetPropertyValue(m_Diameter.Value * 0.6, INTERFACE_IJUARectSolid, RECTANGLE_PROPERTY_Length)
                        oShape.SetPropertyValue(m_Diameter.Value * 0.6, INTERFACE_IJUARectSolid, RECTANGLE_PROPERTY_Width)
                        oShape.SetPropertyValue(m_Diameter.Value + 0.1, INTERFACE_IJUARectSolid, RECTANGLE_PROPERTY_Height)
                    End If

                    ' Rotate shapes 90 degrees
                    oMatrix.Rotate(Math.PI / 4.0, New Vector(0.0, 0.0, 1.0))
                Next

                ' Add all object to the list and set them on the solid
                For iCnt As Integer = 0 To m_objHoles.Count - 1 Step 1
                    oNewChildrenList.Add(New DesignSolidChild(DirectCast(m_objHoles(iCnt), GenericShape), DesignSolidOperationType.Subtract))
                Next

                ' Set the operator order
                oSolid.DesignSolidChildren = oNewChildrenList
            End If

            ' Block to test ToDoRecords (for middle tier ATP and client tier test cmd)
            Dim oEquipmentForToDoRecordTest As Equipment.Middle.Equipment = DirectCast(Me.Occurrence, Equipment.Middle.Equipment)
            Dim sNameOfEquipmentForToDoRecordTest As String
            sNameOfEquipmentForToDoRecordTest = oEquipmentForToDoRecordTest.Name
            ' For middle tier ATP
            If (sNameOfEquipmentForToDoRecordTest = "Equip_CreateTDRForMiddleTierATP") Then
                Throw New SymbolErrorException("DotNetCustomAsmErrorCodeList", 10, Nothing, "ToDoRecord:Equipment name is invalid (MiddleTier ATP)", Nothing)
            Else
                ' For client tier test cmd
                If (sNameOfEquipmentForToDoRecordTest = "Equip_CreateTDRForClientTierTest") Then
                    Throw New SymbolErrorException("DotNetCustomAsmErrorCodeList", 11, Nothing, "ToDoRecord:Equipment name is invalid (ClientTier Test Cmd)", Nothing)
                End If
            End If
        End Sub
#End Region

#Region "Implement property Management methods"
        ''' <summary>
        ''' Preload property pages
        ''' </summary>
        ''' <param name="oBusinessObject"></param>
        ''' <param name="boProperties"></param>
        ''' <remarks></remarks>
        Public Overrides Sub OnPreLoad(ByVal oBusinessObject As BusinessObject, ByVal boProperties As ReadOnlyCollection(Of PropertyDescriptor))
            For Each inputProperty As PropertyDescriptor In boProperties
                If (inputProperty.Property.PropertyInfo.Name = "UseDefaults") Then
                    Dim bUseDefaults As Boolean = DirectCast(inputProperty.Property, PropertyValueBoolean).PropValue.Value
                    FindDescriptor(boProperties, "BaseHeight").ReadOnly = bUseDefaults
                End If
            Next
        End Sub

        ''' <summary>
        ''' Method to verify a custom assembly or output changed property
        ''' </summary>
        ''' <param name="oBusinessObject"></param>
        ''' <param name="CollAllDisplayedValues"></param>
        ''' <param name="oPropToChange"></param>
        ''' <param name="oNewPropValue"></param>
        ''' <param name="strErrorMsg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overrides Function OnPropertyChange(ByVal oBusinessObject As BusinessObject, _
                                  ByVal CollAllDisplayedValues As ReadOnlyCollection(Of PropertyDescriptor), _
                                  ByVal oPropToChange As PropertyDescriptor, _
                                  ByVal oNewPropValue As PropertyValue, _
                                  ByRef strErrorMsg As String) As Boolean
            Dim bSucceeded As Boolean = True
            If oPropToChange.Property.PropertyInfo.Name = "UseDefaults" Then
                Dim bUseDefaults As Boolean = DirectCast(oNewPropValue, PropertyValueBoolean).PropValue.Value
                If bUseDefaults Then
                    DirectCast(FindDescriptor(CollAllDisplayedValues, "BaseHeight").Property, PropertyValueDouble).PropValue = 1.0
                End If
                FindDescriptor(CollAllDisplayedValues, "BaseHeight").ReadOnly = bUseDefaults
            ElseIf oPropToChange.Property.PropertyInfo.Name = "BaseHeight" Then
                Dim baseHeight As Double = DirectCast(oNewPropValue, PropertyValueDouble).PropValue.Value
                If baseHeight < 0.130001 Then
                    strErrorMsg = "Base height is  too small - enter a value greater than 0.13."
                    bSucceeded = False
                ElseIf baseHeight > 12.999999 Then
                    strErrorMsg = "Base height is too big - enter a value less than 13.0."
                    bSucceeded = False
                End If
            End If

            Return bSucceeded
        End Function
#End Region

#Region "Protected methods"
        Protected Function FindDescriptor(ByVal boProperties As ReadOnlyCollection(Of PropertyDescriptor), ByVal propertyName As String) As PropertyDescriptor
            For Each propDescriptor In boProperties
                If propDescriptor.Property.PropertyInfo.Name = propertyName Then
                    Return propDescriptor
                End If
            Next

            Throw New CmnPropertyInfoNotAvailableException()
        End Function
#End Region

#Region "Private Methods /Properties"
        Private Function CreateAndRetrieveSolid(ByVal oParent As ISystem) As DesignSolid
            Dim oSolid As DesignSolid

            ' Check if the output was already created and was not deleted by the user
            If m_objColumn.Output Is Nothing Then
                '=================================================
                ' Construct Solid
                '=================================================
                oSolid = New DesignSolid(oParent)

                ' Default the material
                Dim oSiteMgr As SiteManager = MiddleServiceProvider.SiteMgr
                Dim oPlantModel As Plant = oSiteMgr.ActiveSite.ActivePlant
                Dim oCatalog As Catalog = oPlantModel.PlantCatalog
                Dim oCatStructHelper As CatalogStructHelper = New CatalogStructHelper(DirectCast(oCatalog, SP3DConnection))
                oSolid.Material = oCatStructHelper.GetMaterial("Stone", "A")

                m_objColumn.Output = oSolid ' Make this our solid output
            Else
                oSolid = DirectCast(m_objColumn.Output, DesignSolid)
            End If

            Return oSolid
        End Function

        Private Function CreateAndRetrieveCylinderShape(ByVal oSolid As ISystem) As GenericShape
            Dim oCylinder As GenericShape = Nothing

            ' Have we constructed the cylinder yet
            If m_objCylinder.Output = Nothing Then
                '=================================================
                ' Construct Cylinder
                '=================================================
                oCylinder = New GenericShape("RtCircularCylinder 001", oSolid, False)
                m_objCylinder.Output = oCylinder ' Make this an output so we can control the delete behavior
                oCylinder.Origin = New Position(0.0, 0.0, m_BaseHeight.Value)
                ' Define the orientation of the shape
                oCylinder.SetOrientation(New Vector(0.0, 0.0, 1.0), New Vector(1.0, 0.0, 0.0))
            Else
                oCylinder = DirectCast(m_objCylinder.Output, GenericShape)
            End If

            Return oCylinder
        End Function

        Private Function CreateAndRetrieveRectangleShape(ByVal oSolid As ISystem) As GenericShape
            Dim oRectangle As GenericShape

            ' Have we constructed the rectangle yet?
            If m_objRectangle.Output = Nothing Then
                '=================================================
                ' Construct Rectangle
                '=================================================
                oRectangle = New GenericShape("RectangularSolid 001", oSolid, False)
                m_objRectangle.Output = oRectangle ' Make this an output so we can control the delete behavior
                ' Define origin of shape
                oRectangle.Origin = New Position(0.0, 0.0, m_BaseHeight.Value)
                ' Define the orientation of the shape
                oRectangle.SetOrientation(New Vector(0.0, 0.0, 1.0), New Vector(1.0, 0.0, 0.0))
            Else
                oRectangle = DirectCast(m_objRectangle.Output, GenericShape)
            End If

            Return oRectangle
        End Function

        Private Function CreateAndRetrieveCap(ByVal oSolid As ISystem, ByRef bIsEllipticalHead As Boolean) As GenericShape
            Dim oHead As GenericShape = Nothing

            ' Have we constructed the elliptical head yet. If yes, then make certain
            ' it is the correct type (i.e., removing it if it is not).
            If Not m_objCap.Output Is Nothing Then
                oHead = DirectCast(m_objCap.Output, GenericShape)
                bIsEllipticalHead = oHead.SupportsInterface(INTERFACE_IJUASemiElliptical)
                If m_CapType.Value = 1.0 Then
                    ' Expect an elliptical head (if not we need to replace it with a conical head)
                    If Not bIsEllipticalHead Then
                        ' Need to remove the conical head so we can add a new one
                        m_objCap.Delete()
                    End If
                Else
                    If bIsEllipticalHead Then
                        ' Need to remove the elliptical head so we can add a conical one
                        m_objCap.Delete()
                    End If
                End If
            End If

            ' Verify the cap was not removed by the user
            If Not m_objCap.HasBeenDeletedByUser And m_objCap.Output Is Nothing Then
                If m_CapType.Value = 1.0 Then
                    '=================================================
                    ' Construct the elliptical head
                    '=================================================
                    oHead = New GenericShape("SemiEllipticalHead 001", oSolid, True)
                    ' Define origin of shape
                    oHead.Origin = New Position(0.0, 0.0, m_BaseHeight.Value + m_Height.Value)
                    ' Define the orientation of the shape
                    oHead.SetOrientation(New Vector(0.0, 0.0, 1.0), New Vector(1.0, 0.0, 0.0))
                    m_objCap.Output = oHead ' Make this an output so we can control the delete behavior
                    m_objCap.CanDeleteIndependently = True
                    bIsEllipticalHead = True
                Else
                    '=================================================
                    ' Construct the conical head
                    '=================================================
                    oHead = New GenericShape("RtCircularCone 001", oSolid, True)
                    ' Define origin of shape
                    oHead.Origin = New Position(0.0, 0.0, m_BaseHeight.Value + m_Height.Value)
                    ' Define the orientation of the shape
                    oHead.SetOrientation(New Vector(0.0, 0.0, 1.0), New Vector(1.0, 0.0, 0.0))
                    m_objCap.Output = oHead ' Make this an output so we can control the delete behavior
                    m_objCap.CanDeleteIndependently = True
                    bIsEllipticalHead = False
                End If
            End If

            Return oHead
        End Function
#End Region
    End Class

    ' Currently, the EquipmentAssemblyDefinition is not CLS-compliant so this symbol will not
    ' be either so until this is fixed, this is set to False.
    <CLSCompliant(False)> _
    <SymbolVersion("1.0.0.0")> _
    <CacheOption(CacheOptionType.NonCached)> _
    Public Class EquipmentWithAttributeManagement : Inherits CustomAsmWithDesignSolidOutput

#Region "Implement property Management methods"
        ''' <summary>
        ''' Preload property pages
        ''' </summary>
        ''' <param name="oBusinessObject"></param>
        ''' <param name="boProperties"></param>
        ''' <remarks></remarks>
        Public Overrides Sub OnPreLoad(ByVal oBusinessObject As BusinessObject, ByVal boProperties As ReadOnlyCollection(Of PropertyDescriptor))
            For Each inputProperty As PropertyDescriptor In boProperties
                If (inputProperty.Property.PropertyInfo.Name = "UseDefaults") Then
                    Dim bUseDefaults As Boolean = DirectCast(inputProperty.Property, PropertyValueBoolean).PropValue.Value
                    FindDescriptor(boProperties, "BaseHeight").ReadOnly = bUseDefaults
                    FindDescriptor(boProperties, "Height").ReadOnly = bUseDefaults
                    FindDescriptor(boProperties, "Diameter").ReadOnly = bUseDefaults
                End If
            Next
        End Sub

        ''' <summary>
        ''' Method to verify a custom assembly or output changed property
        ''' </summary>
        ''' <param name="oBusinessObject"></param>
        ''' <param name="CollAllDisplayedValues"></param>
        ''' <param name="oPropToChange"></param>
        ''' <param name="oNewPropValue"></param>
        ''' <param name="strErrorMsg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overrides Function OnPropertyChange(ByVal oBusinessObject As BusinessObject, _
                                  ByVal CollAllDisplayedValues As ReadOnlyCollection(Of PropertyDescriptor), _
                                  ByVal oPropToChange As PropertyDescriptor, _
                                  ByVal oNewPropValue As PropertyValue, _
                                  ByRef strErrorMsg As String) As Boolean
            Dim bSucceeded As Boolean = True
            If oPropToChange.Property.PropertyInfo.Name = "UseDefaults" Then
                Dim bUseDefaults As Boolean = DirectCast(oNewPropValue, PropertyValueBoolean).PropValue.Value
                If bUseDefaults Then
                    DirectCast(FindDescriptor(CollAllDisplayedValues, "BaseHeight").Property, PropertyValueDouble).PropValue = 1.0
                    DirectCast(FindDescriptor(CollAllDisplayedValues, "Height").Property, PropertyValueDouble).PropValue = 5.0
                    DirectCast(FindDescriptor(CollAllDisplayedValues, "Diameter").Property, PropertyValueDouble).PropValue = 2.0
                End If
                FindDescriptor(CollAllDisplayedValues, "BaseHeight").ReadOnly = bUseDefaults
                FindDescriptor(CollAllDisplayedValues, "Height").ReadOnly = bUseDefaults
                FindDescriptor(CollAllDisplayedValues, "Diameter").ReadOnly = bUseDefaults
            ElseIf oPropToChange.Property.PropertyInfo.Name = "BaseHeight" Then
                Dim baseHeight As Double = DirectCast(oNewPropValue, PropertyValueDouble).PropValue.Value
                If baseHeight < 0.130001 Then
                    strErrorMsg = "Dude, that base height is way too small - enter something greater than .13."
                    bSucceeded = False
                ElseIf baseHeight > 12.999999 Then
                    strErrorMsg = "Man, that base height is too big - pick something less than 13."
                    bSucceeded = False
                End If
            ElseIf oPropToChange.Property.PropertyInfo.Name = "Height" Then
                Dim height As Double = DirectCast(oNewPropValue, PropertyValueDouble).PropValue.Value
                If height < 5.0 Then
                    strErrorMsg = "Guy, that height is way too small - enter a value 5.0 or greater."
                    bSucceeded = False
                ElseIf height > 12.999999 Then
                    strErrorMsg = "Boy, that height is a bit too big - enter a value less than 13.0"
                    bSucceeded = False
                End If
            ElseIf oPropToChange.Property.PropertyInfo.Name = "Diameter" Then
                Dim diameter As Double = DirectCast(oNewPropValue, PropertyValueDouble).PropValue.Value
                If diameter < 0.130001 Then
                    strErrorMsg = "Girl, that diameter should be greater than .13."
                    bSucceeded = False
                ElseIf diameter > 12.999999 Then
                    strErrorMsg = "Bro, that diameter is too big - enter a value less than 13.0."
                    bSucceeded = False
                End If
            End If

            Return bSucceeded
        End Function
#End Region

    End Class


    <RuleVersion("1.0.0.0")> _
    <RuleInterface("IEquipmentParameterRule", "Parameter rule interface for equipment test column")> _
    Public Class EquipmentTestColumnParameterRule : Inherits ParameterRule

        <ControlledParameter(1, "Diameter", "Diameter")> _
        Public m_Diameter As ControlledParameterDouble
        <ControlledParameter(2, "Height", "Height")> _
        Public m_Height As ControlledParameterDouble

        Public Overrides Sub Evaluate()
            ' If the value for the diameter is ever greter than or equal to 13 then produce a ToDo error
            If m_Diameter.Value > 12.999999 Then
                ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, "Parameter rule does not allow the diameter to be greater than or equal to 13.")
            Else
                m_Diameter.Value = 1.5
            End If

            m_Height.Value = 7.0
        End Sub
    End Class


    '''
    ''' Symbol class to test optional inputs whose value comes from this definition
    ''' 
    ''' Note: the EquipmentAssemblyDefinition is not CLS-compliant so this symbol will not
    ''' be either so until this is fixed, this is set to False.
    <CLSCompliant(False)> _
    <SymbolVersion("1.0.0.0")> _
    <CacheOption(CacheOptionType.NonCached)> _
    Public Class CustomAsmWithOnlyOptionalInputs : Inherits EquipmentAssemblyDefinition
#Region "Declaration of inputs"
        '   1. "Part"  ( Catalog part )
        <InputCatalogPart(1)> _
        Public m_catalogPart As InputCatalogPart
        <InputDouble(2, "CubeDimension", "Cube dimension", 1.314159)> _
        Public m_CubeDimension As InputDouble
#End Region

#Region "Declaration of Symbol outputs"
        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput("Cube", "Cube")> _
        Public m_physicalAspect As AspectDefinition
#End Region

#Region "Symbol construction"
        ''' <summary>
        ''' Construct the symbol outputs
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overrides Sub ConstructOutputs()
            Dim cubeDimension As Double = m_CubeDimension.Value

            Dim collectionOfPoints As Collection(Of Position) = New Collection(Of Position)()
            collectionOfPoints.Add(New Position(0.0, 0.0, 0.0))
            collectionOfPoints.Add(New Position(cubeDimension, 0.0, 0.0))
            collectionOfPoints.Add(New Position(cubeDimension, cubeDimension, 0.0))
            collectionOfPoints.Add(New Position(0.0, cubeDimension, 0.0))
            collectionOfPoints.Add(New Position(0.0, 0.0, 0.0))
            Dim projection As Projection3d = New Projection3d(Me.OccurrenceConnection, New LineString3d(collectionOfPoints), New Vector(0, 0, 1), cubeDimension, True)
            m_physicalAspect.Outputs("Cube") = projection
        End Sub
#End Region

    End Class
End Namespace

