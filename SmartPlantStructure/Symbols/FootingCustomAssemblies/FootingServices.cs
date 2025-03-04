using Ingr.SP3D.Common.Exceptions;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Structure.Middle.Services;
//----------------------------------------------------------------------------------------------------------------
//Copyright 1992 - 2013 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  FootingServices.cs
//
//Abstract
//	FootingServices is a helper class to have common method implementation for .NET Footing custom assemblies.
//----------------------------------------------------------------------------------------------------------------
using System;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Enumerated values for FootingComponent type.
    /// </summary>
    internal enum FootingComponentType
    {
        Grout = 0,
        Pier = 1,
        Slab = 2
    }

    /// <summary>
    /// Helper class to have common method implementation for .NET Footing custom assemblies.
    /// </summary>
    internal static class FootingServices
    {
        internal const string UserDefined = "User Defined";
        internal const string Clearance = "Clearance";

        public static CatalogStructHelper CatalogStructHelper = new CatalogStructHelper();

        /// <summary>
        /// Get the footing component part name.
        /// </summary>
        /// <param name="footing">The footing.</param>
        /// <param name="footingType">Type of the footing.</param>
        /// <returns>The footing component part name.</returns>
        internal static string GetFootingComponentPartName(Footing footing, FootingComponentType footingType)
        {
            string componentPartName = string.Empty;

            if (footing == null)
            {
                throw new CmnArgumentNullException("footing");
            }
            //depending upon the footing component name we get the component part name
            switch (footingType)
            {
                case FootingComponentType.Grout:
                    componentPartName = StructHelper.GetStringProperty(footing, SPSSymbolConstants.IJUASPSPierAndSlabFootingAsm, SPSSymbolConstants.GroutComponent);
                    break;
                case FootingComponentType.Pier:
                    componentPartName = StructHelper.GetStringProperty(footing, SPSSymbolConstants.IJUASPSPierAndSlabFootingAsm, SPSSymbolConstants.PierComponent);
                    break;
                case FootingComponentType.Slab:
                    componentPartName = StructHelper.GetStringProperty(footing, SPSSymbolConstants.IJUASPSPierAndSlabFootingAsm, SPSSymbolConstants.SlabComponent);
                    break;
            }

            return componentPartName;
        }

        /// <summary>
        /// Places the footing shape.
        /// </summary>
        /// <param name="connection">The connection.</param>
        /// <param name="shape">The shape.</param>
        /// <param name="componentLength">Length of the component.</param>
        /// <param name="componentWidth">Width of the component.</param>
        /// <param name="componentHeight">Height of the component.</param>
        /// <param name="clearance">The clearance.</param>
        /// <param name="rotationAngle">The rotation angle.</param>
        /// <param name="orientation">The orientation.</param>
        /// <param name="memberAngle">The member angle.</param>
        /// <param name="footingType">Type of the footing.</param>
        /// <param name="projectionOffset">The projection offset.</param>
        /// <param name="aspect">The aspect.</param>
        /// <exception cref="CmnArgumentNullException">connection</exception>
        /// <exception cref="InvalidOperationException"></exception>
        /// <exception cref="CmnArgumentNullException">connection</exception>
        internal static void PlaceFootingShape(SP3DConnection connection, int shape, double componentLength, double componentWidth, double componentHeight, double clearance, double rotationAngle, int orientation, double memberAngle, FootingComponentType footingType, double projectionOffset, AspectDefinition aspect)
        {
            if (connection == null)
            {
                throw new CmnArgumentNullException("connection");
            }
            if (aspect == null)
            {
                throw new CmnArgumentNullException("aspect");
            }

            IPlane clippingPlane = null;
            //Get the dimensions for this footing
            Collection<ISurface> surfaces = null;

            if (shape == SPSSymbolConstants.SHAPE_RECTANGULAR)
            {
                surfaces = SymbolHelper.CreateRectangularSolid(connection, componentWidth, componentLength, componentHeight, projectionOffset, clippingPlane);
            }
            else if (shape == SPSSymbolConstants.SHAPE_CIRCULAR)
            {
                surfaces = SymbolHelper.CreateCircularSolid(connection, projectionOffset, componentWidth, componentHeight, clippingPlane);
            }
            else if (shape == SPSSymbolConstants.SHAPE_OCTAGONAL)
            {
                surfaces = SymbolHelper.CreateOctagonalSolid(connection, projectionOffset, componentWidth, componentHeight, clippingPlane);
            }

            if (surfaces == null || surfaces.Count == 0)
            {
                throw new InvalidOperationException(FootingLocalizer.GetString(FootingResourceIDs.ErrNoOutput,
                    "No outputs can be created for the footing on given inputs."));
            }
            Projection3d projectedObject = (Projection3d)surfaces[0];

            //If with clipping plane, need to add multiple surface.
            Matrix4X4 matrix = new Matrix4X4();
            Vector vector = new Vector(0, 0, 1);

            //rotating on a local matrix as the symbol machinery will transform the footing to the global space
            if (orientation == SPSSymbolConstants.ORIENTATION_GLOBAL)
            {
                matrix.Rotate(rotationAngle, vector);
            }
            else if (orientation == SPSSymbolConstants.ORIENTATION_LOCAL)
            {
                matrix.Rotate(rotationAngle + memberAngle, vector);
            }

            projectedObject.Transform(matrix);

            aspect.Outputs.Add(footingType.ToString(), projectedObject);
        }

        /// <summary>
        /// Gets the dimensions.
        /// </summary>
        /// <param name="footingComponent">The footing component.</param>
        /// <param name="footingType">Type of the footing.</param>
        /// <param name="length">The length.</param>
        /// <param name="width">The width.</param>
        /// <param name="height">The height.</param>
        /// <param name="clearance">The clearance.</param>
        /// <param name="shape">The shape.</param>
        internal static void GetDimensions(FoundationBase footingComponent, FootingComponentType footingType, ref double length, ref double width, ref double height, ref double clearance, ref int shape)
        {
            if (footingComponent == null)
            {
                throw new CmnArgumentNullException("footingComponent");
            }
            if (!Enum.IsDefined(typeof(FootingComponentType), footingType))
            {
                throw new IndexOutOfRangeException("footingType");
            }

            string componentInterface = null;
            string componentDimensionInterface = null;
            string lengthProperty = null;
            string widthProperty = null;
            string heightProperty = null;
            string clearanceProperty = null;
            string shapeProperty = null;

            //Retrieve the component interfaces and property names based on the type
            switch (footingType)
            {
                case FootingComponentType.Grout:
                    componentInterface = SPSSymbolConstants.IJUASPSFtgGroutPad;
                    componentDimensionInterface = SPSSymbolConstants.IJUASPSFtgGroutPadDim;
                    lengthProperty = SPSSymbolConstants.GroutLength;
                    widthProperty = SPSSymbolConstants.GroutWidth;
                    heightProperty = SPSSymbolConstants.GroutHeight;
                    clearanceProperty = SPSSymbolConstants.GroutEdgeClearance;
                    shapeProperty = SPSSymbolConstants.GroutShape;
                    break;
                case FootingComponentType.Pier:
                    componentInterface = SPSSymbolConstants.IJUASPSFtgPier;
                    componentDimensionInterface = SPSSymbolConstants.IJUASPSFtgPierDim;
                    lengthProperty = SPSSymbolConstants.PierLength;
                    widthProperty = SPSSymbolConstants.PierWidth;
                    heightProperty = SPSSymbolConstants.PierHeight;
                    clearanceProperty = SPSSymbolConstants.PierEdgeClearance;
                    shapeProperty = SPSSymbolConstants.PierShape;
                    break;
                default:
                    componentInterface = SPSSymbolConstants.IJUASPSFtgSlab;
                    componentDimensionInterface = SPSSymbolConstants.IJUASPSFtgSlabDim;
                    lengthProperty = SPSSymbolConstants.SlabLength;
                    widthProperty = SPSSymbolConstants.SlabWidth;
                    heightProperty = SPSSymbolConstants.SlabHeight;
                    clearanceProperty = SPSSymbolConstants.SlabEdgeClearance;
                    shapeProperty = SPSSymbolConstants.SlabShape;
                    break;
            }

            //Use the retrieved interface and property names to get the component dimensions
            length = StructHelper.GetDoubleProperty(footingComponent, componentDimensionInterface, lengthProperty);
            width = StructHelper.GetDoubleProperty(footingComponent, componentDimensionInterface, widthProperty);
            height = StructHelper.GetDoubleProperty(footingComponent, componentDimensionInterface, heightProperty);
            clearance = StructHelper.GetDoubleProperty(footingComponent, componentInterface, clearanceProperty);
            shape = StructHelper.GetIntProperty(footingComponent, componentInterface, shapeProperty);
        }

        /// <summary>
        /// Gets the dimensions.
        /// </summary>
        /// <param name="footingComponent">The footing component.</param>
        /// <param name="footingType">Type of the footing.</param>
        /// <param name="rotationAngle">The rotation angle.</param>
        /// <param name="orientation">The orientation.</param>
        /// <param name="sizingRule">The sizing rule.</param>
        internal static void GetDimensions(FoundationBase footingComponent, FootingComponentType footingType, out double rotationAngle, out int orientation, out int sizingRule, out int shape)
        {
            if (footingComponent == null)
            {
                throw new CmnArgumentNullException("footingComponent");
            }
            if (!Enum.IsDefined(typeof(FootingComponentType), footingType))
            {
                throw new IndexOutOfRangeException("footingType");
            }

            string componentInterface = null;
            string rotationAngleProperty = null;
            string orientationProperty = null;
            string sizingRuleProperty = null;
            string shapeProperty = null;

            //Retrieve the component interfaces and property names based on the type
            switch (footingType)
            {
                case FootingComponentType.Grout:
                    componentInterface = SPSSymbolConstants.IJUASPSFtgGroutPad;
                    rotationAngleProperty = SPSSymbolConstants.GroutRotationAngle;
                    orientationProperty = SPSSymbolConstants.GroutOrientation;
                    sizingRuleProperty = SPSSymbolConstants.GroutSizingRule;
                    shapeProperty = SPSSymbolConstants.GroutShape;
                    break;
                case FootingComponentType.Pier:
                    componentInterface = SPSSymbolConstants.IJUASPSFtgPier;
                    rotationAngleProperty = SPSSymbolConstants.PierRotationAngle;
                    orientationProperty = SPSSymbolConstants.PierOrientation;
                    sizingRuleProperty = SPSSymbolConstants.PierSizingRule;
                    shapeProperty = SPSSymbolConstants.PierShape;
                    break;
                default:
                    componentInterface = SPSSymbolConstants.IJUASPSFtgSlab;
                    rotationAngleProperty = SPSSymbolConstants.SlabRotationAngle;
                    orientationProperty = SPSSymbolConstants.SlabOrientation;
                    sizingRuleProperty = SPSSymbolConstants.SlabSizingRule;
                    shapeProperty = SPSSymbolConstants.SlabShape;
                    break;
            }

            //Use the retrieved interface and property names to get the component dimensions
            rotationAngle = StructHelper.GetDoubleProperty(footingComponent, componentInterface, rotationAngleProperty);
            orientation = StructHelper.GetIntProperty(footingComponent, componentInterface, orientationProperty);
            sizingRule = StructHelper.GetIntProperty(footingComponent, componentInterface, sizingRuleProperty);
            shape = StructHelper.GetIntProperty(footingComponent, componentInterface, shapeProperty);
        }

        /// <summary>
        /// Validates the code list values. Returns a ToDoMessage if any value is invalid
        /// </summary>
        /// <param name="footingComponentType">Type of the footing component.</param>
        /// <param name="orientationType">Type of the orientation.</param>
        /// <param name="sizingRuleType">Type of the sizing rule.</param>
        /// <param name="shapeType">Type of the shape.</param>
        /// <param name="foundationComponent">The foundation component. By default it is null.</param>
        /// <returns>
        /// A ToDoMessage with appropriate message from error code list, if the codelist item value is invalid
        /// </returns>
        /// <exception cref="IndexOutOfRangeException">footingType</exception>
        internal static ToDoListMessage AreCodeListValuesValid(FootingComponentType footingComponentType, int orientationType, int sizingRuleType, int shapeType, FoundationComponent foundationComponent = null)
        {
            int invalidOrientationErrorNumber = 0;
            int invalidSizingRuleErrorNumber = 0;
            int invalidShapeErrorNumber = 0;

            switch (footingComponentType)
            {
                case FootingComponentType.Grout:
                    invalidOrientationErrorNumber = SPSSymbolConstants.TDL_INVALID_GROUT_ORIENTATION;
                    invalidSizingRuleErrorNumber = SPSSymbolConstants.TDL_INVALID_GROUT_SIZING_RULE;
                    invalidShapeErrorNumber = SPSSymbolConstants.TDL_INVALID_GROUT_SHAPE;               
                    break;
                case FootingComponentType.Pier:
                    invalidOrientationErrorNumber = SPSSymbolConstants.TDL_INVALID_PIER_ORIENTATION;
                    invalidSizingRuleErrorNumber = SPSSymbolConstants.TDL_INVALID_PIER_SIZING_RULE;
                    invalidShapeErrorNumber = SPSSymbolConstants.TDL_INVALID_PIER_SHAPE;                  
                    break;
                case FootingComponentType.Slab:
                    invalidOrientationErrorNumber = SPSSymbolConstants.TDL_INVALID_SLAB_ORIENTATION;
                    invalidSizingRuleErrorNumber = SPSSymbolConstants.TDL_INVALID_SLAB_SIZING_RULE;
                    invalidShapeErrorNumber = SPSSymbolConstants.TDL_INVALID_SLAB_SHAPE;               
                    break;
            }

            if (!StructHelper.IsValidCodeListValue(SPSSymbolConstants.StructPrismaticFootingShapes, SPSSymbolConstants.UDP, shapeType))
            {
                if (foundationComponent == null)
                {
                    return new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, invalidShapeErrorNumber,
                        String.Format("Error while validating code list value as shapeType: {0} does not exist in prismaticFootingShapes: {1} table in Footing Services", shapeType.ToString(), SPSSymbolConstants.StructPrismaticFootingShapes));
                }
                else
                {
                    return new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, invalidShapeErrorNumber,
                        String.Format("Error while validating code list value as shapeType: {0} does not exist in prismaticFootingShapes: {1} table in Footing Services", shapeType.ToString(), SPSSymbolConstants.StructPrismaticFootingShapes), foundationComponent);
                }
            } 
            if (!StructHelper.IsValidCodeListValue(SPSSymbolConstants.StructCoordSysReference, SPSSymbolConstants.UDP, orientationType))
            {
                if (foundationComponent == null)
                {
                    return new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, invalidOrientationErrorNumber,
                        String.Format("Error while validating code list value as orientationType: {0} does not exist in StructCoordSysReference: {1} table in Footing Services", orientationType.ToString(), SPSSymbolConstants.StructCoordSysReference));
                }
                else
                {
                    return new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, invalidOrientationErrorNumber,
                        String.Format("Error while validating code list value as orientationType: {0} does not exist in StructCoordSysReference: {1} table in Footing Services", orientationType.ToString(), SPSSymbolConstants.StructCoordSysReference), foundationComponent);
                }
            }
            if (!StructHelper.IsValidCodeListValue(SPSSymbolConstants.StructFootingCompSizingRule, SPSSymbolConstants.UDP, sizingRuleType))
            {
                if (foundationComponent == null)
                {
                    return new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, invalidSizingRuleErrorNumber,
                        String.Format("Error while validating code list value as sizingRuleType: {0} does not exist in FootignComponentSizingRule: {1} table in Footing Services", sizingRuleType.ToString(), SPSSymbolConstants.StructFootingCompSizingRule));
                }
                else
                {
                    return new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, invalidSizingRuleErrorNumber,
                        String.Format("Error while validating code list value as sizingRuleType: {0} does not exist in FootignComponentSizingRule: {1} table in Footing Services", sizingRuleType.ToString(), SPSSymbolConstants.StructFootingCompSizingRule, foundationComponent), foundationComponent);
                }
            }
            return null;
        }

        /// <summary>
        /// Gets the exposed surface area.
        /// </summary>
        /// <param name="footing">The footing.</param>
        /// <param name="footingType">Type of the footing.</param><returns></returns>
        internal static double GetExposedSurfaceArea(FoundationBase footing, FootingComponentType footingType)
        {
            if (footing == null)
            {
                throw new CmnArgumentNullException("footing");
            }
            if (!Enum.IsDefined(typeof(FootingComponentType), footingType))
            {
                throw new IndexOutOfRangeException("footingType");
            }

            double length = 0.0, width = 0.0, height = 0.0, clearance = 0.0;
            int shape = 0;
            bool withGroutPad = false;

            //check here to see if this is a component making the call or the footing symbol
            //if the footing parameter is a component then get its parent.  The parent will be used
            //to get the other component properties for calculations.  If the footing is a symbol
            //then this is not used as the properties are retrieved from the footing object itself.
            Footing parent = null;
            FoundationComponent footingComponent = footing as FoundationComponent;

            if (footingComponent != null)
            {
                parent = (Footing)footing.SystemParent;
            }

            //Get the dimensions for this footing
            GetDimensions(footing, footingType, ref length, ref width, ref height, ref clearance, ref shape);

            double surfaceArea = 0.0;

            if (footingType == FootingComponentType.Grout)
            {
                surfaceArea = GetExposedSurfaceArea(shape, length, width, height);
            }
            else if (footingType == FootingComponentType.Pier)
            {
                Collection<FoundationBase> matingComponents = new Collection<FoundationBase>();
                FootingComponentType matingComponentType = FootingComponentType.Grout;

                //Check to see if this method was called by the footing symbol or by a component from a footing assembly
                //if the parent is null then this was called by the footing symbol.  If the parent is not null then
                //this was called by a footing component.  
                if (parent == null)
                {
                    //if we don’t have a parent then this is the footing symbol.  Get the grout pad information and get the 
                    //grout dimensions so we can calculate the total surface area of the pier
                    withGroutPad = FootingServices.DoesFootingHaveGrout((Footing)footing);

                    //if there is a grout pad then get its dimensions
                    if (withGroutPad)
                    {
                        matingComponents.Add(footing);
                    }
                }
                else
                {
                    //if we have a parent, then this was called by the footing component.  Get the footing components of the parent
                    //that way we can get the grout components and get its dimensions. The grout pad dimensions will be used
                    //to calculate the total surface area of the pier.
                    withGroutPad = FootingServices.DoesFootingHaveGrout(parent);

                    //if there is a grout pad then get mating grout components
                    if (withGroutPad)
                    {
                        matingComponents = FootingServices.GetAssociatedComponents(footing, GetGroutOrPierComponents(parent, matingComponentType));
                    }
                }

                surfaceArea = GetExposedSurfaceArea(shape, length, width, height, matingComponents, matingComponentType);
            }
            else
            {
                //must be the slab component
                //Check to see if this method was called by the footing symbol or by a component from a footing assembly
                //if the parent is null then this was called by the footing symbol.  If the parent is not null then
                //this was called by a footing component.
                Collection<FoundationBase> matingComponents = new Collection<FoundationBase>();
                FootingComponentType matingComponentType = FootingComponentType.Pier;
                if (parent == null)
                {
                    //if we don’t have a parent then this is the footing symbol.  Get the pier dimensions so we can calculate the 
                    //total surface area of the slab
                    matingComponents.Add(footing);
                }
                else
                {
                    //if we have a parent, then this was called by the footing component.  Get the footing components of the parent
                    //that way we can get the pier component and get its dimensions.  The pier dimensions will be used
                    //to calculate the total surface area of the slab.
                    Collection<FoundationComponent> groutOrPierComponents = GetGroutOrPierComponents(parent, FootingComponentType.Pier);
                    if (groutOrPierComponents.Count == 0) //in case we don't have pier
                    {
                        //Get the footing components of the parent that way we can get the grout component and get its dimensions.
                        //The grout dimensions will be used to calculate the total surface area of the slab.
                        withGroutPad = FootingServices.DoesFootingHaveGrout(parent);
                        if (withGroutPad)
                        {
                            groutOrPierComponents = GetGroutOrPierComponents(parent, FootingComponentType.Grout);
                            matingComponentType = FootingComponentType.Grout;
                        }
                    }

                    foreach (var groutOrPierComponent in groutOrPierComponents)
                    {
                        matingComponents.Add(groutOrPierComponent);
                    }

                }

                surfaceArea = GetExposedSurfaceArea(shape, length, width, height, matingComponents, matingComponentType);
            }

            return surfaceArea;
        }

        /// <summary>
        /// Gets the exposed surface area.
        /// </summary>
        /// <param name="shape">The face shape.</param>
        /// <param name="length">The face length.</param>
        /// <param name="width">The face width.</param>
        /// <returns>The top or bottom face surface area.</returns>
        internal static double GetExposedSurfaceArea(int shape, double length, double width, double height)
        {
            double exposedSurfaceArea = 0.0;

            //get the top face area 
            double topFaceArea = GetTopOrBottomFaceSurfaceArea(shape, length, width);

            if (shape == SPSSymbolConstants.SHAPE_RECTANGULAR)
            {
                //top face area plus the area of each of the 4 sides of the rectangle
                exposedSurfaceArea = topFaceArea + (2 * length * height) + (2 * width * height);
            }
            else if (shape == SPSSymbolConstants.SHAPE_CIRCULAR)
            {
                //top face area plus the area that wraps around
                exposedSurfaceArea = topFaceArea + (2 * Math.PI * width / 2 * height);
            }
            else if (shape == SPSSymbolConstants.SHAPE_OCTAGONAL)
            {
                //top face area plus the area of each of the 8 sides of the octagon
                exposedSurfaceArea = topFaceArea + (length * height * 8);
            }

            return exposedSurfaceArea;
        }

        /// <summary>
        /// Gets the slab exposed surface area.
        /// </summary>
        /// <param name="shape">The footing component shape.</param>
        /// <param name="length">The footing component length.</param>
        /// <param name="width">The footing component width.</param>
        /// <param name="height">The footing component height.</param>
        /// <param name="matingComponents">The mating components.</param>
        /// <param name="matingComponentType">The mating component type.</param>
        /// <returns></returns>
        internal static double GetExposedSurfaceArea(int shape, double length, double width, double height, Collection<FoundationBase> matingComponents, FootingComponentType matingComponentType)
        {
            //get the slab exposed surface area
            double exposedSurfaceArea = GetExposedSurfaceArea(shape, length, width, height);

            //get the mating components total bottom face area 
            double matingComponentsTotalBottomFaceArea = 0.0;
            foreach (FoundationBase matingComponent in matingComponents)
            {
                int matingComponentShape = 0;
                double matingComponentLength = 0.0, matingComponentWidth = 0.0, matingComponentHeight = 0.0, matingComponentClearance = 0.0;
                GetDimensions(matingComponent, matingComponentType, ref matingComponentLength, ref matingComponentWidth, ref matingComponentHeight, ref matingComponentClearance, ref matingComponentShape);

                //get the mating component bottom face area
                matingComponentsTotalBottomFaceArea += GetTopOrBottomFaceSurfaceArea(matingComponentShape, matingComponentLength, matingComponentWidth);
            }

            //subtract the mating component bottom face area from the total slab face as the mating component is placed on top of the slab face
            //get the absolute value as we don't care if the mating component face is bigger
            return Math.Abs(exposedSurfaceArea - matingComponentsTotalBottomFaceArea);
        }

        /// <summary>
        /// Gets the top or bottom face surface area.
        /// </summary>
        /// <param name="shape">The face shape.</param>
        /// <param name="length">The face length.</param>
        /// <param name="width">The face width.</param>
        /// <returns>The top or bottom face surface area.</returns>
        internal static double GetTopOrBottomFaceSurfaceArea(int shape, double length, double width)
        {
            double topOrBottomFaceSurfaceArea = 0.0;
            //get the mating component bottom face area
            if (shape == SPSSymbolConstants.SHAPE_RECTANGULAR)
            {
                topOrBottomFaceSurfaceArea = length * width;
            }
            else if (shape == SPSSymbolConstants.SHAPE_CIRCULAR)
            {
                topOrBottomFaceSurfaceArea = Math.PI * Math.Pow(width / 2, 2);
            }
            else if (shape == SPSSymbolConstants.SHAPE_OCTAGONAL)
            {
                topOrBottomFaceSurfaceArea = (2 + 2 * Math.Sqrt(2)) * Math.Pow(length, 2);
            }

            return topOrBottomFaceSurfaceArea;
        }

        /// <summary>
        /// Sets the physical properties.
        /// </summary>
        /// <param name="footingComponent">The footing component.</param>
        /// <param name="footingType">Type of the footing.</param>
        /// <param name="totalSurfaceArea">The total surface area.</param>
        /// <param name="projectionOffset">The projection offset.</param>
        internal static void SetPhysicalProperties(FoundationBase footingComponent, FootingComponentType footingType, double totalSurfaceArea, double projectionOffset)
        {
            if (footingComponent == null)
            {
                throw new CmnArgumentNullException("footingComponent");
            }
            if (!Enum.IsDefined(typeof(FootingComponentType), footingType))
            {
                throw new IndexOutOfRangeException("footingType");
            }

            double length = 0;
            double width = 0;
            double height = 0;
            double clearance = 0;
            int shape = 0;

            //Get the dimensions for this footing component
            GetDimensions(footingComponent, footingType, ref length, ref width, ref height, ref clearance, ref shape);

            VolumeCG volumeCG = null;

            switch (shape)
            {
                case SPSSymbolConstants.SHAPE_RECTANGULAR:
                    volumeCG = SymbolHelper.GetVolumeCGForRectangularSolid(length, width, height);
                    //We need to adjust the COG.Z by grout and pier height for slab and by grout height for pier
                    if (footingType != FootingComponentType.Grout)
                    {
                        volumeCG.COGZ = volumeCG.COGZ - projectionOffset;
                    }
                    break;
                case SPSSymbolConstants.SHAPE_CIRCULAR:
                    volumeCG = SymbolHelper.GetVolumeCGForCircularSolid(width, height);
                    //We need to adjust the COG.Z by grout and pier height for slab and by grout height for pier
                    if (footingType != FootingComponentType.Grout)
                    {
                        volumeCG.COGZ = volumeCG.COGZ - projectionOffset;
                    }
                    break;
                case SPSSymbolConstants.SHAPE_OCTAGONAL:
                    volumeCG = SymbolHelper.GetVolumeCGForOctagonalSolid(length, height);
                    //We need to adjust the COG.Z by grout and pier height for slab and by grout height for pier
                    if (footingType != FootingComponentType.Grout)
                    {
                        volumeCG.COGZ = volumeCG.COGZ - projectionOffset;
                    }
                    break;
                default:

                    throw new ArgumentException("Cannot set physical properties as the footing shape is not supported.");
            }

            //set the volume and surface area on the given footing
            footingComponent.SetPropertyValue(volumeCG.Volume, SPSSymbolConstants.IJGenericVolume, SPSSymbolConstants.Volume);
            footingComponent.SetPropertyValue(totalSurfaceArea, SPSSymbolConstants.IJSurfaceArea, SPSSymbolConstants.SurfaceArea);
        }

        /// <summary>
        /// Gets the component volume CG.
        /// </summary>
        /// <param name="footingComponent">The footing component.</param>
        /// <param name="footingType">Type of the footing.</param>
        /// <param name="projectionOffset">The projection offset.</param><returns></returns>
        internal static VolumeCG GetComponentVolumeCG(FoundationBase footingComponent, FootingComponentType footingType, double projectionOffset)
        {
            if (footingComponent == null)
            {
                throw new CmnArgumentNullException("footingComponent");
            }
            if (!Enum.IsDefined(typeof(FootingComponentType), footingType))
            {
                throw new IndexOutOfRangeException("footingType");
            }

            double length = 0;
            double width = 0;
            double height = 0;
            double clearance = 0;
            int shape = 0;

            //Get the dimensions for this footing component
            GetDimensions(footingComponent, footingType, ref length, ref width, ref height, ref clearance, ref shape);

            VolumeCG volumeCG = null;

            switch (shape)
            {
                case SPSSymbolConstants.SHAPE_RECTANGULAR:
                    volumeCG = SymbolHelper.GetVolumeCGForRectangularSolid(length, width, height);
                    //We need to adjust the COG.Z by grout and pier height for slab and by grout height for pier
                    if (footingType != FootingComponentType.Grout)
                    {
                        volumeCG.COGZ = volumeCG.COGZ - projectionOffset;
                    }
                    break;
                case SPSSymbolConstants.SHAPE_CIRCULAR:
                    volumeCG = SymbolHelper.GetVolumeCGForCircularSolid(width, height);
                    //We need to adjust the COG.Z by grout and pier height for slab and by grout height for pier
                    if (footingType != FootingComponentType.Grout)
                    {
                        volumeCG.COGZ = volumeCG.COGZ - projectionOffset;
                    }
                    break;
                case SPSSymbolConstants.SHAPE_OCTAGONAL:
                    volumeCG = SymbolHelper.GetVolumeCGForOctagonalSolid(length, height);
                    //We need to adjust the COG.Z by grout and pier height for slab and by grout height for pier
                    if (footingType != FootingComponentType.Grout)
                    {
                        volumeCG.COGZ = volumeCG.COGZ - projectionOffset;
                    }
                    break;
                default:

                    throw new ArgumentException("Cannot set physical properties as the footing shape is not supported.");
            }

            return volumeCG;
        }

        /// <summary>
        /// Evaluates the component. Returns a ToDoMessage if data validation fails
        /// </summary>
        /// <param name="footing">The footing.</param>
        /// <param name="footingComponent">The footing component.</param>
        /// <param name="footingComponentType">Type of the footing component.</param>
        /// <param name="withGroutPad">if set to <c>true</c> [with grout pad].</param>
        /// <param name="sectionDepth">The section depth.</param>
        /// <param name="sectionWidth">Width of the section.</param>
        /// <param name="memberAngle">The member angle.</param>
        /// <returns>ToDoListMessage if a codelist property value is invalid</returns>
        /// <exception cref="CmnArgumentNullException">
        /// footing
        /// or
        /// footingComponent
        /// </exception>
        /// <exception cref="IndexOutOfRangeException">footingType</exception>
        internal static ToDoListMessage EvaluateComponent(FoundationBase footing, FoundationComponent footingComponent, FootingComponentType footingComponentType, bool withGroutPad, double sectionDepth, double sectionWidth, double memberAngle)
        {
            if (footing == null)
            {
                throw new CmnArgumentNullException("footing");
            }
            if (footingComponent == null)
            {
                throw new CmnArgumentNullException("footingComponent");
            }
            if (!Enum.IsDefined(typeof(FootingComponentType), footingComponentType))
            {
                throw new IndexOutOfRangeException("footingType");
            }

            ToDoListMessage toDoListMessage = null;
            double rotationAngle = 0;
            int orientation = 0;
            int sizingRule = 0;
            int pierShape = 0;

            //Get the translation part from footing but not the rotation part TR-231229
            Matrix4X4 footingComponentMatrix = new Matrix4X4();
            Position origin = footing.Matrix.Origin;
            if (footingComponentType == FootingComponentType.Pier || footingComponentType == FootingComponentType.Slab)
            {
                double groutHeight = 0;
                double pierHeight = 0;
                //Adjusting the pier component matrix origin to accommodate the grout height
                if (withGroutPad)
                {
                    groutHeight = GetComponentHeight(footing, FootingComponentType.Grout);
                }
                //Adjusting the slab component matrix origin to accommodate the grout and pier height
                if (footingComponentType == FootingComponentType.Slab)
                {
                    pierHeight = GetComponentHeight(footing, FootingComponentType.Pier);
                }
                origin.Z = origin.Z - (groutHeight + pierHeight);
            }
            footingComponentMatrix.Origin = origin;

            Vector vector = new Vector(0, 0, 1);

            //1. Get component rotation and orientation to set the component matrix
            GetDimensions(footingComponent, footingComponentType, out rotationAngle, out orientation, out sizingRule, out pierShape);

            //2. Validate the properties
            toDoListMessage = AreCodeListValuesValid(footingComponentType, orientation, sizingRule, pierShape, footingComponent);
            if (toDoListMessage != null)
            {
                return toDoListMessage;
            }

            //3. Set the length and width properties based on the sizing rule and angle
            //if the sizing rule is not user defined then take into account the sizing rule and angle
            if (sizingRule != SPSSymbolConstants.SIZING_RULE_USERDEFINED)
            {
                SetLengthAndWidthPropertiesByRule(footingComponent, footingComponentType, sectionDepth, sectionWidth, rotationAngle, withGroutPad);
            }

            //4. Rotate the component based on the rotation angles of the component and member angle
            if (orientation == SPSSymbolConstants.ORIENTATION_GLOBAL)
            {
                footingComponentMatrix.Rotate(rotationAngle, vector);
            }
            else if (orientation == SPSSymbolConstants.ORIENTATION_LOCAL)
            {
                footingComponentMatrix.Rotate(rotationAngle + memberAngle, vector);
            }

            //5.  Setting the footing component matrix
            footingComponent.Matrix = footingComponentMatrix;
            return toDoListMessage;
        }

        /// <summary>
        /// Evaluates the weight CG.
        /// </summary>
        /// <param name="footingComponent">The footing component.</param>
        /// <param name="footingType">Type of the footing.</param>
        internal static void EvaluateWeightCG(FoundationBase footingComponent, FootingComponentType footingType)
        {
            if (footingComponent == null)
            {
                throw new CmnArgumentNullException("footingComponent");
            }
            if (!Enum.IsDefined(typeof(FootingComponentType), footingType))
            {
                throw new IndexOutOfRangeException("footingType");
            }

            //getting weight and centre of gravity origin
            int weightCOGOrigin = StructHelper.GetIntProperty(footingComponent, SPSSymbolConstants.IJWCGValueOrigin, SPSSymbolConstants.DryWCGOrigin);

            //we need to calculate weight and COG as DryWCGOrigin is computed 
            if (weightCOGOrigin == SPSSymbolConstants.DRY_WCOG_ORIGIN_COMPUTED)
            {
                double volume = StructHelper.GetDoubleProperty(footingComponent, SPSSymbolConstants.IJGenericVolume, SPSSymbolConstants.Volume);

                //get correct interface and property name based on footing type. 
                string componentInterface = null;
                string materialProperty = null;
                string materialGradeProperty = null;
                Position globalCOG = default(Position);

                switch (footingType)
                {
                    case FootingComponentType.Grout:
                        componentInterface = SPSSymbolConstants.IJUASPSFtgGroutPad;
                        materialProperty = SPSSymbolConstants.GroutMaterial;
                        materialGradeProperty = SPSSymbolConstants.GroutMaterialGrade;

                        //Transform the centre of gravity (COG) from local to global
                        VolumeCG groutVolumeCG = GetComponentVolumeCG(footingComponent, FootingComponentType.Grout, 0.0);
                        Position groutLocalCOG = new Position(groutVolumeCG.COGX, groutVolumeCG.COGY, groutVolumeCG.COGZ);
                        Matrix4X4 groutMatrix = footingComponent.Matrix;
                        globalCOG = groutMatrix.Transform(groutLocalCOG);
                        break;
                    case FootingComponentType.Pier:
                        componentInterface = SPSSymbolConstants.IJUASPSFtgPier;
                        materialProperty = SPSSymbolConstants.PierMaterial;
                        materialGradeProperty = SPSSymbolConstants.PierMaterialGrade;

                        //Transform the centre of gravity (COG) from local to global                    
                        VolumeCG pierVolumeCG = GetComponentVolumeCG(footingComponent, FootingComponentType.Pier, 0.0);
                        Position pierLocalCOG = new Position(pierVolumeCG.COGX, pierVolumeCG.COGY, pierVolumeCG.COGZ);
                        Matrix4X4 pierMatrix = footingComponent.Matrix;
                        globalCOG = pierMatrix.Transform(pierLocalCOG);
                        break;
                    default:
                        //Must be Slab
                        componentInterface = SPSSymbolConstants.IJUASPSFtgSlab;
                        materialProperty = SPSSymbolConstants.SlabMaterial;
                        materialGradeProperty = SPSSymbolConstants.SlabMaterialGrade;

                        //Transform the centre of gravity (COG) from local to global                    
                        VolumeCG slabVolumeCG = GetComponentVolumeCG(footingComponent, FootingComponentType.Slab, 0.0);
                        Position slabLocalCOG = new Position(slabVolumeCG.COGX, slabVolumeCG.COGY, slabVolumeCG.COGZ);
                        Matrix4X4 slabMatrix = footingComponent.Matrix;
                        globalCOG = slabMatrix.Transform(slabLocalCOG);
                        break;
                }

                //Getting weight from volume
                double weight = SymbolHelper.GetWeightFromVolume(footingComponent, componentInterface, materialProperty, materialGradeProperty, volume);

                //Set the net weight and COG on the footing Component business object using helper method provided in WeightCOGServices
                WeightCOGServices weightCOGServices = new WeightCOGServices();
                weightCOGServices.SetWeightAndCOG(footingComponent, weight, globalCOG.X, globalCOG.Y, globalCOG.Z);
            }
        }

        /// <summary>
        /// Sets the length and width properties by rule.
        /// </summary>
        /// <param name="footing">The footing.</param>
        /// <param name="footingType">Type of the footing.</param>
        /// <param name="sectionDepth">The section depth.</param>
        /// <param name="sectionWidth">Width of the section.</param>
        /// <param name="componentRotationAngle">The component rotation angle.</param>
        /// <param name="isGroutPadNeeded">if set to <c>true</c> [is grout pad needed].</param>
        internal static void SetLengthAndWidthPropertiesByRule(FoundationBase footing, FootingComponentType footingType, double sectionDepth, double sectionWidth, double componentRotationAngle, bool withGroutPad)
        {
            if (footing == null)
            {
                throw new CmnArgumentNullException("footing");
            }
            if (!Enum.IsDefined(typeof(FootingComponentType), footingType))
            {
                throw new IndexOutOfRangeException("footingType");
            }

            double edgeClearance = 0.0;
            double componentLength = 0.0;
            double componentWidth = 0.0;
            double width = 0.0;
            double length = 0.0;
            double height = 0.0;

            string dimensionInterface = string.Empty;
            string widthProperty = string.Empty;
            string lengthProperty = string.Empty;

            int componentShape = 0;

            //check here to see if this is a component making the call or the footing symbol
            //if the footing parameter is a component then get its parent.  The parent will be used
            //to get the other component properties for calculations.  If the footing is a symbol
            //then this is not used as the properties are retrieved from the footing object itself.
            Footing parentFooting = null;
            FoundationComponent footingComponent = footing as FoundationComponent;
            if (footingComponent != null)
            {
                parentFooting = (Footing)footing.SystemParent;
            }

            //Get the dimensions for this footing
            GetDimensions(footing, footingType, ref length, ref width, ref height, ref edgeClearance, ref componentShape);

            //length & width calculated considering orientation angle of component
            //for circular shape no need to consider orientation angle
            if (componentShape == SPSSymbolConstants.SHAPE_CIRCULAR)
            {
                componentRotationAngle = 0.0;
            }

            if (componentRotationAngle > Math.PI / 2)
            {
                componentRotationAngle = Math.Abs(Math.PI - componentRotationAngle);
            }

            if (footingType == FootingComponentType.Grout)
            {
                SymbolHelper.ValidateCodelistValue(footing, SPSSymbolConstants.StructPrismaticFootingShapes, SPSSymbolConstants.UDP, componentShape, SPSSymbolConstants.TDL_INVALID_GROUT_SHAPE);

                if (parentFooting != null)
                {
                    componentLength = sectionDepth * Math.Cos(componentRotationAngle) + sectionWidth * Math.Sin(componentRotationAngle);
                    componentWidth = sectionWidth * Math.Cos(componentRotationAngle) + sectionDepth * Math.Sin(componentRotationAngle);
                }
                else
                {
                    componentLength = sectionDepth;
                    componentWidth = sectionWidth;
                }

                widthProperty = SPSSymbolConstants.GroutWidth;
                lengthProperty = SPSSymbolConstants.GroutLength;
                dimensionInterface = SPSSymbolConstants.IJUASPSFtgGroutPadDim;

            }
            else if (footingType == FootingComponentType.Pier)
            {
                SymbolHelper.ValidateCodelistValue(footing, SPSSymbolConstants.StructPrismaticFootingShapes, SPSSymbolConstants.UDP, componentShape, SPSSymbolConstants.TDL_INVALID_PIER_SHAPE);

                //for assembly case, the parent footing will be something
                if (parentFooting != null)
                {
                    //check if it a combined footing or not, if this is a combined footing with MergePiers then, we don’t required any calculations to determine the length and width of the pier
                    //as the dimensions are already calculated and passed into this method
                    if (parentFooting.SupportsInterface(SPSSymbolConstants.IJUACombinedPiers))
                    {
                        bool isMergePiers = StructHelper.GetBoolProperty(parentFooting, SPSSymbolConstants.IJUACombinedPiers, SPSSymbolConstants.MergePiers);
                        if (isMergePiers)
                        {
                            componentLength = sectionDepth;
                            componentWidth = sectionWidth;
                        }
                        else
                        {
                            componentLength = sectionDepth * Math.Cos(componentRotationAngle) + sectionWidth * Math.Sin(componentRotationAngle);
                            componentWidth = sectionWidth * Math.Cos(componentRotationAngle) + sectionDepth * Math.Sin(componentRotationAngle);
                        }
                    }
                    else//single pier footings
                    {
                        //if there is a grout pad then get the componentDepth and componentWidth based on the grout pad dimensions
                        //the grout pad dimensions have been already set using the member section size.  If no grout pad is needed then
                        //use the section depth and width directly
                        if (withGroutPad)
                        {
                            //Need this check here because if a component is passed to this method then we cannot get the 
                            //grout dimensions via the properties with a pier component but with a grout interfaces.
                            FoundationComponent grout = GetComponent(parentFooting, FootingComponentType.Grout);
                            CalculateDepthAndWidth(grout, FootingComponentType.Grout, componentRotationAngle, ref componentLength, ref componentWidth);
                        }
                        else
                        {
                            if (sectionDepth > sectionWidth)
                            {
                                componentWidth = sectionDepth;
                                componentLength = sectionDepth;
                            }
                            else
                            {
                                componentWidth = sectionWidth;
                                componentLength = sectionWidth;
                            }
                        }
                    }
                }
                else//symbol based footings
                {
                    //In case of symbol based footing calculation of pier length and width is depending on grout dimensions irrespective it is there or not.
                    CalculateDepthAndWidth(footing, FootingComponentType.Grout, componentRotationAngle, ref componentLength, ref componentWidth);
                }

                widthProperty = SPSSymbolConstants.PierWidth;
                lengthProperty = SPSSymbolConstants.PierLength;
                dimensionInterface = SPSSymbolConstants.IJUASPSFtgPierDim;

            }
            else if (footingType == FootingComponentType.Slab)
            {
                SymbolHelper.ValidateCodelistValue(footing, SPSSymbolConstants.StructPrismaticFootingShapes, SPSSymbolConstants.UDP, componentShape, SPSSymbolConstants.TDL_INVALID_SLAB_SHAPE);

                //The slab depth and width is based on the dimensions of the pier.            
                if (parentFooting != null)
                {
                    //for combined footing, the width and depth is already evaluated outside this function, so use the values directly
                    if (parentFooting.SupportsInterface(SPSSymbolConstants.IJUACombinedPiers))
                    {
                        componentWidth = sectionWidth;
                        componentLength = sectionDepth;
                    }
                    else
                    {
                        //Need this check here because if a component is passed to this method then we cannot get the 
                        //pier dimensions via the properties with a slab component but with a pier interfaces.  
                        FoundationComponent pier = GetComponent(parentFooting, FootingComponentType.Pier);
                        CalculateDepthAndWidth(pier, FootingComponentType.Pier, componentRotationAngle, ref componentLength, ref componentWidth);
                    }
                }
                else
                {
                    CalculateDepthAndWidth(footing, FootingComponentType.Pier, componentRotationAngle, ref componentLength, ref componentWidth);
                }

                //in case of octagonal slab we are setting the dimensions on the IJUAOctagonalSlabDim
                if (componentShape == SPSSymbolConstants.SHAPE_OCTAGONAL)
                {
                    widthProperty = SPSSymbolConstants.OctOverallDim;
                    lengthProperty = SPSSymbolConstants.OctFaceLength;
                    dimensionInterface = SPSSymbolConstants.IJUAOctagonalSlabDim;
                }
                else
                {
                    widthProperty = SPSSymbolConstants.SlabWidth;
                    lengthProperty = SPSSymbolConstants.SlabLength;
                    dimensionInterface = SPSSymbolConstants.IJUASPSFtgSlabDim;
                }
            }

            //need to know the sizing rule 
            double rotationAngle = 0.0;
            int orientation = 0;
            int sizingRule = 0;
            GetDimensions(footing, footingType, out rotationAngle, out orientation, out sizingRule, out componentShape);

            CalculateWidthAndLengthWithEdgeClearance(edgeClearance, componentShape, sizingRule, ref componentWidth, ref componentLength);

            //Set the length and width of the component
            footing.SetPropertyValue(componentWidth, dimensionInterface, widthProperty);
            footing.SetPropertyValue(componentLength, dimensionInterface, lengthProperty);
        }

        /// <summary>
        /// Calculate the width and length considering the edge clearance. 
        /// </summary>
        /// <param name="edgeClearance">Edge clearance which need to be considered</param>
        /// <param name="componentShape">Component shape</param>
        /// <param name="sizingRule">Sizing rule</param>
        /// <param name="componentWidth">Width of the component which should consider edge clearance</param>
        /// <param name="componentLength">Length of the component which should consider edge clearance</param>
        private static void CalculateWidthAndLengthWithEdgeClearance(double edgeClearance, int componentShape, int sizingRule, ref double componentWidth, ref double componentLength)
        {
            //Only add the clearance if the sizing rule is not user defined.  
            if (componentShape == SPSSymbolConstants.SHAPE_RECTANGULAR)
            {
                //do nothing if sizing rule is user defined, use same dimensions as it is
                if (sizingRule != SPSSymbolConstants.SIZING_RULE_USERDEFINED)
                {
                    //add component edge clearance to each dimension on both sides
                    componentWidth = componentWidth + (edgeClearance * 2);
                    componentLength = componentLength + (edgeClearance * 2);
                }
            }
            else if (componentShape == SPSSymbolConstants.SHAPE_CIRCULAR)
            {
                //in case of circular shape, the diagonal will be the diameter of the component
                //after calculating the diameter from the length and width, reset its length and width with its diameter
                if (sizingRule != SPSSymbolConstants.SIZING_RULE_USERDEFINED)
                {
                    componentWidth = componentLength = Math.Sqrt(componentLength * componentLength + componentWidth * componentWidth) + (edgeClearance * 2);
                }
                else
                {
                    componentWidth = componentLength = Math.Sqrt(componentLength * componentLength + componentWidth * componentWidth);
                }
            }
            else if (componentShape == SPSSymbolConstants.SHAPE_OCTAGONAL)
            {
                double maximumDimension = (componentWidth > componentLength) ? componentWidth : componentLength;

                //calculate the width from the max dimension (considering this as diagonal of the octagonal).
                componentWidth = maximumDimension / (Math.Sin(Math.PI / 4)) + (edgeClearance * 2); // diagonal/sin(45)

                //calculate the side of the octagonal (i.e., length) from the apothem of the octagonal (distance between two opposite sides)
                componentLength = componentWidth / (1 + Math.Sqrt(2)); //apothem = side * (1 + Sqrt(2))
            }
        }

        /// <summary>
        /// Calculates the depth and width of the footing.
        /// </summary>
        /// <param name="footing">The footing.</param>
        /// <param name="footingType">Type of the footing.</param>
        /// <param name="componentRotationAngle">The component rotation angle.</param>
        /// <param name="componentDepth">The component depth.</param>
        /// <param name="componentWidth">Width of the component.</param>
        internal static void CalculateDepthAndWidth(FoundationBase footing, FootingComponentType footingType, double componentRotationAngle, ref double componentDepth, ref double componentWidth)
        {
            double height = 0;
            double clearance = 0;
            int shape = 0;
            GetDimensions(footing, footingType, ref componentDepth, ref componentWidth, ref height, ref clearance, ref shape);
            componentWidth = componentWidth * Math.Cos(componentRotationAngle) + componentDepth * Math.Sin(componentRotationAngle);
            componentDepth = componentDepth * Math.Cos(componentRotationAngle) + componentWidth * Math.Sin(componentRotationAngle);
        }

        /// <summary>
        /// Gets the component.
        /// </summary>
        /// <param name="footing">The footing.</param>
        /// <param name="componentType">Type of the component.</param>
        /// <returns> The footing component. Null if component doesn't exist</returns>
        internal static FoundationComponent GetComponent(FoundationBase footing, FootingComponentType componentType)
        {
            if (footing == null)
            {
                throw new CmnArgumentNullException("footing");
            }

            //Get all children of the footing
            ReadOnlyCollection<ISystemChild> footingChildren = footing.SystemChildren;

            //Set the supporting interface based on the component type
            string supportingInterface = null;
            switch (componentType)
            {
                case FootingComponentType.Grout:
                    supportingInterface = SPSSymbolConstants.IJUASPSFtgGroutPad;
                    break;
                case FootingComponentType.Pier:
                    supportingInterface = SPSSymbolConstants.IJUASPSFtgPier;
                    break;
                case FootingComponentType.Slab:
                    supportingInterface = SPSSymbolConstants.IJUASPSFtgSlab;
                    break;
                default:
                    throw new CmnException("Foundation type not supported.");
            }

            //Look for the component which supports the component interface
            if (footingChildren != null)
            {
                foreach (BusinessObject component in footingChildren)
                {
                    if (component.SupportsInterface(supportingInterface))
                    {
                        return (FoundationComponent)component;
                    }
                    else if (component.SupportsInterface(supportingInterface))
                    {
                        return (FoundationComponent)component;
                    }
                    else if (component.SupportsInterface(supportingInterface))
                    {
                        return (FoundationComponent)component;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Gets the components of given type.
        /// </summary>
        /// <param name="footing">The footing.</param>
        /// <param name="componentType">Type of the component.</param>
        /// <returns> The footing components of given type</returns>
        internal static Collection<FoundationComponent> GetGroutOrPierComponents(FoundationBase footing, FootingComponentType componentType)
        {
            if (footing == null)
            {
                throw new CmnArgumentNullException("footing");
            }

            //Get all children of the footing
            ReadOnlyCollection<ISystemChild> footingChildren = footing.SystemChildren;

            Collection<FoundationComponent> footingComponentCollection = new Collection<FoundationComponent>();

            //Set the supporting interface based on the component type
            string supportingInterface = null;
            switch (componentType)
            {
                case FootingComponentType.Grout:
                    supportingInterface = SPSSymbolConstants.IJUASPSFtgGroutPad;
                    break;
                case FootingComponentType.Pier:
                    supportingInterface = SPSSymbolConstants.IJUASPSFtgPier;
                    break;
                case FootingComponentType.Slab:
                    supportingInterface = SPSSymbolConstants.IJUASPSFtgSlab;
                    break;
                default:
                    throw new CmnException("Foundation type not supported.");
            }

            //Look for the component which supports the component interface
            if (footingChildren != null)
            {
                foreach (BusinessObject component in footingChildren)
                {
                    if (component.SupportsInterface(supportingInterface))
                    {
                        footingComponentCollection.Add((FoundationComponent)component);
                    }
                }
            }

            return footingComponentCollection;
        }

        /// <summary>
        /// Get associated components from candidate components which are connected to given foundation component.
        /// </summary>
        /// <param name="foundationComponent">The foundation component.</param>
        /// <param name="candidateComponents">The candidate components.</param>
        /// <returns></returns>
        internal static Collection<FoundationBase> GetAssociatedComponents(FoundationBase foundationComponent, Collection<FoundationComponent> candidateComponents)
        {
            if (foundationComponent == null)
            {
                throw new CmnArgumentNullException("foundationComponent");
            }

            //Get components which are with in the range of the given foundation component.
            Collection<FoundationBase> components = new Collection<FoundationBase>();
            RangeBox foundationComponentRange = foundationComponent.Range;

            foreach (FoundationComponent tempComponent in candidateComponents)
            {
                Position tempComponentOrigin = tempComponent.Origin;
                if(tempComponentOrigin.X > foundationComponentRange.Low.X && tempComponentOrigin.X < foundationComponentRange.High.X &&
                    tempComponentOrigin.Y > foundationComponentRange.Low.Y && tempComponentOrigin.Y < foundationComponentRange.High.Y)
                {
                    components.Add(tempComponent);
                }
            }
            return components;
        }

        /// <summary>
        /// Gets the height of the component.
        /// </summary>
        /// <param name="footing">The footing.</param>
        /// <param name="componentType">Type of the component.</param>
        /// <returns>Height of the component</returns>
        internal static double GetComponentHeight(FoundationBase footing, FootingComponentType componentType)
        {
            if (footing == null)
            {
                throw new CmnArgumentNullException("footing");
            }

            BusinessObject component = GetComponent(footing, componentType);

            //Set the supporting interface based on the component type
            string componentInterface = null;
            string propertyName = null;
            switch (componentType)
            {
                case FootingComponentType.Grout:
                    componentInterface = SPSSymbolConstants.IJUASPSFtgGroutPadDim;
                    propertyName = SPSSymbolConstants.GroutHeight;
                    break;
                case FootingComponentType.Pier:
                    componentInterface = SPSSymbolConstants.IJUASPSFtgPierDim;
                    propertyName = SPSSymbolConstants.PierHeight;
                    break;
                case FootingComponentType.Slab:
                    componentInterface = SPSSymbolConstants.IJUASPSFtgSlabDim;
                    propertyName = SPSSymbolConstants.SlabHeight;
                    break;
                default:
                    throw new CmnException("Foundation type not supported.");
            }

            //Also consider the case where no separate component exists, i.e., given Footing is symbol type
            if (component != null)
            {
                return StructHelper.GetDoubleProperty(component, componentInterface, propertyName);
            }
            else
            {
                return StructHelper.GetDoubleProperty(footing, componentInterface, propertyName);
            }
        }

        /// <summary>
        /// Gets the grout sizing rule.
        /// </summary>
        /// <param name="properties">The properties.</param>
        /// <returns>Sizing Rule value for grout</returns>
        internal static int GetGroutSizingRule(ReadOnlyCollection<PropertyDescriptor> properties)
        {
            //Need to get the value of Grout sizing rule to determine the fate of the properties
            int groutSizingRule = 0;
            foreach (PropertyDescriptor propertyDescriptor in properties)
            {
                PropertyValue propertyValue = propertyDescriptor.Property;
                if (propertyValue.PropertyInfo.Name == SPSSymbolConstants.GroutSizingRule)
                {
                    groutSizingRule = Convert.ToInt32(((PropertyValueCodelist)propertyValue).PropValue);
                    break;
                }
            }

            return groutSizingRule;
        }

        /// <summary>
        /// Sets the component sizing properties access.
        /// </summary>
        /// <param name="componentType">Type of the component.</param>
        /// <param name="sizingRule">The sizing rule.</param>
        /// <param name="properties">The properties.</param>
        internal static void SetComponentSizingPropertiesAccess(FootingComponentType componentType, int sizingRule, ReadOnlyCollection<PropertyDescriptor> properties)
        {
            //Set the correct property names
            string lengthProperty = null;
            string widthProperty = null;
            string clearanceProperty = null;
            switch (componentType)
            {
                case FootingComponentType.Grout:
                    lengthProperty = SPSSymbolConstants.GroutLength;
                    widthProperty = SPSSymbolConstants.GroutWidth;
                    clearanceProperty = SPSSymbolConstants.GroutEdgeClearance;
                    break;
                case FootingComponentType.Pier:
                    lengthProperty = SPSSymbolConstants.PierLength;
                    widthProperty = SPSSymbolConstants.PierWidth;
                    clearanceProperty = SPSSymbolConstants.PierEdgeClearance;
                    break;
                case FootingComponentType.Slab:
                    lengthProperty = SPSSymbolConstants.SlabLength;
                    widthProperty = SPSSymbolConstants.SlabWidth;
                    clearanceProperty = SPSSymbolConstants.SlabEdgeClearance;
                    break;
                default:
                    throw new CmnException("Foundation type not supported.");
            }

            if (properties != null)
            {
                //Now find these properties in the properties collection set the read-only or writable behavior on the above properties
                foreach (PropertyDescriptor propertyDescriptor in properties)
                {
                    PropertyValue propertyValue = propertyDescriptor.Property;
                    string propertyNameToReadOnly = propertyValue.PropertyInfo.Name;
                    if (propertyNameToReadOnly == lengthProperty || propertyNameToReadOnly == widthProperty)
                    {
                        propertyDescriptor.ReadOnly = (sizingRule == SPSSymbolConstants.SIZING_RULE_USERDEFINED) ? false : true;
                    }
                    else if (propertyNameToReadOnly == clearanceProperty)
                    {
                        //if the rule is user defined then the edge clearance is disabled as it is not used while evaluating the component
                        propertyDescriptor.ReadOnly = (sizingRule == SPSSymbolConstants.SIZING_RULE_USERDEFINED) ? true : false;
                    }
                }
            }
        }

        /// <summary>
        /// Evaluates the Footing assembly weight and COG.
        /// </summary>
        /// <param name="footing">The footing.</param>
        internal static void EvaluateWeightAndCOG(Footing footing)
        {
            //getting weight and centre of gravity origin
            int weightCOGOrigin = StructHelper.GetIntProperty(footing, SPSSymbolConstants.IJWCGValueOrigin, SPSSymbolConstants.DryWCGOrigin);

            //we need to calculate weight and COG as DryWCGOrigin is computed 
            if (weightCOGOrigin != SPSSymbolConstants.DRY_WCOG_ORIGIN_DEFINED)
            {
                //set the weight and COG on the footing business object using helper method provided on WeightCOGServices
                //for assembly weight and COG values are not calculated as it is calculated at each component level
                WeightCOGServices weightCOGServices = new WeightCOGServices();
                weightCOGServices.SetWeightAndCOG(footing, 0, 0, 0, 0);
            }
        }

        /// <summary>
        /// Gets the actual pier rotation angle by considering the slab angle and member angle
        /// </summary>
        /// <param name="pierComponent">pier component</param>
        /// <param name="slabRotationAngle">rotation angle of the slab</param>
        /// <param name="memberAngle">member angle</param>
        /// <returns>resultant rotation angle of the pier</returns>
        internal static double GetPierRotationAngle(FoundationComponent pierComponent, double slabRotationAngle, double memberAngle)
        {
            int pierSizingRule;
            int shape;
            double pierRotationAngle;
            int pierOrientation;
            GetDimensions(pierComponent, FootingComponentType.Pier, out pierRotationAngle, out pierOrientation, out pierSizingRule, out shape);

            if (pierOrientation == SPSSymbolConstants.ORIENTATION_LOCAL)//add the member angle if the orientation is local
            {
                pierRotationAngle = pierRotationAngle + memberAngle;
            }

            if (shape == SPSSymbolConstants.SHAPE_CIRCULAR)//if pier is circular then no need to consider its rotation.
            {
                pierRotationAngle = 0.0;
            }

            pierRotationAngle = Math.Abs(Math.Abs(pierRotationAngle) - slabRotationAngle);
            if (pierRotationAngle > Math.PI / 2)
            {
                pierRotationAngle = Math.Abs(Math.PI / 2 - pierRotationAngle);
            }
            return pierRotationAngle;
        }

        /// <summary>
        /// Depending upon the grout height, slab height and the supported objects elevation it will decide that the pier height is valid or not.
        /// Which determines whether the pier is needed or not.
        /// </summary>
        /// <param name="groutAssemblyOutputs">The grout assembly outputs.</param>
        /// <param name="lowestSupportedObjectPosition">The lowest supported object position.</param>
        /// <param name="indexOfLowestSupportedObject">The index of lowest supported object.</param>
        /// <param name="isGroutNeeded">Is grout pad needed or not.</param>
        /// <param name="slabHeight">Height of the slab.</param>
        /// <param name="bottomPlaneMethod">The footing bottom plane method.</param>
        /// <param name="projectedPositionOnDatumPlane">The projected position on datum plane.</param>
        /// <returns>True if pier height is valid; otherwise, false.</returns>
        internal static bool IsValidPierHeight(AssemblyOutputs groutAssemblyOutputs, Position lowestSupportedObjectPosition, int indexOfLowestSupportedObject, bool isGroutNeeded, double slabHeight, FoundationBottomPlaneMethod bottomPlaneMethod, Position projectedPositionOnDatumPlane)
        {
            bool isValidPierHeight = true; //initialize with true
            double groutHeight = 0.0;

            //If grout pad is needed then get the height of the grout which is present at the lowest elevation of supported object
            if (isGroutNeeded && groutAssemblyOutputs[indexOfLowestSupportedObject] != null)
            {
                groutHeight = StructHelper.GetDoubleProperty(groutAssemblyOutputs[indexOfLowestSupportedObject], SPSSymbolConstants.IJUASPSFtgGroutPadDim, SPSSymbolConstants.GroutHeight);
            }

            double projectedPositionElevation = 0.0;
            if (projectedPositionOnDatumPlane != null)
            {
                projectedPositionElevation = projectedPositionOnDatumPlane.Z;
            }

            //in case the lowest elevation of supported object is greater than the combination of slab height, grout height and the elevation of the projected point on the datum plane, pier is needed
            if (bottomPlaneMethod != FoundationBottomPlaneMethod.None)
            {
                if ((lowestSupportedObjectPosition.Z - (projectedPositionElevation + groutHeight + slabHeight)) < 0.0)
                {
                    isValidPierHeight = false;
                }
            }

            return isValidPierHeight;
        }

        /// <summary>
        /// Determines whether this Footing have grout pad or not.
        /// </summary>
        /// <param name="footing"></param>
        /// <returns>True if Footing has grout pad; otherwise false.</returns>
        internal static bool DoesFootingHaveGrout(Footing footing)
        {
            bool doesFootingHaveGrout = false;
            if (footing.SupportsInterface(SPSSymbolConstants.IJUASPSPierAndSlabFooting))
            {
                doesFootingHaveGrout = StructHelper.GetBoolProperty(footing, SPSSymbolConstants.IJUASPSPierAndSlabFooting, SPSSymbolConstants.WithGroutPad);
            }
            else if (footing.SupportsInterface(SPSSymbolConstants.IJUASPSPierFootingAsm))
            {
                doesFootingHaveGrout = StructHelper.GetBoolProperty(footing, SPSSymbolConstants.IJUASPSPierFootingAsm, SPSSymbolConstants.WithGroutPad);
            }
            else if (footing.SupportsInterface(SPSSymbolConstants.IJUASPSSlabFootingAsm))
            {
                doesFootingHaveGrout = StructHelper.GetBoolProperty(footing, SPSSymbolConstants.IJUASPSSlabFootingAsm, SPSSymbolConstants.WithGroutPad);
            }
            else if (footing.SupportsInterface(SPSSymbolConstants.IJUASPSPierAndSlabFootingAsm))
            {
                doesFootingHaveGrout = StructHelper.GetBoolProperty(footing, SPSSymbolConstants.IJUASPSPierAndSlabFootingAsm, SPSSymbolConstants.WithGroutPad);
            }

            return doesFootingHaveGrout;
        }

        /// <summary>
        /// Sets the footing or footing component sizing rule.
        /// </summary>
        /// <param name="footingOrFootingComponent">The footing or footing component.</param>
        /// <param name="footingType">Type of the footing.</param>
        /// <param name="codelistName">Name of the codelist.</param>
        internal static void SetSizingRule(BusinessObject footingOrFootingComponent, FootingComponentType footingType, string codelistName)
        {
            //get the existing sizing rule value and set the new one in case they are different 
            MetadataManager metadataManager = footingOrFootingComponent.DBConnection.MetadataMgr;
            CodelistInformation codelistInformation = metadataManager.GetCodelistInfo(SPSSymbolConstants.StructFootingCompSizingRule, SPSSymbolConstants.UDP);
            CodelistItem codelistItem = codelistInformation.GetCodelistItem(codelistName);
            
            int existingSizingRule = GetSizingRule(footingOrFootingComponent, footingType);
            if (existingSizingRule != codelistItem.Value)
            {
                string componentInterface = GetComponentInterfaceName(footingType);
                string sizingRuleProperty = GetSizingRulePropertyName(footingType);

                footingOrFootingComponent.SetPropertyValue(codelistItem, componentInterface, sizingRuleProperty);
            }
        }

        /// <summary>
        /// Gets the footing or footing component sizing rule.
        /// </summary>
        /// <param name="footingOrFootingComponent">The footing or footing component.</param>
        /// <param name="footingType">Type of the footing.</param>        
        internal static int GetSizingRule(BusinessObject footingOrFootingComponent, FootingComponentType footingType)
        {
            string componentInterface = GetComponentInterfaceName(footingType);
            string sizingRuleProperty = GetSizingRulePropertyName(footingType);
            return StructHelper.GetIntProperty(footingOrFootingComponent, componentInterface, sizingRuleProperty);
        }

        /// <summary>
        /// Gets the footing or footing component height.
        /// </summary>
        /// <param name="footingOrFootingComponent">The footing or footing component.</param>
        /// <param name="footingType">Type of the footing.</param>        
        internal static double GetHeight(BusinessObject footingOrFootingComponent, FootingComponentType footingType)
        {
            string componentDimensionInterface = null;
            string heightProperty = null;

            //Retrieve the component interfaces and property names based on the type
            switch (footingType)
            {
                case FootingComponentType.Grout:
                    componentDimensionInterface = SPSSymbolConstants.IJUASPSFtgGroutPadDim;
                    heightProperty = SPSSymbolConstants.GroutHeight;
                    break;
                case FootingComponentType.Pier:
                    componentDimensionInterface = SPSSymbolConstants.IJUASPSFtgPierDim;
                    heightProperty = SPSSymbolConstants.PierHeight;
                    break;
                case FootingComponentType.Slab:
                    componentDimensionInterface = SPSSymbolConstants.IJUASPSFtgSlabDim;
                    heightProperty = SPSSymbolConstants.SlabHeight;
                    break;
            }

            return StructHelper.GetDoubleProperty(footingOrFootingComponent, componentDimensionInterface, heightProperty);
        }

        /// <summary>
        /// Gets the footing or footing component rotation angle.
        /// </summary>
        /// <param name="footingOrFootingComponent">The footing or footing component.</param>
        /// <param name="footingType">Type of the footing.</param>        
        internal static double GetRotationAngle(BusinessObject footingOrFootingComponent, FootingComponentType footingType)
        {
            string componentInterface = GetComponentInterfaceName(footingType);
            string rotationAngleProperty = null;

            //Retrieve the component interfaces and property names based on the type
            switch (footingType)
            {
                case FootingComponentType.Grout:
                    rotationAngleProperty = SPSSymbolConstants.GroutRotationAngle;
                    break;
                case FootingComponentType.Pier:
                    rotationAngleProperty = SPSSymbolConstants.PierRotationAngle;
                    break;
                case FootingComponentType.Slab:
                    rotationAngleProperty = SPSSymbolConstants.SlabRotationAngle;
                    break;
            }
            //Use the retrieved interface and property names to get the component dimensions
            return StructHelper.GetDoubleProperty(footingOrFootingComponent, componentInterface, rotationAngleProperty);
        }

        /// <summary>
        /// Gets the footing component interface name.
        /// </summary>        
        /// <param name="footingType">Type of the footing.</param>        
        internal static string GetComponentInterfaceName(FootingComponentType footingType)
        {
            string componentInterface = string.Empty;
            switch (footingType)
            {
                case FootingComponentType.Grout:
                    componentInterface = SPSSymbolConstants.IJUASPSFtgGroutPad;
                    break;
                case FootingComponentType.Pier:
                    componentInterface = SPSSymbolConstants.IJUASPSFtgPier;
                    break;
                case FootingComponentType.Slab:
                    componentInterface = SPSSymbolConstants.IJUASPSFtgSlab;
                    break;
            }
            return componentInterface;
        }

        /// <summary>
        /// Gets the sizing rule property name.
        /// </summary>        
        /// <param name="footingType">Type of the footing.</param>        
        internal static string GetSizingRulePropertyName(FootingComponentType footingType)
        {
            string sizingRuleProperty = string.Empty;
            switch (footingType)
            {
                case FootingComponentType.Grout:
                    sizingRuleProperty = SPSSymbolConstants.GroutSizingRule;
                    break;
                case FootingComponentType.Pier:
                    sizingRuleProperty = SPSSymbolConstants.PierSizingRule;
                    break;
                case FootingComponentType.Slab:
                    sizingRuleProperty = SPSSymbolConstants.SlabSizingRule;
                    break;
            }
            return sizingRuleProperty;
        }

        /// <summary>
        /// Gets the component material.
        /// </summary>
        /// <param name="footing">The footing.</param>
        /// <param name="componentType">Type of the component.</param>
        /// <returns></returns>
        internal static Material GetComponentMaterial(Footing footing, FootingComponentType componentType)
        {
            //First get the correct interface and property name for the component type
            string interfaceName = FootingServices.GetComponentInterfaceName(componentType);
            string materialPropertyName = string.Empty;
            string materialGradePropertyName = string.Empty;
            switch (componentType)
            {
                case FootingComponentType.Grout:
                    materialPropertyName = SPSSymbolConstants.GroutMaterial;
                    materialGradePropertyName = SPSSymbolConstants.GroutMaterialGrade;
                    break;
                case FootingComponentType.Pier:
                    materialPropertyName = SPSSymbolConstants.PierMaterial;
                    materialGradePropertyName = SPSSymbolConstants.PierMaterialGrade;
                    break;
                case FootingComponentType.Slab:
                    materialPropertyName = SPSSymbolConstants.SlabMaterial;
                    materialGradePropertyName = SPSSymbolConstants.SlabMaterialGrade;
                    break;
            }

            //Now get the actual material and grade name for the component set on the footing
            string materialName = StructHelper.GetStringProperty(footing, interfaceName, materialPropertyName);
            string materialGradeName = StructHelper.GetStringProperty(footing, interfaceName, materialGradePropertyName);

            //Now get the material from the catalog 
            CatalogStructHelper catalogStructHelper = new CatalogStructHelper();
            return catalogStructHelper.GetMaterial(materialName, materialGradeName);
        }


        /// <summary>
        /// Sets the component material.
        /// </summary>
        /// <param name="footing">The footing.</param>
        /// <param name="componentType">Type of the component.</param>
        internal static void SetComponentMaterial(Footing footing, FootingComponentType componentType)
        {
            Material material = FootingServices.GetComponentMaterial(footing, componentType);
            footing.SetMaterial(componentType.ToString(), material);
        }
    }
}