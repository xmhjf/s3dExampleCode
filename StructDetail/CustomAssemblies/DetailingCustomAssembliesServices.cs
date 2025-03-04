//-------------------------------------------------------------------------------------------------------
//Copyright© 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  DetailingCustomAssembliesServices.cs
//
//Abstract
//	DetailingCustomAssembliesServices is a helper class to have common method implementation for .NET StructDetailing custom assemblies.
//-------------------------------------------------------------------------------------------------------
using System;
using System.Linq;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using System.Collections.Generic;

namespace Ingr.SP3D.Content.Structure.Services
{
    /// <summary>
    /// Helper class to have common method implementation for .NET StructDetailing custom assemblies.
    /// </summary>
    internal static class DetailingCustomAssembliesServices
    {
        #region Internal Methods

        /// <summary>
        /// Determines whether the section has top flange or not.
        /// </summary>  
        /// <param name="sectionType">The section type.</param>
        /// <returns>True if the section has top flange; otherwise, false.</returns>
        /// <exception cref="ArgumentNullException">Raised if an argument is null.</exception>
        internal static bool HasTopFlange(string sectionType)
        {
            if (string.IsNullOrWhiteSpace(sectionType))
            {
                throw new ArgumentNullException("sectionType");
            }

            //the section has top flange if it is not among the below
            string[] topFlangeSections = new string[] { "FB", "HalfR", "P", "R", "SB", "SqTu", "RT", "L", "2L", "HSSC", "HSSR", "PIPE", "CS", "RS" };
            return (!topFlangeSections.Contains(sectionType)) ? true : false;
        }

        /// <summary>
        /// Checks if the connnectable has top flange or not.
        /// </summary>
        /// <param name="boundedObject">Bounded object</param>
        /// <returns></returns>
        internal static bool HasTopFlange(ProfilePart boundedObject)
        {
            bool hasTopFlange = false;
            int operatorId = -1;
            string sectionType = DetailingCustomAssembliesServices.GetSectionType(boundedObject);
            SectionFlanges flanges = GetSectionFlanges(sectionType);
            TopologyPort innerPort = boundedObject.GetPort(TopologyGeometryType.Face, operatorId, ContextTypes.Lateral, (int)(SectionFaceType.Inner_Web_Right), GeometryStage.Initial, false);
            if ((flanges & SectionFlanges.TopLeft).HasFlag(SectionFlanges.TopLeft) || (flanges & SectionFlanges.TopRight).HasFlag(SectionFlanges.TopRight))
            {
                hasTopFlange = true;
            }
            if (innerPort != null)
            {
                hasTopFlange = true;
            }

            return hasTopFlange;
        }

        /// <summary>
        /// Determines whether the section has bottom flange or not.
        /// </summary>  
        /// <param name="sectionType">The section type.</param>
        /// <returns>True if the section has bottom flange; otherwise, false.</returns>
        /// <exception cref="ArgumentNullException">Raised if an argument is null.</exception>
        internal static bool HasBottomFlange(string sectionType)
        {
            if (string.IsNullOrWhiteSpace(sectionType))
            {
                throw new ArgumentNullException("sectionType");
            }

            //the section has bottom flange if it is among the below
            string[] bottomFlangeSections = new string[] { "I", "ISType", "C_SS", "CSType", "H", "MC", "HP", "L", "2L", "W", "C", "M", "S" };
            return (bottomFlangeSections.Contains(sectionType)) ? true : false;
        }

        /// <summary>
        ///  Checks if the connnectable has bottom flange or not.
        /// </summary>
        /// <param name="boundedObject">Bounded object</param>
        /// <returns></returns>
        internal static bool HasBottomFlange(ProfilePart boundedObject)
        {
            bool hasBottomFlange = false;
            int operatorId = -1;
            string sectionType = DetailingCustomAssembliesServices.GetSectionType(boundedObject);
            SectionFlanges flanges = GetSectionFlanges(sectionType);
            TopologyPort innerPort = boundedObject.GetPort(TopologyGeometryType.Face, operatorId, ContextTypes.Lateral, (int)(SectionFaceType.Inner_Web_Right), GeometryStage.Initial, false);
            if ((flanges & SectionFlanges.BottomLeft) == SectionFlanges.BottomLeft || (flanges & SectionFlanges.BottomRight) == SectionFlanges.BottomRight)
            {
                hasBottomFlange = true;
            }
            if (innerPort != null)
            {
                hasBottomFlange = true;
            }

            return hasBottomFlange;
        }

        /// <summary>
        /// Get the web thickness of the specified cross section.
        /// </summary>
        /// <param name="crossSection">The crossSection for which the web thickness is requested.</param>
        internal static double GetWebThickness(CrossSection crossSection)
        {
            if (crossSection == null)
            {
                throw new ArgumentNullException("crossSection");
            }

            double webThickness = 0.0;

            if (crossSection.SupportsInterface(MarineSymbolConstants.IStructFlangedSectionDimensions))
            {
                webThickness = StructHelper.GetDoubleProperty(crossSection, MarineSymbolConstants.IStructFlangedSectionDimensions, MarineSymbolConstants.tw);
            }
            else if (crossSection.SupportsInterface(MarineSymbolConstants.IJUAXSectionWeb))
            {
                webThickness = StructHelper.GetDoubleProperty(crossSection, MarineSymbolConstants.IJUAXSectionWeb, MarineSymbolConstants.WebThickness);
            }
            else if (crossSection.SupportsInterface(MarineSymbolConstants.IJUAHSS))
            {
                webThickness = StructHelper.GetDoubleProperty(crossSection, MarineSymbolConstants.IJUAHSS, MarineSymbolConstants.tnom);
            }

            return webThickness;
        }

        /// <summary>
        /// Get the flange thickness of the specified cross section.
        /// </summary>
        /// <param name="crossSection">The crossSection for which the flange thickness is requested.</param>
        internal static double GetFlangeThickness(CrossSection crossSection)
        {
            if (crossSection == null)
            {
                throw new ArgumentNullException("crossSection");
            }

            double flangeThickness = 0.0;

            if (crossSection.SupportsInterface(MarineSymbolConstants.IStructFlangedSectionDimensions))
            {
                flangeThickness = StructHelper.GetDoubleProperty(crossSection, MarineSymbolConstants.IStructFlangedSectionDimensions, MarineSymbolConstants.tf);
            }
            else if (crossSection.SupportsInterface(MarineSymbolConstants.IJUAXSectionFlange))
            {
                flangeThickness = StructHelper.GetDoubleProperty(crossSection, MarineSymbolConstants.IJUAXSectionFlange, MarineSymbolConstants.FlangeThickness);
            }
            else if (crossSection.SupportsInterface(MarineSymbolConstants.IJUAHSS))
            {
                flangeThickness = StructHelper.GetDoubleProperty(crossSection, MarineSymbolConstants.IJUAHSS, MarineSymbolConstants.tnom);
            }

            return flangeThickness;
        }

        /// <summary>
        /// Get the web length of the specified cross section.
        /// </summary>
        /// <param name="crossSection">The crossSection for which the web length is requested.</param>
        internal static double GetWebLength(CrossSection crossSection)
        {
            if (crossSection == null)
            {
                throw new ArgumentNullException("crossSection");
            }

            double webLength = 0.0;

            if (crossSection.SupportsInterface(MarineSymbolConstants.IJUAXSectionWeb))
            {
                webLength = StructHelper.GetDoubleProperty(crossSection, MarineSymbolConstants.IJUAXSectionWeb, MarineSymbolConstants.WebLength);
            }

            return webLength;
        }

        /// <summary>
        /// Get the web depth of the specified cross section.
        /// </summary>
        /// <param name="crossSection">The crossSection for which the web depth is requested.</param>
        internal static double GetWebDepth(CrossSection crossSection)
        {
            if (crossSection == null)
            {
                throw new ArgumentNullException("crossSection");
            }

            double webDepth = 0.0;

            if (crossSection.SupportsInterface(MarineSymbolConstants.IStructFlangedSectionDimensions))
            {
                webDepth = StructHelper.GetDoubleProperty(crossSection, MarineSymbolConstants.IStructFlangedSectionDimensions, "d");
            }
            return webDepth;
        }

        /// <summary>
        /// Gets the section flanges based on whether the given section type has TopLeft, TopRight, BottomLeft and BottomRight flange.
        /// </summary>
        /// <param name="sectionType">The section type.</param>
        internal static SectionFlanges GetSectionFlanges(string sectionType)
        {
            SectionFlanges flanges = SectionFlanges.None;
            switch (sectionType)
            {
                case MarineSymbolConstants.TwoC:
                case MarineSymbolConstants.W:
                case MarineSymbolConstants.M:
                case MarineSymbolConstants.HP:
                case MarineSymbolConstants.S:
                case MarineSymbolConstants.H:
                case MarineSymbolConstants.I:
                case MarineSymbolConstants.ISType:
                    flanges = SectionFlanges.TopLeft | SectionFlanges.TopRight | SectionFlanges.BottomLeft | SectionFlanges.BottomRight;
                    break;
                case MarineSymbolConstants.TwoL:
                    flanges = SectionFlanges.BottomLeft | SectionFlanges.BottomRight;
                    break;
                case MarineSymbolConstants.C:
                case MarineSymbolConstants.MC:
                case MarineSymbolConstants.C_S:
                case MarineSymbolConstants.C_SS:
                case MarineSymbolConstants.CSType:
                    flanges = SectionFlanges.TopRight | SectionFlanges.BottomRight;
                    break;                
                case MarineSymbolConstants.L:
                    flanges = SectionFlanges.BottomRight;
                    break;
                case MarineSymbolConstants.T:
                case MarineSymbolConstants.MT:
                case MarineSymbolConstants.ST:
                case MarineSymbolConstants.WT:
                case MarineSymbolConstants.BUT:
                case MarineSymbolConstants.T_XType:
                case MarineSymbolConstants.TSType:
                    flanges = SectionFlanges.TopLeft | SectionFlanges.TopRight;
                    break;               
                case MarineSymbolConstants.EA:
                case MarineSymbolConstants.UA:
                case MarineSymbolConstants.BUTL3:
                case MarineSymbolConstants.BUTL2:
                    flanges = SectionFlanges.TopRight;
                    break;                
            }
            return flanges;
        }

        /// <summary>
        /// Gets the section type.
        /// </summary>
        /// <param name="businessObject">The ProfilePart or StiffenerSystemBase or PlatePart of BuiltUp.</param>        
        internal static string GetSectionType(BusinessObject businessObject)
        {
            // Get the section type of input object
            string sectionType = string.Empty;
            ICrossSection crossSection = businessObject as ICrossSection;
            PlatePart platePart = businessObject as PlatePart;
            if (crossSection != null)
            {
                sectionType = crossSection.SectionType;
            }
            else if (platePart != null)
            {
                PlateSystem plateSystem = (PlateSystem)platePart.RootPlateSystem;
                if (plateSystem.IsBuiltUp)
                {
                    crossSection = plateSystem.ParentBuiltUp;
                    sectionType = crossSection.SectionType;
                }
            }

            return sectionType;
        }

        #endregion Internal Methods

        #region Private Methods

        /// <summary>
        /// Gets the intersected and coplanar edge for the Web,FlangeLeftAndRightBottomWebs,FlangeLeftAndRightTopWebs,FlangeLeftAndRightWebs cases based on
        /// relative positions and the bounded edges.
        /// </summary>
        /// <param name="boundedEdge">possible edges of the bounded member.</param>
        /// <param name="measurments">measurement data related to connected edge.</param>
        /// <param name="intersectedEdge">intersecting edges of the bounding object.</param>
        /// <param name="coplanarEdge">Coplanar edges of the bounding object.</param>
        private static void GetIntersectedAndCoplanarEdgeForFlangeWebCase(BoundedEdgeSectionFaceType boundedEdge, Dictionary<string, double> measurments, out BoundingEdgeSectionFaceType intersectedEdge, out BoundingEdgeSectionFaceType coplanarEdge)
        {
            intersectedEdge = BoundingEdgeSectionFaceType.None;
            coplanarEdge = BoundingEdgeSectionFaceType.None;

            RelativePointPosition relativePos11 = DetailingCustomAssembliesServices.GetRelativePointPosition(11, boundedEdge, measurments);
            RelativePointPosition relativePos15 = DetailingCustomAssembliesServices.GetRelativePointPosition(15, boundedEdge, measurments);
            RelativePointPosition relativePos23 = DetailingCustomAssembliesServices.GetRelativePointPosition(23, boundedEdge, measurments);
            RelativePointPosition relativePos3 = DetailingCustomAssembliesServices.GetRelativePointPosition(3, boundedEdge, measurments);


            if (relativePos11 == RelativePointPosition.Below && relativePos15 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Above;
            }
            else if ((relativePos11 == RelativePointPosition.Coincident || relativePos11 == RelativePointPosition.Above) && relativePos15 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Top;
            }
            else if (relativePos15 == RelativePointPosition.Coincident)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Web_Right;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Top;
                        break;
                }
                if (relativePos11 == RelativePointPosition.Coincident)
                {
                    coplanarEdge = BoundingEdgeSectionFaceType.Top;
                }
            }
            else if (relativePos15 == RelativePointPosition.Above && relativePos23 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Web_Right;
            }
            else if (relativePos23 == RelativePointPosition.Coincident)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Bottom;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Web_Right;
                        break;
                }
                if (relativePos3 == RelativePointPosition.Coincident)
                {
                    coplanarEdge = BoundingEdgeSectionFaceType.Bottom;
                }
            }
            else if (relativePos23 == RelativePointPosition.Above && (relativePos3 == RelativePointPosition.Below || relativePos3 == RelativePointPosition.Coincident))
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Bottom;
            }
            else if (relativePos23 == RelativePointPosition.Above && relativePos3 == RelativePointPosition.Above)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Below;
            }

        }

        /// <summary>
        /// Gets the intersected and coplanar edge for the WebTopFlangeRight cases based on relative positions and the bounded edges.
        /// </summary>
        /// <param name="boundedEdge">possible edges of the bounded member.</param>
        /// <param name="measurments">measurement data related to connected edge.</param>
        /// <param name="intersectedEdge">intersecting edges of the bounding object.</param>
        /// <param name="coplanarEdge">Coplanar edges of the bounding object.</param>
        private static void GetIntersectedAndCoplanarEdgeForWebTopFlangeRight(BoundedEdgeSectionFaceType boundedEdge, Dictionary<string, double> measurments, out BoundingEdgeSectionFaceType intersectedEdge, out BoundingEdgeSectionFaceType coplanarEdge)
        {
            intersectedEdge = BoundingEdgeSectionFaceType.None;
            coplanarEdge = BoundingEdgeSectionFaceType.None;

            RelativePointPosition relativePos11 = DetailingCustomAssembliesServices.GetRelativePointPosition(11, boundedEdge, measurments);
            RelativePointPosition relativePos15 = DetailingCustomAssembliesServices.GetRelativePointPosition(15, boundedEdge, measurments);
            RelativePointPosition relativePos23 = DetailingCustomAssembliesServices.GetRelativePointPosition(23, boundedEdge, measurments);
            RelativePointPosition relativePos17 = DetailingCustomAssembliesServices.GetRelativePointPosition(17, boundedEdge, measurments);
            RelativePointPosition relativePos18 = DetailingCustomAssembliesServices.GetRelativePointPosition(18, boundedEdge, measurments);
            RelativePointPosition relativePos3 = DetailingCustomAssembliesServices.GetRelativePointPosition(3, boundedEdge, measurments);

            if (relativePos11 == RelativePointPosition.Below && relativePos15 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Above;
            }
            else if (relativePos15 == RelativePointPosition.Below && (relativePos11 == RelativePointPosition.Coincident || relativePos11 == RelativePointPosition.Above))
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Top;
            }
            else if (relativePos15 == RelativePointPosition.Coincident)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Top_Flange_Right;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Top;
                        break;
                }
                if (relativePos11 == RelativePointPosition.Coincident)
                {
                    coplanarEdge = BoundingEdgeSectionFaceType.Top;
                }
            }
            else if (relativePos15 == RelativePointPosition.Above && relativePos17 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Top_Flange_Right;
            }
            else if (relativePos17 == RelativePointPosition.Coincident)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Top_Flange_Right_Bottom;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Top_Flange_Right;
                        break;
                }
                if (relativePos18 == RelativePointPosition.Coincident)
                {
                    coplanarEdge = BoundingEdgeSectionFaceType.Top_Flange_Right_Bottom;
                }
            }
            else if (relativePos17 == RelativePointPosition.Above && relativePos18 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Top_Flange_Right_Bottom;
            }
            else if (relativePos18 == RelativePointPosition.Coincident)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Web_Right;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Top_Flange_Right_Bottom;
                        break;
                }
            }
            else if (relativePos18 == RelativePointPosition.Above && relativePos23 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Web_Right;
            }
            else if (relativePos23 == RelativePointPosition.Coincident)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Bottom;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Web_Right;
                        break;
                }
                if (relativePos3 == RelativePointPosition.Coincident)
                {
                    coplanarEdge = BoundingEdgeSectionFaceType.Bottom;
                }
            }
            else if (relativePos23 == RelativePointPosition.Above && (relativePos3 == RelativePointPosition.Below || relativePos3 == RelativePointPosition.Coincident))
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Bottom;
            }
            else if (relativePos23 == RelativePointPosition.Above && relativePos3 == RelativePointPosition.Above)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Below;
            }
        }

        /// <summary>
        /// Gets the intersected and coplanar edge for the WebBuiltUpTopFlangeRight cases based on relative positions and the bounded edges.
        /// </summary>
        /// <param name="boundedEdge">possible edges of the bounded member.</param>
        /// <param name="measurments">measurement data related to connected edge.</param>
        /// <param name="intersectedEdge">intersecting edges of the bounding object.</param>
        /// <param name="coplanarEdge">Coplanar edges of the bounding object.</param>
        private static void GetIntersectedAndCoplanarEdgeForWebBuiltUpTopFlangeRight(BoundedEdgeSectionFaceType boundedEdge, Dictionary<string, double> measurments, out BoundingEdgeSectionFaceType intersectedEdge, out BoundingEdgeSectionFaceType coplanarEdge)
        {
            intersectedEdge = BoundingEdgeSectionFaceType.None;
            coplanarEdge = BoundingEdgeSectionFaceType.None;

            RelativePointPosition relativePos11 = DetailingCustomAssembliesServices.GetRelativePointPosition(11, boundedEdge, measurments);
            RelativePointPosition relativePos14 = DetailingCustomAssembliesServices.GetRelativePointPosition(14, boundedEdge, measurments);
            RelativePointPosition relativePos15 = DetailingCustomAssembliesServices.GetRelativePointPosition(15, boundedEdge, measurments);
            RelativePointPosition relativePos23 = DetailingCustomAssembliesServices.GetRelativePointPosition(23, boundedEdge, measurments);
            RelativePointPosition relativePos17 = DetailingCustomAssembliesServices.GetRelativePointPosition(17, boundedEdge, measurments);
            RelativePointPosition relativePos18 = DetailingCustomAssembliesServices.GetRelativePointPosition(18, boundedEdge, measurments);
            RelativePointPosition relativePos3 = DetailingCustomAssembliesServices.GetRelativePointPosition(3, boundedEdge, measurments);
            RelativePointPosition relativePos50 = DetailingCustomAssembliesServices.GetRelativePointPosition(50, boundedEdge, measurments);

            if (relativePos11 == RelativePointPosition.Below && relativePos14 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Above;
            }
            else if (relativePos14 == RelativePointPosition.Below && (relativePos11 == RelativePointPosition.Coincident || relativePos11 == RelativePointPosition.Above))
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Top;
            }
            else if (relativePos14 == RelativePointPosition.Coincident && relativePos15 == RelativePointPosition.Below)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Web_Right_Top;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Top;
                        break;
                }
                if (relativePos11 == RelativePointPosition.Coincident)
                {
                    coplanarEdge = BoundingEdgeSectionFaceType.Top;
                }
            }
            else if (relativePos14 == RelativePointPosition.Above && relativePos50 == RelativePointPosition.Below && relativePos15 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Web_Right_Top;
            }
            else if (relativePos50 == RelativePointPosition.Coincident && relativePos15 == RelativePointPosition.Below)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Top_Flange_Right_Top;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Web_Right_Top;
                        break;
                }
            }
            else if (relativePos50 == RelativePointPosition.Above && relativePos15 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Top_Flange_Right_Top;
            }
            if (relativePos15 == RelativePointPosition.Coincident)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Top_Flange_Right;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Top_Flange_Right_Top;
                        break;
                }
                if (relativePos50 == RelativePointPosition.Coincident)
                {
                    coplanarEdge = BoundingEdgeSectionFaceType.Top_Flange_Right_Top;
                }
            }
            else if (relativePos15 == RelativePointPosition.Above && relativePos17 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Top_Flange_Right;
            }
            else if (relativePos17 == RelativePointPosition.Coincident)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Top_Flange_Right_Bottom;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Top_Flange_Right;
                        break;
                }
                if (relativePos18 == RelativePointPosition.Coincident)
                {
                    coplanarEdge = BoundingEdgeSectionFaceType.Top_Flange_Right_Bottom;
                }
            }
            else if (relativePos17 == RelativePointPosition.Above && relativePos18 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Top_Flange_Right_Bottom;
            }
            else if (relativePos18 == RelativePointPosition.Coincident)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Web_Right;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Top_Flange_Right_Bottom;
                        break;
                }
            }
            else if (relativePos18 == RelativePointPosition.Above && relativePos23 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Web_Right;
            }
            else if (relativePos23 == RelativePointPosition.Coincident)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Bottom;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Web_Right;
                        break;
                }
                if (relativePos3 == RelativePointPosition.Coincident)
                {
                    coplanarEdge = BoundingEdgeSectionFaceType.Bottom;
                }
            }
            else if (relativePos23 == RelativePointPosition.Above && (relativePos3 == RelativePointPosition.Below || relativePos3 == RelativePointPosition.Coincident))
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Bottom;
            }
            else if (relativePos23 == RelativePointPosition.Above && relativePos3 == RelativePointPosition.Above)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Below;
            }
        }

        /// <summary>
        /// Gets the intersected and coplanar edge for the WebBottomFlangeRight cases based on relative positions and the bounded edges.
        /// </summary>
        /// <param name="boundedEdge">possible edges of the bounded member.</param>
        /// <param name="measurments">measurement data related to connected edge.</param>
        /// <param name="intersectedEdge">intersecting edges of the bounding object.</param>
        /// <param name="coplanarEdge">Coplanar edges of the bounding object.</param>
        private static void GetIntersectedAndCoplanarEdgeForWebBottomFlangeRight(BoundedEdgeSectionFaceType boundedEdge, Dictionary<string, double> measurments, out BoundingEdgeSectionFaceType intersectedEdge, out BoundingEdgeSectionFaceType coplanarEdge)
        {
            intersectedEdge = BoundingEdgeSectionFaceType.None;
            coplanarEdge = BoundingEdgeSectionFaceType.None;

            RelativePointPosition relativePos11 = DetailingCustomAssembliesServices.GetRelativePointPosition(11, boundedEdge, measurments);
            RelativePointPosition relativePos15 = DetailingCustomAssembliesServices.GetRelativePointPosition(15, boundedEdge, measurments);
            RelativePointPosition relativePos20 = DetailingCustomAssembliesServices.GetRelativePointPosition(20, boundedEdge, measurments);
            RelativePointPosition relativePos21 = DetailingCustomAssembliesServices.GetRelativePointPosition(21, boundedEdge, measurments);
            RelativePointPosition relativePos23 = DetailingCustomAssembliesServices.GetRelativePointPosition(23, boundedEdge, measurments);
            RelativePointPosition relativePos3 = DetailingCustomAssembliesServices.GetRelativePointPosition(3, boundedEdge, measurments);

            if (relativePos11 == RelativePointPosition.Below && relativePos15 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Above;
            }
            else if (relativePos15 == RelativePointPosition.Below && (relativePos11 == RelativePointPosition.Coincident || relativePos11 == RelativePointPosition.Above))
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Top;
            }
            else if (relativePos15 == RelativePointPosition.Coincident && relativePos21 == RelativePointPosition.Below)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Web_Right;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Top;
                        break;
                }
                if (relativePos11 == RelativePointPosition.Coincident)
                {
                    coplanarEdge = BoundingEdgeSectionFaceType.Top;
                }
            }
            else if (relativePos15 == RelativePointPosition.Above && relativePos20 == RelativePointPosition.Below && relativePos21 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Web_Right;
            }
            else if (relativePos20 == RelativePointPosition.Coincident && relativePos21 == RelativePointPosition.Below)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Bottom_Flange_Right_Top;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Web_Right;
                        break;
                }
            }
            else if (relativePos20 == RelativePointPosition.Above && relativePos21 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Bottom_Flange_Right_Top;
            }
            else if (relativePos21 == RelativePointPosition.Coincident)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Bottom_Flange_Right;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Bottom_Flange_Right_Top;
                        break;
                }
                if (relativePos20 == RelativePointPosition.Coincident)
                {
                    coplanarEdge = BoundingEdgeSectionFaceType.Bottom_Flange_Right_Top;
                }
            }
            else if (relativePos21 == RelativePointPosition.Above && relativePos23 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Bottom_Flange_Right;
            }
            else if (relativePos23 == RelativePointPosition.Coincident)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Bottom;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Bottom_Flange_Right;
                        break;
                }
                if (relativePos3 == RelativePointPosition.Coincident)
                {
                    coplanarEdge = BoundingEdgeSectionFaceType.Bottom;
                }
            }
            else if (relativePos23 == RelativePointPosition.Above && (relativePos3 == RelativePointPosition.Below || relativePos3 == RelativePointPosition.Coincident))
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Bottom;
            }
            else if (relativePos23 == RelativePointPosition.Above && relativePos3 == RelativePointPosition.Above)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Below;
            }

        }

        /// <summary>
        /// Gets the intersected and coplanar edge for the WebBuiltUpBottomFlangeRight cases based on relative positions and the bounded edges.
        /// </summary>
        /// <param name="boundedEdge">possible edges of the bounded member.</param>
        /// <param name="measurments">measurement data related to connected edge.</param>
        /// <param name="intersectedEdge">intersecting edges of the bounding object.</param>
        /// <param name="coplanarEdge">Coplanar edges of the bounding object.</param>
        private static void GetIntersectedAndCoplanarEdgeForWebBuiltUpBottomFlangeRight(BoundedEdgeSectionFaceType boundedEdge, Dictionary<string, double> measurments, out BoundingEdgeSectionFaceType intersectedEdge, out BoundingEdgeSectionFaceType coplanarEdge)
        {
            intersectedEdge = BoundingEdgeSectionFaceType.None;
            coplanarEdge = BoundingEdgeSectionFaceType.None;

            RelativePointPosition relativePos11 = DetailingCustomAssembliesServices.GetRelativePointPosition(11, boundedEdge, measurments);
            RelativePointPosition relativePos15 = DetailingCustomAssembliesServices.GetRelativePointPosition(15, boundedEdge, measurments);
            RelativePointPosition relativePos20 = DetailingCustomAssembliesServices.GetRelativePointPosition(20, boundedEdge, measurments);
            RelativePointPosition relativePos21 = DetailingCustomAssembliesServices.GetRelativePointPosition(21, boundedEdge, measurments);
            RelativePointPosition relativePos23 = DetailingCustomAssembliesServices.GetRelativePointPosition(23, boundedEdge, measurments);
            RelativePointPosition relativePos3 = DetailingCustomAssembliesServices.GetRelativePointPosition(3, boundedEdge, measurments);
            RelativePointPosition relativePos51 = DetailingCustomAssembliesServices.GetRelativePointPosition(51, boundedEdge, measurments);
            RelativePointPosition relativePos24 = DetailingCustomAssembliesServices.GetRelativePointPosition(24, boundedEdge, measurments);

            if (relativePos11 == RelativePointPosition.Below && relativePos15 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Above;
            }
            else if (relativePos15 == RelativePointPosition.Below && (relativePos11 == RelativePointPosition.Coincident || relativePos11 == RelativePointPosition.Above))
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Top;
            }
            else if (relativePos15 == RelativePointPosition.Coincident && relativePos21 == RelativePointPosition.Below)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Web_Right;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Top;
                        break;
                }
                if (relativePos11 == RelativePointPosition.Coincident)
                {
                    coplanarEdge = BoundingEdgeSectionFaceType.Top;
                }
            }
            else if (relativePos15 == RelativePointPosition.Above && relativePos20 == RelativePointPosition.Below && relativePos21 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Web_Right;
            }
            else if (relativePos20 == RelativePointPosition.Coincident && relativePos21 == RelativePointPosition.Below)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Bottom_Flange_Right_Top;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Web_Right;
                        break;
                }
            }
            else if (relativePos20 == RelativePointPosition.Above && relativePos21 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Bottom_Flange_Right_Top;
            }
            else if (relativePos21 == RelativePointPosition.Coincident)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Bottom_Flange_Right;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Bottom_Flange_Right_Top;
                        break;
                }
                if (relativePos20 == RelativePointPosition.Coincident)
                {
                    coplanarEdge = BoundingEdgeSectionFaceType.Bottom_Flange_Right_Top;
                }
            }
            else if (relativePos21 == RelativePointPosition.Above && relativePos23 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Bottom_Flange_Right;
            }
            else if (relativePos23 == RelativePointPosition.Coincident)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Bottom_Flange_Right_Bottom;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Bottom_Flange_Right;
                        break;
                }
                if (relativePos51 == RelativePointPosition.Coincident)
                {
                    coplanarEdge = BoundingEdgeSectionFaceType.Bottom_Flange_Right_Bottom;
                }
            }
            else if (relativePos23 == RelativePointPosition.Above && relativePos51 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Bottom_Flange_Right_Bottom;
            }
            else if (relativePos51 == RelativePointPosition.Coincident)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Web_Right_Bottom;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Bottom_Flange_Right_Bottom;
                        break;
                }
                if (relativePos51 == RelativePointPosition.Coincident)
                {
                    coplanarEdge = BoundingEdgeSectionFaceType.Bottom_Flange_Right_Bottom;
                }
            }
            else if (relativePos51 == RelativePointPosition.Above && relativePos24 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Web_Right_Bottom;
            }
            else if (relativePos24 == RelativePointPosition.Coincident)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Bottom;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Web_Right_Bottom;
                        break;
                }
                if (relativePos3 == RelativePointPosition.Coincident)
                {
                    coplanarEdge = BoundingEdgeSectionFaceType.Bottom;
                }
            }
            else if (relativePos24 == RelativePointPosition.Above && (relativePos3 == RelativePointPosition.Below || relativePos3 == RelativePointPosition.Coincident))
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Bottom;
            }
            else if (relativePos24 == RelativePointPosition.Above && relativePos3 == RelativePointPosition.Above)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Below;
            }
        }

        /// <summary>
        /// Gets the intersected and coplanar edge for the WebTopAndBottomRightFlanges cases based on relative positions and the bounded edges.
        /// </summary>
        /// <param name="boundedEdge">possible edges of the bounded member.</param>
        /// <param name="measurments">measurement data related to connected edge.</param>
        /// <param name="intersectedEdge">intersecting edges of the bounding object.</param>
        /// <param name="coplanarEdge">Coplanar edges of the bounding object.</param>
        private static void GetIntersectedAndCoplanarEdgeForWebTopAndBottomRightFlanges(BoundedEdgeSectionFaceType boundedEdge, Dictionary<string, double> measurments, out BoundingEdgeSectionFaceType intersectedEdge, out BoundingEdgeSectionFaceType coplanarEdge)
        {
            intersectedEdge = BoundingEdgeSectionFaceType.None;
            coplanarEdge = BoundingEdgeSectionFaceType.None;

            RelativePointPosition relativePos11 = DetailingCustomAssembliesServices.GetRelativePointPosition(11, boundedEdge, measurments);
            RelativePointPosition relativePos15 = DetailingCustomAssembliesServices.GetRelativePointPosition(15, boundedEdge, measurments);
            RelativePointPosition relativePos17 = DetailingCustomAssembliesServices.GetRelativePointPosition(17, boundedEdge, measurments);
            RelativePointPosition relativePos18 = DetailingCustomAssembliesServices.GetRelativePointPosition(18, boundedEdge, measurments);
            RelativePointPosition relativePos20 = DetailingCustomAssembliesServices.GetRelativePointPosition(20, boundedEdge, measurments);
            RelativePointPosition relativePos21 = DetailingCustomAssembliesServices.GetRelativePointPosition(21, boundedEdge, measurments);
            RelativePointPosition relativePos23 = DetailingCustomAssembliesServices.GetRelativePointPosition(23, boundedEdge, measurments);
            RelativePointPosition relativePos3 = DetailingCustomAssembliesServices.GetRelativePointPosition(3, boundedEdge, measurments);
            if (relativePos11 == RelativePointPosition.Below && relativePos15 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Above;
            }
            else if (relativePos15 == RelativePointPosition.Below && (relativePos11 == RelativePointPosition.Coincident || relativePos11 == RelativePointPosition.Above))
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Top;
            }
            else if (relativePos15 == RelativePointPosition.Coincident)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Top_Flange_Right;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Top;
                        break;
                }
                if (relativePos11 == RelativePointPosition.Coincident)
                {
                    coplanarEdge = BoundingEdgeSectionFaceType.Top;
                }
            }
            else if (relativePos15 == RelativePointPosition.Above && relativePos17 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Top_Flange_Right;
            }
            else if (relativePos17 == RelativePointPosition.Coincident)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Top_Flange_Right_Bottom;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Top_Flange_Right;
                        break;
                }
                if (relativePos18 == RelativePointPosition.Coincident)
                {
                    coplanarEdge = BoundingEdgeSectionFaceType.Top_Flange_Right_Bottom;
                }
            }
            else if (relativePos17 == RelativePointPosition.Above && relativePos18 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Top_Flange_Right_Bottom;
            }
            else if (relativePos18 == RelativePointPosition.Coincident && relativePos21 == RelativePointPosition.Below)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Web_Right;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Top_Flange_Right_Bottom;
                        break;
                }
            }
            else if (relativePos18 == RelativePointPosition.Above && relativePos20 == RelativePointPosition.Below && relativePos21 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Web_Right;
            }
            else if (relativePos20 == RelativePointPosition.Coincident && relativePos21 == RelativePointPosition.Below)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Bottom_Flange_Right_Top;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Web_Right;
                        break;
                }
            }
            else if (relativePos20 == RelativePointPosition.Above && relativePos21 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Bottom_Flange_Right_Top;
            }
            else if (relativePos21 == RelativePointPosition.Coincident)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Bottom_Flange_Right;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Bottom_Flange_Right_Top;
                        break;
                }
                if (relativePos18 == RelativePointPosition.Coincident)
                {
                    coplanarEdge = BoundingEdgeSectionFaceType.Bottom_Flange_Right_Top;
                }
            }
            else if (relativePos21 == RelativePointPosition.Above && relativePos23 == RelativePointPosition.Below)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Bottom_Flange_Right;
            }
            else if (relativePos23 == RelativePointPosition.Coincident)
            {
                switch (boundedEdge)
                {
                    case BoundedEdgeSectionFaceType.Top:
                    case BoundedEdgeSectionFaceType.InsideTopFlange:
                    case BoundedEdgeSectionFaceType.FlangeLeft:
                    case BoundedEdgeSectionFaceType.WebLeft:
                        intersectedEdge = BoundingEdgeSectionFaceType.Bottom;
                        break;
                    default:
                        intersectedEdge = BoundingEdgeSectionFaceType.Bottom_Flange_Right;
                        break;
                }
                if (relativePos3 == RelativePointPosition.Coincident)
                {
                    coplanarEdge = BoundingEdgeSectionFaceType.Bottom;
                }
            }
            else if (relativePos23 == RelativePointPosition.Above && (relativePos3 == RelativePointPosition.Below || relativePos3 == RelativePointPosition.Coincident))
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Bottom;
            }
            else if (relativePos23 == RelativePointPosition.Above && relativePos3 == RelativePointPosition.Above)
            {
                intersectedEdge = BoundingEdgeSectionFaceType.Below;
            }
        }

        /// <summary>
        /// Determines the relative position of a point with respect to a bounded edge.
        /// </summary>
        /// <param name="pointNumber">The number of the point (From the Measurement Symbol)</param>
        /// <param name="boundedEdge">The bounded edge to get the relative position with respect to (IE Top, Bottom)</param>
        /// <param name="measurments">The collection of dimensions retrieved from the measurement symbol.</param>
        /// <returns>The relative position of the point with respect to the given bounded edge.</returns>
        internal static RelativePointPosition GetRelativePointPosition(int pointNumber, BoundedEdgeSectionFaceType boundedEdge, Dictionary<string, double> measurments)
        {
            RelativePointPosition relativePointPosition = RelativePointPosition.Below;
            switch (boundedEdge)
            {
                case BoundedEdgeSectionFaceType.Top:
                case BoundedEdgeSectionFaceType.Bottom:
                    //Measurement Symbol Dimensions
                    double top; //Distance from point to Top
                    double bottom; //Distance from point to Bottom
                    //Dimension to determine the Width/Depth of the Bounded Member
                    double depth;

                    measurments.TryGetValue(DetailingCustomAssembliesConstants.DimPt + pointNumber + DetailingCustomAssembliesConstants.ToTop, out top);
                    measurments.TryGetValue(DetailingCustomAssembliesConstants.DimPt + pointNumber + DetailingCustomAssembliesConstants.ToBottom, out bottom);
                    measurments.TryGetValue(DetailingCustomAssembliesConstants.DepthAtPt + pointNumber, out depth);
                    if (top >= Math3d.DistanceTolerance && bottom >= Math3d.DistanceTolerance)
                    {
                        if (StructHelper.AreEqual((top + bottom), depth))
                        {
                            relativePointPosition = (boundedEdge == BoundedEdgeSectionFaceType.Bottom) ? RelativePointPosition.Above : RelativePointPosition.Below;
                        }
                        else if (top < bottom)
                        {
                            relativePointPosition = RelativePointPosition.Above;
                        }
                        else
                        {
                            relativePointPosition = RelativePointPosition.Below;
                        }
                    }
                    else if (StructHelper.AreEqual(top, 0.0))
                    {
                        relativePointPosition = (boundedEdge == BoundedEdgeSectionFaceType.Bottom) ? RelativePointPosition.Above : RelativePointPosition.Coincident;
                    }
                    else if (StructHelper.AreEqual(bottom, 0.0))
                    {
                        relativePointPosition = (boundedEdge == BoundedEdgeSectionFaceType.Bottom) ? RelativePointPosition.Coincident : RelativePointPosition.Below;
                    }
                    break;
                case BoundedEdgeSectionFaceType.InsideTopFlange:
                case BoundedEdgeSectionFaceType.InsideBottomFlange:
                    //Measurement Symbol Dimensions
                    double insideTopFlangeDistance; //Distance from point to InsideTopFlange (Construction line)
                    double insideBottomFlangeDistance; //Distance from point to InsideBottomFlange (Construction line)
                    //Dimensions to determine the Width/Depth of the Bounded Member
                    double webLength;

                    measurments.TryGetValue(DetailingCustomAssembliesConstants.DimPt + pointNumber + DetailingCustomAssembliesConstants.ToTopInside, out insideTopFlangeDistance);
                    measurments.TryGetValue(DetailingCustomAssembliesConstants.DimPt + pointNumber + DetailingCustomAssembliesConstants.ToBottomInside, out insideBottomFlangeDistance);
                    measurments.TryGetValue(DetailingCustomAssembliesConstants.InsideDepthAtPt + pointNumber, out webLength);
                    if (insideTopFlangeDistance >= Math3d.DistanceTolerance && insideBottomFlangeDistance >= Math3d.DistanceTolerance)
                    {
                        if (StructHelper.AreEqual((insideTopFlangeDistance + insideBottomFlangeDistance), webLength))
                        {
                            relativePointPosition = (boundedEdge == BoundedEdgeSectionFaceType.InsideBottomFlange) ? RelativePointPosition.Above : RelativePointPosition.Below;
                        }
                        else if (insideTopFlangeDistance < insideBottomFlangeDistance)
                        {
                            relativePointPosition = RelativePointPosition.Above;
                        }
                        else
                        {
                            relativePointPosition = RelativePointPosition.Below;
                        }
                    }
                    else if (StructHelper.AreEqual(insideTopFlangeDistance, 0.0))
                    {
                        relativePointPosition = (boundedEdge == BoundedEdgeSectionFaceType.InsideBottomFlange) ? RelativePointPosition.Above : RelativePointPosition.Coincident;
                    }
                    else if (StructHelper.AreEqual(insideBottomFlangeDistance, 0.0))
                    {
                        relativePointPosition = (boundedEdge == BoundedEdgeSectionFaceType.InsideBottomFlange) ? RelativePointPosition.Coincident : RelativePointPosition.Below;
                    }
                    break;
                case BoundedEdgeSectionFaceType.WebLeft:
                case BoundedEdgeSectionFaceType.WebRight:
                    //Measurement Symbol Dimensions
                    double webLeftDistance; //Distance from point to Web Left
                    double webRightDistance; //Distance from point to Web Right
                    //Dimensions to determine the Width/Depth of the Bounded Member
                    double webThickness;

                    measurments.TryGetValue(DetailingCustomAssembliesConstants.DimPt + pointNumber + DetailingCustomAssembliesConstants.ToWL, out webLeftDistance);
                    measurments.TryGetValue(DetailingCustomAssembliesConstants.DimPt + pointNumber + DetailingCustomAssembliesConstants.ToWR, out webRightDistance);
                    measurments.TryGetValue(DetailingCustomAssembliesConstants.WebThkAtPt + pointNumber, out webThickness);
                    if (webLeftDistance >= Math3d.DistanceTolerance && webRightDistance >= Math3d.DistanceTolerance)
                    {
                        if (StructHelper.AreEqual((webLeftDistance + webRightDistance), webThickness))
                        {
                            relativePointPosition = (boundedEdge == BoundedEdgeSectionFaceType.WebRight) ? RelativePointPosition.Above : RelativePointPosition.Below;

                        }
                        else if (webLeftDistance < webRightDistance)
                        {
                            relativePointPosition = RelativePointPosition.Above;
                        }
                        else
                        {
                            relativePointPosition = RelativePointPosition.Below;
                        }
                    }
                    else if (StructHelper.AreEqual(webLeftDistance, 0.0))
                    {
                        relativePointPosition = (boundedEdge == BoundedEdgeSectionFaceType.WebRight) ? RelativePointPosition.Above : RelativePointPosition.Coincident;
                    }
                    else if (StructHelper.AreEqual(webRightDistance, 0.0))
                    {
                        relativePointPosition = RelativePointPosition.Below;
                        if (boundedEdge == BoundedEdgeSectionFaceType.WebRight) relativePointPosition = RelativePointPosition.Coincident;
                    }
                    break;
                case BoundedEdgeSectionFaceType.FlangeLeft:
                case BoundedEdgeSectionFaceType.FlangeRight:
                    //Measurement Symbol Dimensions
                    double topFlangeLeftDistance; //Distance from point to Top Flange Left
                    double topFlangeRightDistance; //Distance from point to Top Flange Right
                    //Dimensions to determine the Width/Depth of the Bounded Member
                    double width;

                    measurments.TryGetValue(DetailingCustomAssembliesConstants.DimPt + pointNumber + DetailingCustomAssembliesConstants.ToFL, out topFlangeLeftDistance);
                    measurments.TryGetValue(DetailingCustomAssembliesConstants.DimPt + pointNumber + DetailingCustomAssembliesConstants.ToFR, out topFlangeRightDistance);
                    measurments.TryGetValue(DetailingCustomAssembliesConstants.WidthAtPt + pointNumber, out width);
                    if (topFlangeLeftDistance >= Math3d.DistanceTolerance && topFlangeRightDistance >= Math3d.DistanceTolerance)
                    {
                        if (StructHelper.AreEqual((topFlangeLeftDistance + topFlangeRightDistance), width))
                        {
                            relativePointPosition = (boundedEdge == BoundedEdgeSectionFaceType.FlangeRight) ? RelativePointPosition.Above : RelativePointPosition.Below;
                        }
                        else if (topFlangeLeftDistance < topFlangeRightDistance)
                        {
                            relativePointPosition = RelativePointPosition.Above;
                        }
                        else
                        {
                            relativePointPosition = RelativePointPosition.Below;
                        }
                    }
                    else if (StructHelper.AreEqual(topFlangeLeftDistance, 0.0))
                    {
                        relativePointPosition = (boundedEdge == BoundedEdgeSectionFaceType.FlangeRight) ? RelativePointPosition.Above : RelativePointPosition.Coincident;
                    }
                    else if (StructHelper.AreEqual(topFlangeRightDistance, 0.0))
                    {
                        relativePointPosition = (boundedEdge == BoundedEdgeSectionFaceType.FlangeRight) ? RelativePointPosition.Coincident : RelativePointPosition.Below;
                    }
                    break;
            }
            return relativePointPosition;
        }

        /// <summary>
        /// Gets the required MeasurementSymbolDimensions.
        /// </summary>
        /// <param name="penetratesWeb">Flag which indicates if it is web penetration.</param>
        /// <param name="boundingAlias">bounding alias.</param>
        /// <returns></returns>
        private static ICollection<string> GetMeasurementSymbolDimensions(bool penetratesWeb, BoundingAlias boundingAlias)
        {
            ICollection<string> measurementSymbolDimensions = new Collection<string>();

            if (penetratesWeb)
            {
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt11ToTop);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt11ToBottom);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt11ToTopInside);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt11ToBottomInside);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DepthAtPt11);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.InsideDepthAtPt11);

                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt15ToTop);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt15ToBottom);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt15ToTopInside);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt15ToBottomInside);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DepthAtPt15);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.InsideDepthAtPt15);

                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt23ToTop);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt23ToBottom);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt23ToTopInside);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt23ToBottomInside);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DepthAtPt23);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.InsideDepthAtPt23);

                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt3ToTop);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt3ToBottom);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt3ToTopInside);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt3ToBottomInside);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DepthAtPt3);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.InsideDepthAtPt3);
                if (boundingAlias == BoundingAlias.WebTopFlangeRight || boundingAlias == BoundingAlias.WebBuiltUpTopFlangeRight || boundingAlias == BoundingAlias.WebTopAndBottomRightFlanges)
                {
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt18ToTop);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt18ToBottom);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt18ToTopInside);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt18ToBottomInside);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DepthAtPt18);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.InsideDepthAtPt18);

                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt17ToTop);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt17ToBottom);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt17ToTopInside);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt17ToBottomInside);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DepthAtPt17);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.InsideDepthAtPt17);

                    if (boundingAlias == BoundingAlias.WebBuiltUpTopFlangeRight)
                    {
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt14ToTop);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt14ToBottom);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt14ToTopInside);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt14ToBottomInside);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DepthAtPt14);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.InsideDepthAtPt14);

                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt50ToTop);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt50ToBottom);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt50ToTopInside);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt50ToBottomInside);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DepthAtPt50);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.InsideDepthAtPt50);
                    }
                }

                if (boundingAlias == BoundingAlias.WebBottomFlangeRight || boundingAlias == BoundingAlias.WebBuiltUpBottomFlangeRight || boundingAlias == BoundingAlias.WebTopAndBottomRightFlanges)
                {
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt20ToTop);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt20ToBottom);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt20ToTopInside);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt20ToBottomInside);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DepthAtPt20);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.InsideDepthAtPt20);

                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt21ToTop);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt21ToBottom);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt21ToTopInside);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DdimPt21ToBottomInside);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DepthAtPt21);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.InsideDepthAtPt21);

                    if (boundingAlias == BoundingAlias.WebBuiltUpBottomFlangeRight)
                    {
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt24ToTop);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt24ToBottom);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt24ToTopInside);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt24ToBottomInside);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DepthAtPt24);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.InsideDepthAtPt24);

                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt51ToTop);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt51ToBottom);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt51ToTopInside);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt51ToBottomInside);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DepthAtPt51);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.InsideDepthAtPt51);
                    }
                }
            }
            else
            {
                //common dimensions from the flange measurement symbol
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt11ToWL);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt11ToWR);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt11ToFL);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt11ToFR);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.WidthAtPt11);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.WebThkAtPt11);

                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt15ToWL);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt15ToWR);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt15ToFL);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt15ToFR);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.WidthAtPt15);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.WebThkAtPt15);

                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt23ToWL);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt23ToWR);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt23ToFL);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt23ToFR);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.WidthAtPt23);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.WebThkAtPt23);

                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt3ToWL);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt3ToWR);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt3ToFL);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt3ToFR);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.WidthAtPt3);
                measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.WebThkAtPt3);

                if (boundingAlias == BoundingAlias.WebTopFlangeRight || boundingAlias == BoundingAlias.WebBuiltUpTopFlangeRight || boundingAlias == BoundingAlias.WebTopAndBottomRightFlanges)
                {
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt18ToWL);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt18ToWR);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt18ToFL);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt18ToFR);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.WidthAtPt18);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.WebThkAtPt18);

                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt17ToWL);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt17ToWR);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt17ToFL);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt17ToFR);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.WidthAtPt17);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.WebThkAtPt17);

                    if (boundingAlias == BoundingAlias.WebBuiltUpTopFlangeRight)
                    {
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt14ToWL);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt14ToWR);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt14ToFL);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt14ToFR);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.WidthAtPt14);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.WebThkAtPt14);

                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt50ToWL);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt50ToWR);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt50ToFL);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt50ToFR);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.WidthAtPt50);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.WebThkAtPt50);
                    }
                }
                if (boundingAlias == BoundingAlias.WebBottomFlangeRight || boundingAlias == BoundingAlias.WebBuiltUpBottomFlangeRight || boundingAlias == BoundingAlias.WebTopAndBottomRightFlanges)
                {
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt20ToWL);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt20ToWR);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt20ToFL);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt20ToFR);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.WidthAtPt20);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.WebThkAtPt20);

                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt21ToWL);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt21ToWR);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt21ToFL);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt21ToFR);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.WidthAtPt21);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.WebThkAtPt21);

                    if (boundingAlias == BoundingAlias.WebBuiltUpBottomFlangeRight)
                    {
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt24ToWL);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt24ToWR);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt24ToFL);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt24ToFR);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.WidthAtPt24);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.WebThkAtPt24);

                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt51ToWL);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt51ToWR);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt51ToFL);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt51ToFR);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.WidthAtPt51);
                        measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.WebThkAtPt51);
                    }
                }
            }
            //get point to point dimensions
            switch (boundingAlias)
            {
                case BoundingAlias.FlangeLeftAndRightBottomWebs:
                case BoundingAlias.FlangeLeftAndRightTopWebs:
                case BoundingAlias.FlangeLeftAndRightWebs:
                case BoundingAlias.Web:
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt11ToPt15);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt15ToPt23);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt3ToPt23);
                    break;
                case BoundingAlias.WebBottomFlangeRight:
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt11ToPt15);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt15ToPt20);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt20ToPt21);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt21ToPt23);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt3ToPt23);
                    break;
                case BoundingAlias.WebBuiltUpBottomFlangeRight:
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt11ToPt15);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt15ToPt20);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt20ToPt21);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt21ToPt23);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt23ToPt51);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt24ToPt51);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt3ToPt24);
                    break;
                case BoundingAlias.WebBuiltUpTopFlangeRight:
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt11ToPt14);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt14ToPt20);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt15ToPt50);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt15ToPt17);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt17ToPt18);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt18ToPt23);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt3ToPt23);
                    break;
                case BoundingAlias.WebTopAndBottomRightFlanges:
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt11ToPt15);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt15ToPt17);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt17ToPt18);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt18ToPt20);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt20ToPt21);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt21ToPt23);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt3ToPt23);
                    break;
                case BoundingAlias.WebTopFlangeRight:
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt11ToPt15);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt15ToPt17);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt17ToPt18);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt18ToPt23);
                    measurementSymbolDimensions.Add(DetailingCustomAssembliesConstants.DimPt3ToPt23);
                    break;
            }
            return measurementSymbolDimensions;
        }

        /// <summary>
        /// Gets the decoded string.
        /// </summary>
        /// <param name="decodedString">The decoded string.</param>
        /// <returns></returns>
        private static string GetDecodedString(string decodedString)
        {
            string inputString = decodedString;
            if (inputString.Contains("*"))
            {
                inputString = inputString.Replace("*", DetailingCustomAssembliesConstants.DimPt);
            }
            if (inputString.Contains("@"))
            {
                inputString = inputString.Replace("@", DetailingCustomAssembliesConstants.DepthAtPt);
            }
            if (inputString.Contains("&"))
            {
                inputString = inputString.Replace("&", DetailingCustomAssembliesConstants.ToTop);
            }
            if (inputString.Contains("#"))
            {
                inputString = inputString.Replace("#", DetailingCustomAssembliesConstants.ToBottom);
            }
            if (inputString.Contains("$"))
            {
                inputString = inputString.Replace("$", DetailingCustomAssembliesConstants.Inside);
            }
            return inputString;
        }

        /// <summary>
        /// Gets the cached edge mapping rule data.
        /// </summary>
        /// <param name="assemblyConnectionOrEndcut">assembly connection or end cut .</param> 
        /// <param name="penetratesWeb">Flag which indicates if it is web penetration.</param>
        /// <returns>returns the cached edge mapping rule data.</returns>
        private static Dictionary<int, int> GetCacheEdgeMappingRule(BusinessObject assemblyConnectionOrEndcut, out int sectionAlias, out bool penetratesWeb)
        {
            Dictionary<int, int> edgeMappings = null;
            sectionAlias = 0;
            penetratesWeb = false;
            PropertyValue propertyValue = null;
            if (assemblyConnectionOrEndcut.SupportsInterface(DetailingCustomAssembliesConstants.IJUAMbrACCacheStorage2))
            {
                edgeMappings = new Dictionary<int, int>();
                string value = StructHelper.GetStringProperty(assemblyConnectionOrEndcut, DetailingCustomAssembliesConstants.IJUAMbrACCacheStorage2, DetailingCustomAssembliesConstants.EdgeMapping);
                if (value != null)
                {
                    propertyValue = assemblyConnectionOrEndcut.GetPropertyValue(DetailingCustomAssembliesConstants.IJUAMbrACCacheStorage2, DetailingCustomAssembliesConstants.SectionAlias);
                    sectionAlias = (int)((PropertyValueInt)propertyValue).PropValue;


                    propertyValue = assemblyConnectionOrEndcut.GetPropertyValue(DetailingCustomAssembliesConstants.IJUAMbrACCacheStorage, DetailingCustomAssembliesConstants.IsWebPenetrated);
                    penetratesWeb = (bool)((PropertyValueBoolean)propertyValue).PropValue;

                    int edgeMapping = 0;
                    int edgeMappingvalue = 0;
                    if (value.Length > 0)
                    {
                        string[] arrayInfo = value.Split('+');
                        string stringInfo = string.Empty;
                        int positionOfEquals;
                        for (int i = 0; i < arrayInfo.Length; i++)
                        {
                            stringInfo = arrayInfo[i];
                            if (stringInfo.Contains("="))
                            {
                                positionOfEquals = stringInfo.LastIndexOf("=");
                                if (positionOfEquals > 0)
                                {
                                    string edgeMapName = stringInfo.Remove(positionOfEquals);
                                    string edgeMapValue = stringInfo.Substring(positionOfEquals + 1);
                                    edgeMapping = int.Parse(edgeMapName);
                                    edgeMappingvalue = int.Parse(edgeMapValue);
                                    edgeMappings.Add(edgeMapping, edgeMappingvalue);
                                }
                            }
                        }
                    }
                }
            }
            return edgeMappings;
        }

        /// <summary>
        /// Gets the web edge overlap and clearance.
        /// </summary>
        /// <param name="isBottomEdge">A boolean which determines if it is bottom edge or not.</param>
        /// <param name="measurements">measurement data related to connected edge.</param>
        /// <param name="boundedObject">Bounded object of end cut.</param>
        /// <param name="insideOverlap">The inside overlap.</param>
        /// <param name="outsideOverlap">The outside overlap.</param>
        /// <param name="edgeLength">edgelength value.</param>
        /// <param name="insideClearance">inside clearence value.</param>
        /// <param name="outsideClearance">inside clearence value.</param>
        /// <param name="isEdgeToEdge">Boolean value which tells if the overlap is between the edge and edge.</param>
        private static void GetWebEdgeOverlapAndClearance(bool isBottomEdge, Dictionary<string, double> measurements, BusinessObject boundedObject, out double insideOverlap, out double outsideOverlap,
           out double edgeLength, out double insideClearance, out double outsideClearance, out bool isEdgeToEdge)
        {
            insideOverlap = 0.0;
            outsideOverlap = 0.0;
            edgeLength = 0.0;
            insideClearance = 0.0;
            outsideClearance = 0.0;
            isEdgeToEdge = false;

            //Determine the relative position of the bounded and bounding objects
            if (isBottomEdge)
            {
                if (measurements.Keys.Contains(DetailingCustomAssembliesConstants.DimPt21ToPt23))
                {
                    edgeLength = measurements[DetailingCustomAssembliesConstants.DimPt21ToPt23];

                    RelativePointPosition topInsidePos21 = DetailingCustomAssembliesServices.GetRelativePointPosition(21, BoundedEdgeSectionFaceType.InsideTopFlange, measurements);
                    RelativePointPosition topInsidePos23 = DetailingCustomAssembliesServices.GetRelativePointPosition(23, BoundedEdgeSectionFaceType.InsideTopFlange, measurements);
                    RelativePointPosition btmInsidePos21 = DetailingCustomAssembliesServices.GetRelativePointPosition(21, BoundedEdgeSectionFaceType.InsideBottomFlange, measurements);
                    RelativePointPosition btmInsidePos23 = DetailingCustomAssembliesServices.GetRelativePointPosition(23, BoundedEdgeSectionFaceType.InsideBottomFlange, measurements);
                    RelativePointPosition topOutsidePos21 = DetailingCustomAssembliesServices.GetRelativePointPosition(21, BoundedEdgeSectionFaceType.Top, measurements);
                    RelativePointPosition topOutsidePos23 = DetailingCustomAssembliesServices.GetRelativePointPosition(23, BoundedEdgeSectionFaceType.Top, measurements);
                    RelativePointPosition btmOutsidePos21 = DetailingCustomAssembliesServices.GetRelativePointPosition(21, BoundedEdgeSectionFaceType.Bottom, measurements);
                    RelativePointPosition btmOutsidePos23 = DetailingCustomAssembliesServices.GetRelativePointPosition(23, BoundedEdgeSectionFaceType.Bottom, measurements);

                    if (topInsidePos21 == RelativePointPosition.Below)
                    {
                        //Inside clearance is distance from Pt 21 to the inside of the top flange
                        insideClearance = measurements[DetailingCustomAssembliesConstants.DimPt21ToTopInside];
                        //If the inside of the bottom flange is below the top corner
                        if (btmInsidePos21 == RelativePointPosition.Above)
                        {
                            // If the inside of the bottom flange is below the bottom corner
                            if (btmInsidePos23 == RelativePointPosition.Above || btmInsidePos23 == RelativePointPosition.Coincident)//Conditions 11 and 12 in BoundingEdgeConditions.sha
                            {
                                //Overlap is distance from Pt21 to Pt23
                                insideOverlap = edgeLength;

                                //Outside clearance is distance from Pt23 to flange right
                                outsideOverlap = measurements[DetailingCustomAssembliesConstants.DimPt23ToBottomInside];

                                //If the inside of the bottom flange intersects the edge, the overlap is from Pt21 to the inside of the bottom flange
                            }
                            else
                            {
                                //Conditions 9 and 10 in BoundingEdgeConditions.sha
                                isEdgeToEdge = DetailingCustomAssembliesServices.HasBottomFlange((ProfilePart)boundedObject);
                                insideOverlap = measurements[DetailingCustomAssembliesConstants.DimPt21ToBottomInside];
                            }
                        }
                        else if (btmOutsidePos21 == RelativePointPosition.Above)// Conditions 3-8 in BoundingEdgeConditions.sha
                        {
                            isEdgeToEdge = DetailingCustomAssembliesServices.HasBottomFlange((ProfilePart)boundedObject);
                            insideOverlap = -measurements[DetailingCustomAssembliesConstants.DimPt21ToBottomInside];
                        }
                    }
                    //If the inside of the top flange is below the bottom corner (cases 4a, 5a, 7a, 8a, 9a, 10a, 11a)
                    else if (topInsidePos23 == RelativePointPosition.Above || topInsidePos23 == RelativePointPosition.Coincident)
                    {
                        if (topOutsidePos23 == RelativePointPosition.Below)//(excludes case 11a)
                        {
                            isEdgeToEdge = DetailingCustomAssembliesServices.HasTopFlange((ProfilePart)boundedObject);
                        }
                        // Outside clearance is distance from Pt23 to flange right
                        outsideOverlap = measurements[DetailingCustomAssembliesConstants.DimPt23ToBottomInside];
                        //Avoid very small values like 1.0e-16.  (cases 5a, 8a, 10a)
                        if (topInsidePos23 == RelativePointPosition.Above)
                        {
                            outsideOverlap = -measurements[DetailingCustomAssembliesConstants.DimPt23ToBottomInside];
                        }
                    }//If the inside of the top flange intersects the edge
                    else//cases 2a, 3a, 6a
                    {
                        if (topInsidePos21 == RelativePointPosition.Above)//(excludes case 11a)
                        {
                            isEdgeToEdge = DetailingCustomAssembliesServices.HasTopFlange((ProfilePart)boundedObject);
                        }
                        // Overlap is distance from Pt23 to the inside of the top flange
                        outsideOverlap = measurements[DetailingCustomAssembliesConstants.DimPt23ToTopInside];
                        // Outside clearance is distance from Pt23 to flange right
                        outsideClearance = measurements[DetailingCustomAssembliesConstants.DimPt23ToBottomInside];
                    }
                }
            }
            else //bottom flange
            {
                if (measurements.Keys.Contains(DetailingCustomAssembliesConstants.DimPt15ToPt17))
                {
                    edgeLength = measurements[DetailingCustomAssembliesConstants.DimPt15ToPt17];

                    // If the inside of the bottom flange is below the bottom corner
                    RelativePointPosition btmInsidePos15 = DetailingCustomAssembliesServices.GetRelativePointPosition(15, BoundedEdgeSectionFaceType.InsideBottomFlange, measurements);
                    RelativePointPosition btmInsidePos17 = DetailingCustomAssembliesServices.GetRelativePointPosition(17, BoundedEdgeSectionFaceType.InsideBottomFlange, measurements);
                    RelativePointPosition topInsidePos15 = DetailingCustomAssembliesServices.GetRelativePointPosition(15, BoundedEdgeSectionFaceType.InsideTopFlange, measurements);
                    RelativePointPosition topInsidePos17 = DetailingCustomAssembliesServices.GetRelativePointPosition(17, BoundedEdgeSectionFaceType.InsideTopFlange, measurements);
                    RelativePointPosition btmOutsidePos15 = DetailingCustomAssembliesServices.GetRelativePointPosition(15, BoundedEdgeSectionFaceType.Bottom, measurements);
                    RelativePointPosition btmOutsidePos17 = DetailingCustomAssembliesServices.GetRelativePointPosition(17, BoundedEdgeSectionFaceType.Bottom, measurements);
                    RelativePointPosition topOutsidePos15 = DetailingCustomAssembliesServices.GetRelativePointPosition(15, BoundedEdgeSectionFaceType.Top, measurements);
                    RelativePointPosition topOutsidePos17 = DetailingCustomAssembliesServices.GetRelativePointPosition(17, BoundedEdgeSectionFaceType.Top, measurements);

                    if (btmInsidePos17 == RelativePointPosition.Above)
                    {
                        //Inside clearance is distance from Pt 17 to the inside of the bottom flange
                        insideClearance = measurements[DetailingCustomAssembliesConstants.DimPt17ToBottomInside];

                        //If the inside of the top flange is above the bottom corner
                        if (topInsidePos17 == RelativePointPosition.Below)
                        {
                            //If the inside of the top flange is above the top corner
                            if (topInsidePos15 == RelativePointPosition.Below || topInsidePos15 == RelativePointPosition.Coincident)//Conditions 11 and 12 in BoundingEdgeConditions.sha
                            {
                                //Overlap is distance from Pt15 to Pt17
                                insideOverlap = edgeLength;

                                //Outside clearance is distance from Pt15 to inside of top flange
                                outsideClearance = measurements[DetailingCustomAssembliesConstants.DimPt15ToTopInside];

                                //If the inside of the top flange intersects the edge, the overlap is from Pt17 to the inside of the top flange

                            }
                            else //' Conditions 9 and 10 in BoundingEdgeConditions.sha
                            {
                                isEdgeToEdge = DetailingCustomAssembliesServices.HasTopFlange((ProfilePart)boundedObject);
                                insideOverlap = measurements[DetailingCustomAssembliesConstants.DimPt17ToTopInside];
                            }
                        }
                        else if (topOutsidePos17 == RelativePointPosition.Below)//Conditions 3-8 in BoundingEdgeConditions.sha
                        {
                            isEdgeToEdge = DetailingCustomAssembliesServices.HasTopFlange((ProfilePart)boundedObject);
                            insideOverlap = -measurements[DetailingCustomAssembliesConstants.DimPt17ToTopInside];
                        }
                    }
                    //If the inside of the bottom edge is above the top corner (cases 4a, 5a, 7a, 8a, 9a, 10a, 11a)
                    else if (btmInsidePos15 == RelativePointPosition.Below || btmInsidePos15 == RelativePointPosition.Coincident)
                    {
                        if (btmOutsidePos15 == RelativePointPosition.Above)//(excludes case 11a)
                        {
                            isEdgeToEdge = DetailingCustomAssembliesServices.HasBottomFlange((ProfilePart)boundedObject);

                        }
                        //Outside clearance is distance from Pt15 to inside of top flange
                        outsideClearance = measurements[DetailingCustomAssembliesConstants.DimPt15ToTopInside];

                        if (btmInsidePos15 == RelativePointPosition.Below)
                        {
                            //Overlap is distance from Pt15 to the inside of the bottom flange
                            outsideOverlap = -measurements[DetailingCustomAssembliesConstants.DimPt15ToBottomInside];
                        }
                    }
                    //If the inside of the bottom flange intersects the edge
                    else
                    {
                        if (btmInsidePos17 == RelativePointPosition.Below)
                        {
                            isEdgeToEdge = DetailingCustomAssembliesServices.HasBottomFlange((ProfilePart)boundedObject);

                        }

                        // Overlap is distance from Pt15 to the inside of the bottom flange
                        outsideOverlap = measurements[DetailingCustomAssembliesConstants.DimPt15ToBottomInside];

                        // Outside clearance is distance from Pt15 to inside of top flange
                        outsideClearance = measurements[DetailingCustomAssembliesConstants.DimPt15ToTopInside];
                    }
                }
            }
        }

        /// <summary>
        ///  Gets the flange over lap data between bounded and bounding objects.  
        /// </summary>
        /// <param name="isBottomEdge">A boolean which determines if it is bottom edge or not.</param>
        /// <param name="isBottomFlange">A boolean which determines if it is bottom flange or not.</param>        
        /// <param name="measurements">measurement data related to connected edge.</param>
        /// <param name="edgeMappings">Edge cut mapping information.</param>
        /// <param name="boundedObject">Bounded object of end cut.</param>
        /// <param name="insideOverlap">Inside overlap value</param>
        /// <param name="outsideOverlap">Outside overlap value</param>
        /// <param name="insideClearance">Inside clearance</param>
        /// <param name="outsideClearance">Outside clearence</param>
        /// <param name="edgeLength">Edge length</param>
        /// <param name="isEdgeToEdge">Boolean value which tells if the overlap is between the edge and edge</param> 
        private static void GetFlangeEdgeOverlapAndClearance(bool isBottomEdge, bool isBottonFlange, Dictionary<string, double> measurements, Dictionary<int, int> edgeMappings, BusinessObject boundedObject, out double insideOverlap,
            out double outsideOverlap, out double insideClearance, out double outsideClearance, out double edgeLength, out bool isEdgeToEdge)
        {

            //        |                                                         ' |
            //        |                                                         ' |
            //        |---------------------------------------------            ' |
            //        |              ^                                          ' |--------------------------------------------
            //        |              |                                          ' |-------------------------------------------
            //        |              |                                          ' |              ^
            //        |        Inside clearance                                 ' |              |
            //        |              |                                          ' |        Inside clearance
            //        |              |                                          ' |              |
            //        |              v                                          ' |              v
            //        +------------------+                                      ' +------------------+
            //                           |InsideOverlap                         '                    |
            //                           |---------------------------           '                    |
            //           Bounding        |---------------------------           '    Bounding        |Inside Overlap
            //                           |                                      '                    |
            //                           |Outside Overlap                       '                    |
            //        -------------------+                                      ' -------------------+
            //                 |        ^                  Bounded              '          |        ^                  Bounded
            //                 |        |                                       '          |        |
            //                 | Outside clearance                              '          | Outside clearance
            //                 |        |                                       '          |        |
            //                 |        v                                       '          |        v
            //                 +----------------------------------------        '          +----------------------------------------

            insideOverlap = 0.0;
            outsideOverlap = 0.0;
            insideClearance = 0.0;
            outsideClearance = 0.0;
            edgeLength = 0.0;
            isEdgeToEdge = false;

            int sectionEdgeMap = (int)SectionFaceType.Unknown;
            if (isBottomEdge)
            {
                edgeMappings.TryGetValue((int)SectionFaceType.Bottom_Flange_Right, out sectionEdgeMap);
            }
            else
            {
                edgeMappings.TryGetValue((int)SectionFaceType.Top_Flange_Right, out sectionEdgeMap);
            }

            if (!(sectionEdgeMap == (int)SectionFaceType.Unknown))
            {
                Dictionary<string, double> endCutMeasurements = measurements;
                if (isBottomEdge)
                {
                    if (!endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt21ToPt23, out edgeLength)) //Key does not exist
                    {
                        return;
                    }
                }
                else
                {
                    if (!endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt15ToPt17, out edgeLength)) //Key does not exist
                    {
                        return;
                    }
                }

                bool hasRightFlange = false;
                bool hasLeftFlange = false;

                // appropriate Section flange are returned based on crossection of input object
                string sectionType = DetailingCustomAssembliesServices.GetSectionType(boundedObject);
                SectionFlanges availableFlanges = DetailingCustomAssembliesServices.GetSectionFlanges(sectionType);
                if (isBottonFlange)
                {
                    if (availableFlanges.HasFlag(SectionFlanges.BottomLeft))
                    {
                        hasLeftFlange = true;
                    }
                    if (availableFlanges.HasFlag(SectionFlanges.BottomRight))
                    {
                        hasRightFlange = true;
                    }
                }
                else
                {
                    if (availableFlanges.HasFlag(SectionFlanges.TopLeft))
                    {
                        hasLeftFlange = true;
                    }
                    if (availableFlanges.HasFlag(SectionFlanges.TopRight))
                    {
                        hasRightFlange = true;
                    }
                }
                // If this is overlap at the bounding bottom edge                        
                if (isBottomEdge)
                {
                    // Get position of the four edges relative to the top and bottom corner of the edge                            
                    RelativePointPosition webRightPos21 = DetailingCustomAssembliesServices.GetRelativePointPosition(21, BoundedEdgeSectionFaceType.WebRight, measurements);
                    RelativePointPosition webRightPos23 = DetailingCustomAssembliesServices.GetRelativePointPosition(21, BoundedEdgeSectionFaceType.WebRight, measurements);

                    RelativePointPosition webLeftPos21 = DetailingCustomAssembliesServices.GetRelativePointPosition(21, BoundedEdgeSectionFaceType.WebLeft, measurements);
                    RelativePointPosition webLeftPos23 = DetailingCustomAssembliesServices.GetRelativePointPosition(23, BoundedEdgeSectionFaceType.WebLeft, measurements);

                    RelativePointPosition flangeRightPos21 = DetailingCustomAssembliesServices.GetRelativePointPosition(21, BoundedEdgeSectionFaceType.FlangeRight, measurements);
                    RelativePointPosition flangeRightPos23 = DetailingCustomAssembliesServices.GetRelativePointPosition(23, BoundedEdgeSectionFaceType.FlangeRight, measurements);

                    RelativePointPosition flangeLeftPos21 = DetailingCustomAssembliesServices.GetRelativePointPosition(21, BoundedEdgeSectionFaceType.FlangeLeft, measurements);
                    RelativePointPosition flangeLeftPos23 = DetailingCustomAssembliesServices.GetRelativePointPosition(23, BoundedEdgeSectionFaceType.FlangeLeft, measurements);

                    // If the web is entirely inside the edge (cases 1 and 2)
                    if (webRightPos21 == RelativePointPosition.Below || webRightPos21 == RelativePointPosition.Coincident)
                    {
                        // There is no overlap or clearance if entire flange is inside                               
                        if (!flangeRightPos21.HasFlag(RelativePointPosition.Below))
                        {
                            // Inside clearance is distance from top corner to web right
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt21ToWR, out insideClearance);

                            // If flange right is below the top corner                                
                            if (flangeRightPos21 == RelativePointPosition.Above)
                            {
                                // If flange right is below the bottom corner                                      
                                if (flangeRightPos23 == RelativePointPosition.Above)
                                {
                                    // Outside overlap is distance from top corner to bottom corner
                                    endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt21ToPt23, out insideOverlap);

                                    // Outside clearance is distance from bottom corner to flange right
                                    endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt23ToFR, out outsideClearance);
                                }
                                else
                                {
                                    // If flange right intersects the edge, the overlap is from top corner to flange right
                                    endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt21ToFR, out insideOverlap);
                                }
                            }
                        }
                    }
                    // If the web overlaps the inside corner(cases 3 and 6)
                    else if ((webLeftPos21 == RelativePointPosition.Below || webLeftPos21 == RelativePointPosition.Coincident) &&
                            (webRightPos21 == RelativePointPosition.Above && webRightPos23 == RelativePointPosition.Below))
                    {
                        isEdgeToEdge = true;

                        // Inside overlap is distance from top corner to web left (negative value)
                        endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt21ToWL, out insideOverlap);
                        insideOverlap = -insideOverlap;

                        // If there is a left flange, the inside clearance is from the top corner to flange left
                        if (hasLeftFlange)
                        {
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt21ToFL, out insideClearance);
                        }

                        // If it has a right flange
                        if (hasRightFlange)
                        {
                            // Overlap is from bottom corner to web right
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt23ToWR, out outsideOverlap);

                            // Outside clearance is distance from bottom corner to flange right
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt23ToFR, out outsideClearance);
                        }
                    }
                    // If the web completely overlaps the edge (cases 4,5,7,8)
                    else if ((webLeftPos21 == RelativePointPosition.Below || webLeftPos21 == RelativePointPosition.Coincident) &&
                            (webRightPos23 == RelativePointPosition.Above || webRightPos23 == RelativePointPosition.Coincident))
                    {
                        isEdgeToEdge = true;

                        // Inside overlap is distance from top corner to web left (negative value)
                        endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt21ToWL, out insideOverlap);
                        insideOverlap = -insideOverlap;

                        // Outside overlap is distance from bottom corner to web right (negative value)
                        endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt23ToWR, out insideOverlap);
                        insideOverlap = -insideOverlap;

                        // If there is a left flange, the inside clearance is from the top corner to flange left
                        if (hasLeftFlange)
                        {
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt21ToFL, out insideClearance);
                        }

                        // If there is a right flange, the outside clearance is from the bottom corner to flange right
                        if (hasRightFlange)
                        {
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt23ToFR, out outsideClearance);
                        }
                    }
                    // If the web is entirely within the edge (case 13)
                    else if (webLeftPos21 == RelativePointPosition.Above && webRightPos23 == RelativePointPosition.Below)
                    {
                        isEdgeToEdge = true;

                        // If there is a left flange
                        if (hasLeftFlange)
                        {
                            // Inside clearance is from top corner to flange left
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt21ToFL, out insideClearance);

                            // Inside overlap is from top corner to web left
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt21ToWL, out insideOverlap);
                        }

                        // If there is a right flange
                        if (hasRightFlange)
                        {
                            // Outside overlap is from bottom corner to web right
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt23ToWR, out outsideOverlap);

                            // Outside clearance is distance from bottom corner to flange right
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt23ToFR, out outsideClearance);
                        }
                    }
                    // If the web overlaps the outside corner (cases 9 and 10)
                    else if ((webLeftPos21 == RelativePointPosition.Above && webLeftPos23 == RelativePointPosition.Below) &&
                             (webRightPos23 == RelativePointPosition.Above || webRightPos23 == RelativePointPosition.Coincident))
                    {
                        isEdgeToEdge = true;

                        // If there is a left flange
                        if (hasLeftFlange)
                        {
                            // Inside overlap is distance from top corner to web left
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt21ToWL, out insideOverlap);

                            // Inside clearance is distance from top corner to flange left
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt21ToFL, out insideClearance);
                        }

                        // Outside overlap is distance from bottom corner to web right (negative value)
                        endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt23ToWR, out outsideOverlap);
                        outsideOverlap = -outsideOverlap;

                        // If it has a right flange, the outside clearance is distance from bottom corner to flange right
                        if (hasRightFlange)
                        {
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt23ToFR, out outsideClearance);
                        }
                    }

                    // If the web is entirely outside the edge (cases 11 and 12)
                    else if (webLeftPos23 == RelativePointPosition.Above)
                    {
                        // If there is a left flange
                        if (hasLeftFlange)
                        {
                            // Outside clearance is from outside corner to web left
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt23ToWL, out outsideClearance);

                            // If flange left is at or above the top corner
                            if (flangeLeftPos21 == RelativePointPosition.Below || flangeLeftPos21 == RelativePointPosition.Coincident)
                            {
                                // Inside overlap is from top corner to bottom corner
                                endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt21ToPt23, out insideOverlap);

                                // Inside clearance is from top corner to flange left
                                endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt21ToFL, out insideClearance);
                            }
                            // If flange left intersects the edge
                            else if (flangeLeftPos23 == RelativePointPosition.Below)
                            {
                                // Inside overlap is from bottom corner to flange left
                                endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt23ToFL, out insideOverlap);
                            }
                        }
                    }
                }
                // If this is overlap at the bounding top edge
                else
                {
                    // Get position of the four edges relative to the top and bottom corner of the edge
                    RelativePointPosition webRightPos17 = DetailingCustomAssembliesServices.GetRelativePointPosition(17, BoundedEdgeSectionFaceType.WebRight, measurements);
                    RelativePointPosition webLeftPos17 = DetailingCustomAssembliesServices.GetRelativePointPosition(17, BoundedEdgeSectionFaceType.WebLeft, measurements);
                    RelativePointPosition flangeRightPos17 = DetailingCustomAssembliesServices.GetRelativePointPosition(17, BoundedEdgeSectionFaceType.FlangeRight, measurements);
                    RelativePointPosition flangeLeftPos17 = DetailingCustomAssembliesServices.GetRelativePointPosition(17, BoundedEdgeSectionFaceType.FlangeLeft, measurements);

                    RelativePointPosition webRightPos15 = DetailingCustomAssembliesServices.GetRelativePointPosition(15, BoundedEdgeSectionFaceType.WebRight, measurements);
                    RelativePointPosition webLeftPos15 = DetailingCustomAssembliesServices.GetRelativePointPosition(15, BoundedEdgeSectionFaceType.WebLeft, measurements);
                    RelativePointPosition flangeRightPos15 = DetailingCustomAssembliesServices.GetRelativePointPosition(15, BoundedEdgeSectionFaceType.FlangeRight, measurements);
                    RelativePointPosition flangeLeftPos15 = DetailingCustomAssembliesServices.GetRelativePointPosition(15, BoundedEdgeSectionFaceType.FlangeLeft, measurements);

                    // If the web is entirely inside the edge (cases 1 and 2)
                    if (webLeftPos17 == RelativePointPosition.Above || webLeftPos17 == RelativePointPosition.Coincident)
                    {
                        // There is no overlap or clearance if entire flange is inside
                        if (flangeLeftPos17 != RelativePointPosition.Above)
                        {
                            // Inside clearance is distance from bottom corner to web left
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt17ToWL, out insideClearance);

                            // If flange left is above the bottom corner
                            if (flangeLeftPos17 == RelativePointPosition.Below)
                            {
                                // If flange left is above the top corner
                                if (flangeLeftPos15 == RelativePointPosition.Below)
                                {
                                    // Outside overlap is distance from top corner to bottom corner
                                    endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt15ToPt17, out insideOverlap);

                                    // Outside clearance is distance from top corner to flange left
                                    endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt15ToFL, out outsideClearance);
                                }
                                // If flange left intersects the edge, the overlap is from bottom corner to flange left
                                else
                                {
                                    endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt17ToFL, out insideOverlap);
                                }
                            }
                        }
                    }
                    // If the web overlaps the bottom corner (cases 3 and 6)
                    else if ((webRightPos17 == RelativePointPosition.Above || webRightPos17 == RelativePointPosition.Coincident) &&
                            (webLeftPos15 == RelativePointPosition.Above))
                    {
                        isEdgeToEdge = true;

                        // Inside overlap is distance from bottom corner to web right (negative value)
                        endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt17ToWR, out insideOverlap);
                        insideOverlap = -insideOverlap;

                        // If there is a right flange, the inside clearance is from the bottom corner to flange right
                        if (hasRightFlange)
                        {
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt17ToFR, out insideClearance);
                        }

                        // If it has a left flange
                        if (hasLeftFlange)
                        {
                            // Overlap is from top corner to web left
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt15ToWL, out outsideOverlap);

                            // Outside clearance is distance from top corner to flange left
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt15ToFL, out outsideClearance);
                        }
                    }
                    // If the web completely overlaps the edge (cases 4,5,7,8)
                    else if ((webRightPos17 == RelativePointPosition.Above || webRightPos17 == RelativePointPosition.Coincident) &&
                             (webLeftPos15 == RelativePointPosition.Below || webLeftPos15 == RelativePointPosition.Coincident))
                    {
                        isEdgeToEdge = true;

                        // Inside overlap is distance from bottom corner to web right (negative value)
                        endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt17ToWR, out insideOverlap);
                        insideOverlap = -insideOverlap;

                        // Outside overlap is distance from top corner to web left (negative value)
                        endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt15ToWL, out insideOverlap);
                        insideOverlap = -insideOverlap;

                        // If there is a right flange, the inside clearance is from the bottom corner to flange right
                        if (hasRightFlange)
                        {
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt17ToFR, out insideClearance);
                        }

                        // If there is a left flange, the outside clearance is from the top corner to flange left
                        if (hasLeftFlange)
                        {
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt15ToFL, out outsideClearance);
                        }
                    }
                    // If the web is entirely within the edge (case 13)
                    else if (webRightPos17 == RelativePointPosition.Below && webLeftPos15 == RelativePointPosition.Above)
                    {
                        isEdgeToEdge = true;

                        // If there is a right flange
                        if (hasRightFlange)
                        {
                            // Inside clearance is distance from bottom corner to flange right
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt17ToFR, out insideClearance);

                            // Inside overlap is from bottom corner to web right
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt17ToWR, out insideOverlap);
                        }

                        // If there is a left flange
                        if (hasLeftFlange)
                        {
                            // Inside overlap is from bottom corner to web left
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt15ToWL, out outsideOverlap);

                            // Outside clearance is from top corner to flange left
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt15ToFL, out outsideClearance);
                        }
                    }
                    // If the web overlaps the outside corner (cases 9 and 10)
                    else if ((webRightPos17 == RelativePointPosition.Below && webRightPos15 == RelativePointPosition.Above) &&
                             (webLeftPos15 == RelativePointPosition.Below || webLeftPos15 == RelativePointPosition.Coincident))
                    {
                        isEdgeToEdge = true;

                        // If it has a right flange
                        if (hasRightFlange)
                        {
                            // Inside overlap is distance from bottom corner to web right
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt17ToWR, out insideOverlap);

                            // Inside clearance is distance from bottom corner to flange right
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt17ToFR, out insideClearance);
                        }

                        // Outside overlap is distance from top corner to web left (negative value)
                        endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt15ToWL, out outsideOverlap);
                        outsideOverlap = -outsideOverlap;

                        // If it has a left flange, the outside clearance is distance from top corner to flange left
                        if (hasLeftFlange)
                        {
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt15ToFL, out outsideClearance);
                        }
                    }
                    // If the web is entirely outside the edge (cases 11 and 12)
                    else if (webRightPos15 == RelativePointPosition.Below)
                    {
                        // If there is a right flange
                        if (hasRightFlange)
                        {
                            // Outside clearance is from outside corner to web right
                            endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt15ToWR, out outsideClearance);

                            // If flange right is at or below the bottom corner
                            if (flangeRightPos17 == RelativePointPosition.Above || flangeRightPos17 == RelativePointPosition.Coincident)
                            {
                                // Inside overlap is from top corner to bottom corner
                                endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt15ToPt17, out insideOverlap);

                                // Inside clearance is from bottom corner to flange right
                                endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt17ToFR, out insideClearance);
                            }
                            // If flange right intersects the edge
                            else if (flangeRightPos15 == RelativePointPosition.Above)
                            {
                                // Inside overlap is from top corner to flange right
                                endCutMeasurements.TryGetValue(DetailingCustomAssembliesConstants.DimPt15ToFR, out insideOverlap);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Determines the whether the cornerfeature is needed on shapeAtEdge value and codelisttable(shapeatedge,shapeatedgeoverlap,shapeatedgeoutside).
        /// </summary>
        /// <param name="codeListAnswer">the collection name which can be shapeatedge,shapestedgeoverlap,shapeatedgeoutside.</param>
        /// <param name="shapeAtEdge">answer for the the selection question shapeAtEdge.</param>
        /// <param name="extendOrSetBackAnswer">answer for the the selection question extendOrSetBackAnswer.</param>
        /// <param name="overlap">overlap betwen edges.</param>
        /// <param name="isEdgeOnly">Flag to specify if the is edge only or not.</param>
        /// <param name="isEdgeToEdge">Boolean value which tells if the overlap is between the edge and edge.</param>
        /// <returns></returns>
        private static bool IsCFNeededForShapeAtEdgeAnswer(string codeListAnswer, int shapeAtEdge, int extendOrSetBackAnswer, double overlap, bool isEdgeOnly, bool isEdgeToEdge)
        {
            bool isCornerFeatureNeeded = false;
            bool isCornerFeatureNeededWithShapeAtEdge = false;
            if (codeListAnswer.Equals(DetailingCustomAssembliesConstants.ShapeAtEdgeCol, StringComparison.OrdinalIgnoreCase))
            {
                switch (shapeAtEdge)
                {
                    //If to the Edge or corner
                    case (int)ShapeAtEdge.FaceToCorner:
                    case (int)ShapeAtEdge.FaceToEdge:
                    case (int)ShapeAtEdge.InsideToEdge:
                        isCornerFeatureNeeded = IsCFNeededIfNotEdgeToEdge(shapeAtEdge, overlap, (int)ShapeAtEdge.InsideToEdge, 0, extendOrSetBackAnswer, isEdgeOnly, isEdgeToEdge);
                        break;

                    // If to the flange starting from face, inside, or outside: a feature is always needed as long as it is edge-to-edge
                    case (int)ShapeAtEdge.FaceToFlange:
                    case (int)ShapeAtEdge.InsideToFlange:
                        isCornerFeatureNeeded = isEdgeToEdge ? true : false;
                        break;
                    //If to the flange starting from the edge: a feature is always needed as long as there is edge-to-edge and a minimum overlap is met
                    case (int)ShapeAtEdge.CornerToFlange:
                    case (int)ShapeAtEdge.EdgeToFlange:
                        isCornerFeatureNeeded = IsCfNeededIfEdgeToEdge(shapeAtEdge, (int)ShapeAtEdge.EdgeToFlange, extendOrSetBackAnswer, overlap, isEdgeOnly, isEdgeToEdge);
                        break;


                    case (int)ShapeAtEdge.EdgeToOutside:
                    case (int)ShapeAtEdge.FaceToOutside:
                    case (int)ShapeAtEdge.InsideToOutside:
                    case (int)ShapeAtEdge.CornerToOutside:
                        //In the code below we are assuming that any extension from the near corner is large enough, or any offset from a location
                        //other than the near corner is small enough, to leave enough material for the corner feature to cut, and leave
                        //a realistic amount of material behind for welding.  We can enhance the logic to compare the actual offsets to the
                        //flange length in the sketching plane at a later time.
                        if (extendOrSetBackAnswer.Equals(ExtentOrOffset.OffsetNearCorner))
                        {
                            isCornerFeatureNeededWithShapeAtEdge = true;
                        }

                        //For "ToEdge" configuration, no feature is needed, even if extended
                        if (isEdgeOnly && (shapeAtEdge.Equals((int)ShapeAtEdge.EdgeToOutside) || shapeAtEdge.Equals((int)ShapeAtEdge.CornerToOutside)))
                        {
                            isCornerFeatureNeededWithShapeAtEdge = true;
                        }
                        if (!isCornerFeatureNeededWithShapeAtEdge)
                        {
                            isCornerFeatureNeeded = true;
                        }
                        break;
                }

            }
            else if (codeListAnswer.Equals(DetailingCustomAssembliesConstants.ShapeAtEdgeOverlapCol, StringComparison.OrdinalIgnoreCase))
            {
                switch (shapeAtEdge)
                {
                    //If to the Edge or corner
                    case (int)ShapeAtEdgeOverlap.FaceToEdge:
                    case (int)ShapeAtEdgeOverlap.InsideToEdge:
                    case (int)ShapeAtEdgeOverlap.FaceToInsideCorner:
                        isCornerFeatureNeeded = IsCFNeededIfNotEdgeToEdge(shapeAtEdge, overlap, (int)ShapeAtEdge.InsideToEdge, 0, extendOrSetBackAnswer, isEdgeOnly, isEdgeToEdge);
                        break;

                    case (int)ShapeAtEdgeOverlap.EdgeToOutside:
                    case (int)ShapeAtEdgeOverlap.FaceToOutsideCorner:
                    case (int)ShapeAtEdgeOverlap.FaceToOutside:
                    case (int)ShapeAtEdgeOverlap.InsideCornerToOutside:
                    case (int)ShapeAtEdgeOverlap.InsideToOutsideCorner:
                    case (int)ShapeAtEdgeOverlap.InsideToOutside:
                        //In the code below we are assuming that any extension from the near corner is large enough, or any offset from a location
                        //other than the near corner is small enough, to leave enough material for the corner feature to cut, and leave
                        //a realistic amount of material behind for welding.  We can enhance the logic to compare the actual offsets to the
                        //flange length in the sketching plane at a later time.
                        if (extendOrSetBackAnswer.Equals(ExtentOrOffset.OffsetNearCorner))
                        {
                            isCornerFeatureNeededWithShapeAtEdge = false;
                        }

                        //For "ToEdge" configuration, no feature is needed, even if extended
                        if (isEdgeOnly && (shapeAtEdge.Equals((int)ShapeAtEdgeOverlap.EdgeToOutside)))
                        {
                            isCornerFeatureNeededWithShapeAtEdge = true;
                        }
                        if (!isCornerFeatureNeededWithShapeAtEdge)
                        {
                            isCornerFeatureNeeded = true;
                        }
                        break;
                }

            }
            else if (codeListAnswer.Equals(DetailingCustomAssembliesConstants.ShapeAtEdgeOutsideCol, StringComparison.OrdinalIgnoreCase))
            {
                switch (shapeAtEdge)
                {
                    //If to the Edge or corner
                    case (int)ShapeAtEdgeOutside.OutsideToEdge:
                        isCornerFeatureNeeded = IsCFNeededIfNotEdgeToEdge(shapeAtEdge, overlap, (int)ShapeAtEdge.InsideToEdge, 0, extendOrSetBackAnswer, isEdgeOnly, isEdgeToEdge);
                        break;

                    case (int)ShapeAtEdgeOutside.OutsideToFlange:
                        isCornerFeatureNeeded = isEdgeToEdge ? true : false;
                        break;

                    //If to the flange starting from the edge: a feature is always needed as long as there is edge-to-edge and a minimum overlap is met
                    case (int)ShapeAtEdgeOutside.CornerToFlange:
                    case (int)ShapeAtEdgeOutside.EdgeToFlange:
                        isCornerFeatureNeeded = IsCfNeededIfEdgeToEdge(shapeAtEdge, (int)ShapeAtEdge.EdgeToFlange, extendOrSetBackAnswer, overlap, isEdgeOnly, isEdgeToEdge);
                        break;
                    //No corner features for cases that always go to the outside
                    case (int)ShapeAtEdgeOutside.OutsideToOutside:
                        isCornerFeatureNeeded = false;
                        break;

                    case (int)ShapeAtEdgeOutside.EdgeToOutside:
                    case (int)ShapeAtEdgeOutside.CornerToOutside:
                        //In the code below we are assuming that any extension from the near corner is large enough, or any offset from a location
                        //other than the near corner is small enough, to leave enough material for the corner feature to cut, and leave
                        //a realistic amount of material behind for welding.  We can enhance the logic to compare the actual offsets to the
                        //flange length in the sketching plane at a later time.
                        if (extendOrSetBackAnswer.Equals(ExtentOrOffset.OffsetNearCorner))
                        {
                            isCornerFeatureNeededWithShapeAtEdge = true;
                        }

                        //For "ToEdge" configuration, no feature is needed, even if extended
                        if (isEdgeOnly && (shapeAtEdge.Equals((int)ShapeAtEdgeOutside.EdgeToOutside) || shapeAtEdge.Equals((int)ShapeAtEdgeOutside.CornerToOutside)))
                        {
                            isCornerFeatureNeededWithShapeAtEdge = true;
                        }
                        if (!isCornerFeatureNeededWithShapeAtEdge)
                        {
                            isCornerFeatureNeeded = true;
                        }
                        break;
                }
            }
            return isCornerFeatureNeeded;
        }

        /// <summary>
        /// Determines the whether the cornerfeature is needed on shapeAtEdge and extendOrSetBackAnswer values
        /// </summary>
        /// <param name="shapeAtEdge">answer for the the selection question shapeAtEdge.</param>
        /// <param name="overlap">overlap betwen edges.</param>
        /// <param name="insideToEdge">insideToEdge value from shapeAtEdge or shapAtOverlap collection.</param>
        /// <param name="outsideToEdge">outsideToEdge value from shapeAtEdge or shapAtOverlap collection.</param>
        /// <param name="extendOrSetBackAnswer">answer for the the selection question extendOrSetBackAnswer.</param>
        /// <param name="isEdgeOnly">Flag to specify if the is edge oly or not.</param>
        /// <param name="isEdgeToEdge">Boolean value which tells if the overlap is between the edge and edge.</param>
        /// <returns></returns>
        private static bool IsCFNeededIfNotEdgeToEdge(int shapeAtEdge, double overlap, int insideToEdge, int outsideToEdge, int extendOrSetBackAnswer, bool isEdgeOnly, bool isEdgeToEdge)
        {
            bool isCornerFeatureNeeded = false;
            double minOverlap = 0.0;
            // If not edge-to-edge, there must be an overlap of at least 5mm, 10mm if from the inside or outside
            // Does not apply if not ToEdge case and bounded is extended past the near corner
            //The tolerance is theoretical (i.e. no user requirements have been seen to drive this)
            // It is intended to prevent unreasonably small features from being placed, and to prevent 'slivers' (areas where
            // one or more cuts leave a long thin piece of material which would be weak and unsuited for welding)

            //Compute minimum overlap per comments above

            minOverlap = Math3d.FitTolerance * 5;

            if (shapeAtEdge.Equals(insideToEdge) || shapeAtEdge.Equals(outsideToEdge))
            {
                minOverlap = 0.01;
            }

            // ToEdge case
            bool isCornerFeatureNeededWithShapeAtEdge = false;
            // ToEdge case
            if (isEdgeOnly)
            {
                // gsOutsideToEdge case needs feature even if minimum overlap is not met, as long as a bounding edge meets a bounded edge
                if ((!shapeAtEdge.Equals(outsideToEdge)) && overlap < minOverlap)
                {
                    isCornerFeatureNeededWithShapeAtEdge = true;
                }

                if ((!isEdgeToEdge) && overlap < minOverlap)
                {
                    isCornerFeatureNeededWithShapeAtEdge = true;
                }
            }
            // All other cases
            else
            {
                if ((!isEdgeToEdge) && (overlap < minOverlap) &&
                    ((extendOrSetBackAnswer.Equals(ExtentOrOffset.OffsetNearCorner))))
                {
                    // For now, assume anything else produces some overlap
                    isCornerFeatureNeededWithShapeAtEdge = true;
                }
            }
            if (!isCornerFeatureNeededWithShapeAtEdge)
            {
                isCornerFeatureNeeded = true;
            }
            return isCornerFeatureNeeded;
        }

        /// <summary>
        /// Determines the whether the cornerfeature is needed on shapeAtEdge and extendOrSetBackAnswer values.
        /// </summary>
        /// <param name="shapeAtEdge">answer for the the selection question shapeAtEdge.</param>
        /// <param name="edgeToFlange">insideToEdge value from shapeAtEdge or ShapeAtEdgeOutside collection.</param>
        /// <param name="extendOrSetBackAnswer">answer for the the selection question extendOrSetBackAnswer.</param>
        /// <param name="overlap">overlap betwen edges.</param>
        /// <param name="isEdgeOnly">Flag to specify if the is edge oly or not.</param>
        /// <param name="isEdgeToEdge">Boolean value which tells if the overlap is between the edge and edge.</param>
        /// <returns></returns>
        private static bool IsCfNeededIfEdgeToEdge(int shapeAtEdge, int edgeToFlange, int extendOrSetBackAnswer, double overlap, bool isEdgeOnly, bool isEdgeToEdge)
        {
            double minOverlap = 0.0;
            bool isCornerFeatureNeeded = false;
            if (isEdgeToEdge)
            {
                // If from corner, there must be an overlap of 5mm
                // If from edge, there must be an overlap of 10mm
                minOverlap = 5 * Math3d.FitTolerance;

                if (shapeAtEdge.Equals(edgeToFlange))
                {
                    minOverlap = 10 * Math3d.FitTolerance;
                }
                bool isCornerFeatureNeededWithShapAtEdge = false;
                //ToEdge case
                if (isEdgeOnly)
                {
                    if (overlap < minOverlap)
                    {
                        isCornerFeatureNeededWithShapAtEdge = true;
                    }
                }
                else // All other cases
                {
                    if (overlap < minOverlap)
                    {
                        //For now, assume anything else produces some overlap
                        if (extendOrSetBackAnswer.Equals(ExtentOrOffset.OffsetNearCorner))
                        {
                            isCornerFeatureNeededWithShapAtEdge = true;
                        }
                    }
                }
                if (!isCornerFeatureNeededWithShapAtEdge)
                {
                    isCornerFeatureNeeded = true;
                }

            }
            return isCornerFeatureNeeded;
        }

        #endregion Private Methods
    }
}
