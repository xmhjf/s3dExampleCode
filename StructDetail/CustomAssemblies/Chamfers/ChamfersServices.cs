//--------------------------------------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  ChamfersServices.cs
//
//Abstract
//	ChamfersServices is a helper class to have commom method implementation for .NET selector rule, parameter rule and definition of the chamfers.
//--------------------------------------------------------------------------------------------------------------------------------------
using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Common.Middle.Services;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using Ingr.SP3D.ReferenceData.Middle;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Helper class to have commom method implementation for .NET selector rule, parameter rule and definition of the chamfers.
    /// </summary>
    internal static class ChamfersServices
    {
        /// <summary>
        /// Gets the slope of the chamfer.
        /// </summary>
        /// <param name="chamferPart">The parent system of the chamfer</param>
        /// <param name="canRule">Can Rule</param>
        /// <returns></returns>
        internal static double GetChamferSlope(PlateSystem chamferParentSystem, CanRule canRule)
        {
            double chamferSlope = 0.0;

            if (chamferParentSystem.IsBuiltUp)
            {
                //if section type is of type BUCan the get the primary membersystem 
                if (chamferParentSystem.ParentBuiltUp.SectionType == MarineSymbolConstants.BUCan)
                {
                    if (canRule.SupportsInterface(DetailingCustomAssembliesConstants.IJUASMCanRuleEnd))
                    {
                        chamferSlope = StructHelper.GetDoubleProperty(canRule, DetailingCustomAssembliesConstants.IJUASMCanRuleEnd, DetailingCustomAssembliesConstants.ChamferSlope);
                    }
                    else if (canRule.SupportsInterface(DetailingCustomAssembliesConstants.IJUASMCanRuleInLine))
                    {
                        chamferSlope = StructHelper.GetDoubleProperty(canRule, DetailingCustomAssembliesConstants.IJUASMCanRuleInLine, DetailingCustomAssembliesConstants.ChamferSlope);
                    }
                    else if (canRule.SupportsInterface(DetailingCustomAssembliesConstants.IJUASMCanRuleStubEnd))
                    {
                        chamferSlope = StructHelper.GetDoubleProperty(canRule, DetailingCustomAssembliesConstants.IJUASMCanRuleStubEnd, DetailingCustomAssembliesConstants.ChamferSlope);
                    }
                }
            }
            return chamferSlope;
        }


        /// <summary>
        /// Sets the parameter values when the chamfer assembly parent is free end cut.
        /// </summary>
        /// <param name="freeEndCut">The free end cut which is the assembly parent of the chamfer.</param>>
        internal static double GetThicknessDifferenceWhenAssemblyParentIsFreeEndCut(FreeEndCut freeEndCut)
        {
            IPort freeEndCutBoundedPort = freeEndCut.EndPort;
            StiffenerPartBase freeEndCutBoundedObject = freeEndCutBoundedPort.Connectable as StiffenerPartBase;
            double profileThickness = 0;

            //Get the web thickness of the boundedObject if it is a stiffener
            if (freeEndCutBoundedObject != null)
            {
                profileThickness = DetailingCustomAssembliesServices.GetWebThickness(freeEndCutBoundedObject.CrossSection);
            }

            //Set the parameter values
            return (profileThickness / 2);
        }

        /// <summary>
        /// Gets the chamfer thickness value if the chamfer assembly parent is assembly connection.
        /// </summary>
        /// <param name="assemblyConnection">The assembly connection which is assembly parent of the chamfer.</param>
        /// <returns></returns>
        internal static double GetThicknessDifferenceWhenAssemblyParentIsAssemblyConnection(AssemblyConnection assemblyConnection)
        {
            double thicknessDifference = 0.0;

            //If the feature occurence parent is assembly connection get the connectable profiles                       
            ReadOnlyCollection<IPort> connectedPorts = assemblyConnection.GetPorts(PortType.Edge);

            if (connectedPorts.Count >= 2)
            {
                StiffenerPartBase penetratingPart = (StiffenerPartBase)connectedPorts[0].Connectable;
                StiffenerPartBase penetratedPart = (StiffenerPartBase)connectedPorts[1].Connectable;

                CrossSection crossSectionOfPenetratingPart = penetratingPart.CrossSection;
                CrossSection crossSectionOfPenetratedPart = penetratedPart.CrossSection;

                //check if this Assembly Connection is created by using a Plate/Profile Kunckle point
                ProfileKnuckle profileKnuckle = ProfileKnuckle.GetProfileKnuckleFromBoundedPort(assemblyConnection);
                double inclined = 1.0;
                //calculate the inclined angle
                if (profileKnuckle != null)
                {
                    double knuckleAngle = Math3d.Deg(profileKnuckle.Angle);
                    double angleToCalculateInclined = (180 - knuckleAngle) / 2;
                    inclined = Math.Cos(Math3d.Rad(angleToCalculateInclined));
                }

                double penetratingPartWebThickness = DetailingCustomAssembliesServices.GetWebThickness(crossSectionOfPenetratingPart);
                double penetratedPartWebThickness = DetailingCustomAssembliesServices.GetWebThickness(crossSectionOfPenetratedPart);

                //Calculate the thickness difference using the profile parts web thickness and the inclined angle
                thicknessDifference = Math.Abs(penetratingPartWebThickness - penetratedPartWebThickness) / Math.Abs(inclined);
            }
            //Set the parameter values
            return thicknessDifference;
        }
    }
}
