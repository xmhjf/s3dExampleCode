//---------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  ChamferParameterRule.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMChamferRules.dll
//  Original Class Name: ‘SingleBaseParm’ in VB content
//
//Abstract
//	ChamferSingleBaseParameterRule is a .NET parameter rule which is defining the custom parameter rule. 
//  This class subclasses from ChamferParameterRule. 
//--------------------------------------------------------------------------------------------------------
using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Structure.Middle;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Parameter rule which defines the custom parameter rule for Chamfers. 
    /// </summary>
    [RuleVersion("1.0.0.0")]
    [DefaultLocalizer(ChamfersResourceIds.ChamfersResources, ChamfersResourceIds.ChamfersAssembly)]
    //New content does not need to specify the interface name, but this is specified here to support code written against old rules which directly uses interface name to get parameter value. 
    [RuleInterface(DetailingCustomAssembliesConstants.IASMChamferRules_SingleBaseParm, DetailingCustomAssembliesConstants.IASMChamferRules_SingleBaseParm)]
    public class ChamferSingleBaseParameterRule : ChamferParameterRule
    {
        //==============================================================================================================
        //DefinitionName/ProgID of this symbol is "Chamfers,Ingr.SP3D.Content.Structure.ChamferSingleBaseParameterRule"
        //==============================================================================================================
        #region Parameters
        [ControlledParameter(1, DetailingCustomAssembliesConstants.Depth, DetailingCustomAssembliesConstants.Depth)]
        public ControlledParameterDouble depth;
        [ControlledParameter(2, DetailingCustomAssembliesConstants.AngleRadius, DetailingCustomAssembliesConstants.AngleRadius)]
        public ControlledParameterDouble angleRadius;
        #endregion Parameters

        #region Public override properties and methods

        /// <summary>
        /// Evaluates the parameter rule and validates the inputs of the chamfer.
        /// It computes the item parameters in the context of the smart occurrence and assigns them.
        /// </summary>
        public override void Evaluate()
        {
            try
            {
                //Validating inputs
                base.ValidateChamferInputs();

                //Check if any ToDo message of type error has been created during the validation of inputs
                if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                {
                    //ToDo list message is created while validating the inputs
                    return;
                }
                double minimumBaseValue = StructHelper.DISTTOL;
                Feature chamfer = (Feature)base.Occurrence;

                BusinessObject chamferCustomAssemblyParent = SymbolHelper.GetCustomAssemblyParent(chamfer);

                double baseValue = 0;
                double chamferOffset;

                //Setting slope to 72 degrees by default
                double slope = Math3d.Rad(72);

                //Check if the assembly parent is plant assembly connection
                if (chamferCustomAssemblyParent != null && chamferCustomAssemblyParent.SupportsInterface(MarineSymbolConstants.IJStructAssemblyConnection))
                {
                    SetParametersWhenAssemblyParentIsAssemblyConnection((AssemblyConnection)chamferCustomAssemblyParent);
                }
                else
                {
                    if (chamferCustomAssemblyParent is Feature)
                    {
                        //Check if the assembly parent is feature
                        BusinessObject featureParent = SymbolHelper.GetCustomAssemblyParent(chamferCustomAssemblyParent);
                        AssemblyConnection assemblyConnection = featureParent as AssemblyConnection;
                        FreeEndCut freeEndCut = featureParent as FreeEndCut;
                        if (assemblyConnection != null)
                        {
                            baseValue = ChamfersServices.GetThicknessDifferenceWhenAssemblyParentIsAssemblyConnection(assemblyConnection);
                        }
                        //Check if the assembly parent is free end cut
                        else if (freeEndCut != null)
                        {
                            baseValue = ChamfersServices.GetThicknessDifferenceWhenAssemblyParentIsFreeEndCut(freeEndCut);
                        }
                    }
                    else
                    {
                        IConnectable boundedConnectable = base.ChamferedPort.Connectable;
                        IConnectable boundingConnectable = base.DrivesChamferedPort.Connectable;

                        if (boundedConnectable is PlatePartBase && !(boundingConnectable is PlatePartBase))
                        {
                            GetStiffenerToPlateEdgeChamferData(out baseValue, out chamferOffset);

                            //As a Chamfer Depth of 0.0 is not valid
                            if (baseValue < Math3d.FitTolerance)
                            {
                                baseValue = Math3d.FitTolerance;
                            }
                        }
                        else
                        {
                            PlatePart boundedPart = (PlatePart)boundedConnectable;
                            PlatePart boundingPart = (PlatePart)boundingConnectable;
                            double thicknessDifference = boundedPart.Thickness - boundingPart.Thickness;

                            string chamferPartName = chamfer.PartName;

                            if (chamferPartName.Equals(DetailingCustomAssembliesConstants.SingleSidedBaseTeePC, StringComparison.OrdinalIgnoreCase))
                            {
                                string chamferThickness = (((PropertyValueString)((IPartSelection)chamfer).GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.ChamferThickness)).PropValue);
                                double chamferThicknessValue = Convert.ToDouble(chamferThickness);
                                baseValue = boundedPart.Thickness - chamferThicknessValue;
                            }
                            else if (chamferPartName.Equals(DetailingCustomAssembliesConstants.SingleSidedBasePC, StringComparison.OrdinalIgnoreCase) || chamferPartName.Equals(DetailingCustomAssembliesConstants.SingleSidedBase, StringComparison.OrdinalIgnoreCase))
                            {
                                GetDepthAndAngleRadiusBasedOnChamferCondition(minimumBaseValue, thicknessDifference, boundedPart, boundingPart, out baseValue, out slope);
                            }
                            else
                            {
                                baseValue = minimumBaseValue;
                            }

                            // Make sure Base is not 0.  If so, this parameter is being updated before the symbol will be 
                            //deleted due to assoc sequence (or there is an error in the rules)
                            if (baseValue < minimumBaseValue)
                            {
                                baseValue = minimumBaseValue;
                            }
                        }
                    }

                    //Set the parameter values
                    this.depth.Value = baseValue;
                    this.angleRadius.Value = slope;
                }
            }
            catch (Exception)
            {
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError,
                        base.GetString(ChamfersResourceIds.ToDoChamferParameterRule,
                    "Error while evaluating chamfer parameter rule."));
                }
            }
        }

        #endregion Public override properties and methods

        #region Private methods

        /// <summary>
        /// Sets the parameter values when assembly parent of the chamfer is assembly connection
        /// </summary>
        /// <param name="assemblyConnection">Assembly connection as parent of the chamfer</param>
        private void SetParametersWhenAssemblyParentIsAssemblyConnection(AssemblyConnection assemblyConnection)
        {
            double thicknessDifference = 0.0;

            PropertyValue chamferMeasurementsPropertyValue = assemblyConnection.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.ChamferMeasurement);
            int chamferMeasurements = (chamferMeasurementsPropertyValue != null) ? ((PropertyValueCodelist)chamferMeasurementsPropertyValue).PropValue : 0;

            PropertyValue chamferAnglePropertyValue = assemblyConnection.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.ChamferValue);
            double angle = (chamferAnglePropertyValue != null) ? Convert.ToDouble(chamferAnglePropertyValue.ToString()) : 0;

            //Calculate the angle radius depending upon the chamfer measurements
            double angleRadiusValue = (chamferMeasurements == (int)ChamferMeasurement.Slope) ? Math.PI / 2 - Math.Atan(angle) : Math.PI / 2 - (Math3d.Rad(angle));
            this.angleRadius.Value = angleRadiusValue;

            //Get the 2 edge ports connected to assemblyConnection            
            ReadOnlyCollection<IPort> connectedPorts = assemblyConnection.GetPorts(PortType.All);

            if (connectedPorts.Count >= 2)
            {
                IPort penetratingPort, penetratedPort;
                penetratingPort = connectedPorts[0];
                PlatePart penetratingPart = penetratingPort.Connectable as PlatePart;

                if (penetratingPart != null)
                {
                    penetratedPort = connectedPorts[1];
                }
                else
                {
                    penetratedPort = connectedPorts[0];
                    penetratingPort = connectedPorts[1];
                    penetratingPart = penetratingPort.Connectable as PlatePart;
                }

                MemberPart penetratedPart = penetratedPort.Connectable as MemberPart;


                if ((penetratedPart != null) && (penetratingPart != null))
                {
                    //Get Thickenss difference between Plate and Member Flange
                    thicknessDifference = GetThicknessDifferenceOfPlatePartOverMemberFlange(penetratedPart, penetratingPart);
                }
            }
            //Set the parameter value
            this.depth.Value = thicknessDifference;
        }

        /// <summary>
        /// Gets the thickness difference between plate part and member flange.
        /// </summary>
        /// <param name="boundedPart">The bounded member part.</param>
        /// <param name="boundingPart">The bounding plate part.</param>
        private double GetThicknessDifferenceOfPlatePartOverMemberFlange(MemberPart boundedPart, PlatePart boundingPart)
        {
            Position position1, position2;
            double thicknessDifference = 0.0;
            double toleranceValue = Math3d.FitTolerance;
            IPort platePort = null, bottomPort = null, basePort = null;
            SurfaceScopeType etype;
            bool intersectingTopFlange = false, intersectingBottomFlange = false;
            Vector plateNormal = null, topPortNormal = null;

            try
            {
                //Validates the top and bottom emulated ports
                platePort = GetEmulatedPort(boundingPart, boundedPart, (int)SectionFaceType.Top);
                basePort = GetEmulatedPort(boundingPart, boundedPart, (int)SectionFaceType.Bottom);

                int operatorId = (int)SectionFaceType.Top;

                IPort topPort = boundedPart.GetPort(TopologyGeometryType.Face, operatorId, GeometryOperationTypes.Cutout, ContextTypes.Lateral, GraphPosition.Before, (int)SectionFaceType.Top, false);

                ISurface plateSurface = (ISurface)platePort;
                //Gets the normals on plate and memberpart ports
                plateSurface.ScopeNormal(out etype, out plateNormal);

                ISurface topSurface = (ISurface)topPort;
                topSurface.ScopeNormal(out etype, out topPortNormal);

                //Check the plate Postion with respect to Member Flange
                if (!StructHelper.AreEqual(Math.Abs(plateNormal.Dot(topPortNormal)), 1.0))
                {
                    return thicknessDifference;
                }

                BusinessObject penetratedPartPort = null;
                Collection<ICurve> intersections;
                GeometryIntersectionType intersectionType;
                PlateSystem penetratingPartParent = (PlateSystem)boundingPart.RootPlateSystem;
                ISurface penetratingPort = (ISurface)penetratingPartParent.GetPorts(Ingr.SP3D.Common.Middle.PortType.Face)[0];
                topSurface.Intersect(penetratingPort, out intersections, out intersectionType);

                //Check if Plate is on Top Flange              
                if (intersections != null)
                {
                    intersectingTopFlange = true;
                }
                //if bIntersectingTopFlange is False, then checks for the tolerance value. If the distance between the plate and member is less than tolerance value,
                // make TopFlangeResultantIntersection as true
                if (!intersectingTopFlange)
                {
                    double memberPlateOffsetDistance = 0.0;
                    topSurface.DistanceBetween(plateSurface, out memberPlateOffsetDistance, out position1, out position2);
                    double memberPlateBaseDistance = 0.0;
                    topSurface.DistanceBetween((ISurface)basePort, out memberPlateBaseDistance, out position1, out position2);
                    if ((memberPlateBaseDistance < toleranceValue) != (memberPlateOffsetDistance < toleranceValue))
                    {
                        intersectingTopFlange = true;
                    }
                }
                if (intersectingTopFlange)
                {
                    operatorId = 772;
                    penetratedPartPort = boundedPart.GetPort(TopologyGeometryType.Face, operatorId, GeometryOperationTypes.Cutout, ContextTypes.Lateral, GraphPosition.Before, (int)SectionFaceType.Top_Flange_Right_Bottom, false);
                }
                else
                {
                    operatorId = 513;
                    bottomPort = boundedPart.GetPort(TopologyGeometryType.Face, operatorId, GeometryOperationTypes.Cutout, ContextTypes.Lateral, GraphPosition.Before, (int)SectionFaceType.Bottom, false);

                    //Check if Plate is on Bottom Flange
                    ((ISurface)bottomPort).Intersect(penetratingPort, out intersections, out intersectionType);
                    if (intersections != null)
                    {
                        intersectingBottomFlange = true;
                    }
                    //if intersectingBottomFlange is False, then checks for the tolerance value. If the distance between the plate and member is less than tolerance value,
                    //make TopFlangeResultantIntersection as true
                    if (!intersectingBottomFlange)
                    {
                        double memberPlateOffsetDistance = 0.0;
                        ((ISurface)bottomPort).DistanceBetween(plateSurface, out memberPlateOffsetDistance, out position1, out position2);
                        double memberPlateBaseDistance = 0.0;
                        ((ISurface)bottomPort).DistanceBetween((ISurface)basePort, out memberPlateBaseDistance, out position1, out position2);
                        if ((memberPlateBaseDistance < toleranceValue) != (memberPlateOffsetDistance < toleranceValue))
                        {
                            intersectingBottomFlange = true;
                        }
                    }
                    if (intersectingBottomFlange)
                    {
                        operatorId = 771;
                        penetratedPartPort = boundedPart.GetPort(TopologyGeometryType.Face, operatorId, GeometryOperationTypes.PartFinalTrim, ContextTypes.Lateral, GraphPosition.Before, (int)SectionFaceType.Bottom_Flange_Right_Top, false);
                    }
                    else
                    {
                        return thicknessDifference;
                    }
                }
                //Get the distance between Plate and Member Flange
                thicknessDifference = 0.0;
                plateSurface.DistanceBetween((ISurface)penetratedPartPort, out thicknessDifference, out position1, out position2);

                CrossSection penetratedPartCrossSection = boundedPart.CrossSection;
                double flangeThickness = DetailingCustomAssembliesServices.GetFlangeThickness(penetratedPartCrossSection);
                double plateThickness = boundingPart.Thickness;
                if (flangeThickness > plateThickness)
                {
                    thicknessDifference = -thicknessDifference;
                }
            }
            catch (Exception)
            {
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(ChamfersResourceIds.GetThicknessDifferenceError,
                    "Unexpected exception in getting the thickness difference of plate part over member flange."));
            }
            return thicknessDifference;
        }

        #endregion Private methods
    }
}