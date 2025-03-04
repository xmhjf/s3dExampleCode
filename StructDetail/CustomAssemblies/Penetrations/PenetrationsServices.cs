//--------------------------------------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  PenetrationsServices.cs
//
//Abstract
//	PenetrationsServices is a helper class to have commom method implementation for .NET selector rule, parameter rule and definition of the penetrations.
//--------------------------------------------------------------------------------------------------------------------------------------
using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Common.Middle.Services;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using Ingr.SP3D.Planning.Middle;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Helper class to have commom method implementation for .NET selector rule, parameter rule and definition of the penetrations.
    /// </summary>
    internal static class PenetrationsServices
    {
        /// <summary>
        /// Checks if left physical connection needed or not.
        /// L indicates that the slot is connected on the left edge.
        /// LT indicates that the slot is connected on the left top edge.
        /// </summary>
        /// <param name="partName">The part name.</param>
        /// <returns></returns>
        internal static bool IsPCNeededOnSlotLeftEdgeOrLeftTopEdge(string partName)
        {
            //Checking if slot is connected on left edge or left top edge.
            //Better way of doing this is to have bulkloaded properties in catalog content 
            //(though a user or system interface) to return the information about whether a PC is 
            //desired on SlotLeftEdge or SlotTopLeftEdge. If a system interface is provided, this 
            //method can be moved to base class, else it should remain in PenetrationServices.
            if (partName.IndexOf("_L_") != -1 || partName.IndexOf("_LT_") != -1)
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// Checks if top physical connection needed or not.
        /// T indicates that the slot is connected on the Top edge.
        /// LT indicates that the slot is connected on the left top edge.
        /// </summary>
        /// <param name="partName">The part name.</param>
        /// <returns></returns>
        internal static bool IsPCNeededOnSlotTopEdgeOrLeftTopEdge(string partName)
        {

            //Checking if slot is connected on top edge or left top edge.
            //Better way of doing this is to have bulkloaded properties in catalog content 
            //(though a user or system interface) to return the information about whether a PC is 
            //desired on SlotTopEdge or SlotTopLeftEdge. If a system interface is provided, this 
            //method can be moved to base class, else it should remain in PenetrationServices.
            if (partName.IndexOf("_T_") != -1 || partName.IndexOf("_LT_") != -1)
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// Gets the slot edge ID to create bottom left corner feature.
        /// </summary>
        /// <param name="partName">The part name.</param>
        /// <returns></returns>
        internal static SectionFaceType GetSlotEdgeIDForBottomLeftCornerFeature(Feature slot, string partName)
        {
            SectionFaceType slotEdgeIDForBottomLeftCornerFeature = SectionFaceType.Unknown;
            if (IsPCNeededOnSlotLeftEdgeOrLeftTopEdge(partName))
            {
                int baseCorners = ((PropertyValueCodelist)slot.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.BaseCorners)).PropValue;
                if (baseCorners == (int)Answer.Yes)
                {
                    slotEdgeIDForBottomLeftCornerFeature = SectionFaceType.Web_Left;
                }
            }
            return slotEdgeIDForBottomLeftCornerFeature;
        }

        /// <summary>
        /// Gets the slot edge ID to top left corner feature.
        /// LT indicates that the slot is connected on the left top edge.
        /// </summary>
        /// <param name="partName">The part name.</param>
        /// <returns></returns>
        internal static SectionFaceType GetSlotEdgeIDForTopLeftCornerFeature(Feature slot, string sectionTypeName, string partName)
        {
            SectionFaceType slotEdgeIDForTopLeftCornerFeature = SectionFaceType.Unknown;
            if (partName.IndexOf("LT") == -1)
            {
                return slotEdgeIDForTopLeftCornerFeature;
            }

            int outsideCorners = ((PropertyValueCodelist)slot.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.OutsideCorners)).PropValue;

            //OutsideCorners answer is 'No' then simply return
            if (outsideCorners == (int)Answer.No)
            {
                return slotEdgeIDForTopLeftCornerFeature;
            }

            //Get slection choices
            //Note: SectionTypeName should be replaced with CrossSectionTypeAlias once there is proper 
            //mapping between these two in place.
            switch (sectionTypeName)
            {
                case "FB":
                case "EA":
                case "UA":
                case "BUTL3":
                case "B":
                    slotEdgeIDForTopLeftCornerFeature = SectionFaceType.Top;
                    break;
                case "BUT":
                case "BUT2":
                    slotEdgeIDForTopLeftCornerFeature = SectionFaceType.Top_Flange_Left_Bottom;
                    break;
                default:
                    break;
            }
            return slotEdgeIDForTopLeftCornerFeature;
        }

        /// <summary>
        /// Gets the section Id along V direction for the given section type name depending uponthe cross section of the stiffener.
        /// </summary>
        /// <param name="sectionTypeName">The section type name corresponding to the stiffener.</param>
        /// <returns></returns>
        internal static int GetSectionEdgeIdAlongV(string sectionTypeName)
        {
            int edgeId = (int)SectionFaceType.Unknown;

            //Note: SectionTypeName should be replaced with CrossSectionTypeAlias once there is proper 
            //mapping between these two in place.
            switch (sectionTypeName)
            {
                case "FB":
                case "EA":
                case "UA":
                case "BUTL3":
                case "B":
                    edgeId = (int)SectionFaceType.Top;
                    break;
                case "BUT":
                case "BUT2":
                    edgeId = (int)SectionFaceType.Top_Flange_Left_Bottom;
                    break;
                default:
                    break;
            }
            return edgeId;
        }

        /// <summary>
        /// Gets the value of the clearance based on the section type. The default clearance value is 0.99m.
        /// </summary>
        /// <param name="sectionTypeName">The section type name.</param>
        /// <returns></returns>
        internal static double GetClearance(string sectionTypeName)
        {
            double clearance = 0.99;

            //Note: SectionTypeName should be replaced with CrossSectionTypeAlias once there is proper 
            //mapping between these two in place.
            switch (sectionTypeName)
            {
                case "EA":
                case "UA":
                    clearance = 0.01;
                    break;
                case "F":
                case "FB":
                case "SB":
                case "ST":
                    clearance = 0.02;
                    break;
                case "T_XType":
                case "TSType":
                case "BUT":
                case "BUTL2":
                    clearance = 0.03;
                    break;

                default:
                    break;
            }
            return clearance;
        }

        /// <summary>
        /// Sets the slot assembly orientation depending upon the orientation of the assembly parent of the penetrated object.
        /// This method is expected to be overridden by the inheriting class to re-evaluate the slot assembly orientation.
        /// <param name="slot">Slot for which assembly orientation is set.</param>
        /// </summary>
        internal static void SetSlotAssemblyOrientation(Feature slot, BusinessObject penetrated)
        {
            Vector assemblyOrientation = null;
            IAssemblyChild penetratedAssemblyChild = penetrated as IAssemblyChild;
            if (penetratedAssemblyChild != null)
            {
                IAssembly penetratedAssemblyParent = penetratedAssemblyChild.AssemblyParent;
                if (penetratedAssemblyParent != null)
                {
                    ILocalCoordinateSystem localCoordinateSystem = penetratedAssemblyParent as ILocalCoordinateSystem;
                    if (localCoordinateSystem != null)
                    {
                        assemblyOrientation = localCoordinateSystem.ZAxis;
                    }
                }
            }
            if (assemblyOrientation != null)
            {
                slot.SetPropertyValue(assemblyOrientation.X, DetailingCustomAssembliesConstants.IJUASlotAssyOrientation, DetailingCustomAssembliesConstants.AssyOrientationX);
                slot.SetPropertyValue(assemblyOrientation.Y, DetailingCustomAssembliesConstants.IJUASlotAssyOrientation, DetailingCustomAssembliesConstants.AssyOrientationY);
                slot.SetPropertyValue(assemblyOrientation.Z, DetailingCustomAssembliesConstants.IJUASlotAssyOrientation, DetailingCustomAssembliesConstants.AssyOrientationZ);
            }
        }

        /// <summary>
        /// Get slot open angle from planning assembly block to set this value as the slot angle when the assembly method is Drop or DropAtAngle or VerticalDrop.
        /// </summary>
        /// <param name="slot">The slot Feature.</param>        
        internal static double GetSlotAngle(Feature slot)
        {
            double slotAngle = (Math.PI / 1800); //0.1 degree

            BusinessObject penetrated, penetrating;
            slot.GetInputs(out penetrated, out penetrating);

            IAssemblyChild penetratedAssemblyChild = penetrated as IAssemblyChild;
            if (penetratedAssemblyChild != null)
            {
                IAssembly penetratedAssemblyParent = penetratedAssemblyChild.AssemblyParent;
                if (penetratedAssemblyParent != null)
                {
                    PropertyValue assemblyMethodPropertyValue = slot.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.AssyMethod);
                    if (assemblyMethodPropertyValue != null)
                    {
                        int assemblyMethod = ((PropertyValueCodelist)assemblyMethodPropertyValue).PropValue;
                        switch (assemblyMethod)
                        {
                            case (int)AssemblyMethod.Drop:
                            case (int)AssemblyMethod.DropAtAngle:
                            case (int)AssemblyMethod.VerticalDrop:
                                PlanningHelper planningHelper = new PlanningHelper();
                                double slotOpenAngle = planningHelper.SlotOpenAngle(slot);
                                if (slotOpenAngle > (Math.PI / 1800) && slotOpenAngle < (Math.PI / 2)) //0.1 dgree to 90 degree
                                {
                                    slotAngle = slotOpenAngle;
                                }
                                break;
                            default:
                                break;
                        }
                    }
                }
            }
            //We round off the value to 11 decimal places since this how they are doing it in VB content. 
            //If we do not round off the value , we get a different collar part at the slot, since the collar selector answer depends upon the slot angle value.
            return Math.Round(slotAngle, 11);
        }

        #region Conditional Methods

        /// <summary>
        /// Determines whether top flange right bottom physical connection needed or not. This is dependent on cross section.
        /// </summary>
        /// <param name="sectionType">Type of the section.</param>
        /// <returns>
        /// true if top flange right bottom physical connection needed; otherwise, false.
        /// </returns>
        internal static bool IsTopFlangeRightBottomPhysicalConnectionNeeded(string sectionType)
        {
            bool pcNeeded = false;

            //Note: SectionTypeName should be replaced with CrossSectionTypeAlias once there is proper 
            //mapping between these two in place.
            switch (sectionType)
            {
                case MarineSymbolConstants.B:
                case MarineSymbolConstants.EA:
                case MarineSymbolConstants.UA:
                case MarineSymbolConstants.BUT:
                case MarineSymbolConstants.BUTL2:
                    pcNeeded = true;
                    break;
            }
            return pcNeeded;
        }

        /// <summary>
        /// Determines whether the top flange right physical connection needed or not. This is dependent on cross section.
        /// </summary>
        /// <param name="sectionType">Type of the section.</param>
        /// <returns>
        /// true if top flange right physical connection needed; otherwise, false.
        /// </returns>
        internal static bool IsTopFlangeRightPhysicalConnectionNeeded(string sectionType)
        {
            bool pcNeeded = false;

            //Note: SectionTypeName should be replaced with CrossSectionTypeAlias once there is proper 
            //mapping between these two in place.
            switch (sectionType)
            {
                case MarineSymbolConstants.EA:
                case MarineSymbolConstants.UA:
                case MarineSymbolConstants.BUT:
                case MarineSymbolConstants.BUTL2:
                    pcNeeded = true;
                    break;
            }
            return pcNeeded;
        }

        /// <summary>
        /// Determines whether right physical connection between collar and penetrated needed or not.
        /// </summary>
        /// <returns>true if right physical connection between collar and penetrated needed; otherwise, false.</returns>
        internal static bool IsPhysicalConnectionBetweenCollarAndPenetratedNeeded(string partName)
        {
            return partName.IndexOf("_A2") > 0 ? true : false;
        }

        /// <summary>
        /// Determines whether bottom flange right bottom free edge treatment needed or not.
        /// </summary>
        /// <returns>true if bottom flange right bottom free edge treatment needed; otherwise, false.</returns>
        internal static bool IsBottomFlangeFreeEdgeTreatmentNeeded(string partName)
        {
            bool needed = false;
            if (partName.IndexOf("_A") > 0)
            {
                if (partName.IndexOf("_A2") > 0 || partName.IndexOf("_A3") > 0)
                {
                    needed = false;
                }
                else
                {
                    needed = true;
                }
            }
            return needed;
        }

        /// <summary>
        /// Determines whether top flange right top corner physical connection needed or not.
        /// </summary>
        /// <returns>true if top flange right top corner physical connection needed; otherwise, false.</returns>
        internal static bool IsTopFlangeRightTopCornerPhysicalConnectionNeeded(string partName)
        {
            //checking for part name CollarBCT_A3.
            return partName.IndexOf("CollarBCT_A3") == 0 ? true : false;
        }

        /// <summary>
        /// Determines whether top physical connection needed or not.
        /// </summary>
        /// <returns>
        /// true if top physical connection needed; otherwise, false.
        /// </returns>
        internal static bool IsTopPhysicalConnectionNeeded(string partName)
        {
            return partName.IndexOf("_A3") > 0 ? true : false;
        }

        /// <summary>
        /// Determines whether the corner PhysicalConnection needed for the specified section type.
        /// </summary>
        /// <param name="sectionType">Type of the section.</param>
        /// <param name="penetratingPlate">The penetrating plate.</param>
        internal static bool IsCornerPhysicalConnectionNeeded(string sectionType, PlatePartBase penetratingPlate)
        {
            bool pcNeeded = false;            
            //Not needed if Penetrating Object is a plate, because there is no radius on the section alias
            if (penetratingPlate == null)
            {
                //Note: SectionTypeName should be replaced with CrossSectionTypeAlias once there is proper 
                //mapping between these two in place.
                switch (sectionType)
                {
                    case MarineSymbolConstants.B:
                    case MarineSymbolConstants.EA:
                    case MarineSymbolConstants.UA:
                        pcNeeded = true;
                        break;
                }
            }

            return pcNeeded;
        }

        /// <summary>
        /// Determines whether corner snipe on the collar is needed or not.
        /// </summary>
        /// <returns>Returns true if corner snipe needed; otherwise, false.</returns>
        internal static bool IsCornerSnipeNeeded(CollarPart collarPart)
        {
            PropertyValue addCornerSnipePropertyValue = collarPart.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.AddCornerSnipe);
            int addCornerSnipe = (addCornerSnipePropertyValue != null) ? ((PropertyValueCodelist)addCornerSnipePropertyValue).PropValue : 0;
            return addCornerSnipe == (int)Answer.Yes ? true : false;
        }

        /// <summary>
        /// Determines whether physical connection on the plate molded side needed or not.
        /// </summary>
        /// <returns>true if physical connection on the plate molded side needed; otherwise, false.</returns>
        internal static bool IsNormalSideAdditonalLapPhysicalConnectionNeeded(string partName, CollarPart collarPart, PlatePartBase penetratedPlatePart, StiffenerPartBase penetratedStiffenerPart)
        {
            bool pcNeeded = false;
            if (partName.IndexOf("_A2") > 0 || partName.IndexOf("_B2") > 0)
            {
                return pcNeeded;
            }

            //StandAloneStiffenerPart or StiffenerPart or ERPart            
            if (penetratedStiffenerPart != null)
            {
                return pcNeeded;
            }

            PropertyValueDouble sideOfPartPropertyValue = (PropertyValueDouble)collarPart.GetParameterValue(DetailingCustomAssembliesConstants.SideOfPart);
            int sideOfPart = (sideOfPartPropertyValue.PropValue != null) ? (int)sideOfPartPropertyValue.PropValue.Value : (int)SideOfPart.Molded;

            if (sideOfPart == (int)SideOfPart.Molded && IsAdditonalLapPhysicalConnectionNeeded(penetratedPlatePart, collarPart))
            {
                pcNeeded = true;
            }

            return pcNeeded;
        }

        /// <summary>
        /// Determines whether physical connection on the plate anti molded side needed or not.
        /// </summary>
        /// <returns>true if physical connection on the plate anti molded side needed; otherwise, false.</returns>
        internal static bool IsOppositeSideAdditonalLapPhysicalConnectionNeeded(string partName, CollarPart collarPart, PlatePartBase penetratedPlatePart, StiffenerPartBase penetratedStiffenerPart)
        {
            bool pcNeeded = false;
            if (partName.IndexOf("_A2") > 0 || partName.IndexOf("_B2") > 0)
            {
                return pcNeeded;
            }

            //This PC is the second lap PC if valid, that would be
            //between collar and the penetrated part(Cases where the collar overlaps with more than one leaf
            //part of the penetrated part. Assumed a max of 2 lap PCs

            //StandAloneStiffenerPart or StiffenerPart or ERPart
            if (penetratedStiffenerPart != null)
            {
                return pcNeeded;
            }

            PropertyValueDouble sideOfPartPropertyValue = (PropertyValueDouble)collarPart.GetParameterValue(DetailingCustomAssembliesConstants.SideOfPart);
            int sideOfPart = (sideOfPartPropertyValue.PropValue != null) ? (int)sideOfPartPropertyValue.PropValue.Value : (int)SideOfPart.Molded;

            if (sideOfPart == (int)SideOfPart.AntiMolded && IsAdditonalLapPhysicalConnectionNeeded(penetratedPlatePart, collarPart))
            {
                pcNeeded = true;
            }

            return pcNeeded;
        }

        /// <summary>
        /// Determines whether the second lap physical connection is valid or not.
        /// </summary>
        /// <returns>
        ///   true if overlapping penetrated plate parts count is one or more than one; otherwise, false.
        /// </returns>
        private static bool IsAdditonalLapPhysicalConnectionNeeded(PlatePartBase penetratedPlatepart, CollarPart collarPart)
        {
            bool pcNeeded = false;
            // get the penetrated plate part associated with collar and then loop through all its connected leaf plates
            // see if any of the leaf plate other than oAC.connectedObject1 overlaps collar
            // if yes, create PC with it

            if (penetratedPlatepart == null)
            {
                return pcNeeded;
            }

            //get lateral connected parts of the first plate part
            double tolerance = 0.01;
            ReadOnlyCollection<PlatePartBase> overLappingPenetratedPlateParts = collarPart.GetOverLappingPenetratedPlateParts(tolerance);
            if (overLappingPenetratedPlateParts.Count > 0)
            {
                pcNeeded = true;
            }

            return pcNeeded;
        }

        /// <summary>
        /// Determines whether a physical connection should be placed on the plate molded side or not.
        /// </summary>
        /// <param name="assemblyOutputName"></param>
        /// <param name="partName"></param>
        /// <param name="collarPart"></param>
        /// <returns>Returns true if clip normal side needed. Otherwise returns false.</returns>
        internal static bool IsClipSideLapPhysicalConnectionNeeded(string assemblyOutputName, string partName, CollarPart collarPart)
        {
            bool pcNeeded = false;

            if (partName.IndexOf("_A2") > 0 || partName.IndexOf("_B2") > 0)
            {
                return pcNeeded;
            }

            //Get the SideOfPart value based upon which it is determined whether the corresponding assembly output is needed or not.
            PropertyValueDouble sideOfPartPropertyValue = (PropertyValueDouble)collarPart.GetParameterValue(DetailingCustomAssembliesConstants.SideOfPart);
            int sideOfPart = (sideOfPartPropertyValue.PropValue != null) ? (int)sideOfPartPropertyValue.PropValue.Value : (int)SideOfPart.Molded;

            //For Lapped physical connection with the penetrated base port to be created, SideOfPart should be plate moulded
            if (string.Equals(assemblyOutputName, "NormalSideLapPC") && sideOfPart == (int)SideOfPart.Molded)
            {
                pcNeeded = true;             
            }
            else if (string.Equals(assemblyOutputName, "OppositeSideLapPC") && sideOfPart == (int)SideOfPart.AntiMolded)
            {
                //For Lapped physical connection with the penetrated offset port to be created, SideOfPart should be plate anti moulded
                pcNeeded = true;                
            }

            return pcNeeded;
        }

        /// <summary>
        /// Determines whether second base physical connection is valid or not.
        /// </summary>
        /// <param name="collarPart"></param>
        /// <returns>True if overlapping base plate ports count is one or more than one; otherwise, false.</returns>
        internal static bool IsBaseRightAdditonalPhysicalConnectionNeeded(CollarPart collarPart)
        {
            bool pcNeeded = false;

            //get lateral connected parts of the first plate part; these are the leaf parts need to be queried if any of them overlap with collar
            double tolerance = 0.004;
            ReadOnlyCollection<PlatePartBase> overLappingBasePlateParts = collarPart.GetOverLappingBasePlateParts(tolerance);
            if (overLappingBasePlateParts.Count > 0)
            {
                pcNeeded = true;
            }

            return pcNeeded;
        }

        #endregion Conditional Methods
    }
}