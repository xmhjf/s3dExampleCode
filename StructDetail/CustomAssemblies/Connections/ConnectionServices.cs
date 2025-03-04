//--------------------------------------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  ConnectionServices.cs
//
//Abstract
//	ConnectionServices is a helper class to have commom method implementation for .NET selector rule, parameter rule and definition of the PhysicalConnection.
//--------------------------------------------------------------------------------------------------------------------------------------

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.Structure.Middle;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Structure
{
    internal static class ConnectionServices
    {
        #region LeafTeeWeld PhysicalConnection Methods


        /// <summary>
        /// Gets the PhysicalConnection  bounded and bounding part thickness for the physical connection  which are not placed manually. 
        /// </summary>
        /// <param name="physicalConnection">Physical Connection </param>
        /// <param name="boundedPartThickness">Returns the bounded part  thickness value.</param>
        /// <param name="boundingPartThickness">Returns the bounding part  thickness value.</param>
        internal static void GetPhysicalConnectionPartsThickness(PhysicalConnection physicalConnection, out double boundedPartThickness, out double boundingPartThickness)
        {
            bool isParentEndcut = false;
            boundedPartThickness = 0;
            boundingPartThickness = 0;
            ISystem designParent = physicalConnection.SystemParent;
            if (designParent == null)
            {
                return;
            }
            FeatureType featureType = default(FeatureType);
            Feature parentFeature = designParent as Feature;

            if (parentFeature != null)
            {
                featureType = parentFeature.FeatureType;
                if (featureType == FeatureType.WebCut || featureType == FeatureType.FlangeCut)
                {
                    isParentEndcut = true;
                }
            }

            if (isParentEndcut)
            {
                //PhysicalConnection under endcut,check if it is known Multiple PCs case
                switch (physicalConnection.Name)
                {
                    case DetailingCustomAssembliesConstants.WebCut_PC_SeamAngleCase1:
                    case DetailingCustomAssembliesConstants.WebCut_PC_SeamAngleCase2:
                    case DetailingCustomAssembliesConstants.WebCut_PC_SeamAngleCase3:
                    case DetailingCustomAssembliesConstants.WebCut_PC_SeamAngleCase4:
                    case DetailingCustomAssembliesConstants.WebCut_PC_SeamAngleCase5:
                        //Note: Xid value can be found in symbol and end cut rule.                            
                        IPort boundedPort = physicalConnection.BoundedPort;
                        TopologyPort port1 = (TopologyPort)boundedPort;
                        if (port1.SectionId == (int)SectionFaceType.Idealized_Boundary)//At web
                        {
                            featureType = FeatureType.WebCut;
                        }
                        else if (port1.SectionId == 8194)
                        {
                            featureType = FeatureType.FlangeCut;
                        }
                        break;
                }
                boundedPartThickness = GetEndCutPartThickness(physicalConnection.BoundedObject, featureType);
                boundingPartThickness = GetEndCutPartThickness(physicalConnection.BoundingObject, featureType);
            }
            else
            {
                boundedPartThickness = GetBoundedPartThickness(physicalConnection.BoundedObject);
                boundingPartThickness = GetBoundingPartThickness(physicalConnection.BoundingObject);
            }
        }

        /// <summary>
        /// Gets the Physical Connection bounded parts thickness 
        /// </summary>
        /// <param name="boundingObject">Physical connection bounding object</param>   
        /// <returns>boundedPartThickness</returns>
        internal static double GetBoundedPartThickness(BusinessObject boundedObject)
        {
            double boundedPartThickness = 0.0;
            if (boundedObject is PlatePartBase)
            {
                boundedPartThickness = GetPlatePartBaseThickness((PlatePartBase)boundedObject);
            }
            else if (boundedObject is Slab)
            {
                boundedPartThickness = ((Slab)boundedObject).Thickness;
            }
            else if (boundedObject is WallPart)
            {
                boundedPartThickness = ((WallPart)boundedObject).Thickness;
            }
            else if (boundedObject is ProfilePart)
            {
                ProfilePart boundedProfile = (ProfilePart)boundedObject;
                boundedPartThickness = DetailingCustomAssembliesServices.GetWebThickness(boundedProfile.CrossSection);
            }
            return boundedPartThickness;

        }

        /// <summary>
        /// Gets the Physical Connection bounding parts thickness 
        /// </summary>
        /// <param name="boundingObject">Physical connection bounded object</param>
        /// <returns>boundingPartThickness</returns>
        internal static double GetBoundingPartThickness(BusinessObject boundingObject)
        {
            double boundingPartThickness = 0.0;
            if (boundingObject is PlatePartBase)
            {
                boundingPartThickness = GetPlatePartBaseThickness((PlatePartBase)boundingObject);
            }
            else if (boundingObject is Slab)
            {
                boundingPartThickness = ((Slab)boundingObject).Thickness;
            }
            else if (boundingObject is WallPart)
            {
                boundingPartThickness = ((WallPart)boundingObject).Thickness;
            }
            else if (boundingObject is ProfilePart)
            {
                ProfilePart boundingProfile = (ProfilePart)boundingObject;
                boundingPartThickness = DetailingCustomAssembliesServices.GetWebThickness(boundingProfile.CrossSection);
            }
            return boundingPartThickness;

        }

        /// <summary>
        /// Gets selections when Plate bounded to Plate
        /// </summary>
        /// <param name="boundedObject">Bouned Object from PhysicalConnection</param>
        /// <param name="category">Cateogory Selector Answer</param>
        /// <param name="squareTrim">Square Trim Attribute</param>
        /// <returns>Arraylist</returns>
        internal static Collection<string> GetSelectionsForPlateToPlate(BusinessObject boundedObject, string category, double boundedObjectThickness, bool squareTrim)
        {
            Collection<string> choices = new Collection<string>();

            //check to see if the welded part is a collar
            //if the feature is a collar, pick the fillet weld or the single bevel
            CollarPart collarPart = boundedObject as CollarPart;
            if (collarPart != null && collarPart.Slot.FeatureType == FeatureType.Slot)
            {
                if (collarPart.PenetratedObject is PlatePart)
                {
                    //check if the penetrated plate is tight or watertight
                    if (((PlatePart)collarPart.PenetratedObject).Tightness.HasFlag(Tightness.NonTight))
                    {
                        choices.Add(DetailingCustomAssembliesConstants.FilletWeld1);
                    }
                    else //it is tight
                    {
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldV);
                    }
                }
                else
                {
                    choices.Add(DetailingCustomAssembliesConstants.FilletWeld1);
                }
                return choices;
            }
            else if (squareTrim)
            {
                // Physical Connection for Material Handling Squrare Trim cases
                // Note:
                //   For the Bevel data Property Page data to be displayed
                //   The SmartItem must contain "TeeWled" in its name
                //   see \StructDetail\Data\Controls\PhysConnPropPageBevel
                choices.Add(DetailingCustomAssembliesConstants.TeeWeldSquare);
            }
            else  //the welded plate is not a collar, so continue
            {
                switch (category)
                {
                    case DetailingCustomAssembliesConstants.OneSidedBevel:
                        {
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldY);

                            //also add the special cases
                            if (boundedObjectThickness <= Math3d.FitTolerance * 25)
                            {
                                choices.Add(DetailingCustomAssembliesConstants.TeeWeldV);
                            }
                            else if (boundedObjectThickness > Math3d.FitTolerance * 25)
                            {
                                choices.Add(DetailingCustomAssembliesConstants.TeeWeldX);
                            }
                        }
                        break;
                    case DetailingCustomAssembliesConstants.TwoSidedBevel:
                        {
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldK);
                        }
                        break;
                    case DetailingCustomAssembliesConstants.Chain:
                        {
                            choices.Add(DetailingCustomAssembliesConstants.ChainWeld);
                        }
                        break;
                    case DetailingCustomAssembliesConstants.Staggered:
                        {
                            choices.Add(DetailingCustomAssembliesConstants.StaggeredWeld);
                        }
                        break;
                    case DetailingCustomAssembliesConstants.Chill:
                        {
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldChill);
                        }
                        break;
                    default:
                        {
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldV);
                        }
                        break;
                }
            }
            return choices;
        }

        /// <summary>
        /// Gets selections when plate bounded to stiffener
        /// </summary>
        /// <param name="physicalConnection">Physicalconnection object</param>
        /// <param name="squareTrim">SquareTrim Attribute</param>
        /// <returns>Arraylist</returns>
        internal static Collection<string> GetSelectionsForPlateToStiffener(BusinessObject boundedObject, string category, double boundedObjectThickness, bool squareTrim, double mountingAngle)
        {
            Collection<string> choices = new Collection<string>();
            //we can assume that this is a bracket that is bounded to the top of a profile
            //Physical Connection for Material Handling should be "TeeWeldSquare"
            //if the corner Trim is set to open Square Trim cases.  

            if (Math.Abs((mountingAngle - StructHelper.DISTTOL)) >= Math.PI / 2)
            {
                mountingAngle = Math.PI - mountingAngle;
            }
            //this is the theta angle in the requirements
            double mountingAngleCompliment = (Math.PI / 2) - mountingAngle;
            if (squareTrim)
            {
                choices.Add(DetailingCustomAssembliesConstants.TeeWeldSquare);
            }
            else if (boundedObjectThickness > Math3d.FitTolerance * 15)  //check the thickness of the tripping bracket
            {
                if (mountingAngleCompliment > (Math.PI / 6)) //check the mounting angle
                {
                    choices.Add(DetailingCustomAssembliesConstants.TeeWeldY);
                }
                else
                {
                    //give both options to let the user choose
                    choices.Add(DetailingCustomAssembliesConstants.FilletWeld1);
                    choices.Add(DetailingCustomAssembliesConstants.TeeWeldY);
                }
            }
            else
            {
                //default selection is a fillet weld
                choices.Add(DetailingCustomAssembliesConstants.FilletWeld1);
            }
            return choices;
        }

        /// <summary>
        /// Gets selections when Stiffener bounded to Stiffener
        /// </summary>
        /// <param name="physicalConnection">Physicalconnection Object</param>
        /// <returns>Arraylist</returns>
        internal static Collection<string> GetSelectionsForStiffenerToStiffener(BusinessObject boundedObject, IPort boundedPort, IPort boundingPort, double mountingAngle)
        {
            Collection<string> choices = new Collection<string>();
            //could either be bounded to web or to flange; check to find out
            int idealizedBoundaryId = Feature.GetIdealizedBoundaryId(boundedPort, boundingPort);
            StiffenerPartBase profilePart = (StiffenerPartBase)boundedObject;
            if (Math.Abs((mountingAngle - StructHelper.DISTTOL)) >= Math.PI / 2)
            {
                mountingAngle = Math.PI - mountingAngle;
            }
            //this is the theta angle in the requirements
            double mountingAngleCompliment = (Math.PI / 2) - mountingAngle;
            if (idealizedBoundaryId.Equals((int)SectionFaceType.Top) || idealizedBoundaryId.Equals((int)SectionFaceType.Bottom))
            {
                //default selection is a fillet weld
                choices.Add(DetailingCustomAssembliesConstants.FilletWeld1);
                //optional selection is a TeeWeldY
                choices.Add(DetailingCustomAssembliesConstants.TeeWeldY);
            }
            else //this profile end is connected to the web of another profile
            {
                double boundedobjectThickness = DetailingCustomAssembliesServices.GetWebThickness(profilePart.CrossSection);

                if (boundedobjectThickness < Math3d.FitTolerance * 16)
                {
                    if (mountingAngleCompliment < Math.PI / 4) //45deg
                    {
                        choices.Add(DetailingCustomAssembliesConstants.FilletWeld2);    //this is the angled version of fillet weld 1
                    }
                    else if (mountingAngleCompliment > Math.PI / 4 && mountingAngleCompliment <= Math.PI / 3) //45 to 60 degrees
                    {
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldV);
                    }
                    else
                    {
                        choices.Add(DetailingCustomAssembliesConstants.FilletWeld1);
                    }
                }
                else
                {
                    if (mountingAngleCompliment <= Math.PI / 12) //<15Deg
                    {
                        choices.Add(DetailingCustomAssembliesConstants.FilletWeld2); //this is the angled version of fillet weld 1
                    }
                    else if (mountingAngleCompliment > Math.PI / 12 && mountingAngleCompliment <= Math.PI / 3)// 15 to 60 Deg
                    {
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldY);
                    }
                }
            }
            return choices;
        }

        /// <summary>
        /// Gets Selections when Stiffener bounded to Plate
        /// </summary>
        /// <param name="physicalConnection">Physicalconnection Object</param>
        /// <param name="squareTrim">SquareTrim Attribute</param>
        /// <returns>ArrayList</returns>
        internal static Collection<string> GetSelectionsForStiffenerToPlate(BusinessObject boundedObject, BusinessObject boundingObject, string category, double boundedObjectThickness, double mountingAngle, bool squareTrim)
        {
            // If this is a physical connection between a stiffener and the plate it is stiffening, and if the mounting
            //angle is well off normal (beyond OFF_NORMAL_ANGLE_TOLERANCE), then stiffener attachment method
            //must be examined in selecting the proper tee weld.
            Collection<string> choices = new Collection<string>();

            if (Math.Abs((mountingAngle - StructHelper.DISTTOL)) >= Math.PI / 2)
            {
                mountingAngle = Math.PI - mountingAngle;
            }
            //this is the theta angle in the requirements
            double mountingAngleCompliment = (Math.PI / 2) - mountingAngle;
            bool needToExamineStiffenerAttachment = false;

            Plate boundingPlate = (Plate)boundingObject;
            //Get root plate system if bounding object is generated plate part
            PlatePart boundingPlatePart = boundingObject as PlatePart;
            if(boundingPlatePart != null)
            {
                boundingPlate = boundingPlatePart.RootPlateSystem;
            }


            //Get stiffened plate or reinforcement plate from bounded object
            Plate stiffenedPlate = null;
            
            StandAloneStiffenerPart standAloneStiffenerPart = boundedObject as StandAloneStiffenerPart;
            EdgeReinforcementPart edgeReinforcementPart = boundedObject as EdgeReinforcementPart;
            if(standAloneStiffenerPart != null)
            {
                stiffenedPlate = standAloneStiffenerPart.PlateToStiffen;
            }
            else if(edgeReinforcementPart != null)
            {
                stiffenedPlate = edgeReinforcementPart.ReinforcedPlate;
            }
            else
            {
                //System generated stiffener part
                stiffenedPlate = ((StiffenerPartBase)boundedObject).RootStiffenerSystem.PlateToStiffen;
            }

            double mountingAngleComplement = (Math.PI / 2) - mountingAngle;
            if (stiffenedPlate == boundingPlate && mountingAngleComplement > Math.PI / 90)
            {
                needToExamineStiffenerAttachment = true;
            }
            double gapFilletWeld2, gapTeeWeldY, gapTeeWeldX, gapTeeWeldK;
            StiffenerPartBase profilePart = (StiffenerPartBase)boundedObject;
            PlatePartBase platePart = (PlatePartBase)boundingObject;

            double webThickness = DetailingCustomAssembliesServices.GetWebThickness(profilePart.CrossSection);

            if (needToExamineStiffenerAttachment)
            {
                int mountingPoint = ((StiffenerSystem)profilePart.RootStiffenerSystem).MountingPoint;
                // Physical Connection for Material Handling should be "TeeWeldSquare"
                // if the corner Trim is set to open Square Trim cases. 
                if (squareTrim)
                {
                    choices.Add(DetailingCustomAssembliesConstants.TeeWeldSquare);
                }
                else if (mountingPoint.Equals((int)StiffenerMountingPoint.NoTrim))
                {
                    if (category.Equals(DetailingCustomAssembliesConstants.Normal) || category.Equals(DetailingCustomAssembliesConstants.OneSidedBevel))
                    {
                        gapFilletWeld2 = webThickness * Math.Tan(Math.Round(mountingAngleCompliment, 9));
                        gapTeeWeldY = gapFilletWeld2 / 2;

                        // FilletWeld1 allowed in all cases, but must check size of GAP for FilletWeld2 and TeeWeldY
                        if (gapTeeWeldY > DetailingCustomAssembliesConstants.GapTolerance) //Both FilletWeld2 and TeeWeldY GAPs > tolerance, so only FilletWeld1 allowed
                        {
                            choices.Add(DetailingCustomAssembliesConstants.FilletWeld1);
                        }
                        else //TeeWeldY GAP <= tolerance, so TeeWeldY allowed
                        {
                            if (gapFilletWeld2 > DetailingCustomAssembliesConstants.GapTolerance)  //FilletWeld2 not allowed
                            {
                                choices.Add(DetailingCustomAssembliesConstants.FilletWeld1);
                                choices.Add(DetailingCustomAssembliesConstants.TeeWeldY);
                            }
                            else                  //FilletWeld2 GAP <= tolerance, so all allowed
                            {
                                choices.Add(DetailingCustomAssembliesConstants.FilletWeld1);
                                choices.Add(DetailingCustomAssembliesConstants.FilletWeld2);
                                choices.Add(DetailingCustomAssembliesConstants.TeeWeldY);
                            }
                        }
                    }
                    else if (category.Equals(DetailingCustomAssembliesConstants.TwoSidedBevel))
                    {
                        gapTeeWeldX = webThickness / 2 * Math.Tan(mountingAngleCompliment);
                        if (gapTeeWeldX > DetailingCustomAssembliesConstants.GapTolerance)
                        {
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldV);           //Only TeeWeldV allowed
                        }
                        else   //TeeWeldX GAP <= tolerance, so both allowed
                        {
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldV);
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldX);
                        }
                    }
                    else if (category.Equals(DetailingCustomAssembliesConstants.OneSidedBevel))
                    {
                        gapTeeWeldK = (((webThickness - (Math3d.FitTolerance * 3) * Math.Cos(mountingAngleCompliment)) / 2) + (Math3d.FitTolerance * 3) * Math.Cos(mountingAngleCompliment)) * Math.Tan(mountingAngleCompliment);
                        if (gapTeeWeldK > DetailingCustomAssembliesConstants.GapTolerance)
                        {
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldY);          //Only TeeWeldY allowed
                        }
                        else             //TeeWeldK GAP <= tolerance, so both allowed
                        {
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldK);
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldY);
                        }
                        // else //Chain or Zigzag -- could we end up here???
                        //{//'we have a problem???}


                    }
                }
                else if (mountingPoint.Equals((int)StiffenerMountingPoint.PartialTrim))
                {
                    double gapFilletWeld1, gapTeeWeldV;
                    if (category.Equals(DetailingCustomAssembliesConstants.Normal) || category.Equals(DetailingCustomAssembliesConstants.OneSidedBevel))
                    {
                        gapFilletWeld1 = webThickness / 2 * Math.Tan(mountingAngleCompliment);

                        if (gapFilletWeld1 > DetailingCustomAssembliesConstants.GapTolerance)
                        {
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldY);     //Only TeeWeldY allowed
                        }
                        else                       //FilletWeld1 GAP <= tolerance, so all allowed
                        {
                            choices.Add(DetailingCustomAssembliesConstants.FilletWeld1);
                            choices.Add(DetailingCustomAssembliesConstants.FilletWeld2);
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldY);
                        }
                    }
                    else if (category.Equals(DetailingCustomAssembliesConstants.TwoSidedBevel))
                    {
                        gapTeeWeldV = webThickness / 2 * Math.Tan(mountingAngleCompliment);

                        if (gapTeeWeldV > DetailingCustomAssembliesConstants.GapTolerance)
                        {
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldX);   //Only TeeWeldX allowed
                        }
                        else              //TeeWeldV GAP <= tolerance, so both allowed
                        {
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldV);
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldX);
                        }
                    }
                    else if (category.Equals(DetailingCustomAssembliesConstants.OneSidedBevel))
                    {
                        gapTeeWeldY = (webThickness / 2 - (Math3d.FitTolerance * 3) * Math.Cos(mountingAngleCompliment)) * Math.Tan(mountingAngleCompliment);

                        if (gapTeeWeldY > DetailingCustomAssembliesConstants.GapTolerance)
                        {
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldK);           //Only TeeWeldK allowed
                        }
                        else                                  //TeeWeldY GAP <= tolerance, so both allowed
                        {
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldK);
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldY);
                        }
                    }

                    else   //Chain or Zigzag -- invalid here
                    {
                        throw new CmnException("Invalid Category value encountered");
                    }

                }

                else if (mountingPoint.Equals((int)StiffenerMountingPoint.FullTrim))
                {
                    double gapFilletWeld1, gapTeeWeldV;
                    if (category.Equals(DetailingCustomAssembliesConstants.Normal) || category.Equals(DetailingCustomAssembliesConstants.OneSidedBevel))
                    {
                        gapFilletWeld1 = webThickness * Math.Tan(mountingAngleCompliment);
                        gapTeeWeldY = gapFilletWeld1 / 2;

                        if (gapTeeWeldY > DetailingCustomAssembliesConstants.GapTolerance) //FilletWeld2 allowed in all cases, but must check size of GAP for FilletWeld1 and TeeWeldY
                        {
                            choices.Add(DetailingCustomAssembliesConstants.FilletWeld2);
                        }
                        else    //TeeWeldY GAP <= tolerance, so TeeWeldY allowed
                        {
                            if (gapFilletWeld1 > DetailingCustomAssembliesConstants.GapTolerance)  //FilletWeld1 not allowed
                            {
                                choices.Add(DetailingCustomAssembliesConstants.FilletWeld2);
                                choices.Add(DetailingCustomAssembliesConstants.TeeWeldY);
                            }
                            else                     //FilletWeld1 GAP <= tolerance, so all allowed
                            {
                                choices.Add(DetailingCustomAssembliesConstants.FilletWeld1);
                                choices.Add(DetailingCustomAssembliesConstants.FilletWeld2);
                                choices.Add(DetailingCustomAssembliesConstants.TeeWeldY);
                            }
                        }
                    }
                    else if (category.Equals(DetailingCustomAssembliesConstants.TwoSidedBevel))
                    {
                        gapTeeWeldV = webThickness * Math.Tan(mountingAngleCompliment);
                        if (gapTeeWeldV > DetailingCustomAssembliesConstants.GapTolerance)  //Only TeeWeldX allowed
                        {
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldX);
                        }
                        else                                     //TeeWeldV GAP <= tolerance, so both allowed
                        {
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldV);
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldX);
                        }
                    }
                    else if (category.Equals(DetailingCustomAssembliesConstants.OneSidedBevel))
                    {
                        gapTeeWeldY = (webThickness - (Math3d.FitTolerance * 3) * Math.Cos(mountingAngleCompliment)) * Math.Tan(mountingAngleCompliment);

                        if (gapTeeWeldY > DetailingCustomAssembliesConstants.GapTolerance)
                        {
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldK);
                        }
                        else
                        {
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldK);
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldY);
                        }
                    }
                    else   //Chain or Zigzag -- invalid here
                    {
                        throw new CmnException("Invalid Category value encountered");
                    }
                }
                else   //Chain or Zigzag -- invalid here
                {
                    throw new CmnException("Profile has undefined attachment method");
                }
            }

            else   //Don't need to examine stiffener attachment.  Either this physical connection is not between a stiffener and
            // its stiffened plate OR the mounting angle between a stiffener and it's stiffened plate is close to normal
            {
                StiffenerPartBase stiffnerPartBase = (StiffenerPartBase)boundedObject;
                webThickness = DetailingCustomAssembliesServices.GetWebThickness(stiffnerPartBase.CrossSection);
                double thickness = GetBoundedPartThickness(boundedObject);
                double opening = boundedObjectThickness * Math.Abs(Math.Tan(mountingAngleCompliment));
                //Physical Connection for Material Handling should be DetailingCustomAssembliesConstants.FILLET_WELD1
                // if the corner Trim is set to open Square Trim cases. Implemented for TR#210195
                if (squareTrim)
                {
                    choices.Add(DetailingCustomAssembliesConstants.TeeWeldSquare);
                }
                else if (opening < Math3d.FitTolerance * 3)
                {
                    choices.Add(DetailingCustomAssembliesConstants.FilletWeld1);
                }
                else if (webThickness < Math3d.FitTolerance * 16)
                {
                    if (mountingAngleCompliment <= Math.PI / 4)            //45 degrees
                    {
                        choices.Add(DetailingCustomAssembliesConstants.FilletWeld2);              //this is the angled version of fillet weld 1
                    }
                    else if (mountingAngleCompliment > Math.PI / 4 && mountingAngleCompliment <= Math.PI / 3)   //45 to 60 degrees
                    {
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldV);
                    }
                    else
                    {
                        choices.Add(DetailingCustomAssembliesConstants.FilletWeld1);
                    }
                }
                else
                {
                    if (mountingAngleCompliment <= Math.PI / 12)     //< 15 degrees
                    {
                        choices.Add(DetailingCustomAssembliesConstants.FilletWeld2);               //this is the angled version of fillet weld 1
                    }
                    else if (mountingAngleCompliment > Math.PI / 12 && mountingAngleCompliment <= Math.PI / 3)  //15 to 60 degrees
                    {
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldY);
                    }
                }
            }
            return choices;

        }

        /// <summary>
        /// Retreive the "Square Trim" attibute status from the (Root) Logical Connection
        /// </summary>
        /// <param name="physicalConnection">Connection Object</param>
        /// <param name="rootLogicalicalConnection">Logical Connection on which squaretrim attribute to be retrieved</param>
        /// <returns></returns>
        internal static bool HasSquareTrim(LogicalConnection rootLogicalicalConnection)
        {
            bool squareTrim = false;
            if (rootLogicalicalConnection != null)
            {
                if (rootLogicalicalConnection.SupportsInterface(MarineSymbolConstants.IJStructCrossSectionAttrs))
                {
                    PropertyValueCodelist trimValue = ((PropertyValueCodelist)rootLogicalicalConnection.GetPropertyValue(MarineSymbolConstants.IJStructCrossSectionAttrs, DetailingCustomAssembliesConstants.CornerTrim));

                    if (trimValue.PropValue == 2)
                    {
                        squareTrim = true;
                    }
                }
            }
            return squareTrim;
        }

        /// <summary>
        /// Gets the thickness value of Bounding/Bounded part on physicalconnection, checks if Bounding/Bounded is a PlatePart or StiffenerPart or MemberPart
        /// </summary>
        /// <param name="connectable"> Bounding/Bounded part of physicalconnection. </param>
        /// <param name="featureType">Feature type on which physicalconnection is created .</param>
        /// <param name="thickness">Holds the thickness value of  bounding/bounded part.</param>
        internal static double GetEndCutPartThickness(BusinessObject connectable, FeatureType featureType)
        {

            double thickness = 0;
            if (connectable is PlatePartBase)
            {
                PlatePartBase platePartBase = (PlatePartBase)connectable;
                //dynamic platePart = platePartBase;
                //thickness = platePart.Thickness;
                thickness = GetPlatePartBaseThickness((PlatePartBase)connectable);
            }
            else if (connectable is StiffenerPartBase)
            {
                StiffenerPartBase profilePart = (StiffenerPartBase)connectable;
                if (featureType.HasFlag(FeatureType.FlangeCut))
                {
                    thickness = DetailingCustomAssembliesServices.GetFlangeThickness(profilePart.CrossSection);
                }
                else if (featureType.Equals(FeatureType.WebCut))
                {
                    if (profilePart.SectionType.Equals(MarineSymbolConstants.HalfR) || profilePart.SectionType.Equals(MarineSymbolConstants.R))
                    {
                        double width;
                        StiffenerSystemBase stiffenerSystemBase = profilePart.RootStiffenerSystem;
                        stiffenerSystemBase.GetNominalSectionSize(out width, out thickness);
                    }
                    else
                    {
                        thickness = DetailingCustomAssembliesServices.GetWebThickness(profilePart.CrossSection);
                    }
                }
                else
                {
                    throw new CmnException("Unknown endcut type found in GetEndCutPCPartThickness");
                }
            }
            else if (connectable is MemberPart)
            {
                MemberPart memberPart = (MemberPart)connectable;
                if (featureType.HasFlag(FeatureType.FlangeCut))
                {
                    thickness = DetailingCustomAssembliesServices.GetFlangeThickness(memberPart.CrossSection);
                }

                else if (featureType.HasFlag(FeatureType.WebCut))
                {
                    thickness = DetailingCustomAssembliesServices.GetWebThickness(memberPart.CrossSection);
                }
                else
                {
                    throw new Exception("Unknown endcut type found in GetEndCutPCPartThickness");
                }
            }
            else
            {
                throw new Exception("Unknown part type found in GetEndCutPCPartThickness");
            }
            return thickness;
        }

        /// <summary>
        /// Gets the thickness of the specific plate part from provided platepartbase
        /// </summary>
        /// <param name="platePartBase">PlatePartBase</param>
        /// <returns>thickness</returns>
        internal static double GetPlatePartBaseThickness(PlatePartBase platePartBase)
        {
            double thickness = 0.0;
            PlatePart platePart = platePartBase as PlatePart;
            CustomPlatePart customPlatePart = platePartBase as CustomPlatePart;
            StandAlonePlatePart standAlonePlatePart = platePartBase as StandAlonePlatePart;
            CollarPart collarPart = platePartBase as CollarPart;
            if (platePart != null)
            {
                thickness = platePart.Thickness;
            }
            else if (customPlatePart != null)
            {
                thickness = customPlatePart.Thickness;
            }
            else if (standAlonePlatePart != null)
            {
                thickness = standAlonePlatePart.Thickness;
            }
            else if (collarPart != null)
            {
                thickness = collarPart.Thickness;
            }

            return thickness;
        }

        /// <summary>
        /// Gets moldedside of BoundedObject . Checks if BoundedObject is platepart or profilepart. If it is profilepart, based on the featuretype on which PC is placed retrieves the molded side
        /// </summary>
        /// <param name="physicalConnection">Physical Connection</param>
        /// <param name="connectedObject">BoundedObject on physicalconnection</param>
        /// <returns></returns>
        internal static string GetReferenceSide(PhysicalConnection physicalConnection, BusinessObject connectedObject)
        {
            string moldedSide = string.Empty;
            PlatePartBase weldPlate = connectedObject as PlatePartBase;
            StiffenerPartBase stiffenerPart = connectedObject as StiffenerPartBase;
            if (weldPlate != null)
            {
                moldedSide = weldPlate.MoldedSide.ToString();

            }
            else if (stiffenerPart != null)
            {
                BusinessObject systemChild = SymbolHelper.GetCustomAssemblyParent(physicalConnection);
                if (systemChild is Feature)
                {
                    Feature feature = systemChild as Feature;
                    if (feature.FeatureType.HasFlag(FeatureType.WebCut))
                    {
                        ProfilePart profilePart = (ProfilePart)connectedObject;
                        if (profilePart.SectionType.Equals(MarineSymbolConstants.HalfR))
                        {
                            moldedSide = DetailingCustomAssembliesConstants.TopSideMoldedSide;
                        }
                        else if (profilePart.SectionType.Equals(MarineSymbolConstants.R))
                        {
                            moldedSide = DetailingCustomAssembliesConstants.OuterMoldedSide;
                        }
                        else
                        {
                            moldedSide = DetailingCustomAssembliesConstants.WebLeftMoldedSide;
                        }
                    }
                    else if (feature.FeatureType.HasFlag(FeatureType.FlangeCut))
                    {
                        if (feature.IsTopFlangeCut)
                        {
                            moldedSide = DetailingCustomAssembliesConstants.TopFlangeTopFaceMoldedSide;
                        }
                        else
                        {
                            moldedSide = DetailingCustomAssembliesConstants.BottomFlangeBottomFaceMoldedSide;
                        }
                    }
                }
                else
                {
                    moldedSide = DetailingCustomAssembliesConstants.WebLeftMoldedSide;
                }
            }
            return moldedSide;
        }
        
        /// <summary>
        /// Gets moldedside of BoundedObject . Checks if BoundedObject is platepart or profilepart. 
        /// If it is profilepart, based on the featuretype on which PC is placed and its LoadPoint  retrieves the molded side
        /// </summary>
        /// <param name="physicalConnection">Physical Connection</param>
        /// <param name="connectedObject">BoundedObject on physicalconnection</param>
        /// <returns></returns>
        internal static string GetMoldedSide(PhysicalConnection physicalConnection, ISystem connectedObject = null)
        {
            string moldedSide = string.Empty;
            if (connectedObject == null)
            {
                connectedObject = (ISystem)physicalConnection.BoundedObject;
            }
            PlatePartBase weldPlate = connectedObject as PlatePartBase;
            StiffenerPartBase stiffenerPart = connectedObject as StiffenerPartBase;
            if (weldPlate != null)
            {
                moldedSide = weldPlate.MoldedSide.ToString();
            }
            else if (stiffenerPart != null)
            {
                string profilePartSectionType = ((ProfilePart)connectedObject).SectionType;
                if (profilePartSectionType.Equals(MarineSymbolConstants.HalfR) || profilePartSectionType.Equals(MarineSymbolConstants.R))
                {
                    return moldedSide;
                }
                ISystemChild systemChild = physicalConnection;
                int loadPoint = stiffenerPart.LoadPoint;
                if (systemChild.SystemParent is Feature)
                {
                    Feature feature = (Feature)systemChild.SystemParent;
                    if (feature.FeatureType.Equals(FeatureType.WebCut))
                    {//Mountig face considered as bottom always
                        switch (loadPoint)
                        {
                            case (int)LoadPoint.Web_Left_Bottom_Top_Corner:
                            case (int)LoadPoint.Bottom_Flange_Left_Bottom_Corner:
                            case 4:
                            case 5:
                            case 6:
                                moldedSide = DetailingCustomAssembliesConstants.WebLeftMoldedSide;
                                break;
                            case 20:
                            case 21:
                            case 22:
                            case (int)LoadPoint.Bottom_Flange_Right_Bottom_Corner:
                            case (int)LoadPoint.Web_Right_Bottom_Top_Corner:
                                moldedSide = DetailingCustomAssembliesConstants.WebRightMoldedSide;
                                break;
                            case 1:
                            case 25://centered
                                moldedSide = DetailingCustomAssembliesConstants.WebLeftMoldedSide;
                                break;
                        }
                    }
                    else if (feature.FeatureType.Equals(FeatureType.FlangeCut))
                    {
                        if (feature.IsTopFlangeCut)
                        {
                            switch (loadPoint)
                            {
                                case (int)LoadPoint.Web_Left_Bottom_Top_Corner:
                                case (int)LoadPoint.Bottom_Flange_Left_Bottom_Corner:
                                case 4:
                                case 5:
                                case 6:
                                    moldedSide = DetailingCustomAssembliesConstants.TopFlangeTopFaceMoldedSide;
                                    break;
                                case 20:
                                case 21:
                                case 22:
                                case (int)LoadPoint.Bottom_Flange_Right_Bottom_Corner:
                                case (int)LoadPoint.Web_Right_Bottom_Top_Corner:
                                    moldedSide = DetailingCustomAssembliesConstants.TopFlangeBottomFaceMoldedSide;
                                    break;
                                case (int)LoadPoint.Bottom:
                                case (int)LoadPoint.Web_Left:
                                    moldedSide = DetailingCustomAssembliesConstants.TopFlangeTopFaceMoldedSide;
                                    break;
                            }
                        }
                        else
                        {
                            switch (loadPoint)
                            {
                                case (int)LoadPoint.Web_Left_Bottom_Top_Corner:
                                case (int)LoadPoint.Bottom_Flange_Left_Bottom_Corner:
                                case 4:
                                case 5:
                                case 6:
                                    moldedSide = DetailingCustomAssembliesConstants.BottomFlangeBottomFaceMoldedSide;
                                    break;
                                case 20:
                                case 21:
                                case 22:
                                case (int)LoadPoint.Bottom_Flange_Right_Bottom_Corner:
                                case (int)LoadPoint.Web_Right_Bottom_Top_Corner:
                                    moldedSide = DetailingCustomAssembliesConstants.BottomFlangeTopFaceMoldedSide;
                                    break;
                                case (int)LoadPoint.Bottom:
                                case (int)LoadPoint.Web_Left:
                                    moldedSide = DetailingCustomAssembliesConstants.BottomFlangeBottomFaceMoldedSide;
                                    break;
                            }

                        }

                    }
                }
                else
                {
                    //Bounded object is a stiffener but its parent is neither a web-cut
                    //nor a flange-cut
                    switch (loadPoint)
                    {
                        case (int)LoadPoint.Bottom:
                        case (int)LoadPoint.Web_Left_Bottom_Top_Corner:
                        case (int)LoadPoint.Bottom_Flange_Left_Bottom_Corner:
                        case 4:
                        case 5:
                        case 6:
                        case (int)LoadPoint.Web_Left:
                            moldedSide = DetailingCustomAssembliesConstants.WebLeftMoldedSide;
                            break;
                        case 20:
                        case 21:
                        case 22:
                        case (int)LoadPoint.Bottom_Flange_Right_Bottom_Corner:
                        case (int)LoadPoint.Web_Right_Bottom_Top_Corner:
                            moldedSide = DetailingCustomAssembliesConstants.WebRightMoldedSide;
                            break;

                    }
                }
            }
            return moldedSide;
        }

        /// <summary>
        /// This method helps to determine if given AssemblyConnection is within CF range on bounded part
        /// </summary>
        /// <param name="assemblyConnection">Assembly Connection Value</param>
        /// <returns></returns>
        internal static bool IsAssemblyConnectionWithinCornerFeatureRange(AssemblyConnection assemblyConnection)
        {
            bool checkIfACisWithinCFRange = false;
            RangeBox acRange = assemblyConnection.Range;
            Collection<IPort> ports = assemblyConnection.BoundedPorts;
            IConnectable boundedObject = ports[0].Connectable;

            ReadOnlyCollection<Feature> featureList = ((IDetailable)boundedObject).Features;
            foreach (Feature feature in featureList)
            {
                if (feature.FeatureType == FeatureType.Corner)
                {
                    RangeBox range = feature.Range;
                    if (range.Low.X < acRange.Low.X && range.Low.Y < acRange.Low.Y && range.Low.Z < acRange.Low.Z)
                    {
                        if (range.High.X > acRange.High.X && range.High.Y > acRange.High.Y && range.High.Z > acRange.High.Z)
                            checkIfACisWithinCFRange = true;
                    }

                }
            }
            return checkIfACisWithinCFRange;
        }

        /// <summary>
        /// Get available selections from the physicalconnection catalog part.
        /// Uses the filter prog Id to get choices or gets the choices from the comma separated list.
        /// If there are no choices, an empty collection will be returned.
        /// </summary>
        /// <param name="PhysicalConnection">The physical connection whose catalog part provides the filter for choices.</param>
        /// <returns>Choices to be returned</returns>
        internal static Collection<string> GetFilteredSelections(PhysicalConnection physicalConnection)
        {
            ArrayList choices = new ArrayList();
            Collection<string> filteredSelections = new Collection<string>();
            string filter = (string)((PropertyValueString)physicalConnection.GetPropertyValue(MarineSymbolConstants.IJUASelectionFilter, DetailingCustomAssembliesConstants.FilterProgID)).PropValue;

            if (!String.IsNullOrEmpty(filter))
            {
                //filter is a .NET/VB6 ProgId
                //Ex: WeldTable22.Selector
                if (filter.Contains("."))
                {
                    string methodName = "GetAllowedItems";
                    choices.AddRange(MarineSelectorRule.GetSelectorChoicesFromMethod(filter, methodName, physicalConnection));
                }
                else
                {
                    //Filter is specified with ',' separated choices. EX: WebFillet,TeeWeldV
                    choices.AddRange(filter.Split(','));
                }
            }

            if (choices.Count > 0)
            {
                foreach (string choice in choices)
                {
                    filteredSelections.Add(choice);
                }
            }
            return filteredSelections;
        }

        /// <summary>
        ///  Calculates the split angles on physical connection based on the  provided category and returns the split angles.
        /// </summary>
        /// <param name="category">weld category</param>
        /// <param name="boundedObject">physical connection bounded object</param>
        /// <param name="boundingObject">physical connection bounding object</param>
        /// <returns>threshold angles where the physical connection has to be split </returns>
        internal static IEnumerable<double> ComputeSplitAngles(string category, BusinessObject boundedObject, BusinessObject boundingObject)
        {
            List<double> splitAngles = new List<double>();

            // We will only split the physical connection between a bounded stiffener
            // and a bounding plate if the plate is on the system stiffened by the
            // stiffener
            StiffenerPartBase boundedStiffenerPartBase = boundedObject as StiffenerPartBase;
            PlatePartBase boundingPlatePartBase = boundingObject as PlatePartBase;
            PlatePartBase boundedPlatePartBase = boundedObject as PlatePartBase;
            if (boundedStiffenerPartBase != null && boundingPlatePartBase != null)
            {
                //This connection may need to be split
                //check if connected object 2 the stiffened plate of connected object 1
                BusinessObject stiffenedPlate = null;
                BusinessObject parentSystem = null;
                if (boundedStiffenerPartBase is StandAloneStiffenerPart)
                {
                    stiffenedPlate = boundedStiffenerPartBase;
                }
                else
                {
                    stiffenedPlate = boundedStiffenerPartBase.RootStiffenerSystem.PlateToStiffen;
                }
                if (boundingPlatePartBase is StandAlonePlatePart)
                {
                    parentSystem = boundingPlatePartBase;
                }
                else
                {
                    parentSystem = ((PlatePart)boundingPlatePartBase).RootPlateSystem;
                }
                //PlateSystemBase parentSystem = (PlateSystemBase)boundingPlatePartBase.AssemblyParent;
                parentSystem = parentSystem != null ? parentSystem : boundingObject;
                if (stiffenedPlate == parentSystem)
                {
                    double firstAngle;
                    //this is the connection between stiffener and stiffened plate
                    //need to compute split angles base on attachment method and weld category
                    int mountingPoint = (((StiffenerSystem)((StiffenerPartBase)boundedObject).RootStiffenerSystem)).MountingPoint;
                    double webThickness = DetailingCustomAssembliesServices.GetWebThickness(boundedStiffenerPartBase.CrossSection);
                    if (mountingPoint == (int)StiffenerMountingPoint.NoTrim)
                    {
                        if (category.Equals(DetailingCustomAssembliesConstants.Normal))
                        {
                            //we will have two split angle values - EQ1 and EQ2 from analysis
                            splitAngles.Add(Math.Atan(DetailingCustomAssembliesConstants.GapTolerance / webThickness)); // EQ2
                            splitAngles.Add(Math.Atan(2 * (DetailingCustomAssembliesConstants.GapTolerance / webThickness))); //EQ1
                        }
                        else if (category.Equals(DetailingCustomAssembliesConstants.OneSidedBevel))
                        {
                            //we will have one split angle value - EQ1 from analysis
                            splitAngles.Add(Math.Atan(2 * (DetailingCustomAssembliesConstants.GapTolerance / webThickness))); //EQ1
                        }
                        else if (category.Equals(DetailingCustomAssembliesConstants.TwoSidedBevel))
                        {
                            // we will have one split angle value - EQ3 from analysis
                            // iterative solution, but we only solve twice -- assume it will converge
                            // well enough in two passes

                            //for first pass, use NOSE_SIZE
                            firstAngle = Math.Atan(2 * DetailingCustomAssembliesConstants.GapTolerance / (webThickness + DetailingCustomAssembliesConstants.NoseSize));//EQ3a
                            //for second pass, use NOSE_SIZE*cos(firstAngle)
                            splitAngles.Add(Math.Atan(2 * DetailingCustomAssembliesConstants.GapTolerance / (webThickness + DetailingCustomAssembliesConstants.NoseSize * Math.Cos(firstAngle))));//EQ3b
                        }
                    }
                    else if (mountingPoint == (int)StiffenerMountingPoint.PartialTrim)
                    {
                        if (category.Equals(DetailingCustomAssembliesConstants.Normal) || category.Equals(DetailingCustomAssembliesConstants.TwoSidedBevel))
                        {
                            //we will have one split angle value - EQ1 from analysis
                            splitAngles.Add(Math.Atan(2 * (DetailingCustomAssembliesConstants.GapTolerance / webThickness))); //EQ1
                        }
                        else if (category.Equals(DetailingCustomAssembliesConstants.OneSidedBevel))
                        {
                            if (((webThickness / 2) - DetailingCustomAssembliesConstants.NoseSize) > StructHelper.DISTTOL * 0.1)
                            {
                                //we will have one split angle value - EQ4 from analysis
                                // iterative solution, but we only solve twice -- assume it will converge
                                // well enough in two passes
                                //for first pass, use NOSE_SIZE
                                firstAngle = Math.Atan(DetailingCustomAssembliesConstants.GapTolerance / ((webThickness / 2) - DetailingCustomAssembliesConstants.NoseSize)); //EQ4a
                                //for second pass, use NOSE_SIZE*cos(firstAngle)
                                splitAngles.Add(Math.Atan(DetailingCustomAssembliesConstants.GapTolerance / ((webThickness / 2) - DetailingCustomAssembliesConstants.NoseSize * Math.Cos(firstAngle)))); //EQ4b
                            }
                        }
                    }
                    else if (mountingPoint == (int)StiffenerMountingPoint.FullTrim)
                    {
                        if (category.Equals(DetailingCustomAssembliesConstants.Normal))
                        {
                            //we will have two split angle values - EQ1 and EQ2 from analysis
                            splitAngles.Add(Math.Atan(DetailingCustomAssembliesConstants.GapTolerance / webThickness)); //EQ2
                            splitAngles.Add(Math.Atan(2 * (DetailingCustomAssembliesConstants.GapTolerance / webThickness))); //EQ1
                        }
                        else if (category.Equals(DetailingCustomAssembliesConstants.TwoSidedBevel))
                        {
                            //we will have one split angle value - EQ2 from analysis
                            splitAngles.Add(Math.Atan(DetailingCustomAssembliesConstants.GapTolerance / webThickness)); //EQ2
                        }
                        else if (category.Equals(DetailingCustomAssembliesConstants.OneSidedBevel))
                        {
                            if ((webThickness - DetailingCustomAssembliesConstants.NoseSize) > StructHelper.DISTTOL * 0.1)
                            {
                                // we will have one split angle value - EQ5 from analysis
                                // iterative solution, but we only solve twice -- assume it will converge
                                // well enough in two passes
                                //for first pass, use NOSE_SIZE
                                firstAngle = Math.Atan(DetailingCustomAssembliesConstants.GapTolerance / (webThickness - DetailingCustomAssembliesConstants.NoseSize)); //EQ5a
                                //for second pass, use NOSE_SIZE*cos(firstAngle)
                                splitAngles.Add(Math.Atan(DetailingCustomAssembliesConstants.GapTolerance / (webThickness - DetailingCustomAssembliesConstants.NoseSize * Math.Cos(firstAngle)))); //EQ5b
                            }
                        }
                    }
                }
            }
            else if (boundedPlatePartBase != null)
            {
                //The bounded part is a plate, only need to split based on angles if weld category is normal
                if (category.Equals(DetailingCustomAssembliesConstants.Normal) || category.Equals(DetailingCustomAssembliesConstants.OneSidedBevel))
                {
                    //first criteria occurs when the root opening hits 3 mm (.003 M)
                    // this is dependent on the plate thickness
                    double rootGapAngle = Math.Atan(DetailingCustomAssembliesConstants.MaximumRootGap / ((PlatePart)boundedPlatePartBase).Thickness);
                    if (rootGapAngle < Math.PI / 4)
                    {
                        splitAngles.Add(rootGapAngle);
                        splitAngles.Add(Math.PI / 4);
                    }
                    else
                    {
                        //this does not seem likely
                        splitAngles.Add(Math.PI / 4);
                        splitAngles.Add(rootGapAngle);
                    }
                }
            }
            return splitAngles;
        }

        /// <summary>
        /// Upadtes  Reference side of BoundedObject based on "Base" or "Offset".        
        /// </summary>
        /// <param name="referenceSideName">Reference side</param>
        internal static string GetMoldedType(string referenceSideName)
        {
            string referenceSide = referenceSideName;
            if (referenceSideName.Equals(DetailingCustomAssembliesConstants.ReferenceSideBase))
            {
                referenceSide = DetailingCustomAssembliesConstants.ReferenceSideMolded;
            }
            else if (referenceSideName.Equals(DetailingCustomAssembliesConstants.ReferenceSideOffset))
            {
                referenceSide = DetailingCustomAssembliesConstants.ReferenceSideAntimolded;
            }
            return referenceSide;
        }
              
        /// <summary>
        /// Returns the pitch and length values for the  provided society
        /// </summary>
        /// <param name="classSociety">class Society</param>
        /// <param name="pitchValue">Pitch Value</param>
        /// <param name="lengthvalue">Length Value</param>
        internal static void GetPitchAndLengthValues(int classSociety, out double pitchValue, out double lengthvalue)
        {
            pitchValue = 0.0;
            lengthvalue = 0.0;
            switch (classSociety)
            {
                case (int)ClassSociety.Lloyds:
                    {
                        pitchValue = 0.3;
                        lengthvalue = 0.075;
                    }
                    break;
                case (int)ClassSociety.ABS:
                    {
                        pitchValue = 0.31;
                        lengthvalue = 0.08;
                    }
                    break;
                case (int)ClassSociety.DNV:
                    {
                        pitchValue = 0.32;
                        lengthvalue = 0.09;
                    }
                    break;
            }
        }

        /// <summary>
        /// Returns the list of choices when there is physicalconnection between a stiffener and the plate it is stiffening
        /// </summary>
        /// <param name="boundedObject">Physical Connection Bounded Object</param>
        /// <param name="category">selector rule answer for category</param>        
        /// <param name="mountingAngleComplement">PhysicalConnection Connection Angle </param>
        /// <param name="webThickness">webthickness of Stiffener</param>
        /// <returns></returns>
        internal static Collection<string> GetSelectionsForChamferWeld(BusinessObject boundedObject, string category, double mountingAngleComplement, double webThickness)
        {
            Collection<string> choices = new Collection<string>();
            double tanofMountingAngleComplement = Math.Abs(Math.Tan(mountingAngleComplement));
            double cosofMountingAngleComplement = Math.Abs(Math.Cos(mountingAngleComplement));

            StiffenerPartBase profilePart = (StiffenerPartBase)boundedObject;
            int mountingPoint = ((StiffenerSystem)profilePart.RootStiffenerSystem).MountingPoint;
            if (mountingPoint == (int)StiffenerMountingPoint.NoTrim)
            {
                if (category.Equals(DetailingCustomAssembliesConstants.Normal))
                {
                    if (tanofMountingAngleComplement * webThickness / 2 > DetailingCustomAssembliesConstants.GapTolerance)
                    {
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer1);
                    }
                    else
                    {
                        if (tanofMountingAngleComplement * webThickness > DetailingCustomAssembliesConstants.GapTolerance)
                        {
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer1);
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer3);
                        }
                        else
                        {
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer1);
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer2);
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer3);
                        }
                    }
                }
                else if (category.Equals(DetailingCustomAssembliesConstants.Full))
                {
                    if (tanofMountingAngleComplement * webThickness / 2 > DetailingCustomAssembliesConstants.GapTolerance)
                    {
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer4);
                    }
                    else
                    {
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer4);
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer5);
                    }
                }
                else if (category.Equals(DetailingCustomAssembliesConstants.Deep))
                {
                    double tolerance6 = (((webThickness - Math3d.FitTolerance * 3 * cosofMountingAngleComplement) / 2) + Math3d.FitTolerance * 3 * cosofMountingAngleComplement) * tanofMountingAngleComplement;
                    if (tolerance6 > DetailingCustomAssembliesConstants.GapTolerance)
                    {
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer7);
                    }
                    else
                    {
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer6);
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer7);
                    }
                }
            }
            else if (mountingPoint == (int)StiffenerMountingPoint.PartialTrim)
            {
                if (category.Equals(DetailingCustomAssembliesConstants.Normal))
                {
                    if ((webThickness / 2 * tanofMountingAngleComplement) > DetailingCustomAssembliesConstants.GapTolerance)
                    {
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer3);
                    }
                    else
                    {
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer1);
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer2);
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer3);
                    }
                }
                else if (category.Equals(DetailingCustomAssembliesConstants.Full))
                {
                    if (webThickness / 2 * tanofMountingAngleComplement > DetailingCustomAssembliesConstants.GapTolerance)
                    {
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer5);
                    }
                    else
                    {
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer4);
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer5);
                    }
                }
                else if (category.Equals(DetailingCustomAssembliesConstants.Deep))
                {
                    if (((webThickness / 2 - Math3d.FitTolerance * 3 * cosofMountingAngleComplement) * tanofMountingAngleComplement) > DetailingCustomAssembliesConstants.GapTolerance)
                    {
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer6);
                    }
                    else
                    {
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer6);
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer7);
                    }
                }
            }
            else if (mountingPoint == (int)StiffenerMountingPoint.FullTrim)
            {
                if (category.Equals(DetailingCustomAssembliesConstants.Normal))
                {
                    if (webThickness * tanofMountingAngleComplement / 2 > DetailingCustomAssembliesConstants.GapTolerance)
                    {
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer2);
                    }
                    else
                    {
                        if (webThickness * tanofMountingAngleComplement > DetailingCustomAssembliesConstants.GapTolerance)
                        {
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer2);
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer3);
                        }
                        else
                        {
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer1);
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer2);
                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer3);
                        }
                    }
                }
                else if (category.Equals(DetailingCustomAssembliesConstants.Full))
                {
                    if (webThickness * tanofMountingAngleComplement > DetailingCustomAssembliesConstants.GapTolerance)
                    {
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer5);
                    }
                    else
                    {
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer4);
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer5);
                    }
                }
                else if (category.Equals(DetailingCustomAssembliesConstants.Deep))
                {
                    if (((webThickness - Math3d.FitTolerance * 3 * cosofMountingAngleComplement) * tanofMountingAngleComplement) > DetailingCustomAssembliesConstants.GapTolerance)
                    {
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer6);
                    }
                    else
                    {
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer6);
                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer7);
                    }
                }
            }
            return choices;
        }

        #endregion LeafTeeWeld PhysicalConnection Methods
    }
}
