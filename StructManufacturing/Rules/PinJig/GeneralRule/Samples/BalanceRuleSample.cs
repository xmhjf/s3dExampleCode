//----------------------------------------------------------------------------------
//      Copyright (C) 2015 Intergraph Corporation.  All rights reserved.
//
//      Purpose: Sample rule to drive the PinJig orientation. It proposes the pinjig balance.
//
//      Author:  Suma Mallena
//
//      History:
//      June 18th, 2015   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Manufacturing.Middle.Services;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Common.Exceptions;
using Ingr.SP3D.Common.Middle.Services.Hidden;
using System.Runtime.InteropServices;
using Ingr.SP3D.Content.Manufacturing.Services;
using Ingr.SP3D.Planning.Middle;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// Sample PinJig Balance Rule.
    /// </summary>
    public class BalanceRuleSample : PinJigBalanceRuleBase
    {
        #region Private Fields

        /// <summary>
        /// Enumerated values for balance methods of a pin jig.
        /// </summary>
        private enum PinJigBalanceType
        {
            MostPlanarNatural = 0,
            TrueNatural = 1,
            AssemblyOrientation = 2,
            UserDefined = 3,
            ParallelAxis = 4,

            AverageOfCornersPlane = 51,
            AverageOfBottomLeftCornersPlane = 52,
            AverageOfTopLeftCornersPlane = 53,
            AverageOfBottomRightCornersPlane = 54,
            AverageOfTopRightCornersPlane = 55,

            BottomLeftCorners = 101,
            TopLeftCorners = 102,
            BottomRightCorners = 103,
            TopRightCorners = 104,

            AverageOfFourPointsPlane = 151,
            ByThreepointsPlane = 152,

            AverageOfReferenceLineAndTwoCornersPlane = 201,
            ReferenceLineAndOneCornerPlane = 202,
        }

        #endregion Private Fields

        /// <summary>
        /// Provides the ability to set the default and allowable values for pin jig balance attributes.
        /// </summary>
        /// <param name="pinJigInformation">Data class that pin jig rules use for getting the inputs and also information from configuration.</param>
        /// <param name="referenceEntities">The reference entities to get the default and allowable values for pin jig balance.</param>
        /// <param name="defaultType">The default value for pin jig balance.</param>
        public override Collection<int> GetAllowableTypes(PinJigInformation pinJigInformation, IEnumerable<BusinessObject> referenceEntities, out int defaultType)
        {
            Collection<int> balanceTypes = new Collection<int>();
            defaultType = -1;
            try
            {
                if (pinJigInformation == null)
                    throw new ArgumentNullException("pinJigInformation.");

                //Add allowable Balance Types
                Dictionary<int, object> allowableList = pinJigInformation.GetArguments("Allowable");       
                foreach (object type in allowableList.Values)
                {
                    balanceTypes.Add(Convert.ToInt32(type));
                }

                //Set the Default Balance Type
                defaultType = Convert.ToInt32(pinJigInformation.GetArguments("Type").FirstOrDefault().Value);         

            }
            catch (Exception e)
            {
                LogForToDoList(e, "SMCustomWarningMessages", 5030, "Call to PinJig Balance Rule to get default/allowable types failed with the error" + e.Message);
            }

            return balanceTypes;
        }

        /// <summary>
        /// Returns the eligible inputs/criteria for user selection to define the pinjig balance.
        /// </summary>
        /// <param name="pinJigInformation">Data class that pin jig rules use for getting the inputs and also information from configuration.</param>
        /// <param name="balancingPlaneType">Specifies the type of balance for the pin jig.</param>
        public override ReadOnlyCollection<object> GetEligibleBalanceInputs(PinJigInformation pinJigInformation, int balancingPlaneType)
        {
            Collection<object> eligibleBalanceInputs = null;

            try
            {
                if (pinJigInformation == null)
                    throw new ArgumentNullException("Input pinJigInformation is empty.");

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //1. Get Inputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region GetInputs

                PinJig mfgPinJig = null;
                if (pinJigInformation.ManufacturingPart != null)
                {
                    mfgPinJig = (PinJig)pinJigInformation.ManufacturingPart;
                }

                if (mfgPinJig == null)
                    return null;

                #endregion GetInputs

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing

                eligibleBalanceInputs = new Collection<object>();

                switch (balancingPlaneType)
                {
                    case (int)PinJigBalanceType.AverageOfCornersPlane:
                    case (int)PinJigBalanceType.BottomLeftCorners:
                    case (int)PinJigBalanceType.TopLeftCorners:
                    case (int)PinJigBalanceType.BottomRightCorners:
                    case (int)PinJigBalanceType.TopRightCorners:
                        {
                            // All Outer Plates Corners
                            ReadOnlyCollection<Position> outerCorners = base.GetOuterCornersOfPlateParts(mfgPinJig.SupportedPlates);
                            foreach (Position corner in outerCorners)
                            {
                                eligibleBalanceInputs.Add(corner);
                            }

                            break;
                        }
                    case (int)PinJigBalanceType.AverageOfBottomLeftCornersPlane:
                    case (int)PinJigBalanceType.AverageOfTopLeftCornersPlane:
                    case (int)PinJigBalanceType.AverageOfBottomRightCornersPlane:
                    case (int)PinJigBalanceType.AverageOfTopRightCornersPlane:
                        {
                            //	Extreme Eligible Corner Points Collection -- Multiple eligible corners for collection of plates representing the input pinjig nature.
                            //	First Corner represents the Pinjig Nature
                            //	Example: First Corner is LowerAft corner when the Pinjig Nature is AverageOfLowerAftCornersPlane.
                            //	Second Corner, Third Corner are the adjacent Points to first Corner. 
                            //	Remaining corners are the Opposite points among the plate corners which are in the second half of the bounding box.

                            ReadOnlyCollection<Position> eligibleExtremeCorners = null;
                            ReadOnlyCollection<Position> extremeCorners = base.GetExtremeCornersOfPlateParts(mfgPinJig.SupportedPlates, balancingPlaneType, out eligibleExtremeCorners);
                            foreach (Position corner in extremeCorners)
                            {
                                eligibleBalanceInputs.Add(corner);
                            }

                            break;
                        }
                    case (int)PinJigBalanceType.AverageOfFourPointsPlane:
                    case (int)PinJigBalanceType.ByThreepointsPlane:
                        {
                            //All Outer Plates Corners
                            ReadOnlyCollection<Position> outerCorners = base.GetOuterCornersOfPlateParts(mfgPinJig.SupportedPlates);
                            foreach (Position corner in outerCorners)
                            {
                                eligibleBalanceInputs.Add(corner);
                            }

                            //End points of Seams, Knuckles, Reference curves, Center Line, User Defined Marking Lines
                            ReadOnlyCollection<Position> entityEndPoints = null;
                            ReadOnlyCollection<BusinessObject> entities = base.GetReferenceEntities(mfgPinJig);
                            foreach (BusinessObject entity in entities)
                            {
                                entityEndPoints = base.GetEndPositionsOfEntityOnSurface(entity, mfgPinJig.RemarkingSurface);
                                foreach (Position endPoint in entityEndPoints)
                                {
                                    eligibleBalanceInputs.Add(endPoint);
                                }
                            }

                            break;
                        }
                    case (int)PinJigBalanceType.AverageOfReferenceLineAndTwoCornersPlane:
                    case (int)PinJigBalanceType.ReferenceLineAndOneCornerPlane:
                        {
                            //All Outer Plates Corners
                            ReadOnlyCollection<Position> outerCorners = base.GetOuterCornersOfPlateParts(mfgPinJig.SupportedPlates);
                            foreach (Position corner in outerCorners)
                            {
                                eligibleBalanceInputs.Add(corner);
                            }

                            //End points of Seams, Knuckles, Reference curves, Center Line, User Defined Marking Lines
                            ReadOnlyCollection<Position> entityEndPoints = null;
                            ReadOnlyCollection<BusinessObject> entities = base.GetReferenceEntities(mfgPinJig);
                            foreach (BusinessObject entity in entities)
                            {
                                entityEndPoints = base.GetEndPositionsOfEntityOnSurface(entity, mfgPinJig.RemarkingSurface);
                                foreach (Position endPoint in entityEndPoints)
                                {
                                    eligibleBalanceInputs.Add(endPoint);
                                }
                            }

                            //Seams, Knuckles, Reference curves, Center Line, User Defined Marking Lines                            
                            foreach (BusinessObject entity in entities)
                            {
                                eligibleBalanceInputs.Add(entity);
                            }

                            break;
                        }
                     default:
                        {
                            return null;
                        }
                }

                #endregion Processing


            }
            catch (Exception e)
            {
                LogForToDoList(e, "SMCustomWarningMessages", 5030, "Call to PinJig Balance Rule to get eligible inputs failed with the error" + e.Message);
            }

            if (eligibleBalanceInputs != null)
            {
                return new ReadOnlyCollection<object>(eligibleBalanceInputs);
            }
            else
            {
                return null;
            }

        }

        /// <summary>
        /// Returns the balancing plane for given balancing type and user inputs.
        /// </summary>
        /// <param name="pinJigInformation">Data class that pin jig rules use for getting the inputs and also information from configuration.</param>  
        /// <param name="balancingPlaneType">Specifies the type of balance for the pin jig.</param>
        /// <param name="balancingPlaneInputs">Specifies the balancing inputs to create the plane.</param>
        /// <param name="supportingSurface">The supporting surface of the pin jig.</param>
        /// <param name="positionsOnPlane">Collection of positions from which the balancing plane is constructed.</param>
        public override IPlane GetBalancingPlane(PinJigInformation pinJigInformation, int balancingPlaneType, ReadOnlyCollection<object> balancingPlaneInputs, ISurface supportingSurface, out ReadOnlyCollection<Position> positionsOnPlane)
        {
            IPlane balancinglane = null;
            positionsOnPlane = null;

            try
            {
                if (pinJigInformation == null)
                    throw new ArgumentNullException("Input pinJigInformation is empty.");

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //1. Get Inputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region GetInputs

                PinJig mfgPinJig = null;
                if (pinJigInformation.ManufacturingPart != null)
                {
                    mfgPinJig = (PinJig)pinJigInformation.ManufacturingPart;
                }

                if (mfgPinJig == null)
                    return null;

                #endregion GetInputs

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing


                ReadOnlyCollection<Position> balancingPlanePositions = null;

                switch (balancingPlaneType)
                {
                    case (int)PinJigBalanceType.MostPlanarNatural:
                        {

                            double planarityRation = 0.0;
                            balancinglane = base.GetBalancingPlane((TopologySurface)supportingSurface, out planarityRation);

                            double percentagePlanarity = planarityRation * 100;
                            if (balancingPlaneInputs.Count() > 0)
                            {
                                if (percentagePlanarity < (double)balancingPlaneInputs[0])
                                {
                                    return null;
                                }
                            }

                            break;
                        }
                    case (int)PinJigBalanceType.TrueNatural:
                        {
                            balancinglane = base.GetBalancingPlane(BalanceCreationType.TrueNatural, (TopologySurface)supportingSurface);
                            break;
                        }
                    case (int)PinJigBalanceType.AssemblyOrientation:
                        {
                            IAssembly parentAssembly = mfgPinJig.SupportedPlates[0].AssemblyParent;
                            balancinglane = base.GetBalancingPlane((AssemblyBase)parentAssembly);
                            break;
                        }
                    case (int)PinJigBalanceType.ParallelAxis:
                        {
                            balancinglane = base.GetBalancingPlane(BalanceCreationType.ParallelToGlobalAxis, (TopologySurface)supportingSurface);
                            break;
                        }
                    case (int)PinJigBalanceType.AverageOfCornersPlane:
                    case (int)PinJigBalanceType.AverageOfBottomLeftCornersPlane:
                    case (int)PinJigBalanceType.AverageOfTopLeftCornersPlane:
                    case (int)PinJigBalanceType.AverageOfBottomRightCornersPlane:
                    case (int)PinJigBalanceType.AverageOfTopRightCornersPlane:
                    case (int)PinJigBalanceType.BottomLeftCorners:
                    case (int)PinJigBalanceType.TopLeftCorners:
                    case (int)PinJigBalanceType.BottomRightCorners:
                    case (int)PinJigBalanceType.TopRightCorners:
                        {
                            BalanceCreationType creationType = BalanceCreationType.AverageOfCorners;
                            if (balancingPlaneType == (int)PinJigBalanceType.AverageOfCornersPlane)
                            {
                                creationType = BalanceCreationType.AverageOfCorners;
                            }
                            else if (balancingPlaneType == (int)PinJigBalanceType.AverageOfBottomLeftCornersPlane)
                            {
                                creationType = BalanceCreationType.AverageOfBottomLeftCorners;
                            }
                            else if (balancingPlaneType == (int)PinJigBalanceType.AverageOfTopLeftCornersPlane)
                            {
                                creationType = BalanceCreationType.AverageOfTopLeftCorners;
                            }
                            else if (balancingPlaneType == (int)PinJigBalanceType.AverageOfBottomRightCornersPlane)
                            {
                                creationType = BalanceCreationType.AverageOfBottomRightCorners;
                            }
                            else if (balancingPlaneType == (int)PinJigBalanceType.AverageOfTopRightCornersPlane)
                            {
                                creationType = BalanceCreationType.AverageOfTopRightCorners;
                            }
                            else if (balancingPlaneType == (int)PinJigBalanceType.BottomLeftCorners)
                            {
                                creationType = BalanceCreationType.BottomLeftCorners;
                            }
                            else if (balancingPlaneType == (int)PinJigBalanceType.TopLeftCorners)
                            {
                                creationType = BalanceCreationType.TopLeftCorners;
                            }
                            else if (balancingPlaneType == (int)PinJigBalanceType.BottomRightCorners)
                            {
                                creationType = BalanceCreationType.BottomRightCorners;
                            }
                            else if (balancingPlaneType == (int)PinJigBalanceType.TopRightCorners)
                            {
                                creationType = BalanceCreationType.TopRightCorners;
                            }


                            if (balancingPlaneInputs.Count() > 0)
                            {
                                // Corners are from user selection.
                                Collection<Position> cornerPositions = new Collection<Position>();
                                foreach (Position corner in balancingPlaneInputs)
                                {
                                    cornerPositions.Add(corner);
                                }

                                balancingPlanePositions = new ReadOnlyCollection<Position>(cornerPositions);
                                balancinglane = base.GetBalancingPlane(creationType, balancingPlanePositions);
                            }
                            else
                            {
                                ReadOnlyCollection<TopologySurface> surfaces = base.GetSurfacesFromPlateParts(mfgPinJig.SupportedPlates, ContextTypes.Base);
                                balancinglane = base.GetBalancingPlane(creationType, surfaces);
                            }

                            break;
                        }
                    case (int)PinJigBalanceType.AverageOfFourPointsPlane:
                    case (int)PinJigBalanceType.ByThreepointsPlane:
                        {
                            BalanceCreationType creationType = BalanceCreationType.AverageOfFourPoints;
                            if (balancingPlaneType == (int)PinJigBalanceType.AverageOfFourPointsPlane)
                            {
                                creationType = BalanceCreationType.AverageOfFourPoints;
                            }
                            else //if (balancingPlaneType == (int)PinJigBalanceType.ByThreepointsPlane)
                            {
                                creationType = BalanceCreationType.ByThreePoints;
                            }

                            Collection<Position> cornerPositions = new Collection<Position>();
                            if (balancingPlaneInputs.Count() > 0)
                            {
                                // Corners are from user selection.                               
                                foreach (Position corner in balancingPlaneInputs)
                                {
                                    cornerPositions.Add(corner);
                                }

                                balancingPlanePositions = new ReadOnlyCollection<Position>(cornerPositions);
                                balancinglane = base.GetBalancingPlane(creationType, balancingPlanePositions);
                            }

                            break;
                        }
                    case (int)PinJigBalanceType.AverageOfReferenceLineAndTwoCornersPlane:
                    case (int)PinJigBalanceType.ReferenceLineAndOneCornerPlane:
                        {
                            Collection<Position> cornerPositions = new Collection<Position>();
                            if (balancingPlaneInputs.Count() > 0)
                            {
                                // Corners are from user selection. 
                                BusinessObject referenceLine = null;
                                foreach (object reference in balancingPlaneInputs)
                                {
                                    if (reference is Position)
                                    {
                                        cornerPositions.Add((Position)reference);
                                    }
                                    else
                                    {
                                        referenceLine = (BusinessObject)reference;
                                    }
                                }

                                balancingPlanePositions = new ReadOnlyCollection<Position>(cornerPositions);
                                balancinglane = base.GetBalancingPlane((TopologySurface)supportingSurface, referenceLine, cornerPositions);
                            }

                            break;
                        }

                }

                #endregion Processing

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //3. Set Outputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Set Outputs

                positionsOnPlane = balancingPlanePositions;               

                #endregion Set Outputs

            }
            catch (Exception e)
            {
                LogForToDoList(e, "SMCustomWarningMessages", 5031, "Call to PinJig Balance Rule to get base plane failed with the error" + e.Message);
            }

            return balancinglane;
        }
    }

}
