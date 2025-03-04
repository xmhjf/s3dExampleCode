//----------------------------------------------------------------------------------
//      Copyright (C) 2015 Intergraph Corporation.  All rights reserved.
//
//      Purpose: Sample rule to drive the PinJig orientation. It proposes the pinjig alignment.
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
using Ingr.SP3D.ReferenceData.Middle;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// Sample PinJig Alignment Rule.
    /// </summary>
    public class AlignmentRuleSample : PinJigAlignmentRuleBase
    {
        #region Private Fields

        /// <summary>
        /// Enumerated values for type of alignment of a pin jig.
        /// </summary>
        private enum PinJigAlignmentType
        {
            Default = 1,
            LongestEdge = 2,
            Global = 3,
            PinsOnCenterLine = 4,
            LeftContour = 5,
            BottomContour = 6,
            RightContour = 7,
            TopContour = 8,
        }

        #endregion Private Fields

        /// <summary>
        /// Provides the ability to set the default and allowable values for pin jig alignment attributes.
        /// </summary>
        /// <param name="pinJigInformation">Data class that pin jig rules use for getting the inputs and also information from configuration.</param>
        /// <param name="referenceEntities">The reference entities to get the default and allowable values for pin jig alignment.</param>
        /// <param name="defaultType">The default value for pin jig alignment.</param>
        public override Collection<int> GetAllowableTypes(PinJigInformation pinJigInformation, IEnumerable<BusinessObject> referenceEntities, out int defaultType)
        {
            Collection<int> alignmentTypes = new Collection<int>();
            defaultType = -1;
            try
            {
                if (pinJigInformation == null)
                    throw new ArgumentNullException("pinJigInformation.");

                //Add allowable Alignment Types
                Dictionary<int, object> allowableList = pinJigInformation.GetArguments("Allowable");
                foreach (object type in allowableList.Values)
                {
                    alignmentTypes.Add(Convert.ToInt32(type));
                }

                //TODO:
                //if (DoesPlatesHaveCenterLineIntersection(referenceEntities) = true)
                //{
                //    alignmentTypes.Add((int)PinJigAlignmentType.PinsOnCenterLine);
                //}

                ////Set the Default Alignment Type
                defaultType = Convert.ToInt32(pinJigInformation.GetArguments("Type").FirstOrDefault().Value);      

            }
            catch (Exception e)
            {
                LogForToDoList(e, "SMCustomWarningMessages", 5026, "Call to PinJig Alignment Rule to get default/allowable types failed with the error" + e.Message);
            }

            return alignmentTypes;
        }

        /// <summary>
        /// Returns various values required for certain purpose of computing PinJig orientation.
        /// 'The tolerances would be used in the following order:
        /// Offset would be applied on the PinBed.
        /// Later, MinOverHang check would be done
        /// Later additonal row/col would be added
        /// 
        /// Properties "TransverseOffsetForOriginPin" and "LongitudinalOffsetForOriginPin" are best explained by below figure
        ///   NEGATIVE OFFSET         POSITIVE OFFSET
        ///   x  | x    x             | x    x
        ///      |                    |
        ///   x  | x    x             | x    x
        ///      +-------             +-------
        ///   x    x    x
        ///  We can apply an offset on the Pin Bed so that we avoid pins along the edge
        ///  or a pin exactly at the Pin Jig's seam origin.
        /// 
        ///  NEGATIVE Offset would move the pin origin outwards from the contours
        ///  POSITIVE Offset would move the pin origin inwards from the contours
        /// 
        ///  These properties allow you to control these offsets along the transverse and longitudinal directions respectively.
        /// 
        ///  These offset values are used ONLY when Pin Jig is created WITHOUT ANY ROTATION/MOVEMENT APPLIED ON JIG FLOOR.
        /// 
        ///  The values returned are FRACTIONS of their respective intervals i.e., between 0.0 and 1.0.
        /// 
        ///  To offset the pin bed by an absolute amount, use "ValForRounding" argument to calculate the ratio.
        ///  (i.e., ValForRounding will contain the corresponding Pin Interval)
        /// 
        ///  Properties "TransverseMinimumOverhang" and "LongitudinalMinimumOverhang" are best explained by below figure
        /// 
        ///        A  B  C  D        Property "TransverseMinimumOverhang" specifies the minimum overhang that 'O' should have
        ///                          beyond column 'C'. (i.e) Minimum overhang that would be respected in the transversal direction.
        ///       \
        ///        k  x  x  x        Property "LongitudinalMinimumOverhang" is similar to "TransverseMinimumOverhang"
        ///         \                except that it specifies the longitudinal minimum overhang.
        ///        x \x  x  x
        ///           \              The values returned are FRACTIONS of their respective intervals i.e., between 0.0 and 1.0.
        ///        x  x\ x  x
        ///             \
        ///        x  x  O  x        The input argument "ValForRounding" will contain the pin interval corresponding to "ToleranceContext".
        ///             /            If ToleranceContext is "TransverseMinimumOverhang", ValForRounding will be Transverse Interval.
        ///        x  x/ x  x        If ToleranceContext is "LongitudinalMinimumOverhang", ValForRounding will be Longitudinal Interval.
        ///           /
        ///        x /x  x  x        E.g., if you desire to have a row ONLY if the overhang exceeds half the
        ///         /                longitudinal interval then return 0.5 for "LongitudinalMinimumOverhang".
        ///        k  x  x  x
        ///       /                  E.g., if you desire to have a column ONLY if the overhang exceeds
        ///                          100 mm (assume DBU is meters), then return (0.1 / ValForRounding)
        ///        A  B  C  D        for "TransverseMinimumOverhang". (i.e) a minimum of 100mm overhang would be maintained in the transversal direction
        /// 
        /// Properties "AddAdditionalRow" and "AddAdditionalColumn" indicate whether to add an additional row, column.
        /// To always have an additional row or column (e.g., like a buffer zone), return ONE.
        /// To never have an additional row or column, return ZERO.
        /// </summary>
        /// <param name="pinJigInformation">Data class that pin jig rules use for getting the inputs and also information from configuration.</param>
        /// <param name="alignmentType">Specifies the alignment type of pin jig.</param>        
        public override PinJigAlignmentInputs GetAlignmentInputs(PinJigInformation pinJigInformation, int alignmentType)
        {
            PinJigAlignmentInputs pinJigAlignmentInputs = new PinJigAlignmentInputs();
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
                    return pinJigAlignmentInputs;

                #endregion GetInputs

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing

                // Get the transversal and longitudinal pin intervals.
                double horizontalPinInterval = mfgPinJig.Report.HorizontalPinInterval;
                double verticalPinInterval = mfgPinJig.Report.VerticalPinInterval;

                string pinAlignmentType = CatalogService.GetCodeListStringValue("StrMfgPinJigAlignmentTypes", alignmentType,"STRMFG",false); //RuleQuery.GetRule(400, 413, alignmentType).Name;

                #endregion Processing

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //3. Set Outputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Set Outputs
           
                //Note:  
                // PinJigOrientationAdjustmentType.BoxPointsCorners is NOT APPLICABLE for "Default" alignment option.
                // PinJigOrientationAdjustmentType.ExtensionFill will have no meaning when it is based on CONTOURS (PinJigOrientationAdjustmentType.BoxPointsOfContours).
                switch (alignmentType)
                {
                    case (int)PinJigAlignmentType.Default:
                    case (int)PinJigAlignmentType.LeftContour:
                    case (int)PinJigAlignmentType.BottomContour:
                    case (int)PinJigAlignmentType.RightContour:
                    case (int)PinJigAlignmentType.TopContour:
                        {
                            pinJigAlignmentInputs.TransverseOffsetForOriginPin = Convert.ToDouble(pinJigInformation.GetArguments("TransverseOffsetForOriginPin", pinAlignmentType).FirstOrDefault().Value) / horizontalPinInterval;
                            pinJigAlignmentInputs.LongitudinalOffsetForOriginPin = Convert.ToDouble(pinJigInformation.GetArguments("LongitudinalOffsetForOriginPin", pinAlignmentType).FirstOrDefault().Value) / verticalPinInterval;
                            pinJigAlignmentInputs.TransverseMinimumOverhang = Convert.ToDouble(pinJigInformation.GetArguments("TransverseMinimumOverhang", pinAlignmentType).FirstOrDefault().Value) / horizontalPinInterval;
                            pinJigAlignmentInputs.LongitudinalMinimumOverhang = Convert.ToDouble(pinJigInformation.GetArguments("LongitudinalMinimumOverhang", pinAlignmentType).FirstOrDefault().Value) / verticalPinInterval;

                            pinJigAlignmentInputs.AddAdditionalRow = Convert.ToBoolean(pinJigInformation.GetArguments("AddAdditionalRow", pinAlignmentType).FirstOrDefault().Value);
                            pinJigAlignmentInputs.AddAdditionalColumn = Convert.ToBoolean(pinJigInformation.GetArguments("AddAdditionalColumn", pinAlignmentType).FirstOrDefault().Value);

                            pinJigAlignmentInputs.PinOriginPositionType = (PinJigPinLocationType)Convert.ToInt32(pinJigInformation.GetArguments("PinOriginPosition", pinAlignmentType).FirstOrDefault().Value);
                            pinJigAlignmentInputs.ViewUpVectorType = (PinJigViewDirectionType)Convert.ToInt32(pinJigInformation.GetArguments("ViewUpVector", pinAlignmentType).FirstOrDefault().Value);
                            pinJigAlignmentInputs.LocalCoordinateSystemOriginType = (PinJigPinLocationType)Convert.ToInt32(pinJigInformation.GetArguments("LocalCoordinateSystemOrigin", pinAlignmentType).FirstOrDefault().Value);

                            pinJigAlignmentInputs.AdjustmentType = GetAdjustmentType(pinJigInformation, pinAlignmentType);

                            break;
                        }
                    case (int)PinJigAlignmentType.LongestEdge:
                    case (int)PinJigAlignmentType.Global:
                    case (int)PinJigAlignmentType.PinsOnCenterLine:
                        {
                            pinJigAlignmentInputs.TransverseOffsetForOriginPin = 0.1 / horizontalPinInterval;
                            pinJigAlignmentInputs.LongitudinalOffsetForOriginPin = 0.1 / verticalPinInterval;
                            pinJigAlignmentInputs.TransverseMinimumOverhang = 0.1 / horizontalPinInterval;
                            pinJigAlignmentInputs.LongitudinalMinimumOverhang = 0.1 / verticalPinInterval;

                            pinJigAlignmentInputs.AddAdditionalRow = Convert.ToBoolean(pinJigInformation.GetArguments("AddAdditionalRow", pinAlignmentType).FirstOrDefault().Value);
                            pinJigAlignmentInputs.AddAdditionalColumn = Convert.ToBoolean(pinJigInformation.GetArguments("AddAdditionalColumn", pinAlignmentType).FirstOrDefault().Value);

                            pinJigAlignmentInputs.PinOriginPositionType = (PinJigPinLocationType)Convert.ToInt32(pinJigInformation.GetArguments("PinOriginPosition", pinAlignmentType).FirstOrDefault().Value);
                            pinJigAlignmentInputs.ViewUpVectorType = (PinJigViewDirectionType)Convert.ToInt32(pinJigInformation.GetArguments("ViewUpVector", pinAlignmentType).FirstOrDefault().Value);
                            pinJigAlignmentInputs.LocalCoordinateSystemOriginType = (PinJigPinLocationType)Convert.ToInt32(pinJigInformation.GetArguments("LocalCoordinateSystemOrigin", pinAlignmentType).FirstOrDefault().Value);
                            pinJigAlignmentInputs.AdjustmentType = GetAdjustmentType(pinJigInformation, pinAlignmentType);

                            break;
                        }
                    case -1:
                        {
                            // This is the case where move/rotate has been applied on the PinJig or in modification of the PinJig
                            // Minimum Overhang values can be specified after move/rotate. Offset values are not applicable
                            // Overriding some of the values of Configuration XML.

                            pinJigAlignmentInputs.TransverseMinimumOverhang = 0.1 / horizontalPinInterval;
                            pinJigAlignmentInputs.LongitudinalMinimumOverhang = 0.1 / verticalPinInterval;

                            break;
                        }
                }

                #endregion Set Outputs

            }
            catch (Exception e)
            {
                LogForToDoList(e, "SMCustomWarningMessages", 5026, "Call to PinJig Alignment Rule to get values for purpose failed with the error" + e.Message);
            }

            return pinJigAlignmentInputs;
        }

        /// <summary>
        /// Returns the adjusted pin bed positions in the following order: origin position, X position, other position, Y position.
        /// It would adjust the pin bed corners based on the different tolerances such as offset, minimim overhang, adding additional rows or columns that are provided by the rule.
        /// The corners would also be ordered to get the origin position at the desired corner.
        /// </summary>
        /// <param name="pinJigInformation">Data class that pin jig rules use for getting the inputs and also information from configuration.</param>
        /// <param name="alignmentType">Specifies the alignment type of pin jig.</param>
        /// <param name="pinJigAlignmentInputs">Data class that holds the pin jig alignment inputs.</param>
        public override ReadOnlyCollection<Position> GetCornerPinPositions(PinJigInformation pinJigInformation, int alignmentType, PinJigAlignmentInputs pinJigAlignmentInputs)
        {
            ReadOnlyCollection<Position> cornerPinPositions = null;
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

                cornerPinPositions = base.GetAdjustedCornerPinPositions(mfgPinJig, alignmentType, pinJigAlignmentInputs);

                #endregion Processing

            }
            catch (Exception e)
            {
                LogForToDoList(e, "SMCustomWarningMessages", 5027, "Call to PinJig Alignment Rule to get Corner Pin Positions failed with the error" + e.Message);
            }

            return cornerPinPositions;
        }

        #region Private Methods

        private PinJigOrientationAdjustmentType GetAdjustmentType(PinJigInformation pinJigInformation, string pinAlignmentType)
        {
            Dictionary<int, object> allowableList = pinJigInformation.GetArguments("OrientationAdjustmentType", pinAlignmentType);
            ArrayList adjustmentTypes = new ArrayList(allowableList.Values);            
       
            int adjustmentType = 0;
            for (int idx = 0; idx < adjustmentTypes.Count; idx++)
            {
                if (idx == 0)
                {
                    adjustmentType = Convert.ToInt32(adjustmentTypes[idx]);                   
                }
                else
                {
                    adjustmentType = (adjustmentType) | (Convert.ToInt32(adjustmentTypes[idx]));
                }
            }

            return (PinJigOrientationAdjustmentType)adjustmentType;
        }

        #endregion Private Methods
    }
}
