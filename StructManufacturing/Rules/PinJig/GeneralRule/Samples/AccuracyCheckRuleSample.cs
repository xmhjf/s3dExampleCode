
//----------------------------------------------------------------------------------
//      Copyright (C) 2015 Intergraph Corporation.  All rights reserved.
//
//      Purpose: Sample PinJig rule that provides the accuracy check points.
//               This serves as an example for customizing the rule with different implementation.
//
//
//      Author:  Suma Mallena
//
//      History:
//      March 1st, 2015   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using Ingr.SP3D.Manufacturing.Middle;
using System;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// Sample PinJig Accuracy Check Rule.
    /// </summary>
    public class AccuracyCheckRuleSample : PinJigAccuracyCheckRuleBase
    {
        /// <summary>
        /// Returns the accuracy check points of the PinJig.
        /// </summary>
        /// <param name="pinJigInformation">Data class that pin jig rules use for getting the inputs and also information from configuration.</param>
        /// <param name="attributeName">The pin jig acurachy check attribute name for which the jig remarking intersection points have to be returned.</param>
        public override ReadOnlyCollection<JigRemarkingIntersection> GetAccuracyPoints(PinJigInformation pinJigInformation, string attributeName)
        {
            ReadOnlyCollection<JigRemarkingIntersection> accuracyCheckPoints = null;
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

                // Get Remarking Geometries
                
                int attributeValue = 0;

                switch (attributeName)
                {
                    case "BendPoints":

                        accuracyCheckPoints = base.GetBendPoints(mfgPinJig, PinJigRemarkingLineTopologyType.All, PinJigRemarkingDataLocation.JigFloor, 0.09);
                        break;

                    case "AftBoundary":

                        attributeValue = mfgPinJig.RemarkingAccuracyCheckSettings.LeftBoundary;
                        switch (attributeValue)
                        {
                            case 1: //Along Girth From Lower
                                accuracyCheckPoints = base.GetAccuracyCheckPoints(mfgPinJig, PinJigRemarkingLineTopologyType.LeftContour, PinJigRemarkingDataLocation.JigFloor, 0.5, PinJigRemarkingLineTopologyType.BottomContour);
                                break;
                            case 2: //Along Girth From Upper
                                accuracyCheckPoints = base.GetAccuracyCheckPoints(mfgPinJig, PinJigRemarkingLineTopologyType.LeftContour, PinJigRemarkingDataLocation.JigFloor, 0.5, PinJigRemarkingLineTopologyType.TopContour);
                                break;
                            case 3: //At Pin Lines
                                accuracyCheckPoints = base.GetAccuracyCheckPointsAtPinGridLines(mfgPinJig, PinJigRemarkingLineTopologyType.LeftContour, PinJigRemarkingDataLocation.JigFloor, false);
                                break;
                            case 4: //At and Middle of Pin Lines
                                accuracyCheckPoints = base.GetAccuracyCheckPointsAtPinGridLines(mfgPinJig, PinJigRemarkingLineTopologyType.LeftContour, PinJigRemarkingDataLocation.JigFloor, true);
                                break;
                        }

                        break;

                    case "LowerBoundary":
                        attributeValue = mfgPinJig.RemarkingAccuracyCheckSettings.BottomBoundary;
                        switch (attributeValue)
                        {
                            case 1: //Along Girth From Left
                                accuracyCheckPoints = base.GetAccuracyCheckPoints(mfgPinJig, PinJigRemarkingLineTopologyType.BottomContour, PinJigRemarkingDataLocation.JigFloor, 0.5, PinJigRemarkingLineTopologyType.LeftContour);
                                break;
                            case 2: //Along Girth From Right
                                accuracyCheckPoints = base.GetAccuracyCheckPoints(mfgPinJig, PinJigRemarkingLineTopologyType.BottomContour, PinJigRemarkingDataLocation.JigFloor, 0.5, PinJigRemarkingLineTopologyType.RightContour);
                                break;
                            case 3: //At Pin Lines
                                accuracyCheckPoints = base.GetAccuracyCheckPointsAtPinGridLines(mfgPinJig, PinJigRemarkingLineTopologyType.BottomContour, PinJigRemarkingDataLocation.JigFloor, false);
                                break;
                            case 4: //At and Middle of Pin Lines
                                accuracyCheckPoints = base.GetAccuracyCheckPointsAtPinGridLines(mfgPinJig, PinJigRemarkingLineTopologyType.BottomContour, PinJigRemarkingDataLocation.JigFloor, true);
                                break;
                        }

                        break;

                    case "ForeBoundary":
                        attributeValue = mfgPinJig.RemarkingAccuracyCheckSettings.RightBoundary;
                        switch (attributeValue)
                        {
                            case 1: //Along Girth From Lower
                                accuracyCheckPoints = base.GetAccuracyCheckPoints(mfgPinJig, PinJigRemarkingLineTopologyType.RightContour, PinJigRemarkingDataLocation.JigFloor, 0.5, PinJigRemarkingLineTopologyType.BottomContour);
                                break;
                            case 2: //Along Girth From Upper
                                accuracyCheckPoints = base.GetAccuracyCheckPoints(mfgPinJig, PinJigRemarkingLineTopologyType.RightContour, PinJigRemarkingDataLocation.JigFloor, 0.5, PinJigRemarkingLineTopologyType.TopContour);
                                break;
                            case 3: //At Pin Lines
                                accuracyCheckPoints = base.GetAccuracyCheckPointsAtPinGridLines(mfgPinJig, PinJigRemarkingLineTopologyType.RightContour, PinJigRemarkingDataLocation.JigFloor, false);
                                break;
                            case 4: //At and Middle of Pin Lines
                                accuracyCheckPoints = base.GetAccuracyCheckPointsAtPinGridLines(mfgPinJig, PinJigRemarkingLineTopologyType.RightContour, PinJigRemarkingDataLocation.JigFloor, true);
                                break;
                        }

                        break;

                    case "UpperBoundary":
                        attributeValue = mfgPinJig.RemarkingAccuracyCheckSettings.TopBoundary;
                        switch (attributeValue)
                        {
                            case 1: //Along Girth From Left
                                accuracyCheckPoints = base.GetAccuracyCheckPoints(mfgPinJig, PinJigRemarkingLineTopologyType.TopContour, PinJigRemarkingDataLocation.JigFloor, 0.5, PinJigRemarkingLineTopologyType.LeftContour);
                                break;
                            case 2: //Along Girth From Right
                                accuracyCheckPoints = base.GetAccuracyCheckPoints(mfgPinJig, PinJigRemarkingLineTopologyType.TopContour, PinJigRemarkingDataLocation.JigFloor, 0.5, PinJigRemarkingLineTopologyType.RightContour);
                                break;
                            case 3: //At Pin Lines
                                accuracyCheckPoints = base.GetAccuracyCheckPointsAtPinGridLines(mfgPinJig, PinJigRemarkingLineTopologyType.TopContour, PinJigRemarkingDataLocation.JigFloor, false);
                                break;
                            case 4: //At and Middle of Pin Lines
                                accuracyCheckPoints = base.GetAccuracyCheckPointsAtPinGridLines(mfgPinJig, PinJigRemarkingLineTopologyType.TopContour, PinJigRemarkingDataLocation.JigFloor, true);
                                break;
                        }

                        break;
                }

                #endregion Processing

            }
            catch (Exception e)
            {
                LogForToDoList(e, "SMCustomWarningMessages", 5029, "Call to PinJig Accuracy Check Rule failed with the error" + e.Message);
            }

            return accuracyCheckPoints;

        }

    }
}
