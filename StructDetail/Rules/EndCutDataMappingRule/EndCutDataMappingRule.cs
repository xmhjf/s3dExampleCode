//-------------------------------------------------------------------------------------------------------------------------------------
//      Copyright (C) 2011-16 Intergraph Corporation.  All rights reserved.
//
//      Author:  
//
//      History:
//      January 27, 2016        mkonduri                CR-CP-273577  Updated the VB parameter rules and added a new project "EndCutDataMappingRule.cs"

//      February 3, 2016        PYK                     DI-CP-287121  Fix coverity defects stated in January 15, 2016 report
//-------------------------------------------------------------------------------------------------------------------------------------

using System;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Interop.ProfileEntity;
using refProfilePart = Ingr.SP3D.Structure.Middle;
using refPortType = Ingr.SP3D.Common.Middle;



namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Returns end cut data on profile part
    /// </summary>
    [WrapperProgID("MarineRuleWrappers.EndCutDataMappingRuleWrapper")]
    public class EndCutDataMappingRule : EndCutDataMappingRuleBase, ICustomEndCutDataMappingRule
    {
        #region  ICustomEndCutDataMappingRule Members

        /// <summary>
        /// Gets the end cut data in strings at both start and end connections of profile part.
        /// </summary>
        /// <param name="profilePart">The input profile part.</param>
        /// <returns>Edge cut Data mapping information.</returns>
        /// <exception cref="ArgumentNullException">The passed profile part is null.</exception>
        public void GetEndCutDataStrings(refProfilePart.ProfilePart profilePart)
        {
            if (profilePart == null)
            {
                throw new ArgumentNullException("profilePart");
            }
            UpdateDrawingFeatureType(profilePart);

        }
        public void AddFeatureEndCutData(Feature featureObject, long endCutRelativePosition, String endCutData)
        {
            int firstDelimiterPos = 0;
            int secDelimiterPos = 0;

            string endCutDataString;
            string[] tempString = null;
            String flangeString;

            IPort boundedPortFlange;
            BusinessObject boundedPortWeb;
            BusinessObject boundingPort;
            Feature webCut;
            IPort boundedPortType = null;
            IJSDEndCutData sdEndCutData = null;

            if (featureObject.FeatureType == FeatureType.FlangeCut)
            {
                featureObject.GetInputs(out boundingPort, out boundedPortFlange, out webCut);
                boundedPortType = boundedPortFlange;
                sdEndCutData = (IJSDEndCutData)Ingr.SP3D.Common.Middle.Services.Hidden.COMConverters.ConvertBOToCOMBO((BusinessObject)boundedPortFlange.Connectable);
            }

            else if (featureObject.FeatureType == FeatureType.WebCut)
            {
                featureObject.GetInputs(out boundedPortWeb, out boundingPort);

                boundedPortType = boundedPortWeb as IPort;
                if (boundedPortType != null)
                {
                    sdEndCutData = (IJSDEndCutData)Ingr.SP3D.Common.Middle.Services.Hidden.COMConverters.ConvertBOToCOMBO((BusinessObject)boundedPortType.Connectable);
                }
                else
                {

                    throw new Exception("Unable to get the boundedPortType as input to the feature object.");
                }
            }
            String startEndCutData = null;
            String endEndCutData = null;
            if (sdEndCutData != null)
            {
                startEndCutData = sdEndCutData.StartEndCutData;
                endEndCutData = sdEndCutData.EndEndCutData;
            }
            else
            {
                throw new Exception("startEndCutData and endEndCutData should not be null.");
            }

            if (startEndCutData == null)
            {
                startEndCutData = " ; ; | ; ; | ; ; | ; ; | ";
            }
            if (endEndCutData == null)
            {
                endEndCutData = " ; ; | ; ; | ; ; | ; ; | ";
            }

            if (boundedPortType.PortType == (refPortType.PortType)ContextTypes.Base)
            {
                tempString = startEndCutData.Split('|');
            }
            else if (boundedPortType.PortType == (refPortType.PortType)ContextTypes.Offset)
            {
                tempString = endEndCutData.Split('|');
            }
            else
                return;
            if (featureObject.FeatureType == FeatureType.WebCut)
            {
                firstDelimiterPos = tempString[0].IndexOf(';', 1);
                secDelimiterPos = tempString[0].LastIndexOf(';', -1);

                switch (endCutRelativePosition)
                {
                    case (long)EndCutRelativePosition.Primary:

                        endCutDataString = tempString[0].Substring(firstDelimiterPos - 1);
                        endCutDataString = endCutData;
                        break;

                    case (long)EndCutRelativePosition.TopOrLeft:

                        endCutDataString = tempString[0].Substring(secDelimiterPos - 1);
                        endCutDataString = endCutData;
                        break;

                    case (long)EndCutRelativePosition.BottomOrRight:

                        endCutDataString = tempString[0].Substring(secDelimiterPos + 1);
                        endCutDataString = endCutData;
                        break;
                }
            }
            else if (featureObject.FeatureType == FeatureType.FlangeCut)
            {

                if (featureObject.IsTopFlangeCut)
                {
                    flangeString = tempString[1];
                }
                else
                {
                    flangeString = tempString[2];
                }
                switch (endCutRelativePosition)
                {
                    case (long)EndCutRelativePosition.Primary:

                        endCutDataString = tempString[0].Substring(firstDelimiterPos - 1);
                        endCutDataString = endCutData;
                        break;

                    case (long)EndCutRelativePosition.TopOrLeft:

                        endCutDataString = tempString[0].Substring(secDelimiterPos - 1);
                        endCutDataString = endCutData;
                        break;

                    case (long)EndCutRelativePosition.BottomOrRight:

                        endCutDataString = tempString[0].Substring(secDelimiterPos + 1);
                        endCutDataString = endCutData;
                        break;
                }
                if (featureObject.IsTopFlangeCut)
                {
                    tempString[1] = flangeString;
                }
                else
                {
                    tempString[2] = flangeString;
                }
            }
            startEndCutData = string.Join("|", tempString);
            endEndCutData = string.Join("|", tempString);
        }
        internal enum EndCutRelativePosition
        {
            Primary = 0,
            TopOrLeft = 1,
            BottomOrRight = 2,
        }
        #endregion

        void ICustomEndCutDataMappingRule.UpdateDrawingFeatureType(refProfilePart.ProfilePart profilePart)
        {
            if (profilePart == null)
            {
                throw new ArgumentNullException("profilePart");
            }
            UpdateDrawingFeatureType(profilePart);


        }
    }
}
