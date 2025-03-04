//**************************************************************************************
//  Copyright (C) 2015, Intergraph Corporation.  All rights reserved.
//
//  File        : ManufacturingProfileKnuckleReport.cs
//
//  Description : 
//
//
//  Author      : Nautilus - HSV
//
//  History     : Created 11/09/2015
//
//
//**************************************************************************************

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections.ObjectModel;
using System.IO;

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Manufacturing.Services;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Manufacturing.Middle.Services;
using Ingr.SP3D.Planning.Middle;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// 
    /// </summary>
    public class ManufacturingMemberCustomReport : ManufacturingProfileCustomReportRuleBase
    {

        #region private members

        CustomReportLogService logService = new CustomReportLogService();
        ProfileReportService profileReport = null;
        int roundTo = 1;

        #endregion

        #region override methods

        /// <summary>
        /// Generates the specified entities.
        /// </summary>
        /// <param name="entities">The entities.</param>
        /// <param name="filePath">The file path.</param>
        /// <exception cref="System.NotImplementedException"></exception>
        public override void Generate(IEnumerable<BusinessObject> entities, string filePath)
        {
            logService.OpenFile(filePath);

            try
            {

                if (entities == null || entities.Count() == 0) return;
                foreach (BusinessObject businessObject in entities)
                {
                    if (businessObject is ManufacturingProfile)
                    {
                        try { ReportMemberPartInformation((ManufacturingProfile)businessObject); }
                        catch { continue; }
                    }
                    else if (businessObject is ProfilePart)
                    {
                        Collection<ManufacturingBase> mfgProfiles =
                            EntityService.GetManufacturingEntity(businessObject, ManufacturingEntityType.ManufacturingProfile);
                        foreach (ManufacturingBase profile in mfgProfiles)
                        {
                            try { ReportMemberPartInformation((ManufacturingProfile)profile); }
                            catch { continue; }
                        }
                    }
                    else if (businessObject is AssemblyBase)
                    {
                        foreach (ManufacturingProfile profile in GetManufacturedProfilesFromAssembly((AssemblyBase)businessObject))
                        {
                            try { ReportMemberPartInformation(profile); }
                            catch { continue; }
                        }
                    }
                }
            }
            catch { /*DO NOTHING*/ }
            finally
            {
                logService.CloseFile();
                if (profileReport != null) profileReport.Dispose();
            }
        }

        #endregion

        #region private methods

        private void ReportMemberPartInformation(ManufacturingProfile profile)
        {
            profileReport = GetProfileReportingService(profile);

            GenerateParentObjectInformation(profile);

            GenerateProcessingLengthInformation(profile);

            GenerateConnectedObjectDistances(profile);

            GenerateFeatureDistances(profile);

            GenerateMarkingLineDistances(profile);

            GenerateSupportValues(profile);
        }

        private void GenerateParentObjectInformation(ManufacturingProfile profile)
        {
            IAssemblyChild child;
            IAssembly parent;
            INamedItem namedItem;

            child = (IAssemblyChild)profile;
            parent = child.AssemblyParent;

            namedItem = (INamedItem) parent;
            logService.WriteToFile(" PROFILE_NAME  : " + namedItem.Name);

            child = (IAssemblyChild)parent;
            parent = child.AssemblyParent;

            if (parent == null) return;

            if (!(parent is Block))
            {
                namedItem = (INamedItem)parent;
                logService.WriteToFile(" ASSEMBLY_NAME  : " + namedItem.Name);

                child = parent as IAssemblyChild;
                if (child != null) parent = child.AssemblyParent;
                else parent = null;
            }

            if (parent == null) return;

            while (parent != null && !(parent is Block))
            {
                child = parent as IAssemblyChild;
                if (child != null) parent = child.AssemblyParent;
                else parent = null;
            }

            if (parent == null) return;

            namedItem = (INamedItem)parent;
            logService.WriteToFile(" BLOCK_NAME  : " + namedItem.Name);

        }

        private void GenerateProcessingLengthInformation(ManufacturingProfile profile)
        {

            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("Test: ManufacturingLength");
            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("");

            #region length variables

            double landingCurveLength;
            double landingCurveStart;
            double detailProfileLength;
            double detailWebLength;
            double detailTopLength;
            double detailBottomLength;
            double beforeFeaturesLength;
            double beforeFeaturesWebLength;
            double beforeFeaturesTopLength;
            double beforeFeaturesBottomLength;
            double afterFeaturesLength;
            double afterFeaturesWebLength;
            double afterFeaturesTopLength;
            double afterFeaturesBottomLength;
            double unfoldedProfileLength;
            double unfoldedWebLength;
            double unfoldedTopLength;
            double unfoldedBottomLength;

            #region get values

            landingCurveLength = profile.ManufacturingLength(ManufacturingProfileLengthType.LandingCurve, SectionFaceType.Unknown);
            landingCurveStart = profile.ManufacturingLength(ManufacturingProfileLengthType.LandingCurveStart, SectionFaceType.Unknown);
            detailProfileLength = profile.ManufacturingLength(ManufacturingProfileLengthType.StructDetailing, SectionFaceType.Unknown);
            detailWebLength = profile.ManufacturingLength(ManufacturingProfileLengthType.StructDetailing, SectionFaceType.Web_Left);
            detailTopLength = profile.ManufacturingLength(ManufacturingProfileLengthType.StructDetailing, SectionFaceType.Top);
            detailBottomLength = profile.ManufacturingLength(ManufacturingProfileLengthType.StructDetailing, SectionFaceType.Bottom);
            beforeFeaturesLength = profile.ManufacturingLength(ManufacturingProfileLengthType.BeforeFeatures, SectionFaceType.Unknown);
            beforeFeaturesWebLength = profile.ManufacturingLength(ManufacturingProfileLengthType.BeforeFeatures, SectionFaceType.Web_Left);
            beforeFeaturesTopLength = profile.ManufacturingLength(ManufacturingProfileLengthType.BeforeFeatures, SectionFaceType.Top);
            beforeFeaturesBottomLength = profile.ManufacturingLength(ManufacturingProfileLengthType.BeforeFeatures, SectionFaceType.Bottom);
            afterFeaturesLength = profile.ManufacturingLength(ManufacturingProfileLengthType.AfterFeatures, SectionFaceType.Unknown);
            afterFeaturesWebLength = profile.ManufacturingLength(ManufacturingProfileLengthType.AfterFeatures, SectionFaceType.Web_Left);
            afterFeaturesTopLength = profile.ManufacturingLength(ManufacturingProfileLengthType.AfterFeatures, SectionFaceType.Top);
            afterFeaturesBottomLength = profile.ManufacturingLength(ManufacturingProfileLengthType.AfterFeatures, SectionFaceType.Bottom);
            unfoldedProfileLength = profile.ManufacturingLength(ManufacturingProfileLengthType.Unfolded, SectionFaceType.Unknown);
            unfoldedWebLength = profile.ManufacturingLength(ManufacturingProfileLengthType.Unfolded, SectionFaceType.Web_Left);
            unfoldedTopLength = profile.ManufacturingLength(ManufacturingProfileLengthType.Unfolded, SectionFaceType.Top);
            unfoldedBottomLength = profile.ManufacturingLength(ManufacturingProfileLengthType.Unfolded, SectionFaceType.Bottom);

            #endregion

            #region convert to milli meters

            landingCurveLength = ConvertToMilliMeters(landingCurveLength);
            landingCurveStart = ConvertToMilliMeters(landingCurveStart);
            detailProfileLength = ConvertToMilliMeters(detailProfileLength);
            detailWebLength = ConvertToMilliMeters(detailWebLength);
            detailTopLength = ConvertToMilliMeters(detailTopLength);
            detailBottomLength = ConvertToMilliMeters(detailBottomLength);
            beforeFeaturesLength = ConvertToMilliMeters(beforeFeaturesLength);
            beforeFeaturesWebLength = ConvertToMilliMeters(beforeFeaturesWebLength);
            beforeFeaturesTopLength = ConvertToMilliMeters(beforeFeaturesTopLength);
            beforeFeaturesBottomLength = ConvertToMilliMeters(beforeFeaturesBottomLength);
            afterFeaturesLength = ConvertToMilliMeters(afterFeaturesLength);
            afterFeaturesWebLength = ConvertToMilliMeters(afterFeaturesWebLength);
            afterFeaturesTopLength = ConvertToMilliMeters(afterFeaturesTopLength);
            afterFeaturesBottomLength = ConvertToMilliMeters(afterFeaturesBottomLength);

            #endregion

            #endregion

            logService.WriteToFile("Landing Curve Length:" + "\t" + "= " + Math.Round(landingCurveLength, roundTo));
            logService.WriteToFile("");
            logService.WriteToFile("Landing Curve Start" + "\t" + "= " + Math.Round(landingCurveStart, roundTo));
            logService.WriteToFile("");

            logService.WriteToFile("Structural Detailing Length:");
            logService.WriteToFile("\t" + "Profile Length" + "\t\t" + "= " + Math.Round(detailProfileLength, roundTo));
            logService.WriteToFile("\t" + "Web Length" + "\t\t" + "= " + Math.Round(detailWebLength, roundTo));
            logService.WriteToFile("\t" + "Top Flange Length" + "\t" + "= " + Math.Round(detailTopLength, roundTo));
            logService.WriteToFile("\t" + "Bottom Flange Length" + "\t" + "= " + Math.Round(detailBottomLength, roundTo));

            logService.WriteToFile("Unfolded Length:");
            logService.WriteToFile("\t" + "Profile Length" + "\t\t" + "= " + Math.Round(unfoldedProfileLength, roundTo));
            logService.WriteToFile("\t" + "Web Length" + "\t\t" + "= " + Math.Round(unfoldedWebLength, roundTo));
            logService.WriteToFile("\t" + "Top Flange Length" + "\t" + "= " + Math.Round(unfoldedTopLength, roundTo));
            logService.WriteToFile("\t" + "Bottom Flange Length" + "\t" + "= " + Math.Round(unfoldedBottomLength, roundTo));

            logService.WriteToFile("Before Features Length:");
            logService.WriteToFile("\t" + "Profile Length" + "\t\t" + "= " + Math.Round(beforeFeaturesLength, roundTo));
            logService.WriteToFile("\t" + "Web Length" + "\t\t" + "= " + Math.Round(beforeFeaturesWebLength, roundTo));
            logService.WriteToFile("\t" + "Top Flange Length" + "\t" + "= " + Math.Round(beforeFeaturesTopLength, roundTo));
            logService.WriteToFile("\t" + "Bottom Flange Length" + "\t" + "= " + Math.Round(beforeFeaturesBottomLength, roundTo));

            logService.WriteToFile("After Features Length:");
            logService.WriteToFile("\t" + "Profile Length" + "\t\t" + "= " + Math.Round(afterFeaturesLength, roundTo));
            logService.WriteToFile("\t" + "Web Length" + "\t\t" + "= " + Math.Round(afterFeaturesWebLength, roundTo));
            logService.WriteToFile("\t" + "Top Flange Length" + "\t" + "= " + Math.Round(afterFeaturesTopLength, roundTo));
            logService.WriteToFile("\t" + "Bottom Flange Length" + "\t" + "= " + Math.Round(afterFeaturesBottomLength, roundTo));

            logService.WriteToFile("");

        }

        private void GenerateConnectedObjectDistances(ManufacturingProfile profile)
        {

            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("Test: Connected Object Distances");
            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("");

            IAssemblyChild mfgChild = profile as IAssemblyChild;
            IAssembly parent = mfgChild.AssemblyParent as IAssembly;
            BusinessObject parentBO = null;
            if (parent == null)
            {
                logService.WriteToFile("No Connected Objects.");
                return;
            }
            else if (parent is StiffenerPartBase) parentBO = parent as StiffenerPartBase;
            else if (parent is MemberPart) parentBO = parent as MemberPart;
            else return;

            IConnectable connectablePart = parent as IConnectable;
            ReadOnlyCollection<BusinessObject> connectedObjects = connectablePart.GetConnectedObjects();

            logService.WriteToFile("Number of connections to other objects: " + connectedObjects.Count);
            if (connectedObjects.Count > 0)
            {
                logService.WriteToFile("Index " + "\t" + "Name" + "\t\t\t\t" + "Distance");
            }

            foreach (BusinessObject connectedObject in connectedObjects)
            {
                string connectedObjectName = string.Empty;
                INamedItem namedItem = connectedObject as INamedItem;
                if (namedItem != null) connectedObjectName = namedItem.Name;
                else connectedObjectName = "No Name";
                double distance = profileReport.GetDistanceBetweenConnectedObjectAndStartOfProfileUnfold(connectedObject);
                if (connectedObjectName.Length <= 15)
                {
                    logService.WriteToFile(connectedObjects.IndexOf(connectedObject) + "\t" + connectedObjectName + "\t\t\t" + Math.Round(distance, roundTo));
                }
                else if (connectedObjectName.Length <= 23)
                {
                    logService.WriteToFile(connectedObjects.IndexOf(connectedObject) + "\t" + connectedObjectName + "\t\t" + Math.Round(distance, roundTo));
                }
                else
                {
                    logService.WriteToFile(connectedObjects.IndexOf(connectedObject) + "\t" + connectedObjectName + "\t" + Math.Round(distance, roundTo));
                }
            }

            logService.WriteToFile("");
        }

        private void GenerateFeatureDistances(ManufacturingProfile profile)
        {
            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("Test: Feature Distances");
            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("");

            #region feature variables

            double landingCurveLength;
            double landingCurveStart;
            ReadOnlyCollection<Feature> featuresCollection = null;

            if (profile.DetailedPart is MemberPart)
            {
                MemberPart member = (MemberPart)profile.DetailedPart;
                featuresCollection = member.Features;
            }
            else
            {
                //create empty collection
                featuresCollection = new ReadOnlyCollection<Feature>(new List<Feature>());
            }

            landingCurveLength = profile.ManufacturingLength(ManufacturingProfileLengthType.LandingCurve, SectionFaceType.Unknown);
            landingCurveStart = profile.ManufacturingLength(ManufacturingProfileLengthType.LandingCurveStart, SectionFaceType.Unknown);

            landingCurveLength = ConvertToMilliMeters(landingCurveLength);
            landingCurveStart = ConvertToMilliMeters(landingCurveStart);

            #endregion

            logService.WriteToFile("Feature Distance(FD) Feature Distance and Height (FD2, FH)");
            logService.WriteToFile("Number of FeatureObject:" + featuresCollection.Count);
            logService.WriteToFile("Landing Curve Start (LCS): " + Math.Round(landingCurveStart, roundTo));
            logService.WriteToFile("Landing Curve Length (LCL): " + Math.Round(landingCurveLength, roundTo));

            if (featuresCollection.Count > 0)
            {
                logService.WriteToFile("Index" + "\t" + "Type" + "\t\t\t" + "FD" + "\t\t" + "(FD-LCS)" + "\t" + "(LCL-FD)" + "\t" +
                    "(LCL-(FD-LCS))" + "\t" + "End Cut At Base");
            }

            foreach (Feature feature in featuresCollection)
            {
                string featureType = string.Empty;
                switch (feature.FeatureType)
                {
                    case FeatureType.Edge:
                        featureType = "Edge Feature" + "\t";
                        break;
                    case FeatureType.Corner:
                        featureType = "Corner Feature" + "t";
                        break;
                    case FeatureType.FlangeCut:
                        featureType = "Flange Cut" + "\t";
                        break;
                    case FeatureType.Slot:
                        featureType = "Slot" + "\t\t";
                        break;
                    case FeatureType.WaterStop:
                        featureType = "Water Stop" + "\t";
                        break;
                    case FeatureType.WebCut:
                        featureType = "Web Cut" + "\t\t";
                        break;
                    default:
                        featureType = "Unidentified" + "\t";
                        break;
                }

                bool featureBoolean = profileReport.CheckIfEndCutIsAtProfileBase(feature);
                double featureDistance = profileReport.GetDistanceBetweenFeatureAndStartOfProfileUnfold(feature);
                featureDistance = ConvertToMilliMeters(featureDistance);
                int index = featuresCollection.IndexOf(feature);
                logService.WriteToFile(index + "\t" + featureType + "\t" + Math.Round(featureDistance, roundTo)
                    + "\t\t" + Math.Round((featureDistance - landingCurveStart), roundTo)
                    + "\t\t" + Math.Round((landingCurveLength - featureDistance), roundTo)
                    + "\t\t" + Math.Round((landingCurveLength - (featureDistance - landingCurveStart)), roundTo)
                    + "\t\t" + featureBoolean);
            }
            logService.WriteToFile("");

        }

        private void GenerateMarkingLineDistances(ManufacturingProfile profile)
        {
            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("Test: Marking Line Distances");
            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("");

            List<MarkingObjects> topMarkings = GetMarkings(profile, SectionFaceType.Top);
            List<MarkingObjects> bottomMarkings = GetMarkings(profile, SectionFaceType.Bottom);
            List<MarkingObjects> webMarkings = GetMarkings(profile, SectionFaceType.Web_Left);

            if (topMarkings.Count <= 0
                && webMarkings.Count <= 0
                && bottomMarkings.Count <= 0)
            {
                logService.WriteToFile("There are no marking lines to report on.");
                logService.WriteToFile("");
                return;
            }

            if (topMarkings.Count > 0)
            {
                logService.WriteToFile("Number of Top Flange Marking Objects: " + topMarkings.Count);
                logService.WriteToFile("Index" + "\t" + "Marking Distance" + "\t\t" + "Marking Type");
                foreach (MarkingObjects obj in topMarkings)
                {
                    double markingDist = profileReport.GetDistancebetweenMarkingLineAndStartOfProfileUnfold(obj.Marking);
                    markingDist = ConvertToMilliMeters(markingDist);
                    string name = obj.Name;
                    logService.WriteToFile(topMarkings.IndexOf(obj) + "\t" + Math.Round(markingDist, roundTo) + "\t\t" + name);
                }
            }

            if (bottomMarkings.Count > 0)
            {
                logService.WriteToFile("Number of Bottom Flange Marking Objects: " + bottomMarkings.Count);
                logService.WriteToFile("Index" + "\t" + "Marking Distance" + "\t\t" + "Marking Type");
                foreach (MarkingObjects obj in bottomMarkings)
                {
                    double markingDist = profileReport.GetDistancebetweenMarkingLineAndStartOfProfileUnfold(obj.Marking);
                    markingDist = ConvertToMilliMeters(markingDist);
                    string name = obj.Name;
                    logService.WriteToFile(bottomMarkings.IndexOf(obj) + "\t" + Math.Round(markingDist, roundTo) + "\t\t" + name);
                }
            }

            if (webMarkings.Count > 0)
            {
                logService.WriteToFile("Number of Web Marking Objects: " + webMarkings.Count);
                logService.WriteToFile("Index" + "\t" + "Marking Distance" + "\t\t" + "Marking Type");
                foreach (MarkingObjects obj in webMarkings)
                {
                    double markingDist = profileReport.GetDistancebetweenMarkingLineAndStartOfProfileUnfold(obj.Marking);
                    markingDist = ConvertToMilliMeters(markingDist);
                    string name = obj.Name;
                    logService.WriteToFile(webMarkings.IndexOf(obj) + "\t" + Math.Round(markingDist, roundTo) + "\t\t" + name);
                }
            }

            logService.WriteToFile("");

        }

        private void GenerateSupportValues(ManufacturingProfile profile)
        {
            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("Test: Support Values");
            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("");

            #region support value variables

            double maxRange = -99;
            double minRange = -99;
            double rangeIncremntor = -1;
            double supportValue = -1;
            double distanceValue = -1;
            string geometryName = string.Empty;
            string message = string.Empty;
            ManufacturingGeometryType geometryType = ManufacturingGeometryType.GeneralMark;//initialize temporarily

            #endregion

            for (int i = 0; i <= 1; i++)
            {
                #region set geometry type

                switch (i)
                {
                    case 0:
                        geometryType = ManufacturingGeometryType.ProfileLength;
                        break;
                    case 1:
                        geometryType = ManufacturingGeometryType.LandingCurve;
                        break;
                }

                #endregion

                geometryName = GetGeometryName(geometryType);

                #region Web Left

                profileReport.GetSupportGeometryRange(geometryType, SectionFaceType.Web_Left, out minRange, out maxRange);
                minRange = ConvertToMilliMeters(minRange);
                maxRange = ConvertToMilliMeters(maxRange);
                logService.WriteToFile(geometryName + "\t" + "Web Left" + "\t" + "Min: " + Math.Round(minRange, roundTo) + "\t" + "Max: " + Math.Round(maxRange, roundTo));
                if (maxRange > minRange)
                {
                    logService.WriteToFile("Support Values:");
                    rangeIncremntor = (maxRange - minRange) / 10;
                    message = "";
                    distanceValue = 0;
                    for (int j = 0; j <= 10; j++)
                    {
                        message += Math.Round(distanceValue, roundTo) + "\t";
                        distanceValue += rangeIncremntor;
                    }
                    logService.WriteToFile(message);
                    message = "";
                    distanceValue = 0;
                    for (int j = 0; j <= 10; j++)
                    {
                        supportValue = profileReport.GetSupportGeometryValue(geometryType, SectionFaceType.Web_Left, distanceValue);
                        distanceValue += rangeIncremntor / 1000;
                        supportValue = ConvertToMilliMeters(supportValue);
                        message += Math.Round(supportValue, roundTo) + "\t";
                    }
                    logService.WriteToFile(message);
                    logService.WriteToFile("");
                }

                #endregion

                #region Web Right

                maxRange = -99;
                minRange = -99;
                rangeIncremntor = -1;
                supportValue = -1;
                distanceValue = -1;
                message = "";

                profileReport.GetSupportGeometryRange(geometryType, SectionFaceType.Web_Right, out minRange, out maxRange);
                minRange = ConvertToMilliMeters(minRange);
                maxRange = ConvertToMilliMeters(maxRange);
                logService.WriteToFile(geometryName + "\t" + "Web Right" + "\t" + "Min: " + Math.Round(minRange, roundTo) + "\t" + "Max: " + Math.Round(maxRange, roundTo));
                if (maxRange > minRange)
                {
                    logService.WriteToFile("Support Values:");
                    rangeIncremntor = (maxRange - minRange) / 10;
                    message = "";
                    distanceValue = 0;
                    for (int j = 0; j <= 10; j++)
                    {
                        message += Math.Round(distanceValue, roundTo) + "\t";
                        distanceValue += rangeIncremntor;
                    }
                    logService.WriteToFile(message);
                    message = "";
                    distanceValue = 0;
                    for (int j = 0; j <= 10; j++)
                    {
                        supportValue = profileReport.GetSupportGeometryValue(geometryType, SectionFaceType.Web_Right, distanceValue);
                        distanceValue += rangeIncremntor / 1000;
                        supportValue = ConvertToMilliMeters(supportValue);
                        message += Math.Round(supportValue, roundTo) + "\t";
                    }
                    logService.WriteToFile(message);
                    logService.WriteToFile("");
                }

                #endregion

                #region top flange

                maxRange = -99;
                minRange = -99;
                rangeIncremntor = -1;
                supportValue = -1;
                distanceValue = -1;
                message = "";

                profileReport.GetSupportGeometryRange(geometryType, SectionFaceType.Top, out minRange, out maxRange);
                minRange = ConvertToMilliMeters(minRange);
                maxRange = ConvertToMilliMeters(maxRange);
                logService.WriteToFile(geometryName + "\t" + "Top Flange" + "\t" + "Min: " + Math.Round(minRange, roundTo) + "\t" + "Max: " + Math.Round(maxRange, roundTo));
                if (maxRange > minRange)
                {
                    logService.WriteToFile("Support Values:");
                    rangeIncremntor = (maxRange - minRange) / 10;
                    message = "";
                    distanceValue = 0;
                    for (int j = 0; j <= 10; j++)
                    {
                        message += Math.Round(distanceValue, roundTo) + "\t";
                        distanceValue += rangeIncremntor;
                    }
                    logService.WriteToFile(message);
                    message = "";
                    distanceValue = 0;
                    for (int j = 0; j <= 10; j++)
                    {
                        supportValue = profileReport.GetSupportGeometryValue(geometryType, SectionFaceType.Top, distanceValue);
                        distanceValue += rangeIncremntor / 1000;                       
                        supportValue = ConvertToMilliMeters(supportValue);
                        message += Math.Round(supportValue, roundTo) + "\t";
                    }
                    logService.WriteToFile(message);
                    logService.WriteToFile("");
                }

                #endregion

                #region bottom flange

                maxRange = -99;
                minRange = -99;
                rangeIncremntor = -1;
                supportValue = -1;
                distanceValue = -1;
                message = "";

                profileReport.GetSupportGeometryRange(geometryType, SectionFaceType.Bottom, out minRange, out maxRange);
                minRange = ConvertToMilliMeters(minRange);
                maxRange = ConvertToMilliMeters(maxRange);
                logService.WriteToFile(geometryName + "\t" + "Bottom Flange" + "\t" + "Min: " + Math.Round(minRange, roundTo) + "\t" + "Max: " + Math.Round(maxRange, roundTo));
                if (maxRange > minRange)
                {
                    logService.WriteToFile("Support Values:");
                    rangeIncremntor = (maxRange - minRange) / 10;
                    message = "";
                    distanceValue = 0;
                    for (int j = 0; j <= 10; j++)
                    {
                        message += Math.Round(distanceValue, roundTo) + "\t";
                        distanceValue += rangeIncremntor;
                    }
                    logService.WriteToFile(message);
                    message = "";
                    distanceValue = 0;
                    for (int j = 0; j <= 10; j++)
                    {
                        supportValue = profileReport.GetSupportGeometryValue(geometryType, SectionFaceType.Bottom, distanceValue);
                        distanceValue += rangeIncremntor / 1000;
                        supportValue = ConvertToMilliMeters(supportValue);
                        message += Math.Round(supportValue, roundTo) + "\t";
                    }
                    logService.WriteToFile(message);
                    logService.WriteToFile("");
                }

                logService.WriteToFile("");
                #endregion

            }

            logService.WriteToFile("");
        }

        #endregion

        #region helper methods

        private List<MarkingObjects> GetMarkings(ManufacturingProfile profile, SectionFaceType faceID)
        {
            List<MarkingObjects> markings = new List<MarkingObjects>();

            ReadOnlyCollection<ManufacturingGeometry> geometries = profile.GetGeometries(ManufacturingContext.Final, ManufacturingGeometryType.All);
            foreach (ManufacturingGeometry geometry in geometries)
            {
                ManufacturingGeometryType type = geometry.GeometryType;
                if (!(type == ManufacturingGeometryType.BlockMark
                    || type == ManufacturingGeometryType.RobotMark
                    || type == ManufacturingGeometryType.UserDefined
                    || type == ManufacturingGeometryType.BracketLocationMark
                    || type == ManufacturingGeometryType.XFrameMark))
                    continue;

                SectionFaceType geometryFaceID = (SectionFaceType)geometry.FaceID;
                BusinessObject marking = geometry.RepresentedEntity;

                switch (geometryFaceID)
                {
                    case SectionFaceType.Top:
                        if (faceID == SectionFaceType.Top)
                        {
                            MarkingObjects markingObject = new MarkingObjects(marking, GetGeometryName(type));
                            markings.Add(markingObject);
                        }
                        break;
                    case SectionFaceType.Bottom:
                        if (faceID == SectionFaceType.Bottom)
                        {
                            MarkingObjects markingObject = new MarkingObjects(marking, GetGeometryName(type));
                            markings.Add(markingObject);
                        }
                        break;
                    default:
                        if (faceID != SectionFaceType.Top || faceID != SectionFaceType.Bottom)
                        {
                            MarkingObjects markingObject = new MarkingObjects(marking, GetGeometryName(type));
                            markings.Add(markingObject);
                        }
                        break;

                }

            }

            return markings;

        }

        private string GetGeometryName(ManufacturingGeometryType geometry)
        {
            switch (geometry)
            {
                case ManufacturingGeometryType.BlockMark:
                    return "Block Mark";
                case ManufacturingGeometryType.RobotMark:
                    return "Robot Mark";
                case ManufacturingGeometryType.XFrameMark:
                    return "X Frame Mark";
                case ManufacturingGeometryType.BracketLocationMark:
                    return "Bracket Location Mark";
                case ManufacturingGeometryType.UserDefined:
                    return "User Defined Mark";
                case ManufacturingGeometryType.BendingLine:
                    return "Bending Line";
                case ManufacturingGeometryType.FittingAngle:
                    return "Fitting Angle";
                case ManufacturingGeometryType.TwistInformation:
                    return "Twist Information";
                case ManufacturingGeometryType.ProfileBendInformation:
                    return "Profile Bend Information";
                case ManufacturingGeometryType.ProfileRollInformation:
                    return "ProfileRollInformation";
                case ManufacturingGeometryType.ProfileDepthAfterUntwist:
                    return "Profile Depth After Untwist";
                case ManufacturingGeometryType.ProfileDepthBeforeUntwist:
                    return "Profile Depth Before Untwist";
                case ManufacturingGeometryType.ProfileLength:
                    return "Profile Length";
                case ManufacturingGeometryType.LandingCurve:
                    return "Landing Curve";
                default:
                    return string.Empty;
            }
        }

        #endregion

        #region conversion methods

        private double ConvertToMilliMeters(double value)
        {
            UOMManager uomMgr = MiddleServiceProvider.UOMMgr;
            return uomMgr.ConvertDBUtoUnit(UnitType.Distance, value, UnitName.DISTANCE_MILLIMETER);
        }

        private double ConvertToDegrees(double value)
        {
            UOMManager uomMgr = MiddleServiceProvider.UOMMgr;
            return uomMgr.ConvertDBUtoUnit(UnitType.Angle, value, UnitName.ANGLE_DEGREE);
        }

        #endregion

        #region private struct

        //structure to house both marking objects and their names.
        private struct MarkingObjects
        {
            private BusinessObject marking;
            private string name;

            /// <summary>
            /// Initializes a new instance of the <see cref="MarkingObjects"/> struct.
            /// </summary>
            /// <param name="marking">The marking.</param>
            /// <param name="name">The name.</param>
            public MarkingObjects(BusinessObject marking, string name)
            {
                this.marking = marking;
                this.name = name;
            }

            /// <summary>
            /// Gets the marking.
            /// </summary>
            /// <value>
            /// The marking.
            /// </value>
            public BusinessObject Marking
            {
                get { return this.marking; }
            }

            /// <summary>
            /// Gets the name of the marking.
            /// </summary>
            /// <value>
            /// The name.
            /// </value>
            public string Name
            {
                get { return this.name; }
            }
        }

        #endregion

    }
}