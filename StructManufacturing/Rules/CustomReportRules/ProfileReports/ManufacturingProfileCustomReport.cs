//**************************************************************************************
//  Copyright (C) 2015, Intergraph Corporation.  All rights reserved.
//
//  File        : ManufacturingProfileCustomReport.cs
//
//  Description : Creates a custom report for Profile Business Objects
//
//
//  Author      : Nautilus - HSV
//
//  History     : Created 7/15/2015
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
    /// Class for generating a custom profile report.
    /// </summary>
    public class ManufacturingProfileCustomReport : ManufacturingProfileCustomReportRuleBase
    {

        #region private members

        CustomReportLogService logService = new CustomReportLogService();
        ProfileReportService profileReport = null;
        int roundTo = 1;

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

        /// <summary>
        /// Generates a report for the specified entities saved to the given path.
        /// </summary>
        /// <param name="entities">The collection of entities.</param>
        /// <param name="filePath">The file path.</param>
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
                        try { ReportProfilePartInformation((ManufacturingProfile)businessObject); }
                        catch { continue; }
                    }
                    else if (businessObject is ProfilePart)
                    {
                        Collection<ManufacturingBase> mfgProfiles =
                            EntityService.GetManufacturingEntity(businessObject, ManufacturingEntityType.ManufacturingProfile);
                        foreach (ManufacturingBase profile in mfgProfiles)
                        {
                            try { ReportProfilePartInformation((ManufacturingProfile)profile); }
                            catch { continue; }
                        }
                    }
                    else if (businessObject is AssemblyBase)
                    {
                        foreach (ManufacturingProfile profile in GetManufacturedProfilesFromAssembly((AssemblyBase)businessObject))
                        {
                            try { ReportProfilePartInformation(profile); }
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

        #region private methods

        /// <summary>
        /// Reports the profile part information given a profile part
        /// and the file path.
        /// </summary>
        /// <param name="profile">The profile.</param>
        private void ReportProfilePartInformation(ManufacturingProfile profile)
        {           

            profileReport = GetProfileReportingService(profile);

            GenerateObjectInformation(profile);

            GenerateProcessingLengthInformation(profile);

            GenerateNeutralAxisInformation(profile);

            GenerateBevelShift(profile);

            GenerateConnectedObjectDistances(profile);

            GenerateFeatureDistances(profile);

            GenerateSlotDistances(profile);

            GenerateMarkingLineDistances(profile);

            GenerateSupportValues(profile);

            GenerateKnuckleInformation(profile);

            GenerateBendingCurveInformation(profile);

            GenerateCurveHeightWebBendingLine(profile);

            GenerateCurveHeightTopFlangeBendingLine(profile);

            GenerateCurveHeightBottomFlangeBendingLine(profile);

            logService.WriteToFile("");
            logService.WriteToFile("============================End of Report============================");
        }

        #region reporting methods

        private void GenerateObjectInformation(ManufacturingProfile profile)
        {
            INamedItem namedItem;
            IAssemblyChild assemblyChild = profile as IAssemblyChild;
            IAssembly assemblyParent = assemblyChild.AssemblyParent;

            string parentName = string.Empty;
            string assemblyName = string.Empty;
            string blockName = string.Empty;

            namedItem = assemblyParent as INamedItem;
            if (namedItem != null)
            {
                parentName = "Profile Name  : " + namedItem.Name;
                logService.WriteToFile(parentName);
            }

            assemblyChild = assemblyParent as IAssemblyChild;           
            if (assemblyChild == null) { logService.WriteToFile(""); return; }
            assemblyParent = assemblyChild.AssemblyParent;
            if (assemblyParent == null) { logService.WriteToFile(""); return; }

            if (assemblyParent is Assembly && !(assemblyParent is Block))
            {
                namedItem = assemblyParent as INamedItem;
                if (namedItem != null)
                {
                    assemblyName = " Assembly Name  : " + namedItem.Name;
                    logService.WriteToFile(assemblyName);
                }               
                if(assemblyParent is IAssemblyChild)
                {
                    assemblyChild = (IAssemblyChild)assemblyParent;
                    assemblyParent = assemblyChild.AssemblyParent;
                }
            }

            if (assemblyParent != null && assemblyParent is IAssemblyChild)
            {
                while (!(assemblyParent is Block))
                {
                    assemblyChild = (IAssemblyChild)assemblyParent;
                    assemblyParent = assemblyChild.AssemblyParent;
                }
                namedItem = assemblyParent as INamedItem;
                if (namedItem != null)
                {
                    blockName = " Block Name  : " + namedItem.Name;
                    logService.WriteToFile(blockName);
                }
            }

            logService.WriteToFile("");

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

        private void GenerateNeutralAxisInformation(ManufacturingProfile profile)
        {
            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("Test: NeutralAxis");
            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("");

            #region neutral axis variables

            double neutralAxisX = -1;
            double neutralAxisY = -1;

            profileReport.GetNuetralAxis(out neutralAxisX, out neutralAxisY);

            #region convert to milli meters

            neutralAxisX = ConvertToMilliMeters(neutralAxisX);
            neutralAxisY = ConvertToMilliMeters(neutralAxisY);

            #endregion

            #endregion

            logService.WriteToFile("Neutral Axis X" + "\t" + "= " + Math.Round(neutralAxisX, roundTo));
            logService.WriteToFile("Neutral Axis Y" + "\t" + "= " + Math.Round(neutralAxisY, roundTo));

            logService.WriteToFile("");
        }

        private void GenerateBevelShift(ManufacturingProfile profile)
        {
            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("Test: BevelShift");
            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("");

            #region bevel shift variables

            double bevelAngle = -1;
            double bevelGap = -1;
            int profileLength = (int)Math.Round(profile.ManufacturingLength(ManufacturingProfileLengthType.AfterFeatures, SectionFaceType.Unknown));
            string gapMessage = "Distance:";
            string angleMessage = "Bevel Angle:";
            string distanceMessage = "Gap Distance:";

            #region calculate bevel shift messages

            for (int i = 0; i <= profileLength; i++)
            {
                double tempDist = ConvertToMilliMeters(i);
                distanceMessage = distanceMessage + "\t" + tempDist;

                bevelAngle = -1;
                bevelGap = -1;

                profileReport.GetBevelShift(i, out bevelAngle, out bevelGap);
                bevelAngle = ConvertToDegrees(bevelAngle);
                bevelGap = ConvertToMilliMeters(bevelGap);
                angleMessage = angleMessage + "\t" + Math.Round(bevelAngle, roundTo);
                gapMessage = gapMessage + "\t" + Math.Round(bevelGap, roundTo);
            }

            #endregion

            #endregion

            logService.WriteToFile(distanceMessage);
            logService.WriteToFile(angleMessage);
            logService.WriteToFile(gapMessage);
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
            else if(parent is StiffenerPartBase) parentBO = parent as StiffenerPartBase;
            else if(parent is MemberPart) parentBO = parent as MemberPart;
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
                    logService.WriteToFile(connectedObjects.IndexOf(connectedObject) + "\t" + connectedObjectName + "\t\t\t" + Math.Round(distance,roundTo));
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

            if (profile.DetailedPart is StiffenerPartBase)
            {
                StiffenerPartBase stiffener = (StiffenerPartBase)profile.DetailedPart;
                featuresCollection = stiffener.Features;
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
                    "(LCL-(FD-LCS))" + "\t" + "End Cut At Base" + "\t" + "FD2" + "\t\t" + "FDH");
            }

            foreach (Feature feature in featuresCollection)
            {
                string featureType = string.Empty;
                switch (feature.FeatureType)
                {
                    case FeatureType.Edge :
                        featureType = "Edge Feature" + "\t";
                        break;
                    case FeatureType.Corner :
                        featureType = "Corner Feature" + "\t";
                        break;
                    case FeatureType.FlangeCut :
                        featureType = "Flange Cut" + "\t";
                        break;
                    case FeatureType.Slot :
                        featureType = "Slot" + "\t\t";
                        break;
                    case FeatureType.WaterStop :
                        featureType = "Water Stop" + "\t";
                        break;
                    case FeatureType.WebCut :
                        featureType = "Web Cut" + "\t\t";
                        break;
                    default :
                        featureType = "Unidentified" + "\t";
                        break;
                }

                bool featureBoolean = profileReport.CheckIfEndCutIsAtProfileBase(feature);
                double featureDistance = profileReport.GetDistanceBetweenFeatureAndStartOfProfileUnfold(feature);
                double featureDistance2 = 0;
                double featureHeight = 0;
                profileReport.GetFeatureDistanceAndHeight(feature, out featureDistance2, out featureHeight);
                featureDistance = ConvertToMilliMeters(featureDistance);
                featureDistance2 = ConvertToMilliMeters(featureDistance2);
                featureHeight = ConvertToMilliMeters(featureHeight);
                int index = featuresCollection.IndexOf(feature);
                logService.WriteToFile(index + "\t" + featureType + "\t" + Math.Round(featureDistance, roundTo)
                    + "\t\t" + Math.Round((featureDistance - landingCurveStart), roundTo)
                    + "\t\t" + Math.Round((landingCurveLength - featureDistance), roundTo)
                    + "\t\t" + Math.Round((landingCurveLength - (featureDistance - landingCurveStart)), roundTo)
                    + "\t\t" + featureBoolean
                    + "\t\t" + Math.Round(featureDistance2, roundTo)
                    + "\t\t" + Math.Round(featureHeight, roundTo));
            }
            logService.WriteToFile("");

            ReadOnlyCollection<Opening> openingCollection = null;
            if (profile.DetailedPart is ProfilePart)
            {
                ProfilePart profilePart = profile.DetailedPart as ProfilePart;
                openingCollection = profilePart.Openings;
            }

            if (openingCollection == null) return;
            if (featuresCollection.Count == 0 && openingCollection.Count == 0)
                return;

            logService.WriteToFile("\n");
            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("Test: Feature with PCs/FETs");
            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("");

            logService.WriteToFile("Feature,PC GUID,PC Name,Connected Object Name");

            foreach (Feature feature in featuresCollection)
            {
                ReadOnlyCollection<BusinessObject> PCorFETobjects, connectedObjects;
                List<String> pcGuidsAsStrings = 
                    profileReport.GetPhysicalConnectionsForFeatureOnProfile(
                                                            feature,
                                                            ( feature.FeatureType == FeatureType.FlangeCut ?
                                                              SectionFaceType.Top : SectionFaceType.Web_Left ),
                                                            out PCorFETobjects, out connectedObjects );
                for ( int i = 0; i < pcGuidsAsStrings.Count; ++i )
                {
                    logService.WriteToFile( feature.ToString() + "," +
                                            pcGuidsAsStrings[i] + "," +
                                            PCorFETobjects[i] + "," +
                                            ( connectedObjects[i] == null ?
                                              "(null)" : connectedObjects[i].ToString() ) );
                }

            }

            foreach (Opening opening in openingCollection)
            {
                ReadOnlyCollection<BusinessObject> PCorFETobjects, connectedObjects;
                List<String> pcGuidsAsStrings = 
                    profileReport.GetPhysicalConnectionsForFeatureOnProfile(
                                                            opening,
                // PenetratedFace fails for sketched features (SectionFaceType) opening.PenetratedFace.SectionId,
                /* So hard-code Web Left */                 SectionFaceType.Web_Left,
                                                            out PCorFETobjects, out connectedObjects );
                for ( int i = 0; i < pcGuidsAsStrings.Count; ++i )
                {
                    logService.WriteToFile( opening.ToString() + "," +
                                            pcGuidsAsStrings[i] + "," +
                                            PCorFETobjects[i] + "," +
                                            ( connectedObjects[i] == null ?
                                              "(null)" : connectedObjects[i].ToString() ) );
                }

            }
            logService.WriteToFile("");
        }

        private void GenerateSlotDistances(ManufacturingProfile profile)
        {
            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("Test: Slot Distances");
            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("");

            List<Feature> slots = new List<Feature>();
            if (profile.DetailedPart is StiffenerPartBase)
            {
                StiffenerPartBase stiffener = (StiffenerPartBase)profile.DetailedPart;
                ReadOnlyCollection<Feature> features = stiffener.Features;
                slots.AddRange(from feat in features
                               where feat.FeatureType == FeatureType.Slot
                               select feat);
            }

            if (slots.Count <= 0)
            {
                logService.WriteToFile("There are no slots on the profile to report on.");
                logService.WriteToFile("");
                return;
            }

            logService.WriteToFile("Nuber of Slot Features = " + slots.Count);
            logService.WriteToFile("Index" + "\t" + "Name" + "\t\t\t\t" + "Distance of countour edge due to slot on Web Left from start.");
            foreach (Feature slot in slots)
            {
                string slotName = slot.Name;
                double distance = profileReport.GetDistanceBetweenSlotAndStartOfProfileUnfold(slot, SectionFaceType.Web_Left);
                if (slotName.Length <= 15)
                {
                    logService.WriteToFile(slots.IndexOf(slot) + "\t" + slotName + "\t\t\t" + Math.Round(distance, roundTo));
                }
                else if (slotName.Length <= 23)
                {
                    logService.WriteToFile(slots.IndexOf(slot) + "\t" + slotName + "\t\t" + Math.Round(distance, roundTo));
                }
                else
                {
                    logService.WriteToFile(slots.IndexOf(slot) + "\t" + slotName + "\t" + Math.Round(distance, roundTo));
                }
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

            if(topMarkings.Count <= 0
                && webMarkings.Count <=0
                && bottomMarkings.Count <= 0)
            {
                logService.WriteToFile("There are no marking lines to report on.");
                logService.WriteToFile("");
                return;
            }

            if(topMarkings.Count > 0)
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

            for (int i = 0; i <= 8; i++)
            {
                #region set geometry type

                switch (i)
                {
                    case 0:
                        geometryType = ManufacturingGeometryType.BendingLine;
                        break;
                    case 1:
                        geometryType = ManufacturingGeometryType.FittingAngle;
                        break;
                    case 2:
                        geometryType = ManufacturingGeometryType.TwistInformation;
                        break;
                    case 3:
                        geometryType = ManufacturingGeometryType.ProfileBendInformation;
                        break;
                    case 4:
                        geometryType = ManufacturingGeometryType.ProfileRollInformation;
                        break;
                    case 5:
                        geometryType = ManufacturingGeometryType.ProfileDepthAfterUntwist;
                        break;
                    case 6:
                        geometryType = ManufacturingGeometryType.ProfileDepthBeforeUntwist;
                        break;
                    case 7:
                        geometryType = ManufacturingGeometryType.ProfileLength;
                        break;
                    case 8:
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
                        ;                      
                        if (geometryType == ManufacturingGeometryType.FittingAngle)
                            supportValue = ConvertToDegrees(supportValue);
                        else
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
                        if (geometryType == ManufacturingGeometryType.FittingAngle)
                            supportValue = ConvertToDegrees(supportValue);
                        else
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
                        if (geometryType == ManufacturingGeometryType.FittingAngle)
                            supportValue = ConvertToDegrees(supportValue);
                        else
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
                        distanceValue += rangeIncremntor/1000;
                        if (geometryType == ManufacturingGeometryType.FittingAngle)
                            supportValue = ConvertToDegrees(supportValue);
                        else
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

        private void GenerateKnuckleInformation(ManufacturingProfile profile)
        {
            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("Test: Knuckle Information");
            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("");

            ReadOnlyCollection<ManufacturingGeometry> geometries = profile.GetGeometries(ManufacturingContext.Final, ManufacturingGeometryType.KnuckleMark);

            if (geometries.Count <= 0)
            {
                logService.WriteToFile("There are no knuckles on the current profile.");
                logService.WriteToFile("");
                return;
            }

            #region knuckle variables

            bool knuckleThisSide;
            double knuckleAngle;
            double knuckleDistance;

            #endregion

            foreach (ManufacturingGeometry geometry in geometries)
            {
                knuckleThisSide = false;
                knuckleAngle = -1;
                knuckleDistance = -1;

                profileReport.GetKnuckleInformation(geometry, out knuckleThisSide, out knuckleAngle, out knuckleDistance);

                knuckleAngle = ConvertToDegrees(knuckleAngle);
                knuckleDistance = ConvertToMilliMeters(knuckleDistance);

                logService.WriteToFile("Knuckle Number: " + geometries.IndexOf(geometry));
                logService.WriteToFile("\t" + "Is the knuckle on this side: " + knuckleThisSide);
                logService.WriteToFile("\t" + "Knuckle Angle: " + Math.Round(knuckleAngle, roundTo));
                logService.WriteToFile("\t" + "Knuckle Distance from start of unfold: " + Math.Round(knuckleDistance, roundTo));
            }

            logService.WriteToFile("");
        }

        private void GenerateBendingCurveInformation(ManufacturingProfile profile)
        {
            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("Test: Bending Curve Information");
            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("");

            #region Bending Curve Information variables

            bool isBentUp = false;
            Position start = new Position();
            Position end = new Position();

            #endregion

            ReadOnlyCollection<ManufacturingGeometry> geometries = profile.GetGeometries(ManufacturingContext.Final, ManufacturingGeometryType.BendingLine);

            if (geometries.Count <= 0)
            {
                logService.WriteToFile("There are no bending curves on the profile to report.");
                logService.WriteToFile("");
                return;
            }

            foreach (ManufacturingGeometry geometry in geometries)
            {
                start = new Position();
                end = new Position();
                isBentUp = false;

                profileReport.GetBendingCurveInfo(geometry, out start, out end, out isBentUp);

                start.X = ConvertToMilliMeters(start.X);
                start.Y = ConvertToMilliMeters(start.Y);
                start.Z = ConvertToMilliMeters(start.Z);

                end.X = ConvertToMilliMeters(end.X);
                end.Y = ConvertToMilliMeters(end.Y);
                end.Z = ConvertToMilliMeters(end.Z);

                logService.WriteToFile("Bending Curve Number: " + geometries.IndexOf(geometry));

                if (isBentUp)
                    logService.WriteToFile("\t" + "The curve is bent up.");
                else
                    logService.WriteToFile("\t" + "The curve is not bent up.");

                logService.WriteToFile("\t" + "Starting Point (X, Y, Z):" + "\t" + "(" + 
                    Math.Round(start.X, roundTo) + ", " + 
                    Math.Round(start.Y, roundTo) + ", " + 
                    Math.Round(start.Z, roundTo) + ")");
                logService.WriteToFile("\t" + "Ending Point (X, Y, Z):" + "\t\t" + "(" +
                    Math.Round(end.X, roundTo) + ", " +
                    Math.Round(end.Y, roundTo) + ", " +
                    Math.Round(end.Z, roundTo) + ")");
                
            }

            logService.WriteToFile("");
        }

        private void GenerateCurveHeightWebBendingLine(ManufacturingProfile profile)
        {
            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("Test: Curve Height By Offset - Web Bending Line");
            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("");

            #region curve height variables

            double offset = .5;
            int curveCount = -1;
            double startOffset = -1;
            double endOffset = -1;
            double startHeight = -1;
            double endHeight = -1;

            double profileLength = profile.ManufacturingLength(ManufacturingProfileLengthType.AfterFeatures, SectionFaceType.Unknown);

            #endregion

            curveCount = (int)profileReport.GetGeometryCurveCount(ManufacturingGeometryType.BendingLine, SectionFaceType.Web_Left);

            if (curveCount <= 0)
            {
                logService.WriteToFile("There are no bending line curves to report on.");
                logService.WriteToFile("");
                return;
            }

            #region Curve Height By Offset

            logService.WriteToFile("Curve Height By Offset");
            logService.WriteToFile("");
            string message = "Curve" + "\t";
            int curveCounter = 0;
            do
            {
                double temp = ConvertToMilliMeters(curveCounter * offset);
                message += temp + "\t";
                curveCounter++;
            }
            while ((curveCounter * offset) < profileLength);
            logService.WriteToFile(message);

            for (int i = 1; i <= curveCount; i++)
            {
                message = i + "\t";
                curveCounter = 0;
                do
                {
                    double temp = profileReport.GetCurveHeightByOffset(ManufacturingGeometryType.BendingLine, SectionFaceType.Web_Left, i, curveCounter * offset);
                    temp = ConvertToMilliMeters(temp);
                    message += Math.Round(temp, roundTo) + "\t";
                    curveCounter++;
                }
                while((curveCounter * offset) < profileLength);
                logService.WriteToFile(message);
            }

            logService.WriteToFile("");
            logService.WriteToFile("");

            #endregion

            #region Curve Height At Ends

            logService.WriteToFile("Get Curve Height At Ends");
            logService.WriteToFile("Curve" + "\t" + "Start" + "\t" + "Height" + "\t" + "End" + "\t" + "Height");

            for (int i = 1; i <= curveCount; i++)
            {
                profileReport.GetCurveHeightAtEnds(ManufacturingGeometryType.BendingLine, SectionFaceType.Web_Left, i, out startOffset, out startHeight, out endOffset, out endHeight);
                startOffset = ConvertToMilliMeters(startOffset);
                startHeight = ConvertToMilliMeters(startHeight);
                endOffset = ConvertToMilliMeters(endOffset);
                endHeight = ConvertToMilliMeters(endHeight);
                logService.WriteToFile(i + "\t" + 
                    Math.Round(startOffset, roundTo) + "\t" + 
                    Math.Round(startHeight, roundTo) + "\t" + 
                    Math.Round(endOffset, roundTo) + "\t" + 
                    Math.Round(endHeight, roundTo));
            }

            logService.WriteToFile("");
            logService.WriteToFile("");

            #endregion

            #region Curve Heights by Offsets

            logService.WriteToFile("Get Curve Heights By Offsets");

            for (int i = 1; i <= curveCount; i++)
            {
                logService.WriteToFile("Curve # " + i);
                Array heights;
                profileReport.GetCurveHeightAtEnds(ManufacturingGeometryType.BendingLine, SectionFaceType.Web_Left, i, out startOffset, out startHeight, out endOffset, out endHeight);
                heights = profileReport.GetCurveHeightsByOffset(ManufacturingGeometryType.BendingLine, SectionFaceType.Web_Left, i, startOffset, offset);
                message = "Offset: " + "\t";
                string heightString = "Height: " + "\t";
                for (int j = heights.GetLowerBound(0); j <= heights.GetUpperBound(0); j++)
                {
                    double temp = startOffset + (j * offset);
                    temp = ConvertToMilliMeters(temp);
                    message += Math.Round(temp, roundTo) + "\t";
                    double tempHeight = (double)heights.GetValue(j);
                    tempHeight = ConvertToMilliMeters(tempHeight);
                    heightString += Math.Round(tempHeight, roundTo) + "\t";
                }
                logService.WriteToFile(message);
                logService.WriteToFile(heightString);
            }

            #endregion

            logService.WriteToFile("");
        }

        private void GenerateCurveHeightTopFlangeBendingLine(ManufacturingProfile profile)
        {
            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("Test: Curve Height By Offset - Top Flange Bending Line");
            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("");

            #region curve height variables

            double offset = .5;
            int curveCount = -1;
            double startOffset = -1;
            double endOffset = -1;
            double startHeight = -1;
            double endHeight = -1;

            double profileLength = profile.ManufacturingLength(ManufacturingProfileLengthType.AfterFeatures, SectionFaceType.Unknown);

            #endregion

            curveCount = (int)profileReport.GetGeometryCurveCount(ManufacturingGeometryType.BendingLine, SectionFaceType.Top);

            if (curveCount <= 0)
            {
                logService.WriteToFile("There are no bending line curves to report on.");
                logService.WriteToFile("");
                return;
            }

            #region Curve Height By Offset

            logService.WriteToFile("Curve Height By Offset");
            logService.WriteToFile("");
            string message = "Curve" + "\t";
            int curveCounter = 0;
            do
            {
                double temp = ConvertToMilliMeters(curveCounter * offset);
                message += temp + "\t";
                curveCounter++;
            }
            while ((curveCounter * offset) < profileLength);
            logService.WriteToFile(message);

            for (int i = 1; i <= curveCount; i++)
            {
                message = i + "\t";
                curveCounter = 0;
                do
                {
                    double temp = profileReport.GetCurveHeightByOffset(ManufacturingGeometryType.BendingLine, SectionFaceType.Top, i, curveCounter * offset);
                    temp = ConvertToMilliMeters(temp);
                    message += Math.Round(temp, roundTo) + "\t";
                    curveCounter++;
                }
                while ((curveCounter * offset) < profileLength);
                logService.WriteToFile(message);
            }

            logService.WriteToFile("");
            logService.WriteToFile("");

            #endregion

            #region Curve Height At Ends

            logService.WriteToFile("Get Curve Height At Ends");
            logService.WriteToFile("Curve" + "\t" + "Start" + "\t" + "Height" + "\t" + "End" + "\t" + "Height");

            for (int i = 1; i <= curveCount; i++)
            {
                profileReport.GetCurveHeightAtEnds(ManufacturingGeometryType.BendingLine, SectionFaceType.Top, i, out startOffset, out startHeight, out endOffset, out endHeight);
                startOffset = ConvertToMilliMeters(startOffset);
                startHeight = ConvertToMilliMeters(startHeight);
                endOffset = ConvertToMilliMeters(endOffset);
                endHeight = ConvertToMilliMeters(endHeight);
                logService.WriteToFile(i + "\t" +
                    Math.Round(startOffset, roundTo) + "\t" +
                    Math.Round(startHeight, roundTo) + "\t" +
                    Math.Round(endOffset, roundTo) + "\t" +
                    Math.Round(endHeight, roundTo));
            }

            logService.WriteToFile("");
            logService.WriteToFile("");

            #endregion

            #region Curve Heights by Offsets

            logService.WriteToFile("Get Curve Heights By Offsets");

            for (int i = 1; i <= curveCount; i++)
            {
                logService.WriteToFile("Curve # " + i);
                Array heights;
                profileReport.GetCurveHeightAtEnds(ManufacturingGeometryType.BendingLine, SectionFaceType.Top, i, out startOffset, out startHeight, out endOffset, out endHeight);
                heights = profileReport.GetCurveHeightsByOffset(ManufacturingGeometryType.BendingLine, SectionFaceType.Top, i, startOffset, offset);
                message = "Offset: " + "\t";
                string heightString = "Height: " + "\t";
                for (int j = heights.GetLowerBound(0); j <= heights.GetUpperBound(0); j++)
                {
                    double temp = startOffset + (j * offset);
                    temp = ConvertToMilliMeters(temp);
                    message += Math.Round(temp, roundTo) + "\t";
                    double tempHeight = (double)heights.GetValue(j);
                    tempHeight = ConvertToMilliMeters(tempHeight);
                    heightString += Math.Round(tempHeight, roundTo) + "\t";
                }
                logService.WriteToFile(message);
                logService.WriteToFile(heightString);
            }

            #endregion

            logService.WriteToFile("");
        }

        private void GenerateCurveHeightBottomFlangeBendingLine(ManufacturingProfile profile)
        {
            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("Test: Curve Height By Offset - Bottom Flange Bending Line");
            logService.WriteToFile("*************************************************************************");
            logService.WriteToFile("");

            #region curve height variables

            double offset = .5;
            int curveCount = -1;
            double startOffset = -1;
            double endOffset = -1;
            double startHeight = -1;
            double endHeight = -1;

            double profileLength = profile.ManufacturingLength(ManufacturingProfileLengthType.AfterFeatures, SectionFaceType.Unknown);

            #endregion

            curveCount = (int)profileReport.GetGeometryCurveCount(ManufacturingGeometryType.BendingLine, SectionFaceType.Bottom);

            if (curveCount <= 0)
            {
                logService.WriteToFile("There are no bending line curves to report on.");
                logService.WriteToFile("");
                return;
            }

            #region Curve Height By Offset

            logService.WriteToFile("Curve Height By Offset");
            logService.WriteToFile("");
            string message = "Curve" + "\t";
            int curveCounter = 0;
            do
            {
                double temp = ConvertToMilliMeters(curveCounter * offset);
                message += temp + "\t";
                curveCounter++;
            }
            while ((curveCounter * offset) < profileLength);
            logService.WriteToFile(message);

            for (int i = 1; i <= curveCount; i++)
            {
                message = i + "\t";
                curveCounter = 0;
                do
                {
                    double temp = profileReport.GetCurveHeightByOffset(ManufacturingGeometryType.BendingLine, SectionFaceType.Bottom, i, curveCounter * offset);
                    temp = ConvertToMilliMeters(temp);
                    message += Math.Round(temp, roundTo) + "\t";
                    curveCounter++;
                }
                while ((curveCounter * offset) < profileLength);
                logService.WriteToFile(message);
            }

            logService.WriteToFile("");
            logService.WriteToFile("");

            #endregion

            #region Curve Height At Ends

            logService.WriteToFile("Get Curve Height At Ends");
            logService.WriteToFile("Curve" + "\t" + "Start" + "\t" + "Height" + "\t" + "End" + "\t" + "Height");

            for (int i = 1; i <= curveCount; i++)
            {
                profileReport.GetCurveHeightAtEnds(ManufacturingGeometryType.BendingLine, SectionFaceType.Bottom, i, out startOffset, out startHeight, out endOffset, out endHeight);
                startOffset = ConvertToMilliMeters(startOffset);
                startHeight = ConvertToMilliMeters(startHeight);
                endOffset = ConvertToMilliMeters(endOffset);
                endHeight = ConvertToMilliMeters(endHeight);
                logService.WriteToFile(i + "\t" +
                    Math.Round(startOffset, roundTo) + "\t" +
                    Math.Round(startHeight, roundTo) + "\t" +
                    Math.Round(endOffset, roundTo) + "\t" +
                    Math.Round(endHeight, roundTo));
            }

            logService.WriteToFile("");
            logService.WriteToFile("");

            #endregion

            #region Curve Heights by Offsets

            logService.WriteToFile("Get Curve Heights By Offsets");

            for (int i = 1; i <= curveCount; i++)
            {
                logService.WriteToFile("Curve # " + i);
                Array heights;
                profileReport.GetCurveHeightAtEnds(ManufacturingGeometryType.BendingLine, SectionFaceType.Bottom, i, out startOffset, out startHeight, out endOffset, out endHeight);
                heights = profileReport.GetCurveHeightsByOffset(ManufacturingGeometryType.BendingLine, SectionFaceType.Bottom, i, startOffset, offset);
                message = "Offset: " + "\t";
                string heightString = "Height: " + "\t";
                for (int j = heights.GetLowerBound(0); j <= heights.GetUpperBound(0); j++)
                {
                    double temp = startOffset + (j * offset);
                    temp = ConvertToMilliMeters(temp);
                    message += Math.Round(temp, roundTo) + "\t";
                    double tempHeight = (double)heights.GetValue(j);
                    tempHeight = ConvertToMilliMeters(tempHeight);
                    heightString += Math.Round(tempHeight, roundTo) + "\t";
                }
                logService.WriteToFile(message);
                logService.WriteToFile(heightString);
            }

            #endregion

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
                    case SectionFaceType.Top :
                        if (faceID == SectionFaceType.Top)
                        {
                            MarkingObjects markingObject = new MarkingObjects(marking, GetGeometryName(type));
                            markings.Add(markingObject);
                        }                           
                        break;
                    case SectionFaceType.Bottom :
                        if (faceID == SectionFaceType.Bottom)
                        {
                            MarkingObjects markingObject = new MarkingObjects(marking, GetGeometryName(type));
                            markings.Add(markingObject);
                        }                           
                        break;
                    default :
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

        #endregion
    }
}
