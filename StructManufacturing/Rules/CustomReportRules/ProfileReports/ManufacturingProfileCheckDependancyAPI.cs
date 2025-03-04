//**************************************************************************************
//  Copyright (C) 2015, Intergraph Corporation.  All rights reserved.
//
//  File        : ManufacturingProfileCheckDependancyAPI.cs
//
//  Description : 
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
using Ingr.SP3D.Common.Middle.Services.Hidden;
using Ingr.SP3D.Content.Manufacturing.Services;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Manufacturing.Middle.Services;
using Ingr.SP3D.Planning.Middle;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.ReferenceData.Middle;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// 
    /// </summary>
    public class ManufacturingProfileCheckDependancyAPI : ManufacturingProfileCustomReportRuleBase
    {

        #region private members

        private CustomReportLogService logService = new CustomReportLogService();
        ProfileReportService profileReport;

        #endregion

        /// <summary>
        /// Generates a report for the specified entities saved to the given path.
        /// </summary>
        /// <param name="entities">The collection of entities.</param>
        /// <param name="filePath">The file path.</param>
        /// <exception cref="NotImplementedException"></exception>
        public override void Generate(IEnumerable<BusinessObject> entities, string filePath)
        {

            logService = new CustomReportLogService();
            logService.OpenFile(filePath);

            if (entities.Count() <= 0)
            {
                logService.CloseFile();
                return;
            }

            try
            {

                #region write csv column headings

                logService.WriteToFile("Assembly Path," +
                              "Profile Name," +
                              "Status," +
                              "Details," +
                              "Section Type," +
                              "Section Name," +
                              "Width," +
                              "Height," +
                              "Web Depth," +
                              "Web Thickness," +
                              "Flange Width," +
                              "Flange Thickness," +
                              "Cross section Area," +
                              "Molded Length," +
                              "Material Type," +
                              "Approximate Length," +
                              "Is Linear," +
                              "Is Twisted,");

                #endregion

                foreach (BusinessObject businessObject in entities)
                {
                    if (businessObject is ProfilePart)
                    {
                        ReportProfilePartInformation((ProfilePart)businessObject);
                    }
                    else if (businessObject is AssemblyBase)
                    {
                        ReadOnlyCollection<ProfilePart> profiles = GetDetailedProfilesFromAssembly((AssemblyBase)businessObject);
                        foreach (ProfilePart profile in profiles) ReportProfilePartInformation(profile);
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

        private void ReportProfilePartInformation(ProfilePart profile)
        {
            #region private variables

            profileReport = GetProfileReportingService(profile);

            string assemblyPath = string.Empty;
            string partInfo = string.Empty;
            string supportProperties = string.Empty;
            string apiSummary = string.Empty;
            string sectionType = string.Empty;
            string sectionName = string.Empty;
            string profileName = string.Empty;
            string profileMaterialType = string.Empty;

            double profileWidth = -1;
            double profileHeight = -1;
            double profileWebLength = -1;
            double profileWebThickness = -1;
            double profileFlangeLength = -1;
            double profileFlangeThickness = -1;
            double profileXSectionArea = -1;
            double profileMoldedLength = -1;

            #endregion

            #region fill variables

            StiffenerPart stiffenerPart = profile as StiffenerPart;
            if (stiffenerPart == null) return;
            profileName = profile.Name;
            profileWidth = profileReport.GetProfileWidth();
            profileHeight = profileReport.GetProfileHeight();
            sectionName = stiffenerPart.SectionName;
            sectionType = stiffenerPart.SectionType;
            profileMaterialType = stiffenerPart.MaterialType;
            profileWebLength = profileReport.GetProfileWebLength();
            profileWebThickness = profileReport.GetProfileWebThickness();
            profileFlangeLength = profileReport.GetProfileFlangeLength();
            profileFlangeThickness = profileReport.GetProfileFlangeThickness();
            profileXSectionArea = profileReport.GetProfileCrossSectionArea();
            profileMoldedLength = profileReport.GetProfileMoldedLength();

            assemblyPath = GetAssemblyPath(profile);
            partInfo = sectionType + ","
                     + sectionName + ","
                     + profileWidth + ","
                     + profileHeight + ","
                     + profileWebLength + ","
                     + profileWebThickness + ","
                     + profileFlangeLength + ","
                     + profileFlangeThickness + ","
                     + profileXSectionArea + ","
                     + profileMoldedLength + ","
                     + profileMaterialType;

            supportProperties = profileReport.GetApproximateLength() + ","
                              + profileReport.IsLinear() + ","
                              + profileReport.IsTwisted() + ",";

            #endregion

            #region API Calls

            bool failed = false;

            ICurve curve = null;
            ISurface surface = null;
            Vector vector, xVec, yVec;
            Position start, end, orig;

            ReadOnlyCollection<BusinessObject> countours;
            ReadOnlyDictionary<BusinessObject,BusinessObject> ports;
            bool center;

            try { profileReport.GetProfileLandingCurve(out curve, out vector, out center, true); }
            catch { failed = true; apiSummary = "GetProfilePartLandingCurve failed for SideA"; }

            if (!failed)
            {               
                curve.EndPoints(out start, out end);

                try { profileReport.GetOrientation(start, out xVec, out yVec, out orig); }
                catch
                {
                    if (string.IsNullOrEmpty(apiSummary)) apiSummary = "GetOrientation failed at start point";
                    else apiSummary += "\n" + ",,,GetOrientation failed at start point";
                }

                try { profileReport.GetOrientation(end, out xVec, out yVec, out orig); }
                catch
                {
                    if (string.IsNullOrEmpty(apiSummary)) apiSummary = "GetOrientation failed at start point";
                    else apiSummary += "\n" + ",,,GetOrientation failed at start point";
                }
            }

            try { profileReport.GetProfileLandingCurve(out curve, out vector, out center, false); }
            catch 
            { 
                if (string.IsNullOrEmpty(apiSummary)) apiSummary = "GetProfilePartLandingCurve failed for SideB";
                else apiSummary += "\n" + ",,,GetProfilePartLandingCurve failed for SideB"; 
            }

            try { profileReport.GetExtendedStiffenerLandingCurve(out curve, out vector, out center, true); }
            catch
            {
                if (string.IsNullOrEmpty(apiSummary)) apiSummary = "GetExtendedStiffPartLandingCurve failed for SideA";
                else apiSummary += "\n" + ",,,GetExtendedStiffPartLandingCurve failed for SideA"; 
            }

            try { profileReport.GetExtendedStiffenerLandingCurve(out curve, out vector, out center, false); }
            catch
            {
                if (string.IsNullOrEmpty(apiSummary)) apiSummary = "GetExtendedStiffPartLandingCurve failed for SideB";
                else apiSummary += "\n" + ",,,GetExtendedStiffPartLandingCurve failed for SideB";
            }

            try { curve = profileReport.GetMoldedStiffenerLandingCurve(); }
            catch
            {
                if (string.IsNullOrEmpty(apiSummary)) apiSummary = "GetMoldedStiffPartLandingCurve failed";
                else apiSummary += "\n" + ",,,GetMoldedStiffPartLandingCurve failed";
            }

            try { profileReport.GetProfileContours(out surface, out countours, out ports, SectionFaceType.Web_Left); }
            catch
            {
                if (string.IsNullOrEmpty(apiSummary)) apiSummary = "GetProfileContours failed for Web Left";
                else apiSummary += "\n" + ",,,GetProfileContours failed for Web Left";
            }

            try { profileReport.GetProfileContours(out surface, out countours, out ports, SectionFaceType.Web_Right); }
            catch
            {
                if (string.IsNullOrEmpty(apiSummary)) apiSummary = "GetProfileContours failed for Web Right";
                else apiSummary += "\n" + ",,,GetProfileContours failed for Web Right";
            }

            #endregion

            #region fill string

            string mfgString = string.IsNullOrEmpty(apiSummary) ? "Okay" : "Bad";

            string finishedString = assemblyPath + ","
                                  + profile.Name + ","
                                  + mfgString + ","
                                  + apiSummary + ","
                                  + partInfo + ","
                                  + supportProperties;

            logService.WriteToFile(finishedString);

            #endregion
        }

        #endregion
    }
}
