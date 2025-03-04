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
//  History     : Created 7/15/2015
//
//
//**************************************************************************************

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Content.Manufacturing.Services;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Manufacturing.Middle.Services;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Planning.Middle;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// 
    /// </summary>
    public class ManufacturingProfileKnuckleReport : ManufacturingProfileCustomReportRuleBase
    {

        #region private members

        private ProfileReportService profileReport;
        private CustomReportLogService logService = new CustomReportLogService();

        #endregion

        /// <summary>
        /// Generates a report for the specified entities saved to the given path.
        /// </summary>
        /// <param name="entities">The collection of entities.</param>
        /// <param name="filePath">The file path.</param>
        /// <exception cref="NotImplementedException"></exception>
        public override void Generate(IEnumerable<BusinessObject> entities, string filePath)
        {
            logService.OpenFile(filePath);

            try
            {

                #region write csv header

                logService.WriteToFile(
                      "Profile Name,"
                    + "AssemblyPath,"
                    + "Section Type,"
                    + "Section Name,"
                    + "Side,"
                    + "Knuckle Angle"
                    );

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

        private void ReportProfilePartInformation(ProfilePart profile)
        {
            profileReport = base.GetProfileReportingService(profile);

            #region declare variables

            string assemblyPath = string.Empty;
            string sectionName = string.Empty;
            string knuckleInfo = string.Empty;
            string sectionType = string.Empty;
            string side = string.Empty;

            bool failed1 = false; bool failed2 = false;

            double knuckleAngle = -1;
            ReadOnlyCollection<BusinessObject> knuckleCollection, knuckleLineCollection; ReadOnlyCollection<double> knuckleAngleCollection;

            #endregion

            #region write report

            assemblyPath = GetAssemblyPath(profile);
            sectionName = profileReport.GetSectionName();
            sectionType = profileReport.GetSectionType();

            profileReport.GetBendingProfileKnuckleData(SectionFaceType.Web_Left, out knuckleCollection, out knuckleLineCollection, out knuckleAngleCollection);

            if (knuckleCollection.Count != 0)
            {
                for (int idx = 0; idx < knuckleCollection.Count; idx++)
                {
                    knuckleAngle = knuckleAngleCollection[idx];
                    if (knuckleAngle < 180) side = "WEB_RIGHT";
                    else side = "WEB_LEFT";
                    knuckleInfo = profile.Name + ","
                                + assemblyPath + ","
                                + sectionType + ","
                                + sectionName + ","
                                + side + ","
                                + knuckleAngle;
                    logService.WriteToFile(knuckleInfo);
                }
            }
            else failed1 = true;

            knuckleCollection = null;
            knuckleLineCollection = null;
            knuckleAngleCollection = null;

            profileReport.GetBendingProfileKnuckleData(SectionFaceType.Top, out knuckleCollection, out knuckleLineCollection, out knuckleAngleCollection);

            if (knuckleCollection.Count != 0)
            {
                for (int idx = 0; idx < knuckleCollection.Count; idx++)
                {
                    knuckleAngle = knuckleAngleCollection[idx];
                    if (knuckleAngle < 180) side = "FLANGE_BOTTOM";
                    else side = "FLANGE_TOP";
                    knuckleInfo = profile.Name + ","
                                + assemblyPath + ","
                                + sectionType + ","
                                + sectionName + ","
                                + side + ","
                                + knuckleAngle;
                    logService.WriteToFile(knuckleInfo);
                }
            }
            else failed2 = true;

            if(failed1 && failed2) {
                knuckleInfo = profile.Name + ","
                                + assemblyPath + ","
                                + sectionType + ","
                                + sectionName + ","
                                + "No Knuckle";                                
                logService.WriteToFile(knuckleInfo);}

            #endregion
        }
    }
}
