//**************************************************************************************
//  Copyright (C) 2015, Intergraph Corporation.  All rights reserved.
//
//  File        : ManufacturingProfileWeightReport.cs
//
//  Description : Populates a CSV file with profile detailing weight and mfg weight
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
    public class ManufacturingProfileWeightReport : ManufacturingProfileCustomReportRuleBase
    {

        #region private members

        private CustomReportLogService logService = new CustomReportLogService();
        private ProfileReportService profileReport = null;

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

                #region write csv headings

                logService.WriteToFile("Unit,kg");
                logService.WriteToFile("Margin is applied (if any).");
                logService.WriteToFile(
                    "PartName," +
                    "Assembly," +
                    "CrossSection," +
                    "Dry Weight," +
                    "Final Mfg Weight," +
                    "Web Weight," +
                    "Flange Weight," +
                    "Non Trimmed Margin Applied");

                #endregion

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

        private void ReportProfilePartInformation(ManufacturingProfile profile)
        {

            profileReport = GetProfileReportingService(profile);

            #region private variables

            double totalWeight = 0; 
            double faceWeight = 0;
            double flangeWeight = 0;
            double webWeight = 0;
            double totalMarginWeight = 0;
            double dryWeight = profileReport.GetDryWeight();


            double flangeThickness = profileReport.GetProfileFlangeThickness();
            double webThickness = profileReport.GetProfileWebThickness();

            string sectionName = profileReport.GetSectionName();
            

            ReadOnlyCollection<ManufacturingGeometry> geometries;
            Collection<int> faces = new Collection<int>();

            string assemblyPath = GetAssemblyPath(profile);

            #endregion

            #region write report

            geometries = profile.GetGeometries(ManufacturingContext.Final, ManufacturingGeometryType.All);
            foreach (ManufacturingGeometry geometry in geometries)
            {
                if (faces.Contains(geometry.FaceID)) continue;
                faces.Add(geometry.FaceID);
            }

            foreach (int face in faces)
            {
                SectionFaceType faceType = (SectionFaceType)Enum.ToObject(typeof(SectionFaceType), face);
                if (faceType == SectionFaceType.Top || faceType == SectionFaceType.Bottom)
                {
                    faceWeight = profile.GetWeight(31, faceType, flangeThickness);
                    flangeWeight += faceWeight;
                }
                else if (faceType == SectionFaceType.Web_Left || faceType == SectionFaceType.Web_Right)
                {
                    faceWeight = profile.GetWeight(31, faceType, webThickness);
                    webWeight += faceWeight;
                }
                totalWeight += faceWeight;
                faceWeight = 0;
            }

            totalMarginWeight = profile.GetWeight(16, SectionFaceType.Unknown, webThickness);

            logService.WriteToFile(
                  profile.Name + ","
                + assemblyPath + ","
                + sectionName + ","
                + dryWeight + ","
                + totalWeight + ","
                + webWeight + ","
                + flangeWeight + ","
                + totalMarginWeight
                );

            #endregion

        }
    }
}
