//**************************************************************************************
//  Copyright (C) 2015, Intergraph Corporation.  All rights reserved.
//
//  File        : ManufacturingProfileCurvatureReport.cs
//
//  Description : Creates a report of the curvature of profile parts
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
    public class ManufacturingProfileCurvatureReport : ManufacturingProfileCustomReportRuleBase
    {

        #region private members

        private CustomReportLogService logService = null;
        private ProfileReportService profileReport;

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

            try
            {

                if (entities.Count() <= 0)
                {
                    throw new Exception("Throw to reach finally block.");                   
                }

                #region write csv column headings

                logService.WriteToFile("Profile Name," +
                    "Assembly Path," +
                    "Manufacturability," +
                    "Section Type," +
                    "Section Name," +
                    "Start Point X," +
                    "Start Point Y," +
                    "Start Point Z," +
                    "End Point X," +
                    "End Point Y," +
                    "End Point Z," +
                    "Curvature Type," +
                    "Depth Of BowString(mm)," +
                    "Radius Of BowString(mm)," +
                    "Maximum Depth Of BowString(mm)," +
                    "Maximum Radius Of BowString(mm)");

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

        #region reporting methods

        private void ReportProfilePartInformation(ProfilePart profile)
        {
            #region private variables

            profileReport = GetProfileReportingService(profile);

            int numberOfSections;
            Array pointCoordinates = Array.CreateInstance(typeof(double), 1);
            Array convexConcaveInformation = Array.CreateInstance(typeof(int), 1);
            Array depthOfBowString = Array.CreateInstance(typeof(double), 1);
            Array radiusOfBowString = Array.CreateInstance(typeof(double), 1);
            Array maxDepthOfBowString = Array.CreateInstance(typeof(double), 1);
            Array maxRadiusOfBowString = Array.CreateInstance(typeof(double), 1);

            string sectionType = string.Empty;
            string viewName = string.Empty;
            string min = string.Empty;
            string max = string.Empty;
            string assemblyPath = string.Empty;
            string sectionName = string.Empty;

            INamedItem namedItem = profile as INamedItem;
            IAssemblyChild assemblyChild = profile as IAssemblyChild;
            IAssembly assemblyParent = null;
            BusinessObject landingCurve = null;
            
            #endregion

            #region construct assembly path

            while (true)
            {
                if (assemblyChild == null) break;
                assemblyParent = assemblyChild.AssemblyParent;
                if (assemblyParent != null)
                {
                    INamedItem parentNamedItem = assemblyParent as INamedItem;
                    if (assemblyPath.Length <= 0)
                        assemblyPath = parentNamedItem.Name;
                    else
                        assemblyPath = parentNamedItem.Name + "|" + assemblyPath;
                }
                if (assemblyParent is IAssemblyChild)
                {
                    assemblyChild = assemblyParent as IAssemblyChild;
                    continue;
                }
                break;
            }

            #endregion

            #region find view name

            StiffenerPart detailedPart = profile as StiffenerPart;
            sectionName = detailedPart.SectionName;
            sectionType = detailedPart.SectionType;
            bool isBuiltUp = profileReport.IsProfileBuiltUp();

            if (isBuiltUp)
            {
                viewName = "JUAMfgProfileCurvatureLimBU";
                min = "BuiltUpSizeMin";
                max = "BuiltUpSizeMax";
            }
            else if (sectionType.Equals("UA") || sectionType.Equals("B"))
            {
                viewName = "JUAMfgProfileCurvatureLimAB";
                min = "WebSizeMin";
                max = "WebSizeMax";
            }

            #endregion

            #region get curvature information

            profileReport.GetConvexAndConcaveCurvatureInformation(
                1, 
                out landingCurve, 
                out numberOfSections, 
                ref pointCoordinates, 
                ref convexConcaveInformation, 
                ref depthOfBowString, 
                ref radiusOfBowString, 
                ref maxDepthOfBowString, 
                ref maxRadiusOfBowString);

            string curvatureInformation = string.Empty;

            #endregion

            #region determine if part is manufactuable

            for (int i = 0; i < numberOfSections; i++)
            {
                string manufacturable = string.Empty;
                if ((int)convexConcaveInformation.GetValue(i) != 1
                    && !string.IsNullOrEmpty(viewName))
                {
                    double webLength;
                    webLength = profileReport.GetProfileWebLength();

                    Array queryValues = Array.CreateInstance(typeof(object), 1);
                    string query = "SELECT BowStringDepth FROM " 
                        + viewName 
                        + " WHERE ( (CurvatureType = " 
                        + (int)convexConcaveInformation.GetValue(i) 
                        + ") and (" 
                        + webLength 
                        + " > " 
                        + min 
                        + ") and (" 
                        + webLength 
                        + " < " 
                        + max 
                        + " ) and (" 
                        + (double)maxRadiusOfBowString.GetValue(i) 
                        + " < RadiusOfCurvature ) )";

                    queryValues = profileReport.GetValuesFromDBQuery(query);
                    if (queryValues.GetUpperBound(0) >= 0)
                    {
                        if ((double)queryValues.GetValue(0) <= (double)maxRadiusOfBowString.GetValue(i))
                        {
                            manufacturable = "No";
                        }
                    }
                    
                }

                if (string.IsNullOrEmpty(manufacturable)) manufacturable = "Yes";

            #endregion

            #region populate string with information

                curvatureInformation = namedItem.Name 
                                     + "," + assemblyPath 
                                     + "," + manufacturable 
                                     + "," + sectionType 
                                     + "," + sectionName 
                                     + "," + Math.Round((double)pointCoordinates.GetValue(3 * i), 6)
                                     + "," + Math.Round((double)pointCoordinates.GetValue(3 * i + 1), 6)
                                     + "," + Math.Round((double)pointCoordinates.GetValue(3 * i + 2), 6)
                                     + "," + Math.Round((double)pointCoordinates.GetValue(3 * i + 3), 6)
                                     + "," + Math.Round((double)pointCoordinates.GetValue(3 * i + 4), 6)
                                     + "," + Math.Round((double)pointCoordinates.GetValue(3 * i + 5), 6);

                switch ((int)convexConcaveInformation.GetValue(i))
                {
                    case 1:
                        curvatureInformation += ",Straight,";
                        break;
                    case 2:
                        curvatureInformation += ",Convex,";
                        break;
                    case 3:
                        curvatureInformation += ",Concave,";
                        break;
                }

                curvatureInformation += Math.Round(1000 * (double)depthOfBowString.GetValue(i), 4)
                    + "," + Math.Round(1000 * (double)radiusOfBowString.GetValue(i), 4)
                    + "," + Math.Round(1000 * (double)maxDepthOfBowString.GetValue(i), 4)
                    + "," + Math.Round(1000 * (double)maxRadiusOfBowString.GetValue(i), 4);

                logService.WriteToFile(curvatureInformation);
            }

            #endregion

        }

        #endregion

        #endregion
    }
}
