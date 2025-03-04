//**************************************************************************************
//  Copyright (C) 2015, Intergraph Corporation.  All rights reserved.
//
//  File        : ManufacturingProfileCatalogSectionReport.cs
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
using System.Xml;

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Manufacturing.Middle.Services;
using Ingr.SP3D.Planning.Middle;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Manufacturing
{

    /// <summary>
    /// 
    /// </summary>
    public class ManufacturingProfileCatalogSectionReport : ManufacturingProfileCustomReportRuleBase
    {
        /// <summary>
        /// Generates a report for the specified entities saved to the given path.
        /// </summary>
        /// <param name="entities">The collection of entities.</param>
        /// <param name="filePath">The file path.</param>
        /// <exception cref="NotImplementedException"></exception>
        public override void Generate(IEnumerable<BusinessObject> entities, string filePath)
        {
            #region write report

            ProfileReportService profileReport = null;

            try
            {               

                profileReport = GetProfileReportingService();
                string xml = profileReport.GetProfileSectionData("", "", "ShipShapes");

                XmlDocument doc = new XmlDocument();
                doc.LoadXml(xml);
                doc.Save(filePath);
            }
            catch { /*DO NOTHING*/ }
            finally
            {
                if (profileReport != null) profileReport.Dispose();
            }

            #endregion
        }
    }
}
