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
using System.Collections.ObjectModel;
using System.Linq;
using System.Xml;
using System.Text;
using System.Threading.Tasks;

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Manufacturing.Services;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Manufacturing.Middle.Services;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Planning.Middle;
using System.Data.SqlClient;
using System.Data;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// 
    /// </summary>
    public class ManufacturingProfileSectionCatalogXML : ManufacturingProfileCustomReportRuleBase
    {

        #region private members

        string sqlQuery    = @"SELECT DISTINCT JDPC.Name As PartClassName, JSCS.SectionName, JSXS.SectionTypeName, RDCST.name As CrossSectionType, JRS.Name As ReferenceStandard, XSW.WebLength, XSW.WebThickness, XSF.FlangeLength, XSF.FlangeThickness, JSFBG.gf As FlangeGage,  JSCSD.Depth As OuterDiameter, JSCSD.Area * JDM.Density As WeightPerUnitLength, JSS.tnom As NominalThickness " +
                             @"FROM JDPartClass JDPC " +
                             @"INNER JOIN XReferenceStdHasPartClasses XRSHPC ON XRSHPC.OidDestination = JDPC.Oid " +
                             @"INNER JOIN JReferenceStandard JRS ON JRS.Oid = XRSHPC.OidOrigin " +
                             @"INNER JOIN JDCrossSection JDCS ON JDCS.Type = JDPC.Name " +
                             @"INNER JOIN JSTRUCTCrossSectionDimensions JSCSD ON JSCSD.Oid = JDCS.Oid " +
                             @"INNER JOIN JSTRUCTCrossSection JSCS ON JSCS.Oid = JDCS.Oid " +
                             @"INNER JOIN JStructXSection JSXS ON JSXS.oid = JDCS.oid " +
                             @"LEFT JOIN XCrossSectionClassToType XCSCT ON XCSCT.oidOrigin = JDPC.oid " +
                             @"LEFT JOIN REFDATCrossSectionType RDCST ON RDCST.oid = XCSCT.oidDestination " +
                             @"LEFT JOIN JUAHSS JSS ON JSS.oid = JDCS.oid " +
                             @"LEFT JOIN JUAHSSC JSSC ON JSSC.oid = JDCS.oid " +
                             @"LEFT JOIN JSTRUCTFlangedBoltGage JSFBG ON JSFBG.oid = JDCS.oid " +
                             @"LEFT JOIN JUAXSectionWeb XSW ON XSW.Oid = JDCS.Oid " +
                             @"LEFT JOIN JUAXSectionFlange XSF ON XSF.Oid = JDCS.Oid, " +
                             @"(SELECT Density FROM JDMaterial M WHERE M.MaterialType = 'Steel - Carbon' AND M.MaterialGrade = 'A') As JDM " +
                             @"WHERE NOT EXISTS(SELECT * FROM COREBoolAttribute CBA WHERE CBA.iid ='0742B82C-EABE-494B-829A-E66C2A6D1CEE' AND JDPC.oid = CBA.oid) " +
                             @"ORDER BY JDPC.Name, JSCS.SectionName";

        string serverName = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantCatalog.Server;
        string catalogName = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantCatalog.Name;

        XmlDocument sectionMap = new XmlDocument();
        XmlDocument outputDocument = new XmlDocument();
        XmlNode rootNode;

        UOMManager uom = MiddleServiceProvider.UOMMgr;

        #endregion

        #region overriden methods

        /// <summary>
        /// Generates the specified entities.
        /// </summary>
        /// <param name="entities">The entities.</param>
        /// <param name="filePath">The file path.</param>
        /// <exception cref="System.NotImplementedException"></exception>
        public override void Generate(IEnumerable<BusinessObject> entities, string filePath)
        {

            try
            {
                string symbolShare = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantCatalog.SymbolShare;
                string fileLocation = symbolShare + @"\StructManufacturing\DSTV_OUTPUT\ATeK_OUTPUT\CATALOG_MAP.xml";
                sectionMap.Load(fileLocation);

                rootNode = outputDocument.CreateNode(XmlNodeType.Element, "CAShapes", "m_xmlns");
                outputDocument.AppendChild(rootNode);

                string sqlConnection = "server=" + serverName + ";Database=" + catalogName + ";Integrated Security=yes";
                DataTable queryResults = new DataTable();

                using (SqlConnection connection = new SqlConnection(sqlConnection))
                using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                using (SqlDataAdapter dataAdapter = new SqlDataAdapter(command))
                    dataAdapter.Fill(queryResults);

                WriteSectionXML(queryResults);

                outputDocument.Save(filePath);
            }
            catch (Exception e) { System.Console.WriteLine(e.Message); }
            finally
            {
                sectionMap = null;
                rootNode = null;
            }

        }

        #endregion

        #region private methods

        private void WriteSectionXML(DataTable queryResults)
        {
            foreach (DataRow row in queryResults.Rows)
            {
                string sectionClass = string.Empty;
                string sectionName = string.Empty;
                string sectionType = string.Empty;
                string crossSectionType = string.Empty;
                int materialtype = -1;
                double webLength = -1;
                double webThickness = -1;
                double flangeLength = -1;
                double flangeThickness = -1;
                double flangeGage = -1;
                double weightPerLength = -1;
                int measuringSystem = -1;
                double outerDiameter = -1;
                double innerDiameter = -1;
                double tubeThickness = -1;
                double tubeDiameter = 0;
                XmlNode mapElement;
                XmlNode sectionNode;
                XmlNode shapeNode;
                XmlNodeList sectionMapNodeList;

                sectionClass = (string)row["PartClassName"];
                sectionName = (string)row["SectionName"];
                sectionType = (string)row["SectionTypeName"];

                if (!(row["CrossSectionType"] is System.DBNull)) crossSectionType = (string)row["CrossSectionType"];
                else crossSectionType = (-1).ToString();

                mapElement = sectionMap.SelectSingleNode("/SMS_CATALOG_MAP/MATERIAL_TYPE_MAPPING/XSECTION_TYPE[@S3D_NAME='" + sectionType + "']");

                if (mapElement != null) { materialtype = int.Parse(mapElement.Attributes["ATek_NAME"].Value); mapElement = null; }
                else materialtype = -1;

                if (!(row["WebLength"] is System.DBNull)) webLength = uom.ConvertDBUtoUnit(UnitType.Distance, (double)row["WebLength"], UnitName.DISTANCE_INCH);
                else webLength = 0;

                if (!(row["WebThickness"] is System.DBNull)) webThickness = uom.ConvertDBUtoUnit(UnitType.Distance, (double)row["WebThickness"], UnitName.DISTANCE_INCH);
                else webThickness = 0;

                if (!(row["FlangeLength"] is System.DBNull)) flangeLength = uom.ConvertDBUtoUnit(UnitType.Distance, (double)row["FlangeLength"], UnitName.DISTANCE_INCH);
                else flangeLength = 0;

                if (!(row["FlangeThickness"] is System.DBNull)) flangeThickness = uom.ConvertDBUtoUnit(UnitType.Distance, (double)row["FlangeThickness"], UnitName.DISTANCE_INCH);
                else flangeThickness = 0;

                if (!(row["FlangeGage"] is System.DBNull)) flangeGage = uom.ConvertDBUtoUnit(UnitType.Distance, (double)row["FlangeGage"], UnitName.DISTANCE_INCH);
                else flangeGage = 0;

                measuringSystem = 2;
                weightPerLength = (double)row["WeightPerUnitLength"];

                if (crossSectionType == "CircTube" && !(row["OuterDiameter"] is System.DBNull)) outerDiameter = uom.ConvertDBUtoUnit(UnitType.Distance, (double)row["OuterDiameter"], UnitName.DISTANCE_INCH);
                else outerDiameter = 0;

                if (!(row["Nominalthickness"] is System.DBNull)) tubeThickness = uom.ConvertDBUtoUnit(UnitType.Distance, (double)row["NominalThickness"], UnitName.DISTANCE_INCH);
                else tubeThickness = 0;

                if (tubeThickness > 0) innerDiameter = outerDiameter - 2 * tubeThickness;
                else innerDiameter = 0;

                sectionNode = outputDocument.SelectSingleNode("//Section[@Section='" + sectionClass + "' and @MaterialType='" + materialtype + "']");
                if (sectionNode == null)
                {
                    sectionMapNodeList = null;
                    sectionMapNodeList = sectionMap.SelectNodes("/SMS_CATALOG_MAP/SECTION_MAP_NAME_MAPPING/XSECTION_TYPE[@S3D_NAME='" + sectionType + "']/XSECTION_MAP_NAME");
                    sectionNode = CreateSectionNode(sectionClass, materialtype, sectionMapNodeList);
                    rootNode.AppendChild(sectionNode);                    
                }
                shapeNode = CreateShapeNode(sectionName, sectionName, webLength, webThickness,
                                                flangeLength, flangeThickness, flangeGage, weightPerLength,
                                                measuringSystem, sectionName, sectionName, outerDiameter,
                                                innerDiameter, tubeThickness, tubeDiameter);
                sectionNode.AppendChild(shapeNode);
            }              
        }

        private XmlNode CreateSectionNode(string sectionClass, int materialType, XmlNodeList sectionMapNodeList)
        {
            XmlNode sectionNode;
            XmlElement sectionElement;
            XmlNode sectionMapNode;
            XmlElement sectionMapChildElement;
            XmlNode sectionNameNode;
            string sectionMapName;

            sectionNode = outputDocument.CreateNode(XmlNodeType.Element, "Section", "m_xmlns");
            sectionElement = (XmlElement)sectionNode;
            sectionElement.SetAttribute("Section", sectionClass);
            sectionElement.SetAttribute("MaterialType", materialType.ToString());
            sectionMapNode = outputDocument.CreateNode(XmlNodeType.Element, "SectionMapping", "m_xmlns");
            sectionElement.AppendChild(sectionMapNode);

            if (sectionMapNodeList != null)
            {
                foreach (XmlNode node in sectionMapNodeList)
                {
                    sectionMapChildElement = (XmlElement)node;
                    sectionMapName = sectionMapChildElement.GetAttribute("NAME");
                    sectionNameNode = outputDocument.CreateNode(XmlNodeType.Element, "Section_Map_Name", "m_xmlns");
                    sectionNameNode.InnerText = sectionMapName;
                    sectionMapNode.AppendChild(sectionNameNode);
                }
            }

            return (XmlNode)sectionElement;

        }

        private XmlNode CreateShapeNode(string sectionName, string displayString, double webLength, double webThickness, 
                                        double flangeLength, double flangeThickness, double flangeGage, double weightPerLength,
                                        int measuringSystem, string winCadName, string cadCadName, double outerDiameter,
                                        double innerDiameter, double nominalThickness, double nominalDiameter)
        {
            XmlNode shapeRootNode = outputDocument.CreateNode(XmlNodeType.Element, "ShapeData", "m_Xmlns");

            XmlNode shapeNode = outputDocument.CreateNode(XmlNodeType.Element, "Shape", "m_Xmlns");
            shapeNode.InnerText = sectionName;
            shapeRootNode.AppendChild(shapeNode);

            XmlNode webDepthNode = outputDocument.CreateNode(XmlNodeType.Element, "WebDepth", "m_Xmlns");
            webDepthNode.InnerText = webLength.ToString();
            shapeRootNode.AppendChild(webDepthNode);

            XmlNode webThicknessNode = outputDocument.CreateNode(XmlNodeType.Element, "WebThickness", "m_Xmlns");
            webThicknessNode.InnerText = webThickness.ToString();
            shapeRootNode.AppendChild(webThicknessNode);

            XmlNode breathFlangeNode = outputDocument.CreateNode(XmlNodeType.Element, "BreathOfFlange", "m_Xmlns");
            breathFlangeNode.InnerText = flangeLength.ToString();
            shapeRootNode.AppendChild(breathFlangeNode);

            XmlNode flangeThicknessNode = outputDocument.CreateNode(XmlNodeType.Element, "FlangeThickness", "m_Xmlns");
            flangeThicknessNode.InnerText = flangeThickness.ToString();
            shapeRootNode.AppendChild(flangeThicknessNode);

            XmlNode flangeGageNode = outputDocument.CreateNode(XmlNodeType.Element, "StandardFlangeGage", "m_Xmlns");
            flangeGageNode.InnerText = flangeGage.ToString();
            shapeRootNode.AppendChild(flangeGageNode);

            XmlNode weightPerNode = outputDocument.CreateNode(XmlNodeType.Element, "WeightPerFoot", "m_Xmlns");
            weightPerNode.InnerText = weightPerLength.ToString();
            shapeRootNode.AppendChild(weightPerNode);

            XmlNode systemNode = outputDocument.CreateNode(XmlNodeType.Element, "MeasuringSystem", "m_Xmlns");
            systemNode.InnerText = measuringSystem.ToString();
            shapeRootNode.AppendChild(systemNode);

            XmlNode displayNode = outputDocument.CreateNode(XmlNodeType.Element, "WinCAd_Name", "m_Xmlns");
            displayNode.InnerText = displayString;
            shapeRootNode.AppendChild(displayNode);

            XmlNode winCadnode = outputDocument.CreateNode(XmlNodeType.Element, "CADCAD_Name", "m_Xmlns");
            winCadnode.InnerText = winCadName;
            shapeRootNode.AppendChild(winCadnode);

            XmlNode odNode = outputDocument.CreateNode(XmlNodeType.Element, "OD", "m_Xmlns");
            odNode.InnerText = outerDiameter.ToString();
            shapeRootNode.AppendChild(odNode);

            XmlNode idNode = outputDocument.CreateNode(XmlNodeType.Element, "ID", "m_Xmlns");
            idNode.InnerText = innerDiameter.ToString();
            shapeRootNode.AppendChild(idNode);

            XmlNode ntNode = outputDocument.CreateNode(XmlNodeType.Element, "Nominal_Thickness", "m_Xmlns");
            ntNode.InnerText = nominalThickness.ToString();
            shapeRootNode.AppendChild(ntNode);

            XmlNode ndNode = outputDocument.CreateNode(XmlNodeType.Element, "Nominal_Diameter", "m_Xmlns");
            ndNode.InnerText = nominalDiameter.ToString();
            shapeRootNode.AppendChild(ndNode);

            XmlNode smNode = outputDocument.CreateNode(XmlNodeType.Element, "Shape_Mapping", "m_Xmlns");
            shapeRootNode.AppendChild(smNode);

            return shapeRootNode;
        }

        #endregion
    }
}
