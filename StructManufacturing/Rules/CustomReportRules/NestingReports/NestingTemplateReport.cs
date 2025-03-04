/***************************************************************************************************
Copyright (c) 2017 Hexagon AB and/or its subsidiaries and affiliates. All rights reserved.

File:
  NestingTemplateReport.cs

Author:
  Nautilus-HSV

Description:
  Geneate Template LoopBack xml files to test MfgNestingService         

***************************************************************************************************/


using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Content.Manufacturing.Services;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Manufacturing.Middle.Services;
using Ingr.SP3D.Planning.Middle;
using Ingr.SP3D.Structure.Middle;
using System.Xml;
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// Generate Nesting LoopBack XML 
    /// </summary>
    public class NestingTemplateReport : ManufacturingProfileCustomReportRuleBase
    {

        #region Public Methods

        /// <summary>
        /// Generates the specified entities.
        /// </summary>
        /// <param name="entities">The entities.</param>
        /// <param name="filePath">The file path.</param>     
        public override void Generate(IEnumerable<BusinessObject> entities, string filePath)
        {

            try
            {

                if (entities == null || entities.Count() == 0) return;

                // Report for each selected object.
                foreach (BusinessObject businessObject in entities)
                {

                    if (businessObject is ManufacturingPlate || businessObject is ManufacturingProfile
                        || businessObject is PlatePartBase || businessObject is ProfilePart)
                    {
                        ManufacturingOutputBase manufacturingBase = businessObject as ManufacturingOutputBase;
                        IManufacturable detailedPart = manufacturingBase != null ? manufacturingBase.DetailedPart : (IManufacturable)businessObject;
                        try
                        {
                            ReadOnlyCollection<NestData> nestDataCollection = GetNestData(detailedPart);
                            SaveTemplateDocument(detailedPart, nestDataCollection, "Import", "CONFIRMED", filePath);
                            SaveTemplateDocument(detailedPart, nestDataCollection, "Nesting", "NESTED", filePath);
                            SaveTemplateDocument(detailedPart, nestDataCollection, "Delete", "", filePath);
                        }
                        catch { continue; }
                    }
                    else if (businessObject is AssemblyBase)
                    {
                        AssemblyBase assembly = (AssemblyBase)businessObject;
                        ReadOnlyCollection<IAssemblyChild> plateParts = assembly.GetAssemblyChildren("IJPlatePart", true);
                        ReadOnlyCollection<IAssemblyChild> profileParts = assembly.GetAssemblyChildren("IJProfilePart", true);

                        Tuple<XmlDocument, XmlElement> nestingImport = GetNestingDocument();
                        Tuple<XmlDocument, XmlElement> nestingNesting = GetNestingDocument();
                        Tuple<XmlDocument, XmlElement> nestingDelete = GetNestingDocument();

                        foreach (IAssemblyChild platePart in plateParts)
                        {
                            IManufacturable detailedPart = (IManufacturable)platePart;
                            try
                            {
                                ReadOnlyCollection<NestData> nestDataCollection = GetNestData(detailedPart);
                                GenerateTemplateDocument(nestingImport.Item1, nestingImport.Item2, detailedPart, nestDataCollection, "Import", "CONFIRMED");
                                GenerateTemplateDocument(nestingNesting.Item1, nestingNesting.Item2, detailedPart, nestDataCollection, "Nesting", "NESTED");
                                GenerateTemplateDocument(nestingDelete.Item1, nestingDelete.Item2, detailedPart, nestDataCollection, "Delete", "");
                            }
                            catch { continue; }
                        }

                        foreach (IAssemblyChild profilePart in profileParts)
                        {
                            IManufacturable detailedPart = (IManufacturable)profilePart;
                            try
                            {
                                ReadOnlyCollection<NestData> nestDataCollection = GetNestData(detailedPart);
                                GenerateTemplateDocument(nestingImport.Item1, nestingImport.Item2, detailedPart, nestDataCollection, "Import", "CONFIRMED");
                                GenerateTemplateDocument(nestingNesting.Item1, nestingNesting.Item2, detailedPart, nestDataCollection, "Nesting", "NESTED");
                                GenerateTemplateDocument(nestingDelete.Item1, nestingDelete.Item2, detailedPart, nestDataCollection, "Delete", "");
                            }
                            catch { continue; }
                        }

                        nestingImport.Item1.Save(filePath.Replace(".xml", "_Import.xml"));
                        nestingNesting.Item1.Save(filePath.Replace(".xml", "_Nesting.xml"));
                        nestingDelete.Item1.Save(filePath.Replace(".xml", "_Delete.xml"));
                    }
                }


            }
            catch { /*DO NOTHING*/ }
            finally
            {
            }
        }

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Save Template XML file based on input actionName
        /// </summary>
        private void SaveTemplateDocument(IManufacturable detailedPart, ReadOnlyCollection<NestData> nestDataCollection, string actionName, string outputStatus, string filePath)
        {
            if (detailedPart == null) return;

            XmlDocument nestingDocument = new XmlDocument();

            //Create Root Elment 
            XmlElement rootElement = nestingDocument.CreateElement("NESTING_SERVICES");
            rootElement.SetAttribute("PROJECT_DB_SERVER_NAME", MiddleServiceProvider.SiteMgr.ActiveSite.Server);
            rootElement.SetAttribute("PROJECT_DB_NAME", MiddleServiceProvider.SiteMgr.ActiveSite.Name);
            rootElement.SetAttribute("SHIP_NUMBER", MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.Name);
            rootElement.SetAttribute("PROVIDER", MiddleServiceProvider.SiteMgr.ActiveSite.DBProvider);
            nestingDocument.AppendChild(rootElement);

            if (nestDataCollection.Count == 0 && detailedPart is PlatePartBase)
            {
                XmlElement actionNode = nestingDocument.CreateElement("ACTION");
                actionNode.SetAttribute("NAME", actionName);
                GeneratePartNode(nestingDocument, actionNode, (BusinessObject)detailedPart, null);
                GeneratePropertyNodes(nestingDocument, actionNode, (BusinessObject)detailedPart, null, actionName, outputStatus);
                rootElement.AppendChild(actionNode);
            }
            else
            {
                foreach (NestData nestData in nestDataCollection)
                {
                    XmlElement actionNode = nestingDocument.CreateElement("ACTION");
                    actionNode.SetAttribute("NAME", actionName);
                    GeneratePartNode(nestingDocument, actionNode, (BusinessObject)detailedPart, nestData);
                    GeneratePropertyNodes(nestingDocument, actionNode, (BusinessObject)detailedPart, nestData, actionName, outputStatus);
                    rootElement.AppendChild(actionNode);
                }
            }
            nestingDocument.Save(filePath.Replace(".xml", "_" + actionName + ".xml"));
        }

        /// <summary>
        /// Add Action Node to Template document file based on input actionName
        /// </summary>
        private void GenerateTemplateDocument(XmlDocument nestingDocument, XmlElement rootElement, IManufacturable detailedPart, ReadOnlyCollection<NestData> nestDataCollection, string actionName, string outputStatus)
        {
            if (detailedPart == null) return;

            if (nestDataCollection.Count == 0 && detailedPart is PlatePartBase)
            {
                XmlElement actionNode = nestingDocument.CreateElement("ACTION");
                actionNode.SetAttribute("NAME", actionName);
                GeneratePartNode(nestingDocument, actionNode, (BusinessObject)detailedPart, null);
                GeneratePropertyNodes(nestingDocument, actionNode, (BusinessObject)detailedPart, null, actionName, outputStatus);
                rootElement.AppendChild(actionNode);
            }
            else
            {
                foreach (NestData nestData in nestDataCollection)
                {
                    XmlElement actionNode = nestingDocument.CreateElement("ACTION");
                    actionNode.SetAttribute("NAME", actionName);
                    GeneratePartNode(nestingDocument, actionNode, (BusinessObject)detailedPart, nestData);
                    GeneratePropertyNodes(nestingDocument, actionNode, (BusinessObject)detailedPart, nestData, actionName, outputStatus);
                    rootElement.AppendChild(actionNode);
                }
            }
        }

        private Tuple<XmlDocument, XmlElement> GetNestingDocument()
        {
            XmlDocument nestingDocument = new XmlDocument();
            XmlElement rootElement = nestingDocument.CreateElement("NESTING_SERVICES");
            rootElement.SetAttribute("PROJECT_DB_SERVER_NAME", MiddleServiceProvider.SiteMgr.ActiveSite.Server);
            rootElement.SetAttribute("PROJECT_DB_NAME", MiddleServiceProvider.SiteMgr.ActiveSite.Name);
            rootElement.SetAttribute("SHIP_NUMBER", MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.Name);
            rootElement.SetAttribute("PROVIDER", MiddleServiceProvider.SiteMgr.ActiveSite.DBProvider);
            nestingDocument.AppendChild(rootElement);

            return new Tuple<XmlDocument, XmlElement>(nestingDocument, rootElement);

        }

        private ReadOnlyCollection<NestData> GetNestData(IManufacturable detailedPart)
        {

            List<NestData> nestDataList = new List<NestData>();

            ManufacturingOutputBase manufacturingPart = null;

            try
            {
                Collection<ManufacturingBase> manufacturingParts;
                if (detailedPart is PlatePartBase)
                {
                    manufacturingParts = EntityService.GetManufacturingEntity((BusinessObject)detailedPart, ManufacturingEntityType.ManufacturingPlate);
                    if (manufacturingParts.Count > 0)
                        manufacturingPart = manufacturingParts[0] as ManufacturingOutputBase ?? null;

                    NestData nestData = null;
                    if (manufacturingPart != null)
                        nestData = manufacturingPart.NestingInformation(false, "") ?? null;
                    if (nestData != null) nestDataList.Add(nestData);
                }
                else
                {
                    manufacturingParts = EntityService.GetManufacturingEntity((BusinessObject)detailedPart, ManufacturingEntityType.ManufacturingProfile);
                    if (manufacturingParts.Count > 0)
                        manufacturingPart = manufacturingParts[0] as ManufacturingOutputBase ?? null;

                    NestData webNestData = null; NestData topFlangeNestData = null;
                    if (manufacturingPart != null)
                    {
                        webNestData = manufacturingPart.NestingInformation("PLATE,WEB") ?? null;
                        topFlangeNestData = manufacturingPart.NestingInformation("PLATE,TOP_FLANGE") ?? null;
                    }
                    if (webNestData != null) nestDataList.Add(webNestData);
                    if (topFlangeNestData != null) nestDataList.Add(topFlangeNestData);


                }
            }
            catch (Exception) { }

            return new ReadOnlyCollection<NestData>(nestDataList);

        }

        private void GeneratePartNode(XmlDocument nestingDocument, XmlElement parentElement, BusinessObject detailedPart, NestData nestData)
        {

            PlatePartBase platePart = detailedPart as PlatePartBase;
            StiffenerPart stiffenerPart = detailedPart as StiffenerPart;
            string partType = platePart != null ? "PLATE" : "";
            partType = nestData != null ? nestData.PartType : partType;

            if(partType.Contains(","))
            {
                //Get second part of partType 
                string[] cellName = partType.Split(',');
                if(cellName.Length> 0)
                    partType = cellName[cellName.Length-1];
            }
           
            XmlElement partNode = nestingDocument.CreateElement("PART");

            partNode.SetAttribute("MODEL_PART_GUID", detailedPart.ObjectIDForQuery + "-" + partType);
            partNode.SetAttribute("BLOCK_NAME", "");
            partNode.SetAttribute("PART_NAME", detailedPart.ToString());
            partNode.SetAttribute("NESTING_GUID", nestData != null ? nestData.ObjectIDForQuery : string.Empty);

            BoardManagementSymmetry boardInformation = platePart != null ? platePart.Symmetry : BoardManagementSymmetry.NotSet;
            if (stiffenerPart != null) boardInformation = stiffenerPart.Symmetry;

            string board = "";
            switch (boardInformation)
            {
                case BoardManagementSymmetry.Centered:
                    board = "C"; break;
                case BoardManagementSymmetry.StarboardOnly:
                    board = "S"; break;
                default:
                    board = "P"; break;
            }
            partNode.SetAttribute("PART_BOARDSIDE", board);
            parentElement.AppendChild(partNode);

        }
        private void GeneratePropertyNodes(XmlDocument nestingDocument, XmlElement parentElement, BusinessObject detailedPart, NestData nestData, string actionName, string outputStatus)
        {

            PlatePartBase platePart = detailedPart as PlatePartBase;
            StiffenerPart stiffenerPart = detailedPart as StiffenerPart;
            ProfileReportService profileReportService = null;

            try
            {
                if (actionName.Equals("Delete")) return;

                // <PROPERTY NAME="ExportTimestamp" TYPE="Date" VALUE="2017-07-23T14:52:00Z" />
                //ExportTeimeStamp
                GeneratePropertyNode(nestingDocument, parentElement, "ExportTimestamp", "Date", nestData != null ? nestData.ExportTimestamp.ToString() : DateTime.Now.ToString());
                GeneratePropertyNode(nestingDocument, parentElement, "RegistrationTimestamp", "Date", nestData != null ? nestData.RegistrationTimestamp.ToString() : DateTime.Now.ToString());
                GeneratePropertyNode(nestingDocument, parentElement, "NestingTimestamp", "Date", nestData != null ? nestData.NestingTimestamp.ToString() : DateTime.Now.ToString());

                //LotNumber
                string lotNumber = "";
                string lotMaterialType = "";
                string partType="";
                if (platePart != null)
                {
                    partType = nestData != null ? nestData.PartType : "PLATE";
                    double thickness=0.0;
                    if (platePart is StandAlonePlatePart) thickness = ((StandAlonePlatePart)platePart).Thickness;
                    else if (platePart is CollarPart) thickness = ((CollarPart)platePart).Thickness;
                    else if (platePart is PlatePart) thickness = ((PlatePart)platePart).Thickness;

                    lotNumber = platePart.MaterialGrade + "_" + thickness.ToString();
                    lotMaterialType = platePart.MaterialType;
                }
                else if (stiffenerPart != null)
                {
                    profileReportService = base.GetProfileReportingService(stiffenerPart);
                    lotNumber = stiffenerPart.MaterialGrade + "_" + profileReportService.GetProfileWebThickness().ToString();
                    partType = nestData != null ? nestData.PartType : "PROFILE";
                    lotMaterialType = stiffenerPart.MaterialType;
                }
                GeneratePropertyNode(nestingDocument, parentElement, "LotNumber", "String", nestData != null ? nestData.LotNumber : lotNumber);
                GeneratePropertyNode(nestingDocument, parentElement, "PartRegistredNumber", "Int", nestData != null ? nestData.PartRegisteredNumber.ToString() : "0");
                GeneratePropertyNode(nestingDocument, parentElement, "OrderNumber", "String", nestData != null ? nestData.OrderNumber.ToString() : "");

                //LotMaterialType
                GeneratePropertyNode(nestingDocument, parentElement, "LotMaterialType", "String", nestData != null ? nestData.LotMaterialType : lotMaterialType);

                //LotMaterialGrade
                string lotMaterialGrade = platePart != null ? platePart.MaterialGrade : "";
                if (stiffenerPart != null) lotMaterialGrade = stiffenerPart.MaterialGrade;

                GeneratePropertyNode(nestingDocument, parentElement, "LotMaterialGrade", "String", nestData != null ? nestData.LotMaterialGrade : lotMaterialGrade);
                GeneratePropertyNode(nestingDocument, parentElement, "LotLength", "Double", nestData != null ? nestData.LotLength.ToString() : "");
                GeneratePropertyNode(nestingDocument, parentElement, "LotWidth", "Double", nestData != null ? nestData.LotWidth.ToString() : "");
                GeneratePropertyNode(nestingDocument, parentElement, "LotThickness", "Double", nestData != null ? nestData.LotThickness.ToString() : "");
                GeneratePropertyNode(nestingDocument, parentElement, "PartType", "String", nestData != null ? nestData.PartType : partType);
                GeneratePropertyNode(nestingDocument, parentElement, "OutputType", "String", nestData != null ? nestData.OutputType : "");
                GeneratePropertyNode(nestingDocument, parentElement, "OutputStatus", "String", outputStatus);
                GeneratePropertyNode(nestingDocument, parentElement, "AssemblyPath", "String", nestData != null ? nestData.AssemblyPath : "");
                GeneratePropertyNode(nestingDocument, parentElement, "PartGUID", "String", nestData != null ? nestData.PartGuid : detailedPart.ObjectIDForQuery);
                GeneratePropertyNode(nestingDocument, parentElement, "NestingSystem", "String", nestData != null ? nestData.NestingSystem : "");
                GeneratePropertyNode(nestingDocument, parentElement, "Routing", "String", nestData != null ? nestData.Routing : "");
                GeneratePropertyNode(nestingDocument, parentElement, "Reference1", "String", nestData != null ? nestData.Reference1 : "");
                GeneratePropertyNode(nestingDocument, parentElement, "Reference2", "String", nestData != null ? nestData.Reference2 : "");
                GeneratePropertyNode(nestingDocument, parentElement, "Reference3", "String", nestData != null ? nestData.Reference3 : "");
                GeneratePropertyNode(nestingDocument, parentElement, "NumberConfirmed", "Int", nestData != null ? nestData.NumberConfirmed.ToString() : "0");
                GeneratePropertyNode(nestingDocument, parentElement, "NumberNested", "Int", nestData != null ? nestData.NumberNested.ToString() : "0");
            }
            catch (Exception) { return; }
            finally
            {
                if (profileReportService != null) profileReportService.Dispose();
            }

        }

        private void GeneratePropertyNode(XmlDocument nestingDocument, XmlElement parentElement, string name, string type, string value)
        {
            XmlElement propertyNode = nestingDocument.CreateElement("PROPERTY");
            propertyNode.SetAttribute("NAME", name);
            propertyNode.SetAttribute("TYPE", type);
            propertyNode.SetAttribute("VALUE", value);
            parentElement.AppendChild(propertyNode);
        }

        #endregion Private Methods
    }
}
