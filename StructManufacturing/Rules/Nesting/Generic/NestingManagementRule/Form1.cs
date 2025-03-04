using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Schema;
using System.Collections.ObjectModel;
using System.Threading;
using System.Reflection;
using Microsoft.Win32;
using System.IO;
using Ingr.SP3D.Common.Middle.Services.Hidden;

using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Manufacturing.Middle.Services;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// The Form.
    /// </summary>
    public partial class Form1 : Form
    {
        /// <summary>
        /// Initializes a new instance of the form.
        /// </summary>
        /// <param name="inputXML">The input xml.</param>
        /// <param name="outputXML">The output xml.</param>
        public Form1(string inputXML, out string outputXML)
        {
            NestingService nestingService = new NestingService();
            nestingService.Run(inputXML, out outputXML);
            //InitializeComponent();
            this.Visible = false;
            this.Close();
        }
    }

    /// <summary>
    /// The nesting service class.
    /// </summary>
    public class NestingService: NestingServiceBase
    {
        const long MyBaseExitCode = -10000;

        private MetadataManager m_MetaDataManager;
        private SiteManager siteMgr;
        private Plant activePlant;
        private Model modelDB;
        private Catalog catalogDB;
        private Site site;
        private string strStartTime = DateTime.Now.ToLongTimeString().Replace(":", "-");//   Replace(Format(Now, "s"), ":", "-") & "-" & strUser

        /// <summary>
        /// Initializes a new instance of the nesting service class.
        /// </summary>
        public NestingService() {
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(MyResolveEventHandler);
        }

        /// <summary>
        /// Runs the nesting service.
        /// </summary>
        /// <param name="inputXML">The input xml.</param>
        /// <param name="outputXML">The output xml.</param>
        public void Run(string inputXML, out string outputXML)
        {
            //System.Windows.Forms.MessageBox.Show(inputXML);
            outputXML = "";
            try
            {
                EnsureLogFolderExists();
                CommonLogInfo("Nesting Service called", strStartTime + "-MfgNestingService");

                XmlDocument inputXmlDocument;
                siteMgr = MiddleServiceProvider.SiteMgr;
                inputXmlDocument = new XmlDocument();

                inputXmlDocument.LoadXml(inputXML);
                ConnectToModel(inputXmlDocument);
                string sharedContentPath = catalogDB.SymbolShare;
                if (File.Exists(sharedContentPath + "\\StructManufacturing\\SMS_SCHEMA\\NestingReport\\NestingReport.xsd"))
                {
                    inputXmlDocument.Schemas.Add(null, sharedContentPath + "\\StructManufacturing\\SMS_SCHEMA\\NestingReport\\NestingReport.xsd");
                    ValidationEventHandler eventHandler = new ValidationEventHandler(ValidationEventHandler);
                    inputXmlDocument.Validate(eventHandler);
                }

                ProcessInputXML(inputXmlDocument);
                outputXML = inputXmlDocument.OuterXml;

                //MiddleServiceProvider.TransactionMgr.Commit("Update Nesting Information");
                m_MetaDataManager = null;
                MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.Dispose();
                MiddleServiceProvider.SiteMgr.ActiveSite.Dispose();
                MiddleServiceProvider.Cleanup();
                return;
            }
            catch (Exception ex)
            {
                CommonLogExceptionInfo(ex, strStartTime + "-MfgNestingServiceError");
            }
        }

        private void ProcessInputXML(XmlDocument inputXmlDocument)
        {
            try
            {
                XmlNodeList actionNodeList = inputXmlDocument.SelectNodes("/NESTING_SERVICES/ACTION");
                for (int i = 0; i < actionNodeList.Count; i++)
                {
                    XmlNode actionNode = actionNodeList.Item(i);
                    if (actionNode.Attributes.GetNamedItem("NAME").Value == "Confirmation") Confirmation(actionNode, inputXmlDocument);
                    if (actionNode.Attributes.GetNamedItem("NAME").Value == "Import") Confirmation(actionNode, inputXmlDocument);
                    if (actionNode.Attributes.GetNamedItem("NAME").Value == "Nesting") Nesting(actionNode, inputXmlDocument);
                    if (actionNode.Attributes.GetNamedItem("NAME").Value == "Delete") Delete(actionNode, inputXmlDocument);
                    if (actionNode.Attributes.GetNamedItem("NAME").Value == "GetParts") GetParts(actionNode, inputXmlDocument);
                    if (actionNode.Attributes.GetNamedItem("NAME").Value == "GetStatus") GetStatus(actionNode, inputXmlDocument);
                    if (actionNode.Attributes.GetNamedItem("NAME").Value == "GetXML") GetXML(actionNode, inputXmlDocument);
                    if (actionNode.Attributes.GetNamedItem("NAME").Value == "SetProperties") SetProperties(actionNode, inputXmlDocument);
                    if (actionNode.Attributes.GetNamedItem("NAME").Value == "Reset") Reset(actionNode, inputXmlDocument);
                    if (actionNode.Attributes.GetNamedItem("NAME").Value == "Update") Update(actionNode, inputXmlDocument);
                    if (actionNode.Attributes.GetNamedItem("NAME").Value == "GetDBInformation") GetDBInformation(actionNode, inputXmlDocument);
                    if (actionNode.Attributes.GetNamedItem("NAME").Value == "GetProperties") { }//GetProperties(actionNode, inputXmlDocument);
                    if (actionNode.Attributes.GetNamedItem("NAME").Value == "NestingCommand") { GetRemnantReport(actionNode, inputXmlDocument);}
                    actionNode.ParentNode.RemoveChild(actionNode);
                    MiddleServiceProvider.TransactionMgr.Commit("Update Nesting Information");

                }
            }
            catch (Exception ex)
            {
                CommonLogExceptionInfo(ex, strStartTime + "-MfgNestingServiceError");
            }
        }

        private void GetRemnantReport(XmlNode actionNode, XmlDocument inputXmlDocument)
        {
            CommonLogInfo("Getting Nesting System Information", strStartTime + "-MfgNestingService");
            if (actionNode.Attributes.GetNamedItem("COMMAND") != null)
            {
                string commandName = actionNode.Attributes.GetNamedItem("COMMAND").Value;
                string nestingOutput = CallNestingSystem(commandName, modelDB.Name, "");
                XmlDocument nestingOutputDocument = new XmlDocument();
                try
                {
                    nestingOutputDocument.LoadXml(nestingOutput);
                    XmlElement nestingReportElement = nestingOutputDocument.DocumentElement;
                    XmlNode importedNode = inputXmlDocument.ImportNode(nestingReportElement, true);
                    inputXmlDocument.DocumentElement.AppendChild(importedNode);
                }
                catch (Exception ex)
                {
                    CommonLogExceptionInfo(ex, strStartTime + "-MfgNestingServiceError");
                }
            }
            else {
                CommonLogInfo("Missing COMMAND attribute while getting Nesting System Information", strStartTime + "-MfgNestingService");
            }
        }

        private void GetDBInformation(XmlNode actionNode, XmlDocument inputXmlDocument){
            //Get the DB connections 
            try
            {
                CommonLogInfo("Getting DB Information", strStartTime + "-MfgNestingService");
                siteMgr = MiddleServiceProvider.SiteMgr;
                //site = MiddleServiceProvider.SiteMgr.ConnectSite();
                if (!(site == null))
                {
                    for (int i = 0; i < siteMgr.Sites.Count; i++)
                    {
                        XmlElement siteNode = inputXmlDocument.CreateElement("SITE");
                        siteNode.SetAttribute("NAME", siteMgr.Sites[i].Name);
                        siteNode.SetAttribute("SERVER_NAME", siteMgr.Sites[i].Server.ToString());
						siteNode.SetAttribute("PROVIDER", siteMgr.Sites[i].DBProvider);
                        inputXmlDocument.DocumentElement.AppendChild(siteNode);
                        site = siteMgr.Sites[i];
                        for (int j = 0; j < site.Plants.Count; j++)
                        {
                            try
                            {
                                int refCount = site.Plants[j].References.Count;
                                XmlElement modelNode = inputXmlDocument.CreateElement("MODEL");
                                site.OpenPlant(site.Plants[j]);
                                activePlant = siteMgr.ActiveSite.ActivePlant;
                                modelDB = activePlant.PlantModel;
                                catalogDB = activePlant.PlantCatalog;
                                modelNode.SetAttribute("NAME", site.Plants[j].Name);
                                modelNode.SetAttribute("SYMBOL_SHARE", catalogDB.SymbolShare);
                                siteNode.AppendChild(modelNode);
                            }
                            catch (Exception ex)
                            {
                                CommonLogExceptionInfo(ex, strStartTime + "-MfgNestingServiceError");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                CommonLogExceptionInfo(ex, strStartTime + "-MfgNestingServiceError");
            }
        }

        private void SetAttributeValue(BusinessObject bo, string propertyName, string interfaceName, string propertyType, string propertyValue) {
            try {
                if (propertyName == "OutputStatus")
                {
                    int outputStatusValue = 0;
                    switch (propertyValue)
                    {
                        case "DELETED":
                        case "FAILED":
                        case "1":
                            outputStatusValue = 1;
                            break;
                        case "OUTOFDDATE":
                        case "UPTODATE":
                        case "CONFIRMED":
                        case "2":
                            outputStatusValue = 2;
                            break;
                        case "NESTED":
                        case "4":
                            outputStatusValue = 4;
                            break;
                        case "CUT":
                        case "8":
                            outputStatusValue = 8;
                            break;
                        case "WELDED":
                        case "16":
                            outputStatusValue = 16;
                            break;
                    }
                    bo.SetPropertyValue(outputStatusValue, interfaceName, propertyName);
                }
                else if (propertyType == "Date")
                {
                    DateTime dateTime = DateTime.Parse(propertyValue);
                    bo.SetPropertyValue(dateTime, interfaceName, propertyName);
                }
                else if (propertyType == "Int" || propertyType == "CodelistIndex")
                {
                    int intValue = int.Parse(propertyValue);
                    bo.SetPropertyValue(intValue, interfaceName, propertyName);
                }
                else if (propertyType == "Double")
                {
                    double dblValue = double.Parse(propertyValue);
                    bo.SetPropertyValue(dblValue, interfaceName, propertyName);
                }
                else if (propertyType == "CodelistShort")
                {
                    //bo.
                    InterfaceInformation interfaceInformation = m_MetaDataManager.GetInterfaceInfo(interfaceName, "UDP");
                    PropertyInformation propertyInfo = interfaceInformation.GetPropertyInfo(propertyName);
                    CodelistItem codelistItem = propertyInfo.CodeListInfo.GetCodelistItem(propertyValue);
                    int intValue = codelistItem.Value;
                    bo.SetPropertyValue(intValue, interfaceName, propertyName);
                }
                else if (propertyType == "CodelistLong")
                {
                    InterfaceInformation interfaceInformation = m_MetaDataManager.GetInterfaceInfo(interfaceName, "UDP");
                    PropertyInformation propertyInfo = interfaceInformation.GetPropertyInfo(propertyName);
                    ReadOnlyDictionary<CodelistItem> codelistDictionary = propertyInfo.CodeListInfo.GetCodelistItemsForLongString(propertyValue);
                    int intValue = codelistDictionary.Values.ElementAt(0).Value;
                    bo.SetPropertyValue(intValue, interfaceName, propertyName);
                }
                else
                {
                    bo.SetPropertyValue(propertyValue, interfaceName, propertyName);
                }

            }
            catch (Exception ex) {
                CommonLogExceptionInfo(ex, strStartTime + "-MfgNestingServiceError");
            }
        }

        private void Confirmation(XmlNode actionNode, XmlDocument inputXmlDocument)
        {
            try
            {
                CommonLogInfo("Updating NestData for confirmation", strStartTime + "-MfgNestingService");
                ReadOnlyCollection<BusinessObject> oNestDataByFilter;
                oNestDataByFilter = GetParts(actionNode, inputXmlDocument);
                for (int i = 0; i < oNestDataByFilter.Count(); i++)
                {
                    try
                    {
                        BusinessObject bo = oNestDataByFilter[i];

                        CommonLogInfo("Updating NestData for confirmation for NestData: " + bo.ObjectID, strStartTime + "-MfgNestingService");
                        XmlNodeList propertyNodeList = actionNode.SelectNodes(".//PROPERTY");
                        NestData nestData = bo as NestData;
                        if (!(nestData == null))
                        {
                            nestData.NumberConfirmed = 1;
                            for (int j = 0; j < propertyNodeList.Count; j++)
                            {
                                string propertyName = propertyNodeList.Item(j).Attributes.GetNamedItem("NAME").Value;
                                string interfaceName = "";
                                if (propertyNodeList.Item(j).Attributes.GetNamedItem("INTERFACE") != null)
                                {
                                    interfaceName = propertyNodeList.Item(j).Attributes.GetNamedItem("INTERFACE").Value;
                                }
                                if (interfaceName == "") interfaceName = "IJMfgNestData";
                                string propertyType = propertyNodeList.Item(j).Attributes.GetNamedItem("TYPE").Value;
                                string propertyValue = propertyNodeList.Item(j).Attributes.GetNamedItem("VALUE").Value;
                                CommonLogInfo("     Giving the attribute: " + propertyName + " the value: " + propertyValue, strStartTime + "-MfgNestingService");
                                SetAttributeValue(bo, propertyName, interfaceName, propertyType, propertyValue);
                            }
						}
                        //CallNestingRule(nestData, inputXmlDocument.DocumentElement);
                    }
                    catch (Exception ex)
                    {
                        CommonLogExceptionInfo(ex, strStartTime + "-MfgNestingServiceError");
                    }
                }
            }
            catch (Exception ex)
            {
                CommonLogExceptionInfo(ex, strStartTime + "-MfgNestingServiceError");
            }
        }

        private void Nesting(XmlNode actionNode, XmlDocument inputXmlDocument)
        {
            CommonLogInfo("Updating NestData for Nesting", strStartTime + "-MfgNestingService");
            try
            {
                ReadOnlyCollection<BusinessObject> oNestDataByFilter;
                oNestDataByFilter = GetParts(actionNode, inputXmlDocument);
                for (int i = 0; i < oNestDataByFilter.Count(); i++)
                {
                    try
                    {
                        BusinessObject bo = oNestDataByFilter[i];
                        CommonLogInfo("Updating NestData for Nesting for NestData: " + bo.ObjectID, strStartTime + "-MfgNestingService");
                        XmlNodeList propertyNodeList = actionNode.SelectNodes(".//PROPERTY");
                        NestData nestData = bo as NestData;
                        if (!(nestData == null))
                        {
                            nestData.NumberNested = 1;
                            for (int j = 0; j < propertyNodeList.Count; j++)
                            {
                                string propertyName = propertyNodeList.Item(j).Attributes.GetNamedItem("NAME").Value;
                                string interfaceName = propertyNodeList.Item(j).Attributes.GetNamedItem("INTERFACE").Value;
                                string propertyType = propertyNodeList.Item(j).Attributes.GetNamedItem("TYPE").Value;
                                string propertyValue = "";
                                if (propertyNodeList.Item(j).Attributes.GetNamedItem("VALUE") != null)
                                {
                                    propertyValue = propertyNodeList.Item(j).Attributes.GetNamedItem("VALUE").Value;
                                }
                                CommonLogInfo("     Giving the attribute: " + propertyName + " the value: " + propertyValue, strStartTime + "-MfgNestingService");
                                SetAttributeValue(bo, propertyName, interfaceName, propertyType, propertyValue);
                                //CallNestingRule(nestData, inputXmlDocument.DocumentElement);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        CommonLogExceptionInfo(ex, strStartTime + "-MfgNestingServiceError");
                    }
                }
            }
            catch (Exception ex)
            {
                CommonLogExceptionInfo(ex, strStartTime + "-MfgNestingServiceError");
            }
        }

        private ReadOnlyCollection<BusinessObject> GetParts(XmlNode actionNode, XmlDocument inputXmlDocument)
        {
            CommonLogInfo("Getting Parts", strStartTime + "-MfgNestingService");
            try
            {
                XmlNode xmlFilterNode = actionNode;
                ReadOnlyCollection<BusinessObject> oNestDataByFilter;
                if (xmlFilterNode != null)
                {
                    if (xmlFilterNode.Name == "ACTION")
                    {
                        oNestDataByFilter = GetPartsByFilter(xmlFilterNode, inputXmlDocument);
                        //xmlFilterNode.ParentNode.RemoveChild(xmlFilterNode);
                    }
                    else
                    {
                        oNestDataByFilter = new ReadOnlyCollection<BusinessObject>(new List<BusinessObject>());
                    }
                }
                else
                {
                    oNestDataByFilter = new ReadOnlyCollection<BusinessObject>(new List<BusinessObject>());
                }
                #region "Add objects to XML"
                if (oNestDataByFilter != null)
                {
                    for (int i = 0; i < oNestDataByFilter.Count; i++)
                    {
                        NestData nestData = (oNestDataByFilter[i] as Object) as NestData;
                        ManufacturingOutputBase moBase = (oNestDataByFilter[i] as Object) as ManufacturingOutputBase;
                        //PlatePart
                        IManufacturable detailedPart = (oNestDataByFilter[i] as IManufacturable);
                        string nestOid = "";
                        string mfgPartOid = "";
                        string mfgPartName = "";
                        string detailedPartOid = "";
                        string lotName = "";
                        string exportTimestamp = "";
                        string registrationTimestamp = "";
                        if (nestData != null)
                        {
                            moBase = nestData.ManufacturingPart as ManufacturingOutputBase;
                            //if (moBase != null)
                            //{
                            //    detailedPart = moBase.DetailedPart as IManufacturable;
                            //}
                            //else
                            //{
                            //    detailedPart = modelDB.WrapSP3DBO(modelDB.GetBOMonikerFromDbIdentifier(nestData.PartGuid)) as IManufacturable;
                            //}

                        }
                        else if (moBase != null)
                        {
                            nestData = moBase.NestingInformation();
                            //detailedPart = moBase.DetailedPart as IManufacturable;
                        }
                        else if (detailedPart != null)
                        {
                            moBase = detailedPart.Manufacture() as ManufacturingOutputBase;
                            if (moBase != null)
                            {
                                nestData = moBase.NestingInformation() as NestData;
                            }
                            else
                            {
                                if ((oNestDataByFilter[i] as PlatePartBase) != null)
                                {
                                    nestData = new NestData((PlatePart)detailedPart);
                                }
                                else if ((oNestDataByFilter[i] as ProfilePart) != null)
                                {
                                    nestData = new NestData((ProfilePart)detailedPart);
                                }
                            }
                        }

                        if (moBase != null)
                        {
                            mfgPartOid = moBase.ObjectID;
                            mfgPartName = moBase.Name;
                        }

                        if (detailedPart != null)
                        {
                            detailedPartOid = ((BusinessObject)detailedPart).ObjectID;
                        }

                        if (nestData != null)
                        {
                            nestOid = nestData.ObjectID;
                            detailedPartOid = nestData.PartGuid;
                            lotName = nestData.LotNumber;
                            exportTimestamp = nestData.ExportTimestamp.ToLongDateString();
                            registrationTimestamp = nestData.RegistrationTimestamp.ToLongDateString();
                        }

                        if (moBase != null || detailedPart != null || nestData != null)
                        {

                            XmlNode nestObjectNode = inputXmlDocument.SelectSingleNode("//PART[@GUID='" + nestOid + "']");
                            XmlElement rootElement = (XmlElement)inputXmlDocument.SelectSingleNode("/NESTING_SERVICES");
                            if (nestObjectNode == null)
                            {
                                nestObjectNode = inputXmlDocument.CreateElement("PART");
                            }
                            XmlAttribute partGuidAttribute = inputXmlDocument.CreateAttribute("PART_GUID");
                            XmlAttribute nestGuidAttribute = inputXmlDocument.CreateAttribute("NEST_GUID");
                            XmlAttribute detailedGuidAttribute = inputXmlDocument.CreateAttribute("DETAILED_GUID");
                            XmlAttribute mfgNameAttribute = inputXmlDocument.CreateAttribute("MFG_NAME");
                            XmlAttribute lotNameAttribute = inputXmlDocument.CreateAttribute("LOT_NAME");
                            XmlAttribute exportTimeStampAttribute = inputXmlDocument.CreateAttribute("EXPORT_TIMESTAMP");
                            XmlAttribute registrationTimeStampAttribute = inputXmlDocument.CreateAttribute("REGISTRATION_TIMESTAMP");
                            partGuidAttribute.Value = mfgPartOid;
                            nestGuidAttribute.Value = nestOid;
                            detailedGuidAttribute.Value = detailedPartOid;
                            mfgNameAttribute.Value = mfgPartName;
                            lotNameAttribute.Value = lotName;
                            exportTimeStampAttribute.Value = exportTimestamp;
                            registrationTimeStampAttribute.Value = registrationTimestamp;
                            nestObjectNode.Attributes.Append(partGuidAttribute);
                            nestObjectNode.Attributes.Append(nestGuidAttribute);
                            nestObjectNode.Attributes.Append(detailedGuidAttribute);
                            nestObjectNode.Attributes.Append(mfgNameAttribute);
                            nestObjectNode.Attributes.Append(lotNameAttribute);
                            nestObjectNode.Attributes.Append(exportTimeStampAttribute);
                            nestObjectNode.Attributes.Append(registrationTimeStampAttribute);
                            rootElement.AppendChild(nestObjectNode);
                        }
                        moBase = null;
                        detailedPart = null;
                        nestData = null;
                    }
                }
                XmlNodeList partList = actionNode.SelectNodes("./PART");
                Collection<BusinessObject> boCol = new Collection<BusinessObject>(new List<BusinessObject>());
                for (int i = 0; i < partList.Count; i++)
                {
                    BusinessObject bo = null;
                    XmlNode partAttribute = partList.Item(i).Attributes.GetNamedItem("NESTING_GUID");
                    if (partAttribute != null)
                    {
                        if (partAttribute.Value != "")
                        {
                            bo = modelDB.WrapSP3DBO(modelDB.GetBOMonikerFromDbIdentifier(partAttribute.Value));
                            try
                            {
                                NestData nd = bo as NestData;
                                if (!(nd == null)){
                                    string assemblyPath = nd.AssemblyPath;
                                }
                            }
                            catch { bo = null; }
                        }
                    }
                    if (bo == null)
                    {
                        partAttribute = partList.Item(i).Attributes.GetNamedItem("PART_GUID");
                        XmlNode partAttributeModel = partList.Item(i).Attributes.GetNamedItem("MODEL_PART_GUID");

                        string modelPartGuid = null;
                        string modelPartType = null;
                        if (partAttributeModel != null)
                        {
                            modelPartGuid = partAttributeModel.Value;
                            if (modelPartGuid.Contains("}-")) {
                                modelPartType = modelPartGuid.Substring(39);
                                modelPartGuid = modelPartGuid.Substring(0, 38);
                            }
                        } 
                        ManufacturingOutputBase moBase = null;
                        if (partAttribute != null && partAttribute.Value != null)
                        {
                            BusinessObject mfgBO = modelDB.WrapSP3DBO(modelDB.GetBOMonikerFromDbIdentifier(partAttribute.Value));
                            moBase = mfgBO as ManufacturingOutputBase;
                        }
                        BusinessObject modelBO = null;
                        if (partAttributeModel != null && modelPartGuid !=null)
                        {
                            modelBO = modelDB.WrapSP3DBO(modelDB.GetBOMonikerFromDbIdentifier(modelPartGuid));
                        }
                        //if (moBase == null && modelBO != null)
                        //{ 
                        //    IManufacturable imModelBO;
                        //    imModelBO = (IManufacturable)modelBO;
                        //    BusinessObject mfgBO = Ingr.SP3D.Manufacturing.Middle.Services.EntityService.NavigateToManufacturingPartByRelation(modelBO);
                        //    moBase = mfgBO as ManufacturingOutputBase;
                        //}
                        PlatePartBase platePart = modelBO as PlatePartBase;
                        ProfilePart stiffenerPart = modelBO as ProfilePart;
                        IDetailable detailPlatePart = modelBO as IDetailable;
                        
                        if (detailPlatePart != null && moBase != null)
                        {
                            try
                            {
                                string name = moBase.Name;
                                XmlNode profileType = partList.Item(i).Attributes.GetNamedItem("PART_TYPE");
                                if (profileType != null)
                                {
                                    bo = moBase.NestingInformation(detailPlatePart, profileType.Value, true, "");
                                }
                                else
                                {
                                    if (platePart != null)
                                    {
                                        bo = moBase.NestingInformation(detailPlatePart, "PLATE", true, "");
                                    }
                                    else if (stiffenerPart != null)
                                    {
                                        bo = moBase.NestingInformation(detailPlatePart, "PROFILE", true, "");
                                    }
                                }
                            }
                            catch { }
                        }
                        if (bo == null && platePart != null)
                        {
                            bo = (BusinessObject)new NestData(platePart);
                        }
                        if (bo == null && stiffenerPart != null)
                        {
                            XmlNode profileType = partList.Item(i).Attributes.GetNamedItem("PART_TYPE");
                            if (profileType != null)
                            {
                                bo = (BusinessObject)new NestData(stiffenerPart, profileType.Value);
                            }
                            else if (modelPartType != null)
                            {
                                bo = (BusinessObject)new NestData(stiffenerPart, modelPartType);
                            }
                            else
                            {
                                bo = (BusinessObject)new NestData(stiffenerPart, "PROFILE");
                            }
                        }
                        if (bo == null && moBase != null)
                        {
                            bo = moBase.NestingInformation();
                        }
                    }
                    if (bo != null)
                    {
                        boCol.Add(bo);
                    }
                }
                return new ReadOnlyCollection<BusinessObject>(boCol);
            }
            catch (Exception ex)
            {
                CommonLogExceptionInfo(ex, strStartTime + "-MfgNestingServiceError");
            }
            return new ReadOnlyCollection<BusinessObject>(new List<BusinessObject>());
        }

        private void Reset(XmlNode actionNode, XmlDocument inputXmlDocument)
        {
            try
            {
                CommonLogInfo("Updating NestData for Reset", strStartTime + "-MfgNestingService");
                ReadOnlyCollection<BusinessObject> oNestDataByFilter;
                oNestDataByFilter = GetParts(actionNode, inputXmlDocument);
                for (int i = 0; i < oNestDataByFilter.Count; i++)
                {
                    BusinessObject businessObject = oNestDataByFilter[i];
                    if (businessObject.SupportsInterface("IJMfgNestData"))
                    {
                        try
                        {
                            NestData nestData = (NestData)businessObject;
                            CommonLogInfo("Resetting NestData for NestData: " + nestData.ObjectID, strStartTime + "-MfgNestingService");
                            nestData.OrderNumber = "";
                            nestData.LotNumber = "";
                            nestData.LotMaterialType = "";
                            nestData.LotMaterialGrade = "";
                            nestData.OutputStatus = 1;

                            nestData.NestingTimestamp = new DateTime(1900, 1, 1);
                            nestData.PartRegisteredNumber = 0;
                            nestData.LotLength = 0.0;
                            nestData.LotWidth = 0.0;
                            nestData.LotThickness = 0.0;
                            nestData.NumberNested = 0;
                            //CallNestingRule(nestData, inputXmlDocument.DocumentElement);
                        }
                        catch (Exception ex) {
                            CommonLogExceptionInfo(ex, strStartTime + "-MfgNestingServiceError");                        
                        }
                    }
                    businessObject = null;
                }
            }
            catch (Exception ex)
            {
                CommonLogExceptionInfo(ex, strStartTime + "-MfgNestingServiceError");
            }
        }

        private void Delete(XmlNode actionNode, XmlDocument inputXmlDocument)
        {
            CommonLogInfo("Deleting NestData", strStartTime + "-MfgNestingService");
            try
            {
                ReadOnlyCollection<BusinessObject> oNestDataByFilter;
                oNestDataByFilter = GetParts(actionNode, inputXmlDocument);
                for (int i = 0; i < oNestDataByFilter.Count; i++)
                {
                    BusinessObject businessObject = oNestDataByFilter[i];
                    if (businessObject.SupportsInterface("IJMfgNestData"))
                    {
                        try
                        {
                            NestData pNestData = businessObject as NestData;
                            string outputPath = String.Empty;

                            if (pNestData != null)
                                outputPath = base.NestingFilePath(pNestData.OutputType);

                            if (outputPath != null && outputPath.Length > 0)
                            {
                                XmlDocument outputXmlDocument = new XmlDocument();
                                XmlElement rootElement = outputXmlDocument.CreateElement("IS2NEST_DELETE");
                                rootElement.SetAttribute("PROJECT_DB_SERVER_NAME", site.DatabaseID);
                                rootElement.SetAttribute("PROJECT_DB_NAME", site.Name);
                                rootElement.SetAttribute("SHIP_NUMBER", modelDB.Name);
                                XmlElement platesElement = outputXmlDocument.CreateElement("PLATES");

                                outputXmlDocument.AppendChild(rootElement);
                                rootElement.AppendChild(platesElement);
                                XmlElement plateElement = outputXmlDocument.CreateElement("PLATE");
                                platesElement.AppendChild(plateElement);
                                string partName = "";
                                BusinessObject partBO = null;
                                if (pNestData != null)
                                {
                                    partBO = pNestData.ManufacturingPart;
                                }
                                ManufacturingOutputBase mfgOutput = (ManufacturingOutputBase)partBO;
                                if (mfgOutput != null)
                                {
                                    partName = mfgOutput.Name;
                                }
                                if (pNestData!=null)
                                {
                                    plateElement.SetAttribute("PART_GUID", pNestData.ObjectIDForQuery);
                                }
                                else {
                                    plateElement.SetAttribute("PART_GUID", "");
                                }
                                plateElement.SetAttribute("BLOCK_NAME", "");
                                plateElement.SetAttribute("PART_NAME", partName);
                                try
                                {
                                    CommonLogInfo("Saving Delete XML to: " + outputPath + "\\" + pNestData.ObjectIDForQuery + partName + "DELETE.xml", strStartTime + "-MfgNestingService");
                                    outputXmlDocument.Save(outputPath + "\\" + pNestData.ObjectIDForQuery + partName + "DELETE.xml");
                                }
                                catch (Exception ex)
                                {
                                    CommonLogExceptionInfo(ex, strStartTime + "-MfgNestingServiceError");
                                }
                            }
                            else if (pNestData != null)
                            {
                                CommonLogInfo("No Path Specified for Delete XML for: " + pNestData.ObjectIDForQuery, strStartTime + "-MfgNestingService");
                            }
                            businessObject.Delete();
                            businessObject = null;
                        }
                        catch (Exception ex)
                        {
                            CommonLogExceptionInfo(ex, strStartTime + "-MfgNestingServiceError");
                        }
                    }
                    businessObject = null;
                }
                oNestDataByFilter = null;

            }
            catch (Exception ex)
            {
                CommonLogExceptionInfo(ex, strStartTime + "-MfgNestingServiceError");
            }
        }

        private void Update(XmlNode actionNode, XmlDocument inputXmlDocument)
        {
            CommonLogInfo("Getting Up-to-date XML", strStartTime + "-MfgNestingService");
            try
            {
                ReadOnlyCollection<BusinessObject> oNestDataByFilter;
                oNestDataByFilter = GetParts(actionNode, inputXmlDocument);
                for (int i = 0; i < oNestDataByFilter.Count; i++)
                {
                    BusinessObject businessObject = oNestDataByFilter[i];
                    if (businessObject.SupportsInterface("IJMfgNestData"))
                    {
                        try
                        {
                            NestData pNestData = businessObject as NestData;
                            string partName = "";
                            BusinessObject partBO = pNestData.ManufacturingPart;
                            ManufacturingOutputBase mfgOutput = (ManufacturingOutputBase)partBO;
                            if (mfgOutput != null)
                            {
                                CommonLogInfo("Getting Up-to-date XML for: " + mfgOutput.ObjectID, strStartTime + "-MfgNestingService");
                                partName = mfgOutput.Name;
                                string outputPath = base.NestingFilePath(pNestData.OutputType);
                                if (outputPath == null || outputPath.Length == 0)
                                {
                                    outputPath = System.IO.Path.GetTempPath();
                                }
                                if (outputPath.Last() == '\\') {
                                    outputPath = outputPath.Remove(outputPath.Length - 1);
                                }
                                string xmlOutput = "";
                                string outputType = "CMfgPlateOutputCmd_DEFAULT";
                                if (pNestData.OutputType != "")
                                {
                                    outputType = pNestData.OutputType;
                                }
                                if (outputPath!= null && outputPath.Length > 0)
                                {
                                    CommonLogInfo("Outputting XML to : " + outputPath + "\\" + pNestData.ObjectIDForQuery + partName + ".xml", strStartTime + "-MfgNestingService");
                                    xmlOutput = mfgOutput.OutputAsString(outputType, outputPath + "\\" + pNestData.ObjectIDForQuery + partName + ".xml");
                                }
                                else
                                {
                                    xmlOutput = mfgOutput.OutputAsString(outputType);
                                }
                                if (xmlOutput.Length > 0)
                                {
                                    XmlDocument xmlOutputDoc = new XmlDocument();
                                    xmlOutputDoc.LoadXml(xmlOutput);
                                    XmlNodeList partNodeList = xmlOutputDoc.SelectNodes("//SMS_PLATE[not(./SMS_PROD_INFO/SMS_PART_INFO/@NEST_GUID='" + pNestData.ObjectID.Replace("{","").Replace("}","") + "')]");
                                    for (int j = 0; j < partNodeList.Count; j++) {
                                        partNodeList[j].ParentNode.RemoveChild(partNodeList[j]);
                                    }
                                    XmlNode importedNode = inputXmlDocument.ImportNode(xmlOutputDoc.DocumentElement, true);
                                    inputXmlDocument.DocumentElement.AppendChild(importedNode);
                                }
                            }
                        }
                        catch (Exception ex) {
                            CommonLogExceptionInfo(ex, strStartTime + "-MfgNestingServiceError");
                        }
                    }
                    businessObject = null;
                }
            }
            catch (Exception ex)
            {
                CommonLogExceptionInfo(ex, strStartTime + "-MfgNestingServiceError");
            }
        }

        private void GetStatus(XmlNode actionNode, XmlDocument inputXmlDocument)
        {
            try
            {
                ReadOnlyCollection<BusinessObject> oNestDataByFilter = GetParts(actionNode, inputXmlDocument);
                XmlNodeList partNodeList = inputXmlDocument.SelectNodes(".//PART");
                for (int i = 0; i < partNodeList.Count; i++)
                {
                    XmlNode partNode = partNodeList.Item(i);
                    string partGUID = partNode.SelectSingleNode("./@NEST_GUID").Value;
                    Model oPlant = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel;
                    BOMoniker partMoniker = oPlant.GetBOMonikerFromDbIdentifier(partGUID);
                    BusinessObject bo = oPlant.WrapSP3DBO(partMoniker);
                    if (bo.SupportsInterface("IJMfgNestData"))
                    {
                        PropertyValue statusValue = bo.GetPropertyValue("IJMfgNestData", "OutputStatus");
                        XmlAttribute statusAttribute = inputXmlDocument.CreateAttribute("STATUS");
                        statusAttribute.Value = statusValue.ToString();
                        partNode.Attributes.Append(statusAttribute);
                    }
                }
            }
            catch (Exception ex)
            {
                CommonLogExceptionInfo(ex, strStartTime + "-MfgNestingServiceError");
            }
        }

        private void GetXML(XmlNode actionNode, XmlDocument inputXmlDocument)
        {
            try
            {
                ReadOnlyCollection<BusinessObject> oNestDataByFilter;
                oNestDataByFilter = GetParts(actionNode, inputXmlDocument);
                for (int i = 0; i < oNestDataByFilter.Count(); i++)
                {
                    BusinessObject bo = oNestDataByFilter[i];
                    NestData nestData = (NestData)bo;
                    ManufacturingOutputBase manufacturingOutput = (ManufacturingOutputBase)nestData.ManufacturingPart;
                    XmlDocument partDocument = new XmlDocument();
                    partDocument.LoadXml(manufacturingOutput.OutputAsString(nestData.OutputType));
                    XmlNode importedNode = inputXmlDocument.ImportNode(partDocument.DocumentElement, true);
                    inputXmlDocument.DocumentElement.AppendChild(importedNode);
                }
            }
            catch (Exception ex)
            {
                CommonLogExceptionInfo(ex, strStartTime + "-MfgNestingServiceError");
            }
        }

        private void SetProperties(XmlNode actionNode, XmlDocument inputXmlDocument)
        {
            CommonLogInfo("Setting Nesting Properties", strStartTime + "-MfgNestingService");
            try
            {
                ReadOnlyCollection<BusinessObject> oNestDataByFilter;
                oNestDataByFilter = GetParts(actionNode, inputXmlDocument);
                for (int i = 0; i < oNestDataByFilter.Count(); i++)
                {
                    BusinessObject bo = oNestDataByFilter[i];
                    CommonLogInfo("Setting Nesting properties for: " + bo.ObjectID, strStartTime + "-MfgNestingService");
                    XmlNodeList propertyNodeList = actionNode.SelectNodes(".//PROPERTY");
                    for (int j = 0; j < propertyNodeList.Count; j++)
                    {
                        string propertyName = propertyNodeList.Item(j).Attributes.GetNamedItem("NAME").Value;
                        string interfaceName = propertyNodeList.Item(j).Attributes.GetNamedItem("INTERFACE").Value;
                        string propertyType = propertyNodeList.Item(j).Attributes.GetNamedItem("TYPE").Value;
                        string propertyValue = propertyNodeList.Item(j).Attributes.GetNamedItem("VALUE").Value;
                        CommonLogInfo("    Updating attribute: " + propertyName + " With value: " + propertyValue, strStartTime + "-MfgNestingService");
                        SetAttributeValue(bo, propertyName, interfaceName, propertyType, propertyValue);
                    }
                }
            }
            catch (Exception ex)
            {
                CommonLogExceptionInfo(ex, strStartTime + "-MfgNestingServiceError");
            }
        }

        private ReadOnlyCollection<BusinessObject> GetPartsByFilter(XmlNode actionNode, XmlDocument inputXmlDocument)
        {
            ReadOnlyCollection<BusinessObject> oNestDataByFilter = null;
            try
            {
                XmlNode xmlFilter = actionNode.ChildNodes.Item(0);
                if (xmlFilter != null)
                {
                    if (xmlFilter.Name != "FILTER")
                    {
                        return new ReadOnlyCollection<BusinessObject>(new List<BusinessObject>());
                    }
                }
                else { 
                    return new ReadOnlyCollection<BusinessObject>(new List<BusinessObject>());
                }
                XmlNode parentNode = xmlFilter.ParentNode;
                parentNode.RemoveChild(xmlFilter);
                string type = xmlFilter.SelectSingleNode("./@TYPE").Value;
                switch (type)
                {
                    #region "system"
                    case "system":
                        string filterName = xmlFilter.SelectSingleNode("./@QUERY").Value;
                        Collection<SP3DFolder> folderCollection = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.Folders;
                        FilterFolder oFilterFolder = null; //new FilterFolder("temp");
                        for (int i = 0; i < folderCollection.Count; i++)
                        {
                            if (folderCollection[i].Name == "My Filters")
                            {
                                oFilterFolder = (FilterFolder)folderCollection[i];
                                break;
                            }
                        }
                        if (oFilterFolder != null)
                        {
                            Filter oFilter = null;
                            for (int i = 0; i < oFilterFolder.ChildFilters.Count; i++)
                            {
                                if (oFilterFolder.ChildFilters[i].Name == filterName)
                                {
                                    oFilter = (Filter)oFilterFolder.ChildFilters[i];
                                }
                            }
                            if (oFilter != null)
                            {
                                //oFilter.
                                oNestDataByFilter = oFilter.Apply();
                            }
                        }
                        break;
                    #endregion
                    #region "custom"
                    case "custom":
                    default:
                        XmlNodeList filterConditionList = xmlFilter.SelectNodes("./CONDITION");
                        Filter oPropertyFilter = new Filter();
                        for (int i = 0; i < filterConditionList.Count; i++)
                        {
                            XmlNode filterCondition = filterConditionList[i];
                            PropertyValue oProperty;
                            string interfaceName = "IJMfgNestData";
                            string propertyName = filterCondition.SelectSingleNode("./@ATTRIBUTE").Value;
                            if (propertyName.IndexOf("::") != -1)
                            {
                                interfaceName = propertyName.Substring(0, propertyName.IndexOf("::"));
                                propertyName = propertyName.Substring(propertyName.IndexOf("::") + 2);
                            }
                            string propertyValue = filterCondition.SelectSingleNode("./@VALUE").Value;
                            string comparisonValue = filterCondition.SelectSingleNode("./@COMPARISON").Value;
                            PropertyComparisonOperators comparisonOperator;
                            switch (comparisonValue)
                            {
                                case "EQ": comparisonOperator = PropertyComparisonOperators.EQ; break;
                                case "NE": comparisonOperator = PropertyComparisonOperators.NE; break;
                                case "GT": comparisonOperator = PropertyComparisonOperators.GT; break;
                                case "LE": comparisonOperator = PropertyComparisonOperators.LE; break;
                                case "LT": comparisonOperator = PropertyComparisonOperators.LT; break;
                                case "GE": comparisonOperator = PropertyComparisonOperators.GE; break;
                                case "LIKE": comparisonOperator = PropertyComparisonOperators.LIKE; break;
                                case "NOTLIKE": comparisonOperator = PropertyComparisonOperators.NOTLIKE; break;
                                default: comparisonOperator = PropertyComparisonOperators.EQ; break;
                            }
                            InterfaceInformation interfaceInfo = null;
                            try
                            {
                                interfaceInfo = m_MetaDataManager.GetInterfaceInfo(interfaceName, "STRMFG");
                            }
                            catch { }
                            if (interfaceInfo == null)
                            {
                                interfaceInfo = m_MetaDataManager.GetInterfaceInfo(interfaceName, "STRUCT");
                            }
                            PropertyInformation propertyInfo = interfaceInfo.GetPropertyInfo(propertyName);
                            switch (propertyInfo.PropertyType)
                            {
                                case SP3DPropType.PTString:
                                    oProperty = new PropertyValueString(interfaceName, propertyName, propertyValue);
                                    break;
                                case SP3DPropType.PTDate:
                                    DateTime dateTime;
                                    DateTime.TryParse(propertyValue, out dateTime);
                                    oProperty = new PropertyValueDateTime(interfaceName, propertyName, dateTime);
                                    break;
                                case SP3DPropType.PTDouble:
                                    double dVal;
                                    double.TryParse(propertyValue, out dVal);
                                    oProperty = new PropertyValueDouble(interfaceName, propertyName, dVal);
                                    break;
                                case SP3DPropType.PTInteger:
                                    int intVal;
                                    int.TryParse(propertyValue, out intVal);
                                    oProperty = new PropertyValueInt(interfaceName, propertyName, intVal);
                                    break;
                                case SP3DPropType.PTCodelist:
                                    int codelistVal;
                                    int.TryParse(propertyValue, out codelistVal);
                                    oProperty = new PropertyValueCodelist(interfaceName, propertyName, codelistVal);
                                    break;
                                default:
                                    oProperty = new PropertyValueString(interfaceName, propertyName, propertyValue);
                                    break;
                            }
                            oPropertyFilter.Definition.AddWhereProperty(oProperty, comparisonOperator);
                        }
                        if (filterConditionList.Count > 1)
                        {
                            XmlNode Node = filterConditionList.Item(0);
                            if (Node.Attributes.GetNamedItem("CONDITION").Value == "AND")
                            {
                                oPropertyFilter.Definition.MatchAllProperties = true;
                            }
                            else
                            {
                                oPropertyFilter.Definition.MatchAllProperties = false;
                            }
                        }
                        oNestDataByFilter = oPropertyFilter.Apply();
                        break;
                    #endregion
                    #region "sql"
                    case "sql":
                        string query = xmlFilter.SelectSingleNode("./@QUERY").Value;
                        SQLFilter oSQLFilter = new SQLFilter();
                        oSQLFilter.SetSQLFilterString(query.ToUpper());
                        try
                        {
                            oNestDataByFilter = oSQLFilter.Apply();
                        }
                        catch (Exception e) { throw e; }
                        break;
                    #endregion
                }
                //#region "Add objects to XML"
                //if (oNestDataByFilter != null)
                //{
                //    for (int i = 0; i < oNestDataByFilter.Count; i++)
                //    {
                //        NestData nestData = (oNestDataByFilter[i] as Object) as NestData;
                //        ManufacturingOutputBase moBase = (oNestDataByFilter[i] as Object) as ManufacturingOutputBase;
                //        //PlatePart
                //        IManufacturable detailedPart = (oNestDataByFilter[i] as IManufacturable);
                //        string nestOid = "";
                //        string mfgPartOid = "";
                //        string mfgPartName = "";
                //        string detailedPartOid = "";
                //        string lotName = "";
                //        string exportTimestamp = "";
                //        string registrationTimestamp = "";
                //        if (nestData != null)
                //        {
                //            moBase = nestData.ManufacturingPart as ManufacturingOutputBase;
                //            //if (moBase != null)
                //            //{
                //            //    detailedPart = moBase.DetailedPart as IManufacturable;
                //            //}
                //            //else
                //            //{
                //            //    detailedPart = modelDB.WrapSP3DBO(modelDB.GetBOMonikerFromDbIdentifier(nestData.PartGuid)) as IManufacturable;
                //            //}

                //        }
                //        else if (moBase != null)
                //        {
                //            nestData = moBase.NestingInformation();
                //            //detailedPart = moBase.DetailedPart as IManufacturable;
                //        }
                //        else if (detailedPart != null )
                //        {
                //            moBase = detailedPart.Manufacture() as ManufacturingOutputBase;
                //            if (moBase != null)
                //            {
                //                nestData = moBase.NestingInformation() as NestData;
                //            }
                //            else
                //            {
                //                if ((oNestDataByFilter[i] as PlatePartBase) != null)
                //                {
                //                    nestData = new NestData((PlatePart)detailedPart);
                //                }
                //                else if ((oNestDataByFilter[i] as StiffenerPartBase) != null)
                //                {
                //                    nestData = new NestData((StiffenerPartBase)detailedPart);
                //                }
                //            }
                //        }

                //        if (moBase != null)
                //        {
                //            mfgPartOid = moBase.ObjectID;
                //            mfgPartName = moBase.Name;
                //        }

                //        if (detailedPart != null)
                //        {
                //            detailedPartOid = ((BusinessObject)detailedPart).ObjectID;
                //        }

                //        if (nestData != null)
                //        {
                //            nestOid = nestData.ObjectID;
                //            detailedPartOid = nestData.PartGuid;
                //            lotName = nestData.LotNumber;
                //            exportTimestamp = nestData.ExportTimestamp.ToLongDateString();
                //            registrationTimestamp = nestData.RegistrationTimestamp.ToLongDateString();
                //        }

                //        if (moBase != null || detailedPart != null || nestData != null)
                //        {

                //            XmlNode nestObjectNode = inputXmlDocument.SelectSingleNode("//PART[@GUID='" + nestOid + "']");
                //            XmlElement rootElement = (XmlElement)inputXmlDocument.SelectSingleNode("/NESTING_SERVICES");
                //            if (nestObjectNode == null)
                //            {
                //                nestObjectNode = inputXmlDocument.CreateElement("PART");
                //            }
                //            XmlAttribute partGuidAttribute = inputXmlDocument.CreateAttribute("PART_GUID");
                //            XmlAttribute nestGuidAttribute = inputXmlDocument.CreateAttribute("NEST_GUID");
                //            XmlAttribute detailedGuidAttribute = inputXmlDocument.CreateAttribute("DETAILED_GUID");
                //            XmlAttribute mfgNameAttribute = inputXmlDocument.CreateAttribute("MFG_NAME");
                //            XmlAttribute lotNameAttribute = inputXmlDocument.CreateAttribute("LOT_NAME");
                //            XmlAttribute exportTimeStampAttribute = inputXmlDocument.CreateAttribute("EXPORT_TIMESTAMP");
                //            XmlAttribute registrationTimeStampAttribute = inputXmlDocument.CreateAttribute("REGISTRATION_TIMESTAMP");
                //            partGuidAttribute.Value = mfgPartOid;
                //            nestGuidAttribute.Value = nestOid;
                //            detailedGuidAttribute.Value = detailedPartOid;
                //            mfgNameAttribute.Value = mfgPartName;
                //            lotNameAttribute.Value = lotName;
                //            exportTimeStampAttribute.Value = exportTimestamp;
                //            registrationTimeStampAttribute.Value = registrationTimestamp;
                //            nestObjectNode.Attributes.Append(partGuidAttribute);
                //            nestObjectNode.Attributes.Append(nestGuidAttribute);
                //            nestObjectNode.Attributes.Append(detailedGuidAttribute);
                //            nestObjectNode.Attributes.Append(mfgNameAttribute);
                //            nestObjectNode.Attributes.Append(lotNameAttribute);
                //            nestObjectNode.Attributes.Append(exportTimeStampAttribute);
                //            nestObjectNode.Attributes.Append(registrationTimeStampAttribute);
                //            rootElement.AppendChild(nestObjectNode);
                //        }
                //        moBase = null;
                //        detailedPart = null;
                //        nestData = null;
                //    }
                //}
                return oNestDataByFilter;
                #endregion
            }
            catch (Exception ex)
            {
                CommonLogExceptionInfo(ex, strStartTime + "-MfgNestingServiceError");
            }

            return new ReadOnlyCollection<BusinessObject>(new List<BusinessObject>());
        }

        private void ConnectToModel(XmlDocument inputXmlDocument)
        {
            CommonLogInfo("Connecting to model", strStartTime + "-MfgNestingService");
            try
            {
                site = MiddleServiceProvider.SiteMgr.ActiveSite;
                activePlant = site.ActivePlant;
                modelDB = site.ActivePlant.PlantModel;
                catalogDB = site.ActivePlant.PlantCatalog;
                m_MetaDataManager = modelDB.MetadataMgr;

            }
            catch (Exception ex)
            {
                CommonLogExceptionInfo(ex, strStartTime + "-MfgNestingServiceError");
            }
        }

        private void ValidationEventHandler(object sender, ValidationEventArgs e)
        {
            switch (e.Severity)
            {
                case XmlSeverityType.Error:
                    Console.WriteLine("Error: {0}", e.Message);
                    break;
                case XmlSeverityType.Warning:
                    Console.WriteLine("Warning {0}", e.Message);
                    break;
            }
        }

        private Site ConnectToSite(string siteDetails)
        {
            CommonLogInfo("Connecting to Site", strStartTime + "-MfgNestingService");
            Site oSite = null;
            try
            {
                if (siteDetails == "Active")
                {
                    oSite = MiddleServiceProvider.SiteMgr.ConnectSite();
                }
                else
                {
                    string[] strParams = siteDetails.Split(";".ToCharArray());
                    SiteManager.eDBProviderTypes eDBType;
                    if (strParams[0].ToUpper() == "MSSQL")
                    {
                        eDBType = SiteManager.eDBProviderTypes.MSSQL;
                    }
                    else
                    {
                        eDBType = SiteManager.eDBProviderTypes.Oracle;
                    }
                    oSite = MiddleServiceProvider.SiteMgr.ConnectSite(strParams[1], strParams[2], eDBType, strParams[3]);
                }
            }
            catch (Exception ex)
            {
                CommonLogExceptionInfo(ex, strStartTime + "-MfgNestingServiceError");
                CommonLogInfo("Exception Occurred While Connecting to Site - " + ex.Message, strStartTime + "-MfgNestingService");
            }

            return oSite;
        }

        private Plant OpenPlant(Site oSite, string plantName)
        {
            CommonLogInfo("Opening Model", strStartTime + "-MfgNestingService");
            try
            {
                Plant oPlant;
                for (int i = 0; i < oSite.Plants.Count(); i++)
                { // Each oPlant In oSite.Plants
                    oPlant = oSite.Plants[i];
                    if (oPlant.Name == plantName)
                    {
                        oSite.OpenPlant(oPlant);
                    }
                }
            }
            catch (Exception ex)
            {
                CommonLogExceptionInfo(ex, strStartTime + "-MfgNestingServiceError");
                CommonLogInfo("Exception Occurred While Opening Plant - " + plantName + "\n" + ex.Message, strStartTime + "-MfgNestingService");
            }
            return MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant;
        }

        private bool SetActivePG(Model oModel, String pgName)
        {
            CommonLogInfo("Setting Permission Group", strStartTime + "-MfgNestingService");
            try
            {
                Console.WriteLine("Starting to set active Permission Group");
                PermissionGroup oPG;
                for (int i = 0; i < oModel.PermissionGroups.Count; i++)
                {
                    oPG = oModel.PermissionGroups[i];
                    AccessRule oAR;
                    for (int j = 0; j < oPG.AccessRules.Count; j++)
                    {
                        oAR = oPG.AccessRules[j];
                        string thisPGName = oPG.ToString();
                        Console.WriteLine("Evaluating " + oPG.Parent.ToString() + "\\" + thisPGName + "." + oAR.User + " - " + oAR.AccessRight.ToString());
                        if ((oAR.AccessRight != PGAccessRights.Read) || Thread.CurrentPrincipal.IsInRole(oAR.User))
                        {
                            if (oPG.Name == pgName || (oPG.Parent.ToString() + "\\" + oPG.Name == pgName))
                            {
                                oModel.ActivePermissionGroup = oPG;
                                MiddleServiceProvider.TransactionMgr.Commit("");
                                return true;
                            }
                        }
                    }
                }
            }
            catch (Exception ex) {
                CommonLogExceptionInfo(ex, strStartTime + "-MfgNestingServiceError");
            }
            return false;
        }

        /// <summary>
        /// The event handler for assemblies that fail to load.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="args">The event args.</param>
        /// <returns></returns>
        public static Assembly MyResolveEventHandler(object sender, ResolveEventArgs args)
        {
            //This handler is called only when the common language runtime tries to bind to the assembly and fails.
            //Retrieve the list of referenced assemblies in an array of AssemblyName.
            //Loop through the array of referenced assemblies.
            foreach (AssemblyName strAssmbName in Assembly.GetExecutingAssembly().GetReferencedAssemblies())
            {
                //Look for the assembly names that have raised the "AssemblyResolve" event.
                if ((((strAssmbName.Name.EndsWith("CommonMiddle") || strAssmbName.Name.EndsWith("ManufacturingMiddle")) &&
                    (strAssmbName.FullName.Substring(0, strAssmbName.FullName.IndexOf(",")) ==
                    args.Name.Substring(0, args.Name.IndexOf(",")))) ? 1 : 0) != 0)
                {
                    //We only have this handler to deal with loading of CommonMiddle. Rest everything we dont bother.
                    AppDomain.CurrentDomain.AssemblyResolve -= new
                    ResolveEventHandler(MyResolveEventHandler);
                    //Load the assembly from the specified path and return it.
                    string sInstallPath;
#if DEBUG
                    sInstallPath = @"X:\Container\Bin\Assemblies\Release\";
#else
                    sInstallPath = GetInstallDir();
                    if (!sInstallPath.EndsWith(@"\"))
                    {
                        sInstallPath = sInstallPath + @"\";
                    }
                    sInstallPath = sInstallPath + @"Core\Container\Bin\Assemblies\Release\";
#endif
                    return Assembly.LoadFrom(sInstallPath + strAssmbName.Name + ".dll");
                }
            }
            return null;
        }

        /// <summary>
        /// Gets the install directory for the current instance of Smart3D.
        /// </summary>
        /// <returns>The install dirrectory path.</returns>
        public static string GetInstallDir()
        {
            object regKeyValue;
            //Build the path of the assembly from where it has to be loaded.
            //Check CurrentUser,LocalMachine, 64bit CurrentUser,LocalMachine
            regKeyValue = Registry.GetValue(Registry.CurrentUser +
            @"\Software\Intergraph\SP3D\Installation", "INSTALLDIR", "");

            if (regKeyValue == null)
            {
                regKeyValue = Registry.GetValue(Registry.LocalMachine +
                @"\Software\Intergraph\SP3D\Installation", "INSTALLDIR", "");
            }
            if (regKeyValue == null)
            {
                regKeyValue = Registry.GetValue(Registry.CurrentUser +
                @"\Software\Wow6432Node\Intergraph\SP3D\Installation", "INSTALLDIR", "");
            }


            if (regKeyValue == null)
            {
                regKeyValue = Registry.GetValue(Registry.LocalMachine +
                @"\Software\Wow6432Node\Intergraph\SP3D\Installation", "INSTALLDIR", "");
            }
            string sInstallPath = @"";
            if (regKeyValue == null)
            {
#if DEBUG
                sInstallPath = @"X:\Container\Bin\Assemblies\Debug\";
#else
                throw new Exception("Error Reading Smart 3D Installation Directory from Registry !!! Exiting");
#endif
            }
            else
            {
                sInstallPath = regKeyValue.ToString();
            }
            return sInstallPath;
        }

        private bool EnsureLogFolderExists(){
            string strLogDir = GetLogDirectoryPath();
            if (! Directory.Exists(strLogDir)){
                Directory.CreateDirectory(strLogDir); // if exception, caller catches & deals
                return true; // if it reached here, then directory created w/o exception
            }
            return true; // if it reached here, then directory exists already
        }
        
        private string GetLogDirectoryPath(){
            //return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Log\\";
            return Path.GetTempPath() + "\\NestingServiceLog\\";
        }

        private void CommonLogExceptionInfo(Exception ex , string strName){
            string strErrLogFile = GetLogDirectoryPath() + strName + ".log";
            using (StreamWriter oErrLogFile = new System.IO.StreamWriter(strErrLogFile, true))
            {
	            oErrLogFile.WriteLine(ex.Message);
	            oErrLogFile.Flush();
	            oErrLogFile.Close();
	        }
        }

        private void CommonLogInfo(string message, string strName)
        {
            string strCurrentTime = DateTime.Now.ToLongTimeString(); //String.Replace(String.Format(DateTime.Now, "s"), ":", "-");
            string strErrLogFile = GetLogDirectoryPath() + strName + ".log";
            using (StreamWriter oErrLogFile = new System.IO.StreamWriter(strErrLogFile, true))
            {
            oErrLogFile.WriteLine(strCurrentTime + " - " + message);
            oErrLogFile.Flush();
            oErrLogFile.Close();
            }
        }

        private string CallNestingSystem(string commandName, string DBName, string fileNameOut)
        {
            #region Call Nesting System
            //Create process
            System.Diagnostics.Process pProcess = new System.Diagnostics.Process();
            //strCommand is path and file name of command to run
            string fileName = "";
            string strWorkingDirectory = "";
            string installDir = GetInstallDir();
            if (File.Exists(installDir + "S3DProductionClient.exe"))
            {
                fileName = installDir + "S3DProductionClient.exe";
            }
            else {
                CommonLogInfo("Missing client program to call nesting system", strStartTime + "-MfgNestingService");
                return "";
            }

            pProcess.StartInfo.FileName = fileName;
            //strCommandParameters are parameters to pass to program
            string strCommandParameters = "";
            if (fileName == installDir + "S3DProductionClient.exe")
            {
                if (!CreateClientConfigXML())
                {
                    CommonLogInfo("Missing configuration file for Remote access to nesting system", strStartTime + "-MfgNestingService");
                    return "";
                }

                strCommandParameters = "NestingSystem" + " " + commandName + " " + DBName;
                pProcess.StartInfo.Arguments = strCommandParameters;
                pProcess.StartInfo.UseShellExecute = false;
                //Set output of program to be written to process output stream
                pProcess.StartInfo.RedirectStandardOutput = true;
                //Optional
                pProcess.StartInfo.WorkingDirectory = strWorkingDirectory;
                //Start the process
                pProcess.Start();
                //Get program output
                string strOutput = pProcess.StandardOutput.ReadToEnd();
                //Wait for process to finish
                pProcess.WaitForExit();
                return strOutput;
            }
            #endregion Call Nesting Service
            return "";
        }

        private bool CreateClientConfigXML()
        {
            if (File.Exists(GetInstallDir() + "S3DProductionClient.exe.config"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
