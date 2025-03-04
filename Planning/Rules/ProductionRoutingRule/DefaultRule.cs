using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Common.Middle.Services.Hidden;
using Ingr.SP3D.Planning.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace Ingr.SP3D.Content.Planning
{
    /// <summary>
    /// 
    /// </summary>
    public class DefaultRule : PlanningProductionRoutingRuleBase
    {
        private bool xmlParsed = false;
        /// <summary>
        /// Constructor
        /// </summary>
        public DefaultRule()
        {

        }

        #region Methods

        /// <summary>
        /// To compute production data
        /// </summary>
        /// <param name="part"></param>
        /// <param name="prodRouting"></param>
        public override void ComputeProductionData(BusinessObject part, ProductingRouting prodRouting)
        {
            #region Input Error Handling
            if (part == null && prodRouting == null)
            {
                throw new ArgumentNullException("Input arguments are null");
            }
            #endregion //Input Error Handling

            try
            {
                if (xmlParsed == false)
                {
                    string xmlFilePath = string.Empty;
                    string sharedContentPath = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantCatalog.SymbolShare;

                    if (!string.IsNullOrEmpty(base.Argument) && (Directory.Exists(sharedContentPath + "\\" + base.Argument) == true) && (File.Exists(sharedContentPath + "\\" + base.Argument) == true))
                    {
                        xmlFilePath = sharedContentPath + "\\" + base.Argument;
                    }
                    else
                    {
                        xmlFilePath = sharedContentPath + "\\" + "Production\\ProductionRouting\\ProductionRoutingData.xml";
                    }

                    Readxml(xmlFilePath);
                    xmlParsed = true;
                }
                EvaluateProductionData(prodRouting,part);
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
        }

        /// <summary>
        /// To compute action data
        /// </summary>
        /// <param name="part"></param>
        /// <param name="prodRouting"></param>
        /// <param name="prodRoutingAction"></param>
        public override void ComputeActionData(BusinessObject part, ProductingRouting prodRouting, RoutingAction prodRoutingAction)
        {
            #region Input Error Handling
            if (part == null && prodRouting == null && prodRoutingAction == null)
            {
                throw new ArgumentNullException("Input arguments are null");
            }
            #endregion //Input Error Handling

            try
            {
                if (xmlParsed == false)
                {
                    string xmlFilePath = string.Empty;
                    string sharedContentPath = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantCatalog.SymbolShare;

                    if (!string.IsNullOrEmpty(base.Argument) && (Directory.Exists(sharedContentPath + "\\" + base.Argument) == true) && (File.Exists(sharedContentPath + "\\" + base.Argument) == true))
                    {
                        xmlFilePath = sharedContentPath + "\\" + base.Argument;
                    }
                    else
                    {
                        xmlFilePath = sharedContentPath + "\\" + "Production\\ProductionRouting\\ProductionRoutingData.xml";
                    }

                    Readxml(base.Argument);
                    xmlParsed = true;
                }
                EvaluateActionData(part, prodRoutingAction);
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="prodRoutingData"></param>
        public override void GetProductingRoutingData(ProductionRoutingInfo prodRoutingData)
        {
            #region Input Error Handling
            if (prodRoutingData == null)
            {
                throw new ArgumentNullException("Input Data is null");
            }
            #endregion //Input Error Handling

            try
            {
                if (xmlParsed == false)
                {
                    string xmlFilePath = string.Empty;
                    string sharedContentPath = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantCatalog.SymbolShare;

                    if (!string.IsNullOrEmpty(base.Argument) && (Directory.Exists(sharedContentPath + "\\" + base.Argument) == true) && (File.Exists(sharedContentPath + "\\" + base.Argument) == true))
                    {
                        xmlFilePath = sharedContentPath + "\\" + base.Argument;

                    }
                    else
                    {
                        xmlFilePath = sharedContentPath + "\\" + "Production\\ProductionRouting\\ProductionRoutingData.xml";
                    }

                    Readxml(base.Argument);
                    xmlParsed = true;
                }

                string Action = null;
                string machine = "";

                BusinessObject part = prodRoutingData.Part;
                if (prodRoutingData.RequestType == InputType.Machines)
                {
                    Action = prodRoutingData.ActionName;
                    Dictionary<string, List<string>> machinesinfo = GetMachines(Action, part);
                    prodRoutingData.machineInfo = machinesinfo;
                }
                else if (prodRoutingData.RequestType == InputType.MachineCodes)
                {
                    Action = prodRoutingData.ActionName;                    
                    Dictionary<string, List<string>> machinesinfo = GetCodes(Action, machine);
                    prodRoutingData.machineInfo = machinesinfo;
                }
                else if (prodRoutingData.RequestType == InputType.Actions)
                {                    
                    prodRoutingData.Actions = GetActions();
                }
                else if (prodRoutingData.RequestType == InputType.StageCodes)
                {                    
                    prodRoutingData.StageCodes = GetStageCodes(part);
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
        }

        # endregion
    }
}
