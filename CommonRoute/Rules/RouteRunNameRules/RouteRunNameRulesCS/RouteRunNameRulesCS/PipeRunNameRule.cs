//***************************************************************************************
//  Copyright (C) 2008-2009, Intergraph Corporation.  All rights reserved.
//  Project  : M:\SharedContent\Src\CommonRoute\Rules\RouteRunNameRules\RouteRunNameRulesCS
//
//  Class    : PipeRunNameRule
//
//  Abstract : The file contains implementation of the naming rules for PipeRuns.
//
//***************************************************************************************
using System.Windows.Forms;
using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;              // Add references
using Ingr.SP3D.Systems.Middle;
using Ingr.SP3D.Route.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;

// Define user namespace
namespace Ingr.SP3D.Route.Examples.RunNameRulesCS
{
    /// <summary>
    /// PipeRunNameRule: should be derived form NameRuleBase class. 
    /// </summary>
    public class PipeRunNameRule : NameRuleBase 
    {
        private const string strCountFormat = "0000";

        /// <summary>
        /// Creates a name for the object passed in. The name is based on the parents
        /// name and object name.The Naming Parents are added in AddNamingParents().
        /// Both these methods are called from naming rule semantic.
        /// </summary>
        /// <param name="oEntity">Child object that needs to have the naming rule naming.</param>
        /// <param name="oParents">Naming parents collection.</param>
        /// <param name="oActiveEntity">Naming rules active entity on which the NamingParentsString is stored.</param>

        public override void ComputeName(BusinessObject oEntity, ReadOnlyCollection<BusinessObject> oParents, Ingr.SP3D.Common.Middle.BusinessObject oActiveEntity)
        {
            try
            {
                //Get the Parent of the PipeRun
                Pipeline oPipeline = (Pipeline)oParents[0];
                string strPipelineName = oPipeline.Name;

                //Get the name of the PipeRun
                PipeRun oPipeRun = (PipeRun)oEntity;
                string strRunName = oPipeRun.Name;

                //Get NPD from PipeRun
                string strNPD = oPipeRun.NPD.Size.ToString();

                //Get the Fluid Code from the PipeLine
                PropertyValueCodelist oCLPropValue = (PropertyValueCodelist)oPipeline.GetPropertyValue("IJPipelineSystem", "FluidCode");
                string strFluidCode = oCLPropValue.PropertyInfo.CodeListInfo.GetCodelistItem(oCLPropValue.PropValue).Name;

                //Get Pipe Spec from PipeRun
                PipeSpec oPipeSpec = (PipeSpec)oPipeRun.Specification;
                PropertyValueString oStrPropValue;
                string strPipeSpec = "";
                if (null != oPipeSpec)
                {
                    oStrPropValue = (PropertyValueString)oPipeSpec.GetPropertyValue("IJDPipeSpec", "SpecName");
                    strPipeSpec = oStrPropValue.PropValue;
                }

                //To get the UnitSystem
                ISystem oParentSystem = (ISystem)oParents[0];

                string strUnitSystem = null;
                while (oParentSystem is ISystemChild)
                {
                    if (oParentSystem is UnitSystem)
                    {
                        //Get the Unit system string.
                        Ingr.SP3D.Systems.Middle.System oSys = (Ingr.SP3D.Systems.Middle.System)oParentSystem;
                        oStrPropValue = (PropertyValueString)oSys.GetPropertyValue("IJNamedItem", "Name");
                        strUnitSystem = oStrPropValue.PropValue;
                        // Presently we are Displaying the Name of the Unitsystem but if we want to change
                        // it to display UnitCode of the Unit system then use the below code
                        //oStrPropValue = oSys.GetPropertyValue("IJUnitSystem", "UnitCode")
                        //strUnitSystem = oStrPropValue.PropValue
                        break;
                    }
                    else
                    {
                        ISystemChild oSysChild = (ISystemChild)oParentSystem;
                        oParentSystem = oSysChild.SystemParent;
                    }
                }
                string strValidateName = null;
                string strOldName = null;
                string[] arr = null;
                int intUpperBound = 0;
                int intLowerBound = 0;
                string strOldPipeSpec = null;
                string strOldSeqNo = null;
                int RunNameLength = 0;
                int Checklength = 0;
                 
                if (strUnitSystem != null)
                {
                    strValidateName = strUnitSystem + "-" + strNPD + "-" + strFluidCode + "-";
                }
                else
                {
                    strValidateName = strNPD + "-" + strFluidCode + "-";
                }
                
                arr = strRunName.Split('-');
                arr.GetUpperBound(intUpperBound);
                if (intUpperBound > 0)
                {
                    strOldPipeSpec = arr[intUpperBound];
                    intLowerBound = intUpperBound - 1;
                    strOldSeqNo = arr[intLowerBound];
                    strOldName = strRunName;
                    strOldName.Trim();
                    RunNameLength = strOldName.Length;
                    strOldPipeSpec.Trim();
                    strOldSeqNo.Trim();
                    Checklength = strOldPipeSpec.Length + strOldSeqNo.Length;
                    //Deleting "1" more from RunNameLength to Include "-" also .
                    Checklength = RunNameLength - Checklength - 1;
                    strOldName = strOldName.PadLeft(Checklength);
                }
                //We Compare the NewString (strUnitSystem + "-" + strNPD + "-" + strFluidCode + "-" )and Oldstring from the RunName and
                //if they are Different we generate a New Name.we also Compare whethere the PipeRunSpec has Changed or Not.
                if ((String.Compare(strOldName,strValidateName) != 0) || (String.Compare(strOldPipeSpec,strPipeSpec) != 0))
                {
                    long lCount = 0;
                    string strSeqNo = null;
                    string strLocation = null;
                    string strName = null;

                    GetCountAndLocationID(strPipelineName, out lCount, out  strLocation);
                    strSeqNo = String.Format(strCountFormat, lCount);
                    if (strUnitSystem !=  null)
                    {
                        strName = strUnitSystem + "-" + strNPD + "-" + strFluidCode + "-" + strSeqNo + "-" + strPipeSpec;
                    }
                    else
                    {
                        strName = strNPD + "-" + strFluidCode + "-" + strSeqNo + "-" + strPipeSpec;
                    }
                    SetNamingParentsString(oActiveEntity, strName);
                    oEntity.SetPropertyValue(strName, "IJNamedItem", "Name");
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// All the Naming Parents that need to participate in an objects naming are added here to the
        /// Collection(Of BusinessObject). The parents added here are used in computing the name of the object in
        /// ComputeName(). Both these methods are called from naming rule semantic.
        /// </summary>
        /// <param name="oEntity">Child object that needs to have the naming rule naming.</param>
        /// <returns> Collection of parents that participate in an objects naming.</returns>

        public override Collection<BusinessObject> GetNamingParents(BusinessObject oEntity)
        {
            Collection<BusinessObject> oParentsColl = new Collection<BusinessObject>();

            try
            {
                BusinessObject oParent = GetParent(HierarchyTypes.System, oEntity); // Get the System Parent
                if (oParent != null)
                {
                    oParentsColl.Add(oParent); //Add the Parent to the ParentColl
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return oParentsColl;

        }
    }
}
