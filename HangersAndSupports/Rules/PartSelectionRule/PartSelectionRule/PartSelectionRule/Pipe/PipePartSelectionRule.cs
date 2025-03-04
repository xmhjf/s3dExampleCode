//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   PipePartSelectionRule.cs
//   
//   Author       :  BS
//   Creation Date:  01.May.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   01.May.2013     BS      CR233076  Convert HgrPipePartSelRule to C# .Net  
//   14-03-2014 B.Chethan    DM-251135  [TR] Rod-Size Part Selection rule fails
//   22-Jan-2015     PVK     TR-CP-264951  Resolve coverity issues found in November 2014 report
//   07-Jul-2015     PVK     DI-CP-274352	Anvil SmartParts are missing Part Selection Rules
//   09-Sep-2015     PVK     DI-CP-279143	Replace Anvil & HgrBeam used in HS_Assembly with Anvil2010 Parts & Rich HgrBeam
//   30-Nov-2015     PVK     DI-CP-281220	Malleable Beam Clamp SmartPart needs New Part Selection Rules
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//   21-Mar-2016     PVK     TR-CP-288920	Issues found in HS_Assembly_V2
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Content.Support.Symbols;
using Ingr.SP3D.Content.Support.Rules;

namespace Ingr.SP3D.Content.Support.Rules
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Rules
    //It is recommended that customers specify namespace of their Rules to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this Rule to suit their
    //requirements, they should change namespace/Rule name so the identity of the modified
    //Rule will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    
    //----------------------------------------------------------------------
    //This Rule returns part by pipe size and Insulation 
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByPipeAndInsul
    //----------------------------------------------------------------------  
    public class PartByPipeAndInsul : SupportPartSelectionRule
    {
        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string sPartClass)
        {
            Part selectedPart = null;
            if (!SupportHelper.Support.IsDesignSupportAssembly())
            {
                double insulationThickness = (double)((PropertyValueDouble)SupportHelper.Support.GetPropertyValue("IJUAhsInsulationTh", "InsulationTh")).PropValue;
                NominalDiameter pipeDiameter = new NominalDiameter();
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                UnitName unitName = MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Distance, pipeInfo.NominalDiameter.Units);

                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass(sPartClass);  

                foreach (BusinessObject part in partClass.Parts)
                {
                    string NDUnitType = (string)((PropertyValueString)part.GetPropertyValue("IJHgrDiameterSelection", "NDUnitType")).PropValue;
                    pipeDiameter.Size = pipeInfo.NominalDiameter.Size;

                    if (NDUnitType.ToLower() == "in")
                        pipeDiameter.Size = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeDiameter.Size, unitName, UnitName.NPD_INCH);
                    else if (NDUnitType.ToLower() == "mm")
                        pipeDiameter.Size = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeDiameter.Size, unitName, UnitName.NPD_MILLIMETER);

                    if ((HgrCompareDoubleService.cmpdbl(((double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsInsulationTh", "InsulationTh")).PropValue), insulationThickness) == true && ((double)((PropertyValueDouble)part.GetPropertyValue("IJHgrDiameterSelection", "NDFrom")).PropValue <= pipeDiameter.Size) && ((double)((PropertyValueDouble)part.GetPropertyValue("IJHgrDiameterSelection", "NDTo")).PropValue >= pipeDiameter.Size)))
                    {
                        selectedPart = (Part)part;
                        break;
                    }
                }
            }
            return selectedPart;
        }
    }
    //----------------------------------------------------------------------
    //This Rule returns part by load factor
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByLoadFactor
    //----------------------------------------------------------------------  
    public class PartByLoadFactor : SupportPartSelectionRule
    {
        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string sPartClass)
        {
            Collection<object> sectionInfo = new Collection<object>();
            GenericHelper.GetDataByRule("HgrSupAngleByLF", null, out sectionInfo);          
            string[] sectionData = new string[3];
            if (sectionInfo != null)
            {
                sectionData[0] = (string)sectionInfo[0];
                sectionData[1] = (string)sectionInfo[1];
                sectionData[2] = (string)sectionInfo[2];
            }
            return SupportPartByAssociation(sPartClass, 3, sectionData);
        }
    }
    //----------------------------------------------------------------------
    //This Rule returns part by pipe size.
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByPipeSize
    //---------------------------------------------------------------------- 
    public class PartByPipeSize : SupportPartSelectionRule
    {
        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string sPartClass)
        {
            NominalDiameter nominalPipeDiameter = new NominalDiameter();
            Part selectedPartFromPartClass=null;
            if (SupportedHelper.SupportedObjectInfo(ReferenceIndex).SupportedObjectType == SupportedObjectType.Pipe)
            {
                PipeObjectInfo pipe = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(ReferenceIndex);
               
                if (HgrCompareDoubleService.cmpdbl(pipe.InsulationThickness , 0)==false)
                    nominalPipeDiameter = GenericHelper.GetEquivalentNominalPipeDiameter(pipe.OutsideDiameter, pipe.InsulationThickness, EquivalentNominalPipeDiameterSelectionType.Upper);
                else
                    nominalPipeDiameter = pipe.NominalDiameter;

                selectedPartFromPartClass= SupportPartBySize(sPartClass, nominalPipeDiameter);
            }
            else if (SupportedHelper.SupportedObjectInfo(ReferenceIndex).SupportedObjectType == SupportedObjectType.Conduit)
            {
                ConduitObjectInfo conduit = (ConduitObjectInfo)SupportedHelper.SupportedObjectInfo(ReferenceIndex);
                
                selectedPartFromPartClass= SupportPartBySize(sPartClass, conduit.NominalDiameter);
            }
            else if (SupportedHelper.SupportedObjectInfo(ReferenceIndex).SupportedObjectType == SupportedObjectType.CableTray)
            {
                CableTrayObjectInfo cable = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(ReferenceIndex);
                
                selectedPartFromPartClass= SupportPartBySize(sPartClass, cable.Width, cable.Depth, ComparisonOperatorType.EQUAL);
            }
            else if (SupportedHelper.SupportedObjectInfo(ReferenceIndex).SupportedObjectType == SupportedObjectType.HVAC)
            {
                DuctObjectInfo duct = (DuctObjectInfo)SupportedHelper.SupportedObjectInfo(ReferenceIndex);

                if (HgrCompareDoubleService.cmpdbl(duct.InsulationThickness , 0)==false)
                    nominalPipeDiameter = GenericHelper.GetEquivalentNominalPipeDiameter(duct.OutsideDiameter, duct.InsulationThickness, EquivalentNominalPipeDiameterSelectionType.Upper);
                else
                    nominalPipeDiameter = duct.NominalDiameter;

                selectedPartFromPartClass = SupportPartBySize(sPartClass, nominalPipeDiameter);
            }
            return selectedPartFromPartClass;
        }
    }
    //----------------------------------------------------------------------
    //This Rule retuns part when pipe size is equal. 
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByPipeSizeEqual
    //---------------------------------------------------------------------- 
    public class PartByPipeSizeEqual : SupportPartSelectionRule
    {
        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string sPartClass)
        {
            PipeObjectInfo pipe = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(ReferenceIndex);
            return SupportPartBySize(sPartClass, pipe.NominalDiameter, NDComparisonOperatorType.NDFrom_EQUAL);            
        }
    }
    //----------------------------------------------------------------------
    //This Rule returns part by structure size.
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByPipeStrutSize
    //---------------------------------------------------------------------- 
    public class PartByPipeStrutSize : SupportPartSelectionRule
    {
        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string sPartClass)
        {
            Part selectedPart = null;
            if (!SupportHelper.Support.IsDesignSupportAssembly())
            {
                int routeIndex = 1;
                string strutSize = string.Empty;
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                BusinessObject supportPart = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
                if (ReferenceIndex == 1)
                {
                    if (SupportHelper.Support.SupportsInterface("IJUAHgrURS_RS1Size1"))
                        strutSize = (string)((PropertyValueString)SupportHelper.Support.GetPropertyValue("IJUAHgrURS_RS1Size1", "Strut1Size1")).PropValue;
                    else if (SupportHelper.Support.SupportsInterface("IJUAHgrURS_Strut1Size1"))
                        strutSize = (string)((PropertyValueString)SupportHelper.Support.GetPropertyValue("IJUAHgrURS_Strut1Size1", "Strut1Size1")).PropValue;
                    else if (SupportHelper.Support.SupportsInterface("IJUAhsStrut_Strut1"))
                        strutSize = (string)((PropertyValueString)SupportHelper.Support.GetPropertyValue("IJUAhsStrut_Strut1", "Strut1Size1")).PropValue;
                    else if (supportPart.SupportsInterface("IJUAhsStrut1_Size1"))
                        strutSize = (string)((PropertyValueString)supportPart.GetPropertyValue("IJUAhsStrut1_Size1", "Strut1Size1")).PropValue;
                    else if (supportPart.SupportsInterface("IJUAhsStrutSize"))
                        strutSize = (string)((PropertyValueString)supportPart.GetPropertyValue("IJUAhsStrutSize", "StrutSize")).PropValue;

                }
                else
                {
                    if (SupportHelper.Support.SupportsInterface("IJUAhsStrut2_Size1"))
                        strutSize = (string)((PropertyValueString)SupportHelper.Support.GetPropertyValue("IJUAhsStrut2_Size1", "Strut2Size1")).PropValue;
                    else if (SupportHelper.Support.SupportsInterface("IJUAhsStrut_Strut2"))
                        strutSize = (string)((PropertyValueString)SupportHelper.Support.GetPropertyValue("IJUAhsStrut_Strut2", "Strut2Size1")).PropValue;
                    else if (supportPart.SupportsInterface("IJUAhsStrut2_Size1"))
                        strutSize = (string)((PropertyValueString)supportPart.GetPropertyValue("IJUAhsStrut2_Size1", "Strut2Size1")).PropValue;
                    else if (supportPart.SupportsInterface("IJUAhsStrutSize"))
                        strutSize = (string)((PropertyValueString)supportPart.GetPropertyValue("IJUAhsStrutSize", "StrutSize")).PropValue;

                }

                double pipeNominalDiameter = ((PipeObjectInfo)SupportedHelper.SupportedObjectInfo(routeIndex)).NominalDiameter.Size;

                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass(sPartClass);

                foreach (BusinessObject part in partClass.Parts)
                { 
                    if (((string)((PropertyValueString)part.GetPropertyValue("IJUAhsSize1", "Size1")).PropValue == strutSize && ((double)((PropertyValueDouble)part.GetPropertyValue("IJHgrDiameterSelection", "NDFrom")).PropValue <= (pipeNominalDiameter + 0.0001)) && ((double)((PropertyValueDouble)part.GetPropertyValue("IJHgrDiameterSelection", "NDTo")).PropValue >= (pipeNominalDiameter - 0.0001))))
                    {
                        if (SupportHelper.Support.SupportsInterface("IJUAhsStrut_Strut1") || SupportHelper.Support.SupportsInterface("IJUAhsStrut_Strut2"))
                        {
                            PropertyValueCodelist catalogCodeList = (PropertyValueCodelist)part.GetPropertyValue("IJUAhsCatalog", "Catalog");
                            CodelistItem catalogCodeListItem = catalogCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(catalogCodeList.PropValue);
                            if (catalogCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(catalogCodeList.PropValue).DisplayName == catalogCodeListItem.DisplayName)
                            {
                                selectedPart = (Part)part;
                                break;
                            }
                        }
                        else
                        {
                            selectedPart = (Part)part;
                            break;
                        }
                    }
                }
            }
            return selectedPart;
        }
    }
    //----------------------------------------------------------------------
    //This Rule returns by rod size.
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByRodSize
    //---------------------------------------------------------------------- 
    public class PartByRodSize : SupportPartSelectionRule
    {
        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string sPartClass)
        {
            return SupportPartByCriteria(sPartClass, "JUAHgrRod_Dia", "Rod_Dia", ComparisonOperatorType.GREATER_OR_EQUAL, 0.0001);
        }
    }
    //----------------------------------------------------------------------
    //This Rule returns part by rod size.
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByRodSizeS3DPrt
    //---------------------------------------------------------------------- 
    public class PartByRodSizeS3DPrt : SupportPartSelectionRule
    {
        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string sPartClass)
        {
            Part selectedPart = null;
            if (!SupportHelper.Support.IsDesignSupportAssembly())
            {
                double rodDiameter = (double)((PropertyValueDouble)RuleService.GetPropertyValue(SupportHelper.Support, "IJUAhsRodDiameter", "RodDiameter")).PropValue;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass(sPartClass);

                foreach (BusinessObject part in partClass.Parts)
                {
                    double RodDiameter = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsRodDiameter", "RodDiameter")).PropValue;
                    if ((RodDiameter >= rodDiameter - 0.001) && (RodDiameter <= (rodDiameter + 0.001)))
                    {
                        selectedPart = (Part)part;
                        break;
                    }

                }
            }

            return selectedPart;
        }
    }
    //----------------------------------------------------------------------
    //This Rule returns part by struct size 1.
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByStrutSize1
    //---------------------------------------------------------------------- 
    public class PartByStrutSize1 : SupportPartSelectionRule
    {
        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string sPartClass)
        {
            Part selectedPart = null;
            Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
            BusinessObject supportPart = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
            if (!SupportHelper.Support.IsDesignSupportAssembly())
            {
                string strutSize = string.Empty;
                if (ReferenceIndex == 1)
                {
                    if (SupportHelper.Support.SupportsInterface("IJUAHgrURS_RS1Size1"))
                        strutSize = (string)((PropertyValueString)SupportHelper.Support.GetPropertyValue("IJUAHgrURS_RS1Size1", "Strut1Size1")).PropValue;
                    else if (SupportHelper.Support.SupportsInterface("IJUAHgrURS_Strut1Size1"))
                        strutSize = (string)((PropertyValueString)SupportHelper.Support.GetPropertyValue("IJUAHgrURS_Strut1Size1", "Strut1Size1")).PropValue;
                    else if (SupportHelper.Support.SupportsInterface("IJUAhsStrut_Strut1"))
                        strutSize = (string)((PropertyValueString)SupportHelper.Support.GetPropertyValue("IJUAhsStrut_Strut1", "Strut1Size1")).PropValue;
                    else if (supportPart.SupportsInterface("IJUAhsStrut1_Size1"))
                        strutSize = (string)((PropertyValueString)supportPart.GetPropertyValue("IJUAhsStrut1_Size1", "Strut1Size1")).PropValue;
                    else if (supportPart.SupportsInterface("IJUAhsStrutSize"))
                        strutSize = (string)((PropertyValueString)supportPart.GetPropertyValue("IJUAhsStrutSize", "StrutSize")).PropValue;
                }
                else
                {
                    if (SupportHelper.Support.SupportsInterface("IJUAhsStrut2_Size1"))
                        strutSize = (string)((PropertyValueString)SupportHelper.Support.GetPropertyValue("IJUAhsStrut2_Size1", "Strut2Size1")).PropValue;
                    else if (SupportHelper.Support.SupportsInterface("IJUAhsStrut_Strut2"))
                        strutSize = (string)((PropertyValueString)SupportHelper.Support.GetPropertyValue("IJUAhsStrut_Strut2", "Strut2Size1")).PropValue;
                    else if (supportPart.SupportsInterface("IJUAhsStrut2_Size1"))
                        strutSize = (string)((PropertyValueString)supportPart.GetPropertyValue("IJUAhsStrut2_Size1", "Strut2Size1")).PropValue;
                    else if (supportPart.SupportsInterface("IJUAhsStrutSize"))
                        strutSize = (string)((PropertyValueString)supportPart.GetPropertyValue("IJUAhsStrutSize", "StrutSize")).PropValue;
                }

                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass(sPartClass);

                foreach (BusinessObject part in partClass.Parts)
                {
                    if (((string)((PropertyValueString)part.GetPropertyValue("IJUAhsSize1", "Size1")).PropValue == strutSize))
                    {
                        selectedPart = (Part)part;
                        break;
                    }
                }

            }

            return selectedPart;
        }
        
    }
    //----------------------------------------------------------------------
    //This Rule returns part by struct size 2.
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByStrutSize2
    //---------------------------------------------------------------------- 
    public class PartByStrutSize2 : SupportPartSelectionRule
    {
        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string sPartClass)
        {
            Part selectedPart = null;
            Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
            BusinessObject supportPart = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
            if (!SupportHelper.Support.IsDesignSupportAssembly())
            {
                string strutSize = string.Empty;
                if (ReferenceIndex == 1)
                {
                    if (SupportHelper.Support.SupportsInterface("IJUAHgrURS_RS1Size2"))
                        strutSize = (string)((PropertyValueString)SupportHelper.Support.GetPropertyValue("IJUAHgrURS_RS1Size1", "Strut1Size2")).PropValue;
                    else if (SupportHelper.Support.SupportsInterface("IJUAHgrURS_Strut1Size2"))
                        strutSize = (string)((PropertyValueString)SupportHelper.Support.GetPropertyValue("IJUAHgrURS_Strut1Size2", "Strut1Size2")).PropValue;
                    else if (SupportHelper.Support.SupportsInterface("IJUAhsStrut_Strut1"))
                        strutSize = (string)((PropertyValueString)SupportHelper.Support.GetPropertyValue("IJUAhsStrut_Strut1", "Strut1Size2")).PropValue;
                    else if (supportPart.SupportsInterface("IJUAhsStrut1_Size2"))
                        strutSize = (string)((PropertyValueString)supportPart.GetPropertyValue("IJUAhsStrut1_Size2", "Strut1Size2")).PropValue;
                    else if (supportPart.SupportsInterface("IJUAhsStrutSize"))
                        strutSize = (string)((PropertyValueString)supportPart.GetPropertyValue("IJUAhsStrutSize", "StrutSize")).PropValue;
                }
                else
                {
                    if (SupportHelper.Support.SupportsInterface("IJUAhsStrut2_Size2"))
                        strutSize = (string)((PropertyValueString)SupportHelper.Support.GetPropertyValue("IJUAhsStrut2_Size2", "Strut2Size2")).PropValue;
                    else if (SupportHelper.Support.SupportsInterface("IJUAhsStrut_Strut2"))
                        strutSize = (string)((PropertyValueString)SupportHelper.Support.GetPropertyValue("IJUAhsStrut_Strut2", "Strut2Size2")).PropValue;
                    else if (supportPart.SupportsInterface("IJUAhsStrut2_Size2"))
                        strutSize = (string)((PropertyValueString)supportPart.GetPropertyValue("IJUAhsStrut2_Size2", "Strut2Size2")).PropValue;
                    else if (supportPart.SupportsInterface("IJUAhsStrutSize"))
                        strutSize = (string)((PropertyValueString)supportPart.GetPropertyValue("IJUAhsStrutSize", "StrutSize")).PropValue;
                }

                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass(sPartClass);

                foreach (BusinessObject part in partClass.Parts)
                {
                    if (((string)((PropertyValueString)part.GetPropertyValue("IJUAhsSize2", "Size2")).PropValue == strutSize))
                    {
                        selectedPart = (Part)part;
                        break;
                    }
                }

            }

            return selectedPart;
        }
    }
    //----------------------------------------------------------------------
    //This Rule returns part by size.
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartBySize
    //---------------------------------------------------------------------- 
    public class PartBySize : SupportPartSelectionRule
    {
        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string sPartClass)
        {
            Part selectedPart = null;
            if (!SupportHelper.Support.IsDesignSupportAssembly())
            {
                string Size = string.Empty;
                Size = (string)((PropertyValueString)RuleService.GetPropertyValue(SupportHelper.Support,"IJUAhsSize", "Size")).PropValue;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass(sPartClass);

                foreach (BusinessObject part in partClass.Parts)
                {
                    if (((string)((PropertyValueString)part.GetPropertyValue("IJUAhsSize", "Size")).PropValue == Size))
                    {
                        selectedPart = (Part)part;
                        break;
                    }
                }
            }
            return selectedPart;
        }
    }
 
    //----------------------------------------------------------------------
    //This Rule returns part by Pin Size
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByPinSize
    //----------------------------------------------------------------------  
    public class PartByPinSize : SupportPartSelectionRule
    {
        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string sPartClass)
        {
            Part selectedPart = null;
            if (!SupportHelper.Support.IsDesignSupportAssembly())
            {
                double pinSize = (double)((PropertyValueDouble)RuleService.GetPropertyValue(SupportHelper.Support, "IJUAhsPin1", "Pin1Diameter")).PropValue;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass(sPartClass);

                foreach (BusinessObject part in partClass.Parts)
                {
                    if ((HgrCompareDoubleService.cmpdbl(((double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsPin1", "Pin1Diameter")).PropValue), pinSize) == true))
                    {
                        selectedPart = (Part)part;
                        break;
                    }
                }
            }
            return selectedPart;
        }
    }
    //----------------------------------------------------------------------
    //This Rule returns part by Flage Size
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules.

    //----------------------------------------------------------------------  
    public class PartByFlangeSize : SupportPartSelectionRule
    {
        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string sPartClass)
        {
            Part selectedPart = null;
            if (!SupportHelper.Support.IsDesignSupportAssembly())
            {
                double flangeSize = 0;
                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    if (SupportingHelper.SupportingObjectInfo(ReferenceIndex).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam)
                        flangeSize = SupportingHelper.SupportingObjectInfo(ReferenceIndex).FlangeWidth;
                } 

                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass(sPartClass);
                bool isPartSelected = false;
                foreach (BusinessObject part in partClass.Parts)
                {
                    if ((HgrCompareDoubleService.cmpdbl(((double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsFlangeSize", "FlangeSize")).PropValue), flangeSize) == true))
                    {
                        selectedPart = (Part)part;
                        isPartSelected = true;
                        break;
                    }
                }
                if (isPartSelected == false)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "PartSelectionRule" + ": " + "WARNING: " + "No Part found in the part class " + sPartClass + " with the given parameters. Returning the best available Size", "", "PipePartSelectionRule.cs", 520);
                    int partsCount = partClass.Parts.Count;

                    double[] FlangeSize = new double[partsCount] ;
                    int i = 0;
                    double selectValue = 0;
                    foreach (BusinessObject part in partClass.Parts)
                    {
                        FlangeSize[i] = (double)((PropertyValueDouble)RuleService.GetPropertyValue(part, "IJUAhsFlangeSize", "FlangeSize")).PropValue;
                        i++;
                    }

                    Array.Sort(FlangeSize);
                    for (i = 0; i < partsCount; i++)
                    {
                        if (FlangeSize[i] >= flangeSize)
                        {
                            selectValue = FlangeSize[i];
                            break;
                        }
                    }

                    foreach (BusinessObject part in partClass.Parts)
                    {
                        double partFlangeSize = (double)((PropertyValueDouble)RuleService.GetPropertyValue(part, "IJUAhsFlangeSize", "FlangeSize")).PropValue;
                        if (HgrCompareDoubleService.cmpdbl(partFlangeSize, selectValue) == true)
                        {
                            selectedPart = (Part)part;
                            break;
                        }
                    }
                }

            }
            return selectedPart;
        }
    }
    //----------------------------------------------------------------------
    //This Rule returns part by Size and Flange Size
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartBySizeAndFlangeSize
    //----------------------------------------------------------------------  
    public class PartBySizeAndFlangeSize : SupportPartSelectionRule
    {
        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string sPartClass)
        {
             Part selectedPart = null;
             if (!SupportHelper.Support.IsDesignSupportAssembly())
             {
                 double flangeSize = 0;
                 if ((SupportHelper.SupportingObjects.Count != 0))
                 {
                     if (SupportingHelper.SupportingObjectInfo(ReferenceIndex).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam)
                         flangeSize = SupportingHelper.SupportingObjectInfo(ReferenceIndex).FlangeWidth;
                 }
                 string Size = string.Empty;
                 Size = (string)((PropertyValueString)RuleService.GetPropertyValue(SupportHelper.Support, "IJUAhsSize", "Size")).PropValue;
                 CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                 PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass(sPartClass);
                 bool isPartSelected = false;
                 foreach (BusinessObject part in partClass.Parts)
                 {
                     if ((HgrCompareDoubleService.cmpdbl(((double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsFlangeSize", "FlangeSize")).PropValue), flangeSize) == true) && ((string)((PropertyValueString)part.GetPropertyValue("IJUAhsSize", "Size")).PropValue == Size))
                     {
                         selectedPart = (Part)part;
                         isPartSelected = true;
                         break;
                     }
                 }
                 if (isPartSelected == false)
                 {
                     //Log a warning message 
                     MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "PartSelectionRule" + ": " + "WARNING: " + "No Part found in the part class " + sPartClass + " with the given parameters. Returning the best available Size", "", "PipePartSelectionRule.cs", 520);
                     int partsCount = partClass.Parts.Count;

                     double[] FlangeSize = new double[partsCount];
                     int i = 0;
                     double selectValue = 0;
                     foreach (BusinessObject part in partClass.Parts)
                     {
                         FlangeSize[i] = (double)((PropertyValueDouble)RuleService.GetPropertyValue(part, "IJUAhsFlangeSize", "FlangeSize")).PropValue;
                         i++;
                     }
                     Array.Sort(FlangeSize);
                     for (i = 0; i < partsCount;i++ )
                     {
                         if (FlangeSize[i]>=flangeSize)
                         {
                             selectValue = FlangeSize[i];
                             break;
                         }
                     }
                     foreach (BusinessObject part in partClass.Parts)
                     {
                         double partFlangeSize = (double)((PropertyValueDouble)RuleService.GetPropertyValue(part, "IJUAhsFlangeSize", "FlangeSize")).PropValue;
                         string partSize = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsSize", "Size")).PropValue;
                         if (HgrCompareDoubleService.cmpdbl(partFlangeSize , selectValue) == true)
                         {
                             selectedPart = (Part)part;
                             if (partSize == Size)
                             {
                                 break;
                             }
                         }
                     }   
                 }
             }
                return selectedPart;
        }
    }
    //----------------------------------------------------------------------
    //This Rule returns part by pipe size and Stanchion Size 
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules..PartByStanchionSizeAndPipeSizeEqual
    //----------------------------------------------------------------------  
    public class PartByStanchionSizeAndPipeSizeEqual : SupportPartSelectionRule
    {
        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string sPartClass)
        {
            Part selectedPart = null;
            if (!SupportHelper.Support.IsDesignSupportAssembly())
            {
                double StanchionSize = (double)((PropertyValueDouble)RuleService.GetPropertyValue(SupportHelper.Support,"IJUAhsStanchionSize", "StanchionSize")).PropValue;

                NominalDiameter pipeDiameter = new NominalDiameter();
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                UnitName unitName = MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units);

                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass(sPartClass);

                foreach (BusinessObject part in partClass.Parts)
                {
                    string NDUnitType = (string)((PropertyValueString)part.GetPropertyValue("IJHgrDiameterSelection", "NDUnitType")).PropValue;
                    pipeDiameter.Size = pipeInfo.NominalDiameter.Size;

                    if (NDUnitType.ToLower() == "in")
                        pipeDiameter.Size = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeDiameter.Size, unitName, UnitName.NPD_INCH);
                    else if (NDUnitType.ToLower() == "mm")
                        pipeDiameter.Size = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeDiameter.Size, unitName, UnitName.NPD_MILLIMETER);

                    if ((HgrCompareDoubleService.cmpdbl(((double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsStanchionSize", "StanchionSize")).PropValue), StanchionSize) == true && ((double)((PropertyValueDouble)part.GetPropertyValue("IJHgrDiameterSelection", "NDFrom")).PropValue <= pipeDiameter.Size) && ((double)((PropertyValueDouble)part.GetPropertyValue("IJHgrDiameterSelection", "NDTo")).PropValue >= pipeDiameter.Size)))
                    {
                        selectedPart = (Part)part;
                        break;
                    }
                }
            }
            return selectedPart;
        }
    }
    public static class RuleService
    {
        //----------------------------------------------------------------------
        //This method returns propertyValue from Assembly
        //----------------------------------------------------------------------
        public static PropertyValue GetPropertyValue(BusinessObject businessObject, string interfaceName, string propertyName)
        {
            PropertyValue propertyValue = null;
            try
            {
                if (businessObject.SupportsInterface(interfaceName))
                    propertyValue = businessObject.GetPropertyValue(interfaceName, propertyName);
                else
                {
                    BusinessObject supportPart = businessObject.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
                    if (supportPart.SupportsInterface(interfaceName))
                        propertyValue = supportPart.GetPropertyValue(interfaceName, propertyName);
                }
            }
            catch (Exception ex)
            {
                CmnException e1 = new CmnException("Invalid Attribute : " + propertyName + " Queried on Interface: " + interfaceName + " in Part Selection Rule "+  ex.Message, ex);
                throw e1;
            }
            return (propertyValue);
        }
    }
}



