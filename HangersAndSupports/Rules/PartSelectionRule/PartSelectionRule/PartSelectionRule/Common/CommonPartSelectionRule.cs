//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014, Intergraph Corporation. All rights reserved.
//
//   CommonPartSelectionRule.cs
//   Author       :Vijaya
//   Creation Date:26.Sep.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  26.Sep.2013     Vijaya   CR-CP-233071  Convert HgrPartSelRule to C# .Net    
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using Ingr.SP3D.Support.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Route.Middle;

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
    //This Rule retuns part based on cross section data.
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules.CrossSectionByCW
    //----------------------------------------------------------------------  
    public class CrossSectionByCW : SupportPartSelectionRule
    {
        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string PartClass)
        {
            Collection<object> valuesCollection = new Collection<object>();
            //string versionAISC = string.Empty;
            string[] sectionData = new string[3];
            try
            {
                bool value = GenericHelper.GetDataByRule("HgrAISCversion", null, out valuesCollection);
            }
            catch { }
            if (valuesCollection != null)
            {
                if (string.IsNullOrEmpty((string)valuesCollection[0]))
                    sectionData[0] = "AISC-LRFD-3.0";
                else
                    sectionData[0] = (string)valuesCollection[0];
            }
            sectionData[1] = "L";
            sectionData[2] = "L2x2x3/16";

            return SupportPartByAssociation(PartClass, 3, sectionData);
        }

    }
    //----------------------------------------------------------------------
    //This Rule retuns first part from a collection.
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules.CFirstPart
    //----------------------------------------------------------------------  
    public class CFirstPart : SupportPartSelectionRule
    {
        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string PartClass)
        {
            CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
            PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass(PartClass);

            return (Part)partClass.Parts[0];
        }
    }
    //----------------------------------------------------------------------
    //This Rule retuns part based on cross section.
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules.CPartByCrossSection
    //----------------------------------------------------------------------  
    public class CPartByCrossSection : SupportPartSelectionRule
    {
        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string PartClass)
        {
            Collection<object> valuesCollection = new Collection<object>();
            string versionAISC = string.Empty;
            string[] sectionData = new string[3];

            bool value = GenericHelper.GetDataByRule("HgrSupPrimaryCSSelection", null, out valuesCollection);
            if (valuesCollection != null)
            {
                sectionData[0] = (string)valuesCollection[0];
                sectionData[1] = (string)valuesCollection[1];
                sectionData[2] = (string)valuesCollection[2];
            }
            return SupportPartByAssociation(PartClass, 3, sectionData);
        }

    }
    //----------------------------------------------------------------------
    //This Rule retuns part based on cable tray width.
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules.CPartByCWWidth
    //----------------------------------------------------------------------  
    public class CPartByCWWidth : SupportPartSelectionRule
    {
        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string PartClass)
        {
            CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
            CableTrayObjectInfo cableTray = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(1);
            double width = 0.0;
            if (cableTray.Width <= 0)
                width = cableTray.BendRadius * 2;
            else
                width = cableTray.Width;
            PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass(PartClass);
            Part partByWidth = null;
            foreach (BusinessObject part in partClass.Parts)
            {
                if (((double)((PropertyValueDouble)part.GetPropertyValue("IJHgrDiameterSelection", "NDTo")).PropValue) >= width)
                {
                    partByWidth = (Part)part;
                    break;
                }
            }
            return partByWidth;
        }

    }
    //----------------------------------------------------------------------
    //This Rule retuns part based on profile type.
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules.CPartByProfile
    //----------------------------------------------------------------------  
    public class CPartByProfile : SupportPartSelectionRule
    {
        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string PartClass)
        {
            CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
            PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass(PartClass);
            Part partByProfile = null;

            foreach (BusinessObject part in partClass.Parts)
            {
                if (((string)((PropertyValueString)part.GetPropertyValue("IJUAHgrBracketType", "BracketType")).PropValue).Equals("TwoProfiles"))
                {
                    partByProfile = (Part)part;
                    break;
                }
            }
            return partByProfile;
        }
    }
    //----------------------------------------------------------------------
    //This Rule retuns part based on secondary cross section.
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules.CPartBySecondaryCS
    //----------------------------------------------------------------------  
    public class CPartBySecondaryCS : SupportPartSelectionRule
    {
        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string PartClass)
        {
            Collection<object> valuesCollection = new Collection<object>();
            string versionAISC = string.Empty;
            string[] sectionData = new string[3];
            try
            {

                bool value = GenericHelper.GetDataByRule("HgrAISCversion", null, out valuesCollection);
            }
            catch { }
            if (valuesCollection != null)
            {
                if (string.IsNullOrEmpty((string)valuesCollection[0]))
                    sectionData[0] = "AISC-LRFD-3.0";
                else
                    sectionData[0] = (string)valuesCollection[0];
            }
            sectionData[1] = "L";
            sectionData[2] = "L2x2x3/16";

            return SupportPartByAssociation(PartClass, 3, sectionData);
        }

    }
    //----------------------------------------------------------------------
    //This Rule retuns part based on width and height values.
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules.CPartByWidthHeight
    //----------------------------------------------------------------------  
    public class CPartByWidthHeight : SupportPartSelectionRule
    {
        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string PartClass)
        {
            CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
            BusinessObject supportedObject = SupportHelper.SupportedObjects[0];
            double width = 0.0;
            if (SupportHelper.SupportedObjects.Count > 0)
            {
                IPipePathFeature route = supportedObject as IPipePathFeature;
                if (route != null)                
                    width = route.NPD.Size;               
                IConduitPathFeature conduitInfo = supportedObject as IConduitPathFeature;
                if (conduitInfo != null)
                 width =  conduitInfo.NCD.Size;
                IRouteFeatureWithCrossSection ductInfo = supportedObject as IRouteFeatureWithCrossSection;
                if (ductInfo != null)
                    width = ductInfo.Width;
            }
            PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass(PartClass);
            Part partByWidth = null;

            foreach (BusinessObject part in partClass.Parts)
            {
                if (((double)((PropertyValueDouble)part.GetPropertyValue("IJHgrDiameterSelection", "NDTo")).PropValue) >= width)
                {
                    partByWidth = (Part)part;
                    break;
                }
            }
            return partByWidth;
        }
    }
    //----------------------------------------------------------------------
    //This Rule retuns design support part.
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules.CPartForAllDiscipline
    //----------------------------------------------------------------------  
    public class CPartForAllDiscipline : SupportPartSelectionRule
    {

        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string PartClass)
        {
            CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
            PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass(PartClass);
            Part finalPart = null;
            foreach (BusinessObject part in partClass.Parts)
            {
                finalPart = (Part)part;
                if (finalPart.PartNumber == "HgrDesignSup_P01")
                {
                    finalPart = (Part)part;
                    break;
                }
            }
            return finalPart;
        }

    }
}
