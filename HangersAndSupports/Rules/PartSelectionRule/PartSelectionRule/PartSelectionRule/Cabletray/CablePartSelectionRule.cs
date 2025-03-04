//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014, Intergraph Corporation. All rights reserved.
//
//   CablePartSelectionRule.cs
//   Author       :Vijaya
//   Creation Date:25.Sep.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  25.Sep.2013     Vijaya   CR-CP-233070  Convert HgrCablePartSelRule to C# .Net 
//  22-Jan-2015     PVK      TR-CP-264951  Resolve coverity issues found in November 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using Ingr.SP3D.Support.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Support.Symbols;

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
    //This Rule retuns Bline part based on cable tray size.
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules.BlinePartByCTSize
    //----------------------------------------------------------------------  
    public class BlinePartByCTSize : SupportPartSelectionRule
    {
        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string PartClass)
        {
            CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
            CableTrayObjectInfo cableTray = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(1);
            double diameter = 0.0;
            if (cableTray.Width <= 0 || cableTray.Depth <= 0)
                diameter = cableTray.BendRadius * 2;
            else
                diameter = cableTray.Width;

            //Make sure the Unit type is either inch or mm
            diameter = Math.Ceiling(MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, diameter, UnitName.DISTANCE_METER, UnitName.DISTANCE_INCH));
            PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass(PartClass);
            Part partByWidth = null;

            foreach (BusinessObject part in partClass.Parts)
            {
                if (HgrCompareDoubleService.cmpdbl(((double)((PropertyValueDouble)part.GetPropertyValue("IJHgrDiameterSelection", "NDTo")).PropValue) , diameter)==true)
                {
                    partByWidth = (Part)part;
                    break;
                }
            }
            return partByWidth;
        }
    }
    //----------------------------------------------------------------------
    //This Rule retuns Bline part based on width and depth values.
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules.BlinePartByWdtDpt
    //----------------------------------------------------------------------  
    public class BlinePartByWdtDpt : SupportPartSelectionRule
    {
        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string PartClass)
        {
            CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
            CableTrayObjectInfo cableTray = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(1);
            double width = 0.0, depth = 0.0;
            if (cableTray.Width <= 0 || cableTray.Depth <= 0)
            {
                width = cableTray.BendRadius * 2;
                depth = cableTray.BendRadius * 2;
            }
            else
            {
                width = cableTray.Width;
                depth = cableTray.Depth;
            }

            string auxilaryTable = string.Empty, interfaceName = string.Empty, partName = string.Empty, partNumber = string.Empty;
            if (PartClass.ToLower().Equals("hgrrodclamp9_532x"))
            {
                auxilaryTable = "HgrRodClamp_AUX";
                interfaceName = "IJUAHgrBlineHRClampAUX";
            }
            else if (PartClass.ToLower().Equals("verthgr_9_122x"))
            {
                auxilaryTable = "VertHanger_AUX";
                interfaceName = "IJUAHgrBlineVertHgrAUX";
            }

            PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass(auxilaryTable);
            ReadOnlyCollection<BusinessObject> classItems = partClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

            foreach (BusinessObject item in classItems)
            {
                if (((double)((PropertyValueDouble)item.GetPropertyValue(interfaceName, "CTWidth")).PropValue) > (width - 0.0001) && ((double)((PropertyValueDouble)item.GetPropertyValue(interfaceName, "CTWidth")).PropValue) < (width + 0.0001) && ((double)((PropertyValueDouble)item.GetPropertyValue(interfaceName, "CTDepth")).PropValue) > (depth - 0.0001) && ((double)((PropertyValueDouble)item.GetPropertyValue(interfaceName, "CTDepth")).PropValue) < (depth + 0.0001))
                {
                    partName = (string)((PropertyValueString)item.GetPropertyValue(interfaceName, "PartNo")).PropValue;
                    break;
                }
            }
            partClass = (PartClass)catalogBaseHelper.GetPartClass(PartClass);
            Part partByWidthDepth = null, part = null;
            foreach (BusinessObject item in partClass.Parts)
            {
                part = (Part)item;
                if (part.PartNumber == partName)
                {
                    partByWidthDepth = part;
                    break;
                }
            }

            return partByWidthDepth;

        }
    }
    //----------------------------------------------------------------------
    //This Rule retuns part by cable tray size.
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByCTSize
    //----------------------------------------------------------------------  
    public class PartByCTSize : SupportPartSelectionRule
    {
        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string PartClass)
        {
            CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
            CableTrayObjectInfo cableTray = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(1);
            double diameter = 0.0;
            if (cableTray.Width <= 0 || cableTray.Depth <= 0)
                diameter = cableTray.BendRadius * 2;
            else
                diameter = cableTray.Width;

            //Make sure the Unit type is either inch or mm
            diameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, diameter, UnitName.DISTANCE_METER, UnitName.DISTANCE_INCH);
            PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass(PartClass);
            Part partByWidth = null;

            foreach (BusinessObject part in partClass.Parts)
            {
                if (((double)((PropertyValueDouble)part.GetPropertyValue("IJHgrDiameterSelection", "NDTo")).PropValue) >= diameter)
                {
                    partByWidth = (Part)part;
                    break;
                }
            }
            return partByWidth;
        }
    }
    //----------------------------------------------------------------------
    //This Rule retuns part by cable tray width.
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules.CPartByCTWidth
    //----------------------------------------------------------------------  
    public class CPartByCTWidth : SupportPartSelectionRule
    {
        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string PartClass)
        {
            CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
            CableTrayObjectInfo cableTray = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(1);
            double width = 0.0;
            if (cableTray.Width <= 0 || cableTray.Depth <= 0)
                width = cableTray.BendRadius * 2;
            else
                width = cableTray.Width;

            PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass(PartClass);
            Part partByWidth = null;

            foreach (BusinessObject part in partClass.Parts)
            {
                if (((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrCTNominalDim", "NominalTrayWidthTo")).PropValue) >= width)
                {
                    partByWidth = (Part)part;
                    break;
                }
            }
            return partByWidth;
        }
    }
}