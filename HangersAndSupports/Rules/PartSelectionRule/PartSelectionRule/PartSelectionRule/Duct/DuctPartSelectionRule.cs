//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014, Intergraph Corporation. All rights reserved.
//
//   DuctPartSelectionRule.cs
//   Author       :Vijaya
//   Creation Date:25.Sep.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  25.Sep.2013     Vijaya   CR-CP-233073  Convert HgrDuctPartSelRule to C# .Net  
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;

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
    //This Rule retuns Part based on duct diameter.
    //ProgId : PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByDuctDiam
    //----------------------------------------------------------------------  
    public class PartByDuctDiam : SupportPartSelectionRule
    {
        public override Ingr.SP3D.ReferenceData.Middle.Part SelectedPartFromPartClass(string PartClass)
        {
            CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
            DuctObjectInfo duct = (DuctObjectInfo)SupportedHelper.SupportedObjectInfo(1);
            double width = 0.0;
            if (duct.BendRadius > 0.0 )
                width = duct.BendRadius * 2;
            else
                width = duct.Width;

            //Make sure the Unit type is either inch or mm
            width = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, width, UnitName.DISTANCE_METER, UnitName.DISTANCE_INCH);
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
}
