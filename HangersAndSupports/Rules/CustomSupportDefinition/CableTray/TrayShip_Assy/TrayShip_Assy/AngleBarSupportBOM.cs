//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   AngleBarSupp.cs
//   TrayShip_Assy,Ingr.SP3D.Content.Support.Rules.AngleBarSupportBOM
//   Author       :  Vijay
//   Creation Date:  12/07/2013   
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   12/07/2013     Vijay    CR-CP-224487  Convert HS_TrayShip_Assy to C# .Net  
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
//-----------------------------------------------
// BOM Description
//-----------------------------------------------
#region "ICustomHgrBOMDescription Members"
namespace Ingr.SP3D.Content.Support.Rules
{
    public class AngleBarSupportBOM : ICustomHgrBOMDescription
    {

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string BOMString = "";
            try
            {
                double length11 = Ingr.SP3D.Content.Support.Rules.AngleBarSupp.Length11;
                double length22 = Ingr.SP3D.Content.Support.Rules.AngleBarSupp.Length22;

                bool includeSecsizeInBOM = (bool)((PropertyValueBoolean)oSupportOrComponent.GetPropertyValue("IJUAHgrAngleBar", "InclSecsizeInBOM")).PropValue;
                PropertyValueCodelist sectionSizeCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJUAHgrSecSize", "SectionSize");
                string sectionSize = sectionSizeCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(sectionSizeCodelist.PropValue).ShortDisplayName;
                
                BOMString = ((Ingr.SP3D.Support.Middle.Support)oSupportOrComponent).SupportDefinition.PartNumber + ", L1 = " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, length11, UnitName.DISTANCE_MILLIMETER) + ", L2 = " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, length22, UnitName.DISTANCE_MILLIMETER);

                if (includeSecsizeInBOM)
                    BOMString = BOMString + " " + sectionSize;

                return BOMString;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
    }
}
#endregion