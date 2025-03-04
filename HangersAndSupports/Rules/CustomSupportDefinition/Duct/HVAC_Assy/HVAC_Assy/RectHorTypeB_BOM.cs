//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   RectHorTypeB_BOM.cs
//   HVAC_Assy,Ingr.SP3D.Content.Support.Rules.RectHorTypeB_BOM
//   Author       : Hema
//   Creation Date: 04-Dec-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   04-Dec-2013    Hema    CR-CP-224486 Convert HS_HVAC_Assy to C# .Net
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;


#region "ICustomHgrBOMDescription Members"
//-----------------------------------------------
// BOM Description
//-----------------------------------------------
namespace Ingr.SP3D.Content.Support.Rules
{
    public class RectHorTypeB_BOM : ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject SupportOrComponent)
        {
            string bomString = string.Empty;
            try
            {
                string bomDescription = (string)((PropertyValueString)SupportOrComponent.GetPropertyValue("IJOAHgrHVACAssyBOMDesc", "BOM_DESC")).PropValue;
                
                if (string.IsNullOrEmpty(bomDescription))
                    bomString = "Horizontal Rectangular Duct Assembly Type B for Single or Multiple Ducts";
                else
                    bomString = bomDescription;

                return bomString;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly  RectHorTypeB_BOM." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
    }
}
#endregion