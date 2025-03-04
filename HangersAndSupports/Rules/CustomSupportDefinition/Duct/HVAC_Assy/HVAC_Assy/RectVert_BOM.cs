﻿//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   RectVert_BOM.cs
//   HVAC_Assy,Ingr.SP3D.Content.Support.Rules.RectVert_BOM
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
    public class RectVert_BOM : ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            try
            {
                string bomDescription;
                try
                {
                    bomDescription = (string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrHVACAssyBOMDesc", "BOM_DESC")).PropValue;
                }
                catch
                {
                    bomDescription = string.Empty;
                }

                if (string.IsNullOrEmpty(bomDescription))
                    bomDescription = "Vertical Rectangular Duct Assembly for Single or Multiple Ducts";

                return bomDescription;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOMDescription RectVert_BOM" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
    }
}
#endregion