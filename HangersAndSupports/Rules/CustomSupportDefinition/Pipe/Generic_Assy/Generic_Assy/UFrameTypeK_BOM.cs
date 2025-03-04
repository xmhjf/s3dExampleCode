//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   LFieldSupports.cs
//   Generic_Assy,Ingr.SP3D.Content.Support.Rules.UFrameBOM
//   Author       : Vijay
//   Creation Date: 19-Sep-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   19-Sep-2013    Vijay    CR-CP-224494 Convert HS_Generic_Assy to C# .Net 
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
    public class UFrameTypeK_BOM : ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject supportOrComponent)
        {
            string bomDescription = "";
            try
            {
                string bom = (string)((PropertyValueString)supportOrComponent.GetPropertyValue("IJOAHgrGenAssyBOMDesc", "BOM_DESC")).PropValue;
                if (string.IsNullOrEmpty(bom))
                    bomDescription = "Assembly Type K for Single or Multiple Pipes";
                else
                    bomDescription = bom;
                return bomDescription;
            }
            catch (Exception e)  //General Unhandled exception 
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error BOM Description." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
    }
}
#endregion