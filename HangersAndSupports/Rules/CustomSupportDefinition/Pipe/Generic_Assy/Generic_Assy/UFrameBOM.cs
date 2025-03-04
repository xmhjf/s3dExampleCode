//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   UFrameBOM.cs
//   Generic_Assy,Ingr.SP3D.Content.Support.Rules.UFrameBOM
//   Author       : Rajeswari
//   Creation Date: 03-Sep-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 03-Sep-2013  Rajeswari CR-CP-224494 Convert HS_Generic_Assy to C# .Net 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;

#region "ICustomHgrBOMDescription Members"
namespace Ingr.SP3D.Content.Support.Rules
{
    public class UFrameBOM: ICustomHgrBOMDescription
    {
        public string BOMDescription(BusinessObject supportOrComponent)
        {
            string bomDescription="";
            try
            {
                string bom = (string)((PropertyValueString)supportOrComponent.GetPropertyValue("IJOAHgrGenAssyBOMDesc", "BOM_DESC")).PropValue;
                BusinessObject part = supportOrComponent.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
                string assemblyInfoRule = (string)((PropertyValueString)part.GetPropertyValue("IJHgrSupportDefinition", "AssmInfoRule")).PropValue;
                if (assemblyInfoRule == "Generic_Assy,Ingr.SP3D.Content.Support.Rules.PenPlateAssy")
                {
                    if (string.IsNullOrEmpty(bom))
                        bomDescription = "Generic Penetration Plate Assembly";
                    else
                        bomDescription = bom;
                }
                else if (assemblyInfoRule == "Generic_Assy,Ingr.SP3D.Content.Support.Rules.UFrameTypeA")
                {
                    if (string.IsNullOrEmpty(bom))
                        bomDescription = "Assembly Type A for Single Pipe";
                    else
                        bomDescription = bom;
                }
                else if (assemblyInfoRule == "Generic_Assy,Ingr.SP3D.Content.Support.Rules.UFrameTypeB")
                {
                    if (string.IsNullOrEmpty(bom))
                        bomDescription = "Assembly Type B for Single or Multiple Pipes";
                    else
                        bomDescription = bom;
                }
                else if (assemblyInfoRule == "Generic_Assy,Ingr.SP3D.Content.Support.Rules.UFrameTypeC")
                {
                    if (string.IsNullOrEmpty(bom))
                        bomDescription = "Assembly Type C for Single or Multiple Pipes";
                    else
                        bomDescription = bom;
                }
                else if (assemblyInfoRule == "Generic_Assy,Ingr.SP3D.Content.Support.Rules.UFrameTypeD")
                {
                    if (!string.IsNullOrEmpty(bom) && bom.ToUpper() == "NONE")
                        bomDescription = "";
                    else
                        bomDescription = "Assembly Type D for Single Pipe";
                }
                else if (assemblyInfoRule == "Generic_Assy,Ingr.SP3D.Content.Support.Rules.UFrameTypeE")
                {
                    if (!string.IsNullOrEmpty(bom) && bom.ToUpper() == "NONE")
                        bomDescription = "";
                    else
                        bomDescription = "Assembly Type E for Single Pipe";
                }
                else if (assemblyInfoRule == "Generic_Assy,Ingr.SP3D.Content.Support.Rules.UFrameTypeF")
                {
                    if (!string.IsNullOrEmpty(bom) && bom.ToUpper() == "NONE")
                        bomDescription = "";
                    else
                        bomDescription = "Assembly Type F for Single Pipe";
                }
                else if (assemblyInfoRule == "Generic_Assy,Ingr.SP3D.Content.Support.Rules.UFrameTypeG")
                {
                    if (string.IsNullOrEmpty(bom))
                        bomDescription = "Assembly Type G for Single or Multiple Pipes";
                    else
                        bomDescription = bom;
                }
                else if (assemblyInfoRule == "Generic_Assy,Ingr.SP3D.Content.Support.Rules.UFrameTypeH")
                {
                    if (!string.IsNullOrEmpty(bom) && bom.ToUpper() == "NONE")
                        bomDescription = "";
                    else
                        bomDescription = "Assembly Type H for Single Pipe";
                }
                else if (assemblyInfoRule == "Generic_Assy,Ingr.SP3D.Content.Support.Rules.UFrameTypeI")
                {
                    if (!string.IsNullOrEmpty(bom) && bom.ToUpper() == "NONE")
                        bomDescription = "";
                    else
                        bomDescription = "Assembly Type I for Single Pipe";
                }
                else if (assemblyInfoRule == "Generic_Assy,Ingr.SP3D.Content.Support.Rules.UFrameTypeJ")
                {
                    if (string.IsNullOrEmpty(bom))
                        bomDescription = "Assembly Type J for Single or Multiple Pipes";
                    else
                        bomDescription = bom;
                }
                else if (assemblyInfoRule == "Generic_Assy,Ingr.SP3D.Content.Support.Rules.UFrameTypeK")
                {
                    if (string.IsNullOrEmpty(bom))
                        bomDescription = "Assembly Type K for Single or Multiple Pipes";
                    else
                        bomDescription = bom;
                }
                else if (assemblyInfoRule == "Generic_Assy,Ingr.SP3D.Content.Support.Rules.UFrameTypeL")
                {
                    if (string.IsNullOrEmpty(bom))
                        bomDescription = "Assembly Type L for Single or Multiple Pipes";
                    else
                        bomDescription = bom;
                }
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