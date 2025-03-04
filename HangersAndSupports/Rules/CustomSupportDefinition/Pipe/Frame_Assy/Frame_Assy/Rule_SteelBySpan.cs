//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   RULEGuideSecSize.cs
//   Frame_Assy,Ingr.SP3D.Content.Support.Rules.Rule_SteelBySpan
//   Author       :Hema
//   Creation Date:30.July.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  30.July.2013     Hema    CR-CP-224474 Convert HS_S3DFrame to C# .Net
//  06.May.2015      PVK     CR-CP-270669	Modify the exsisting .Net Frame_Assy to honor attributes which are not in OA
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Support.Middle;

namespace Ingr.SP3D.Content.Support.Rules
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Rules
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    public class Rule_SteelBySpan : IHgrRule
    {
        public Array AttributeValue(BusinessObject SupportOrComponent)
        {
            Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)SupportOrComponent;

            SupportHelper supportHelper = new SP3D.Support.Middle.SupportHelper(support);
            RefPortHelper refPortHelper=new SP3D.Support.Middle.RefPortHelper(support);
            int numberOfStructs = supportHelper.SupportingObjects.Count;
            BusinessObject catalogpart = SupportOrComponent.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
            double span;
            string steelIID = string.Empty; 
            string[] partNumber=new string[1];
            if (supportHelper.SupportingObjects.Equals(null))
            {
                if (numberOfStructs == 1)
                    span = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsFrameSpan", "SpanValue")).PropValue;
                else
                    span = refPortHelper.DistanceBetweenPorts("Structure", "Struct_2", PortDistanceType.Horizontal);
            }
            else
                span = refPortHelper.DistanceBetweenPorts("Structure", "Struct_2", PortDistanceType.Horizontal);

            PropertyValueCodelist steelStandardCodeList = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(catalogpart,"IJUAhsSteelStandard", "SteelStandard");
            string standard = steelStandardCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(steelStandardCodeList.PropValue).DisplayName;

           CatalogBaseHelper cataloghelper = new CatalogBaseHelper();
           PartClass frameSteelBySpan = (PartClass)cataloghelper.GetPartClass("hsS3D_FrameSteelBySpan");
           ReadOnlyCollection<BusinessObject> partsFrameSteelBySpan = frameSteelBySpan.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

           foreach (BusinessObject part in partsFrameSteelBySpan)
            {
                if ((double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrameSteelBySpan", "MaxSpan")).PropValue > span && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsFrameSteelBySpan", "MinSpan")).PropValue <= span)
                {
                    steelIID = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrameSteelBySpan", "SteelID")).PropValue;
                    break;
                }
            }
           PartClass frameSteelEqual = (PartClass)cataloghelper.GetPartClass("hsS3D_FrameSteelEqual");
           ReadOnlyCollection<BusinessObject> parts1FrameSteelEqual = frameSteelEqual.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

            foreach (BusinessObject part in parts1FrameSteelEqual)
            {
                PropertyValueCodelist frameSteelEquivalent = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrameSteelEquivalent", "SteelStandard");
                if ((steelIID.Equals((string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrameSteelEquivalent", "SteelID")).PropValue) && frameSteelEquivalent.PropValue == steelStandardCodeList.PropValue))
                {
                    partNumber[0] = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrameSteelEquivalent", "PartNumber")).PropValue;
                    break;
                }
            }
            return partNumber;
        }
    }
}