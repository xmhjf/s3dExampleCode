//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014, Intergraph Corporation. All rights reserved.
//
//   RULEPAttachSel.cs
//   Marine_Assy,Ingr.SP3D.Content.Support.Rules.RULEPAttachSel
//   Author       :Vijay
//   Creation Date:05.Aug.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  05.Aug.2013     Vijay   CR-CP-224470 Convert HS_Marine_Assy to C# .Net
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;
using System.Collections;
using System.Linq;
using Ingr.SP3D.ReferenceData.Middle.Services;
using System.Collections.Generic;
using System;

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
    public class RULEPAttachSel : IHgrRule
    {
        public Array AttributeValue(BusinessObject SupportOrComponent)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)SupportOrComponent;

                SupportedHelper supportedhelper = new SupportedHelper(support);

                GenericHelper genericHelper = new GenericHelper(support);
                double[] pipeDiameter = new double[support.SupportedObjects.Count];
                int[] uBoltType = new int[support.SupportedObjects.Count];
                string[] uBoltPart = new string[support.SupportedObjects.Count];
                string supType = (string)((PropertyValueString)support.GetPropertyValue("IJOAhsMrnSupType", "SupType")).PropValue;
                string unitType = string.Empty;
                for (int i = 0; i < support.SupportedObjects.Count; i++)
                {
                    PipeObjectInfo pipe = (PipeObjectInfo)supportedhelper.SupportedObjectInfo(i + 1);
                    pipeDiameter[i] = pipe.NominalDiameter.Size;
                    unitType = pipe.NominalDiameter.Units;
                    string specName = pipe.Spec.DisplayName;
                    string pipeMaterialType = pipe.MaterialType;

                    MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;

                    //Set Frame Orientation attribute
                    CodelistItem codeList = metadataManager.GetCodelistInfo("MaterialsType", "REFDAT").GetCodelistItem(pipeMaterialType);

                    string unitCode = "1";

                    //Get UBolt Type
                    IEnumerable<BusinessObject> marinePAttachSelectionParts = null;
                    CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                    ReadOnlyCollection<BusinessObject> classItems;
                    PartClass marineServiceSectionSizeClass = (PartClass)catalogBaseHelper.GetPartClass("hsMrnSrv_PAttachType");
                    classItems = marineServiceSectionSizeClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

                    foreach (BusinessObject classItem in classItems)
                    {
                        if (((string)((PropertyValueString)classItem.GetPropertyValue("IJUAhsMrnSrvPAttachTyp", "SupLocation")).PropValue == unitCode) && (int)((PropertyValueCodelist)classItem.GetPropertyValue("IJUAhsMrnSrvPAttachTyp", "PipeMaterial")).PropValue == codeList.Value && (string)((PropertyValueString)classItem.GetPropertyValue("IJUAhsMrnSrvPAttachTyp", "SpecName")).PropValue == specName && (string)((PropertyValueString)classItem.GetPropertyValue("IJUAhsMrnSrvPAttachTyp", "NomDiaUnitType")).PropValue == unitType && (((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsMrnSrvPAttachTyp", "NomDiaFrom")).PropValue <= (pipeDiameter[i])) && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsMrnSrvPAttachTyp", "NomDiaTo")).PropValue >= (pipeDiameter[i]))))
                        {
                            uBoltType[i] = (int)((PropertyValueCodelist)classItem.GetPropertyValue("IJUAhsMrnSrvPAttachTyp", "PipeAttachment")).PropValue;
                            break;
                        }
                    }

                    PartClass marineServicePAttachSelection = (PartClass)catalogBaseHelper.GetPartClass("hsMrnSrv_PAttachSel");
                    marinePAttachSelectionParts = marineServicePAttachSelection.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                    marinePAttachSelectionParts = marinePAttachSelectionParts.Where(part => ((int)((PropertyValueCodelist)part.GetPropertyValue("IJUAhsMrnSrvPAttachSel", "PipeAttachType")).PropValue == uBoltType[i]));
                    if (marinePAttachSelectionParts.Count() > 0)
                        uBoltPart[i] = (string)((PropertyValueString)marinePAttachSelectionParts.ElementAt(0).GetPropertyValue("IJUAhsMrnSrvPAttachSel", "PipeAttachPart")).PropValue;
                }
                return uBoltPart;
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in Get Supports of RULEPAttachSel class" + ". Error:" + e.Message, e);
                throw e1;
            }
        }

    }
}