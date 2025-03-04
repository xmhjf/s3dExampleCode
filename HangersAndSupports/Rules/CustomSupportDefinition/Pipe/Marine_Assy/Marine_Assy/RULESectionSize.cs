//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014, Intergraph Corporation. All rights reserved.
//
//   RULESectionSize.cs
//   Marine_Assy,Ingr.SP3D.Content.Support.Rules.RULESectionSize
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
    public class RULESectionSize : IHgrRule
    {
        public Array AttributeValue(BusinessObject SupportOrComponent)
        {
            Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)SupportOrComponent;

            SupportedHelper supportedhelper = new SupportedHelper(support);

            GenericHelper genericHelper = new GenericHelper(support);
            double[] pipeDiameter = new double[support.SupportedObjects.Count];
            string supType = (string)((PropertyValueString)support.GetPropertyValue("IJOAhsMrnSupType", "SupType")).PropValue;
            PropertyValueCodelist sectionTypeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnSection", "SectionType");
            string sectionType = sectionTypeCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(sectionTypeCodeList.PropValue).ShortDisplayName;
            string unitType = string.Empty;
            for (int i = 0; i < support.SupportedObjects.Count; i++)
            {
                PipeObjectInfo pipe = (PipeObjectInfo)supportedhelper.SupportedObjectInfo(i + 1);
                pipeDiameter[i] = pipe.NominalDiameter.Size;
                unitType = pipe.NominalDiameter.Units;
            }
            double largePipeDiameter = pipeDiameter.Max();
            string section = string.Empty;
            string steelStanded = string.Empty;
            IEnumerable<BusinessObject> marineServicClassParts = null;
            try
            {
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                ReadOnlyCollection<BusinessObject> classItems;
                PartClass marineServiceSectionSizeClass = (PartClass)catalogBaseHelper.GetPartClass("hsMrnSrv_SecSize");
                classItems = marineServiceSectionSizeClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

                foreach (BusinessObject classItem in classItems)
                {
                    if (((string)((PropertyValueString)classItem.GetPropertyValue("IJUAhsMrnSrvSecSize", "NPDUnitType")).PropValue == unitType) && ((int)((PropertyValueCodelist)classItem.GetPropertyValue("IJUAhsMrnSrvSecSize", "SectionType")).PropValue == sectionTypeCodeList.PropValue) && (((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsMrnSrvSecSize", "NPDMax")).PropValue >= largePipeDiameter) && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsMrnSrvSecSize", "NPDMin")).PropValue <= largePipeDiameter)))
                    {
                        section = (string)((PropertyValueString)classItem.GetPropertyValue("IJUAhsMrnSrvSecSize", "Section")).PropValue;
                        break;
                    }
                }
                genericHelper.GetDataByRule("hsMrnSteelStandardName", null, out steelStanded);

                PartClass marineServiceFromCorrespondenceClass = (PartClass)catalogBaseHelper.GetPartClass("hsMrnSrv_FrmCorresp");
                marineServicClassParts = marineServiceFromCorrespondenceClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

                marineServicClassParts = marineServicClassParts.Where(part => ((string)((PropertyValueString)part.GetPropertyValue("IJUAhsMrnFrmCorresp", "Size")).PropValue == section) && (string)((PropertyValueString)part.GetPropertyValue("IJUAhsMrnFrmCorresp", "StdName")).PropValue == steelStanded);
                string[] sectionSize = new string[1];
                if (marineServicClassParts.Count() > 0)
                    sectionSize[0] = ((string)((PropertyValueString)marineServicClassParts.ElementAt(0).GetPropertyValue("IJUAhsMrnFrmCorresp", "SectionSize")).PropValue);

                return sectionSize;

            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in Get Supports of RULESectionSize class" + ". Error:" + e.Message, e);
                throw e1;
            }
            finally
            {
                if (marineServicClassParts is IDisposable)
                {
                    ((IDisposable)marineServicClassParts).Dispose(); // This line will be executed
                }
            }
        }

    }
}