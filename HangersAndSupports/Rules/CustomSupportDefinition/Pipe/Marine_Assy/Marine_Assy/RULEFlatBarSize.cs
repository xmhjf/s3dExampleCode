//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014, Intergraph Corporation. All rights reserved.
//
//   RULEFlatBarSize.cs
//   Marine_Assy,Ingr.SP3D.Content.Support.Rules.RULEFlatBarSize
//   Author       :BS
//   Creation Date:23.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  23.Jun.2013     BS   CR-CP-224470 Convert HS_Marine_Assy to C# .Net
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Common.Middle;
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
    public class RULEFlatBarSize : IHgrRule
    {
        public Array AttributeValue(BusinessObject SupportOrComponent)
        {
            Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)SupportOrComponent;

            SupportedHelper supportedhelper = new SupportedHelper(support);

            GenericHelper genericHelper = new GenericHelper(support);
            double[] pipeDiameter = new double[support.SupportedObjects.Count];
            string supType = (string)((PropertyValueString)support.GetPropertyValue("IJOAhsMrnSupType", "SupType")).PropValue;
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
                PartClass marineServiceSectionSizeClass = (PartClass)catalogBaseHelper.GetPartClass("hsMrnSrv_FBSecSize");
                marineServicClassParts = marineServiceSectionSizeClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

                marineServicClassParts = marineServicClassParts.Where(part => ((string)((PropertyValueString)part.GetPropertyValue("IJUAhsMrnSrvFBSize", "NPDUnitType")).PropValue == unitType) && (string)((PropertyValueString)part.GetPropertyValue("IJUAhsMrnSrvFBSize", "SupType")).PropValue == supType && ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsMrnSrvFBSize", "NPDMax")).PropValue >= largePipeDiameter && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsMrnSrvFBSize", "NPDMin")).PropValue <= largePipeDiameter));
                if (marineServicClassParts.Count() > 0)
                    section = ((string)((PropertyValueString)marineServicClassParts.ElementAt(0).GetPropertyValue("IJUAhsMrnSrvFBSize", "Section")).PropValue);
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
                CmnException e1 = new CmnException("Error in Get Supports of RULEFlatBarSize class" + ". Error:" + e.Message, e);
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