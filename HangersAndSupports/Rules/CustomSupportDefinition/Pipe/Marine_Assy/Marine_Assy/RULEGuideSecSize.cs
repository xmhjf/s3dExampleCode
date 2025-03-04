//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014, Intergraph Corporation. All rights reserved.
//
//   RULEGuideSecSize.cs
//   Marine_Assy,Ingr.SP3D.Content.Support.Rules.RULEGuideSecSize
//   Author       :Vijaya
//   Creation Date:30.July.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  30.July.2013     Vijaya   CR-CP-224470 Convert HS_Marine_Assy to C# .Net
//  04.Dec.2013   Rajeswari  DI-CP-241804 Modified the code as part of hardening
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
    public class RULEGuideSecSize : IHgrRule
    {
        public Array AttributeValue(BusinessObject SupportOrComponent)
        {
            Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)SupportOrComponent;

            SupportedHelper supportedhlpr = new SupportedHelper(support);

            GenericHelper genHelper = new GenericHelper(support);
            double[] pipeDiameter = new double[support.SupportedObjects.Count];

            PropertyValueCodelist verSectionTypeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnVerSection", "VerSectionType");
            PropertyValueCodelist horSectionTypeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnHorSection", "HorSectionType");

            bool includeHorSection = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnIncHorSec", "IncludeHorSec")).PropValue;
            string unitType = string.Empty, section = string.Empty, steelStandard = string.Empty;
            for (int i = 0; i < support.SupportedObjects.Count; i++)
            {
                PipeObjectInfo pipe = (PipeObjectInfo)supportedhlpr.SupportedObjectInfo(i + 1);
                pipeDiameter[i] = pipe.NominalDiameter.Size;
                unitType = pipe.NominalDiameter.Units;
            }
            double largePipeDiameter = pipeDiameter.Max();

            IEnumerable<BusinessObject> marineSrvClassParts = null, frmOrientClassParts = null, marineVerticalClassParts = null, marineHorizontalClassParts = null, frmOrientSectionClassParts = null;
            try
            {
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();

                PartClass marineSrv_SecSizeClass = (PartClass)catalogBaseHelper.GetPartClass("hsMrnSrv_SecSize");
                marineSrvClassParts = marineSrv_SecSizeClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

                marineVerticalClassParts = marineSrvClassParts.Where(part => ((string)((PropertyValueString)part.GetPropertyValue("IJUAhsMrnSrvSecSize", "NPDUnitType")).PropValue == unitType) && ((PropertyValueCodelist)part.GetPropertyValue("IJUAhsMrnSrvSecSize", "SectionType")).PropValue == verSectionTypeCodeList.PropValue && ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsMrnSrvSecSize", "NPDMax")).PropValue >= largePipeDiameter && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsMrnSrvSecSize", "NPDMin")).PropValue <= largePipeDiameter));

                if (marineVerticalClassParts.Count() > 0)
                    section = ((string)((PropertyValueString)marineVerticalClassParts.ElementAt(0).GetPropertyValue("IJUAhsMrnSrvSecSize", "Section")).PropValue);
                genHelper.GetDataByRule("hsMrnSteelStandardName", null, out steelStandard);

                PartClass marineSrv_FrmCorrespClass = (PartClass)catalogBaseHelper.GetPartClass("hsMrnSrv_FrmCorresp");
                frmOrientClassParts = marineSrv_FrmCorrespClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

                frmOrientSectionClassParts = frmOrientClassParts.Where(part => ((string)((PropertyValueString)part.GetPropertyValue("IJUAhsMrnFrmCorresp", "Size")).PropValue == section) && (string)((PropertyValueString)part.GetPropertyValue("IJUAhsMrnFrmCorresp", "StdName")).PropValue == steelStandard);
                string[] sectionSize = new string[2];

                if (frmOrientSectionClassParts.Count() > 0)
                    sectionSize[0] = ((string)((PropertyValueString)frmOrientSectionClassParts.ElementAt(0).GetPropertyValue("IJUAhsMrnFrmCorresp", "SectionSize")).PropValue);

                if (includeHorSection == true)
                {
                    marineHorizontalClassParts = marineSrvClassParts.Where(part => ((string)((PropertyValueString)part.GetPropertyValue("IJUAhsMrnSrvSecSize", "NPDUnitType")).PropValue == unitType) && ((PropertyValueCodelist)part.GetPropertyValue("IJUAhsMrnSrvSecSize", "SectionType")).PropValue == horSectionTypeCodeList.PropValue && ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsMrnSrvSecSize", "NPDMax")).PropValue >= largePipeDiameter && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsMrnSrvSecSize", "NPDMin")).PropValue <= largePipeDiameter));
                    if (marineHorizontalClassParts.Count() > 0)
                        section = ((string)((PropertyValueString)marineHorizontalClassParts.ElementAt(0).GetPropertyValue("IJUAhsMrnSrvSecSize", "Section")).PropValue);

                    frmOrientSectionClassParts = frmOrientClassParts.Where(part => ((string)((PropertyValueString)part.GetPropertyValue("IJUAhsMrnFrmCorresp", "Size")).PropValue == section) && (string)((PropertyValueString)part.GetPropertyValue("IJUAhsMrnFrmCorresp", "StdName")).PropValue == steelStandard);
                    if (frmOrientClassParts.Count() > 0)
                        sectionSize[1] = ((string)((PropertyValueString)frmOrientSectionClassParts.ElementAt(0).GetPropertyValue("IJUAhsMrnFrmCorresp", "SectionSize")).PropValue);
                }
                return sectionSize;

            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in Get Supports of RULEGuideSecSize class" + ". Error:" + e.Message, e);
                throw e1;
            }
            finally
            {
                if (marineSrvClassParts is IDisposable)
                {
                    ((IDisposable)marineSrvClassParts).Dispose(); // This line will be executed
                }
                if (frmOrientSectionClassParts is IDisposable)
                {
                    ((IDisposable)frmOrientSectionClassParts).Dispose(); // This line will be executed
                }
            }
        }

    }
}