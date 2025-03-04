//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014, Intergraph Corporation. All rights reserved.
//
//   RULEFrmOH.cs
//   Marine_Assy,Ingr.SP3D.Content.Support.Rules.RULEFrmOH
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
    public class RULEFrmOH : IHgrRule
    {
        public Array AttributeValue(BusinessObject SupportOrComponent)
        {
            Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)SupportOrComponent;

            SupportedHelper supportedhelper = new SupportedHelper(support);

            GenericHelper genericHelper = new GenericHelper(support);
            double[] pipeDiameter = new double[support.SupportedObjects.Count];
            string unitType = string.Empty;
            for (int i = 0; i < support.SupportedObjects.Count; i++)
            {
                PipeObjectInfo pipe = (PipeObjectInfo)supportedhelper.SupportedObjectInfo(i + 1);
                pipeDiameter[i] = pipe.NominalDiameter.Size;
                unitType = pipe.NominalDiameter.Units;
            }
            double largePipeDiameter = pipeDiameter.Max();
            double[] hangerLeftandRightOH = new double[2];
            IEnumerable<BusinessObject> marineServicClassParts = null;
            try
            {
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass marineServiceSectionSizeClass = (PartClass)catalogBaseHelper.GetPartClass("hsMrnSrv_FrmOH");
                marineServicClassParts = marineServiceSectionSizeClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
               
                marineServicClassParts = marineServicClassParts.Where(part => ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsMrnSrvFrmOH", "NPD")).PropValue < largePipeDiameter + 0.001) && (double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsMrnSrvFrmOH", "NPD")).PropValue > largePipeDiameter - 0.001 && ((string)((PropertyValueString)part.GetPropertyValue("IJUAhsMrnSrvFrmOH", "NPDUnitType")).PropValue == unitType));
                if (marineServicClassParts.Count() > 0)
                {
                    hangerLeftandRightOH[0] = ((double)((PropertyValueDouble)marineServicClassParts.ElementAt(0).GetPropertyValue("IJUAhsMrnSrvFrmOH", "A")).PropValue);
                    hangerLeftandRightOH[1] = ((double)((PropertyValueDouble)marineServicClassParts.ElementAt(0).GetPropertyValue("IJUAhsMrnSrvFrmOH", "B")).PropValue);
                }

                return hangerLeftandRightOH;

            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in Get Supports of RULEFrmOH class" + ". Error:" + e.Message, e);
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