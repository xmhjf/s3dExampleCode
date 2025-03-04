//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_GUSSET2.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_GUSSET2
//   Author       :  Rajeswari
//   Creation Date:  31/10/1012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   31/10/1012   Rajeswari  CR-CP-222288  Converted HS_Utility VB Project to C# .Net  
//	 27/03/2013		Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   04/June/2013   Manikanth TR-CP-234520 Implemented TDL Issues
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;

namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------

    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class Utility_GUSSET2 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_GUSSET2"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "THICKNESS", "THICKNESS", 0.999999)]
        public InputDouble m_THICKNESS;
        [InputDouble(3, "W", "W", 0.999999)]
        public InputDouble m_W;
        [InputDouble(4, "H", "H", 0.999999)]
        public InputDouble m_H;
        [InputDouble(5, "VERTICAL_INSET", "VERTICAL_INSET", 0.999999)]
        public InputDouble m_VERTICAL_INSET;
        [InputDouble(6, "HORIZONTAL_INSET", "HORIZONTAL_INSET", 0.999999)]
        public InputDouble m_HORIZONTAL_INSET;
        [InputString(7, "BOM_DESC", "BOM_DESC","No Value")]
        public InputString m_BOM_DESC;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("BODY", "BODY")]
        public AspectDefinition m_Symbolic;

        #endregion

        #region "Construct Outputs"

        /// <summary>
        /// Construct symbol outputs in aspects.
        /// </summary>
        /// <remarks></remarks>
        protected override void ConstructOutputs()
        {
            try
            {
                Part part =(Part)m_PartInput.Value ;

                Double T = m_THICKNESS.Value;
                Double W = m_W.Value;
                Double H = m_H.Value;
                Double verticalInset = m_VERTICAL_INSET.Value;
                Double horizontalInset = m_HORIZONTAL_INSET.Value;

                if (T == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidT, "Thickness cannot be zero"));
                    return;
                }
                if (H == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidHeight, "Height cannot be zero"));
                    return;
                }
                if (W == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidWidth, "Width cannot be zero"));
                    return;
                }
                Port port1 = new Port(OccurrenceConnection, part, "Other", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Collection<Position> pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(-T / 2, 0, 0));
                pointCollection.Add(new Position(-T / 2, W, -H + verticalInset));
                pointCollection.Add(new Position(-T / 2, W, -H));
                pointCollection.Add(new Position(-T / 2, horizontalInset, -H));
                pointCollection.Add(new Position(-T / 2, 0, -H + horizontalInset));
                pointCollection.Add(new Position(-T / 2, 0, 0));

                Vector projectionVector = new Vector(T, 0, 0);
                Projection3d body = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                m_Symbolic.Outputs["BODY"] = body;
            }
            catch     //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_GUSSET2"));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomString = "";
            try
            {
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                string bomDescription = (string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GUSSET2", "BOM_DESC")).PropValue;
                if (bomDescription == "")
                {
                    bomString = part.PartDescription;
                }
                else
                {
                    bomString = bomDescription;
                }

                return bomString;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Utility_GUSSET2"));
                }
                return "";
            }
        }
        #endregion

        #region "ICustomWeightCG Members"

        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {

                Double weight, cogX, cogY, cogZ;
                const int getSteelDensityKGPerM = 7900;
                double H = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GUSSET2", "H")).PropValue;
                double W = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GUSSET2", "W")).PropValue;
                double verticalInset = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GUSSET2", "VERTICAL_INSET")).PropValue;
                double horizontalInset = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GUSSET2", "HORIZONTAL_INSET")).PropValue;

                Part part = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                double T = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrUtility_GUSSET2", "THICKNESS")).PropValue;

                weight = ((W * ((H - verticalInset) / 2) * T) + (W * verticalInset * T) - (horizontalInset * (horizontalInset / 2) * T)) * getSteelDensityKGPerM;
                cogX = 0;
                cogY = 0;
                cogZ = 0;
                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWeightCG, "Error in WeightCG of Utility_GUSSET2"));
                }
            }
        }

        #endregion
    }
}
