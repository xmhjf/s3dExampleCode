//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_GENERIC_T.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_GENERIC_T
//   Author       :  Rajeswari
//   Creation Date:  31/10/2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   30/10/2012  Rajeswari   CR-CP-222288  Converted HS_Utility VB Project to C# .Net  
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
    public class Utility_GENERIC_T : HangerComponentSymbolDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_GENERIC_T"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "L", "L", 0.999999)]
        public InputDouble m_L;
        [InputDouble(3, "WIDTH", "WIDTH", 0.999999)]
        public InputDouble m_WIDTH;
        [InputDouble(4, "DEPTH", "DEPTH", 0.999999)]
        public InputDouble m_DEPTH;
        [InputDouble(5, "T_FLANGE", "T_FLANGE", 0.999999)]
        public InputDouble m_T_FLANGE;
        [InputDouble(6, "T_WEB", "T_WEB", 0.999999)]
        public InputDouble m_T_WEB;
        [InputString(7, "BOM_DESC", "BOM_DESC", "No Value")]
        public InputString m_BOM_DESC;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
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

                Double L = m_L.Value;
                Double width = m_WIDTH.Value;
                Double depth = m_DEPTH.Value;
                Double flangeThickness = m_T_FLANGE.Value;
                Double webThickness = m_T_WEB.Value;

                if (L == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidLength, "Length cannot be zero"));
                    return;
                }
                if (width == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidWidth, "Width cannot be zero"));
                    return;
                }
                if (depth == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidDepth, "Depth cannot be zero"));
                    return;
                }
                Port port1 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Other", new Position(0, 0, depth), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                Collection<Position> pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(0, -width / 2, 0));
                pointCollection.Add(new Position(0, width / 2, 0));
                pointCollection.Add(new Position(0, width / 2, flangeThickness));
                pointCollection.Add(new Position(0, webThickness / 2, flangeThickness));
                pointCollection.Add(new Position(0, webThickness / 2, depth));
                pointCollection.Add(new Position(0, -webThickness / 2, depth));
                pointCollection.Add(new Position(0, -webThickness / 2, flangeThickness));
                pointCollection.Add(new Position(0, -width / 2, flangeThickness));
                pointCollection.Add(new Position(0, -width / 2, 0));

                Vector projectionVector = new Vector(L, 0, 0);
                Projection3d body = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                m_Symbolic.Outputs["BODY"] = body;
            }
            catch     //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_GENERIC_T"));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomWeightCG Members"

        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {

                Double weight, cogX, cogY, cogZ, area;
                const int getSteelDensityKGPerM = 7900;
                double L = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GENERIC_T", "L")).PropValue;
                double width = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GENERIC_T", "WIDTH")).PropValue;
                double flangeThickness = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GENERIC_T", "T_FLANGE")).PropValue;
                double depth = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GENERIC_T", "DEPTH")).PropValue;
                double webThickness = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GENERIC_T", "T_WEB")).PropValue;

                area = flangeThickness * width + webThickness * (depth - flangeThickness);

                weight = area * L * getSteelDensityKGPerM;

                cogX = L / 2;
                cogY = 0;
                cogZ = depth / 2;
                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWeightCG, "Error in WeightCG of Utility_GENERIC_T"));
                }
            }
        }

        #endregion
    }
}
