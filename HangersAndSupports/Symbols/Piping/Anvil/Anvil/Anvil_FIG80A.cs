//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG80A.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG80A
//   Author       :  Manikanth
//   Creation Date:  16-05-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   15-05-2013    Manikanth CR-CP-233113-Convert HS_Anvil VB Project to C# .Net 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Support.Middle;

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
    public class Anvil_FIG80A : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG80A"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "LOAD", "LOAD", 0.999999)]
        public InputDouble m_LOAD;
        [InputDouble(3, "Q", "Q", 0.999999)]
        public InputDouble m_Q;
        [InputDouble(4, "TOTAL_TRAV", "TOTAL_TRAV", 0.999999)]
        public InputDouble m_TOTAL_TRAV;
        [InputDouble(5, "P", "P", 0.999999)]
        public InputDouble m_P;
        [InputString(6, "SIZE", "SIZE", "No Value")]
        public InputString m_SIZE;
        [InputDouble(7, "G", "G", 0.999999)]
        public InputDouble m_G;
        [InputDouble(8, "A", "A", 0.999999)]
        public InputDouble m_A;
        [InputDouble(9, "D", "D", 0.999999)]
        public InputDouble m_D;
        [InputDouble(10, "F", "F", 0.999999)]
        public InputDouble m_F;
        [InputDouble(11, "I", "I", 0.999999)]
        public InputDouble m_I;
        [InputDouble(12, "M", "M", 0.999999)]
        public InputDouble m_M;
        [InputDouble(13, "N", "N", 0.999999)]
        public InputDouble m_N;
        [InputDouble(14, "ACT_TRAVEL", "ACT_TRAVEL", 0.999999)]
        public InputDouble m_ActTravel;
        [InputDouble(15, "TRAVEL_DIR", "TRAVEL_DIR", 1)]
        public InputDouble m_TravelDir;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("SPRING", "SPRING")]
        [SymbolOutput("BOTTOM", "BOTTOM")]
        [SymbolOutput("THINGY", "THINGY")]
        public AspectDefinition m_Symbolic;

        #endregion

        #region "Construct Outputs"

        /// <summary>
        /// Construct symbol outputs in aspects.
        /// </summary>
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;

                Double load = m_LOAD.Value;
                Double Q = m_Q.Value;
                Double totalTravel = m_TOTAL_TRAV.Value;
                Double P = m_P.Value;
                Double G = m_G.Value;
                Double A = m_A.Value;
                Double D = m_D.Value;
                Double F = m_F.Value;
                Double I = m_I.Value;
                Double M = m_M.Value;
                Double N = m_N.Value;
                Double actTravel = m_ActTravel.Value;
                double J = 0, B = 0, factor = 0;
                string size = m_SIZE.Value;

                CatalogBaseHelper cataloghelper = new CatalogBaseHelper();
                PartClass anvil80AJ = (PartClass)cataloghelper.GetPartClass("Anvil_FIG80A_J");
                ReadOnlyCollection<BusinessObject> anvil80AJclassItems = anvil80AJ.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

                foreach (BusinessObject classItem in anvil80AJclassItems)
                {
                    if (((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG80A_J", "LOAD_MAX")).PropValue >= load) && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG80A_J", "LOAD_MIN")).PropValue <= load))
                    {
                        totalTravel = 0.5 * ((int)((totalTravel / 0.0254 + 0.499) * 2));
                        J = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG80A_J", "J")).PropValue;
                        break;
                    }
                }

                PartClass anvil80AB = (PartClass)cataloghelper.GetPartClass("Anvil_FIG80A_B");
                ReadOnlyCollection<BusinessObject> anvil80ABclassItems = anvil80AB.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

                foreach (BusinessObject classItem in anvil80ABclassItems)
                {
                    if (((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG80A_B", "TOTAL_TRAVEL")).PropValue > (totalTravel * 0.0254 - 0.001)) && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG80A_B", "TOTAL_TRAVEL")).PropValue < (totalTravel * 0.0254 + 0.001)))
                    {
                        B = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG80A_B", "B")).PropValue;
                        break;
                    }
                }

                Port port1 = new Port(OccurrenceConnection, part, "BotInThdRH", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                if (int.Parse(size) >= 10 && int.Parse(size) <= 18 && totalTravel < 3.75)
                    factor = 16.9375;
                if (int.Parse(size) >= 10 && int.Parse(size) <= 18 && totalTravel > 3.75)
                    factor = 19.25;
                if (int.Parse(size) >= 19 && int.Parse(size) <= 34 && totalTravel < 5.25)
                    factor = 27.9375;
                if (int.Parse(size) >= 19 && int.Parse(size) <= 34 && totalTravel > 5.25)
                    factor = 30.0625;
                if (int.Parse(size) >= 35 && int.Parse(size) <= 49 && totalTravel < 6.25)
                    factor = 32.375;
                if (int.Parse(size) >= 35 && int.Parse(size) <= 49 && totalTravel > 6.25)
                    factor = 37;
                if (int.Parse(size) >= 50 && int.Parse(size) <= 63 && totalTravel < 11.25)
                    factor = 46.5;
                if (int.Parse(size) >= 50 && int.Parse(size) <= 63 && totalTravel > 11.25)
                    factor = 51.75;
                if (int.Parse(size) >= 64 && int.Parse(size) <= 74 && totalTravel < 10.75)
                    factor = 77.625;
                if (int.Parse(size) >= 64 && int.Parse(size) <= 74 && totalTravel > 10.75)
                    factor = 77.75;
                if (int.Parse(size) >= 75 && int.Parse(size) <= 83 && totalTravel < 10.75)
                    factor = 78.1875;
                if (int.Parse(size) >= 75 && int.Parse(size) <= 83 && totalTravel > 10.75)
                    factor = 78.3125;

                double takeOut = (factor - (totalTravel / 2)) * 0.0254;

                Port port2 = new Port(OccurrenceConnection, part, "TopInThdRH", new Position(0, -B + G, takeOut), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (int.Parse(size) < 64)
                    I = B + Q;

                //Validating Inputs
                if (M <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidZeroMvalue, "M value should be greater than zero"));
                    return;
                }
                if (A <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidZeroAvalue, "A value should be greater than zero"));
                    return;
                }
                if (J <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTotalLoad, "Total load value must be between LOAD_MAX and LOAD_MIN"));
                    return;
                }
                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidZeroDvalue, "D value should be greater than zero"));
                    return;
                }
                if (N <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidZeroNvalue, "N value should be greater than zero"));
                    return;
                }

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), (new Vector(0, 0, 1)).GetOrthogonalVector());
                matrix.Translate(new Vector(-B + G, 0, takeOut + P - D));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d springCylinder = symbolGeometryHelper.CreateCylinder(null, M / 2, D);
                springCylinder.Transform(matrix);
                m_Symbolic.Outputs["SPRING"] = springCylinder;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(-B + G - M / 2 - F, -N / 2, takeOut + P - A));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d bottomBox = symbolGeometryHelper.CreateBox(null, F + M / 2 - G + I, N, A - D, 9);
                bottomBox.Transform(matrix);
                m_Symbolic.Outputs["BOTTOM"] = bottomBox;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), (new Vector(0, 0, 1)).GetOrthogonalVector());
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(0, 0, -J * 1.5));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d thingyCylinder = symbolGeometryHelper.CreateCylinder(null, J, takeOut + P - D + J * 1.5);
                thingyCylinder.Transform(matrix);
                m_Symbolic.Outputs["THINGY"] = thingyCylinder;

            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG80A"));
                    return;
                }
            }
        }
        #endregion
    }
}
