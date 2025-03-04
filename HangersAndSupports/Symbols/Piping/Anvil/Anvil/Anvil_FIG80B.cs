//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG80B.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG80B
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
    public class Anvil_FIG80B : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG80B"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "LOAD", "LOAD", 0.999999)]
        public InputDouble m_LOAD;
        [InputDouble(3, "R", "R", 0.999999)]
        public InputDouble m_R;
        [InputDouble(4, "TOTAL_TRAV", "TOTAL_TRAV", 0.999999)]
        public InputDouble m_TOTAL_TRAV;
        [InputDouble(5, "Q", "Q", 0.999999)]
        public InputDouble m_Q;
        [InputString(6, "SIZE", "SIZE", "No Value")]
        public InputString m_SIZE;
        [InputDouble(7, "FACTOR", "FACTOR", 0.999999)]
        public InputDouble m_FACTOR;
        [InputDouble(8, "T", "T", 0.999999)]
        public InputDouble m_T;
        [InputDouble(9, "A", "A", 0.999999)]
        public InputDouble m_A;
        [InputDouble(10, "D", "D", 0.999999)]
        public InputDouble m_D;
        [InputDouble(11, "F", "F", 0.999999)]
        public InputDouble m_F;
        [InputDouble(12, "G", "G", 0.999999)]
        public InputDouble m_G;
        [InputDouble(13, "H", "H", 0.999999)]
        public InputDouble m_H;
        [InputDouble(14, "I", "I", 0.999999)]
        public InputDouble m_I;
        [InputDouble(15, "M", "M", 0.999999)]
        public InputDouble m_M;
        [InputDouble(16, "N", "N", 0.999999)]
        public InputDouble m_N;
        [InputDouble(17, "ACT_TRAVEL", "ACT_TRAVEL", 0.999999)]
        public InputDouble m_ACT_TRAVEL;
        [InputDouble(18, "TRAVEL_DIR", "TRAVEL_DIR", 1)]
        public InputDouble m_TravelDir;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("LUG", "LUG")]
        [SymbolOutput("SPRING", "SPRING")]
        [SymbolOutput("BOTTOM", "BOTTOM")]
        [SymbolOutput("THINGY", "THINGY")]
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
                Part part = (Part)m_PartInput.Value;

                Double load = m_LOAD.Value;
                Double R = m_R.Value;
                Double totalTravel = m_TOTAL_TRAV.Value;
                Double Q = m_Q.Value;
                Double factor = m_FACTOR.Value;
                Double T = m_T.Value;
                Double A = m_A.Value;
                Double D = m_D.Value;
                Double F = m_F.Value;
                Double G = m_G.Value;
                Double H = m_H.Value;
                Double I = m_I.Value;
                Double M = m_M.Value;
                Double N = m_N.Value;
                Double actTravel = m_ACT_TRAVEL.Value;
                Double B = 0, K = 0, J = 0;
                string size = m_SIZE.Value;

                CatalogBaseHelper cataloghelper = new CatalogBaseHelper();
                PartClass anvil80BJ = (PartClass)cataloghelper.GetPartClass("Anvil_FIG80B_J");
                ReadOnlyCollection<BusinessObject> anvil80BJclassItems = anvil80BJ.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

                foreach (BusinessObject classItem in anvil80BJclassItems)
                {
                    if (((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG80B_J", "LOAD_MAX")).PropValue > load) && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG80B_J", "LOAD_MIN")).PropValue <= load))
                    {
                        totalTravel = 0.5 * ((int)((totalTravel / 0.0254 + 0.499) * 2));
                        K = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG80B_J", "K")).PropValue;
                        J = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG80B_J", "J")).PropValue;
                        break;
                    }
                }

                PartClass anvil80BB = (PartClass)cataloghelper.GetPartClass("Anvil_FIG80B_B");
                ReadOnlyCollection<BusinessObject> anvil80BBclassItems = anvil80BB.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

                foreach (BusinessObject classItem in anvil80BBclassItems)
                {
                    if (((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG80B_B", "TOTAL_TRAVEL")).PropValue > (totalTravel * 0.0254 - 0.001)) && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG80B_B", "TOTAL_TRAVEL")).PropValue <= (totalTravel * 0.0254 + 0.001)))
                    {
                        B = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG80B_B", "B")).PropValue;
                        break;
                    }
                }


                Port port1 = new Port(OccurrenceConnection, part, "InThdRH", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_Symbolic.Outputs["Port1"] = port1;

                if (int.Parse(size) >= 10 && int.Parse(size) <= 18 && totalTravel > 3.75)
                    factor = 21.625;
                if (int.Parse(size) >= 19 && int.Parse(size) <= 34 && totalTravel > 5.25)
                    factor = 33.1875;
                if (int.Parse(size) >= 35 && int.Parse(size) <= 49 && totalTravel > 6.25)
                    factor = 41.5;
                if (int.Parse(size) >= 50 && int.Parse(size) <= 63 && totalTravel > 11.25)
                    factor = 57.75;
                if (int.Parse(size) >= 64 && int.Parse(size) <= 74 && totalTravel > 10.75)
                    factor = 77.375;
                if (int.Parse(size) >= 75 && int.Parse(size) <= 83 && totalTravel > 10.75)
                    factor = 78.0625;

                double takeOut = (factor - (totalTravel / 2)) * 0.0254;

                Port port2 = new Port(OccurrenceConnection, part, "Hole", new Position(0, 0, takeOut + K / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (int.Parse(size) < 64)
                    I = B + Q;

                if ((int.Parse(size) >= 35 && int.Parse(size) <= 49 && J < 0.02794) || (int.Parse(size) >= 50 && int.Parse(size) <= 63 && J < 0.02286))
                    R = 0.0381;

                if (int.Parse(size) >= 50 && int.Parse(size) <= 63 && J > 0.0381)
                    R = 0.0762;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                //Validating Inputs
                if (J <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTotalLoad, "Total load value must be between LOAD_MAX and LOAD_MIN"));
                    return;
                }

                if (A <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidZeroAvalue, "A value should be greater than zero"));
                    return;
                }
                if (M <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidZeroMvalue, "M value should be greater than zero"));
                    return;
                }
                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidZeroDvalue, "D value should be greater than zero"));
                    return;
                }
                if (T <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidZeroTvalue, "T value should be greater than zero"));
                    return;
                }
                if (N <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidZeroNvalue, "N value should be greater than zero"));
                    return;
                }
                if ((H + R) <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidHAndR, "H + R value should be greater than zero"));
                    return;
                }

                matrix.Translate(new Vector(-B + G - M / 2, -T / 2, takeOut - H));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d lugBox = symbolGeometryHelper.CreateBox(null, M / 2 - G + B + H, T, H + R, 9);
                lugBox.Transform(matrix);
                m_Symbolic.Outputs["LUG"] = lugBox;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(-B + G, 0, takeOut - H - D));
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1), new Position(0, 0, 0));
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), (new Vector(0, 0, 1)).GetOrthogonalVector());
                Projection3d springCylinder = symbolGeometryHelper.CreateCylinder(null, M / 2, D);
                springCylinder.Transform(matrix);
                m_Symbolic.Outputs["SPRING"] = springCylinder;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(-B + G - M / 2 - F, -N / 2, takeOut - H - A));
                matrix.Rotate((3 * Math.PI) / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d bottomBox = symbolGeometryHelper.CreateBox(null, F + M / 2 - G + I, N, A - D, 9);
                bottomBox.Transform(matrix);
                m_Symbolic.Outputs["BOTTOM"] = bottomBox;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), (new Vector(0, 0, 1)).GetOrthogonalVector());
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(0, 0, -J * 1.5));
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d thingyCylinder = symbolGeometryHelper.CreateCylinder(null, J, takeOut - D - H + J * 1.5);
                thingyCylinder.Transform(matrix);
                m_Symbolic.Outputs["THINGY"] = thingyCylinder;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG80B"));
                    return;
                }
            }
        }
        #endregion

    }

}
