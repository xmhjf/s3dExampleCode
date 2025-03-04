//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG81F.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG81F
//   Author       :  Manikanth
//   Creation Date:  16-05-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   15-05-2013    Manikanth CR-CP-233113-Convert HS_Anvil VB Project to C# .Net 
//   11-Dec-2014     PVK     TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
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
    public class Anvil_FIG81F : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG81F"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "M", "M", 0.999999)]
        public InputDouble m_M;
        [InputDouble(3, "ACT_TRAVEL", "ACT_TRAVEL", 0.999999)]
        public InputDouble m_ACT_TRAVEL;
        [InputDouble(4, "TOTAL_TRAV", "TOTAL_TRAV", 0.999999)]
        public InputDouble m_TOTAL_TRAV;
        [InputDouble(5, "DIR", "DIR", 1)]
        public InputDouble m_DIR;
        [InputString(6, "SIZE", "SIZE", "No Value")]
        public InputString m_SIZE;
        [InputDouble(7, "B", "B", 0.999999)]
        public InputDouble m_B;
        [InputDouble(8, "C", "C", 0.999999)]
        public InputDouble m_C;
        [InputDouble(9, "J", "J", 0.999999)]
        public InputDouble m_J;
        [InputDouble(10, "K", "K", 0.999999)]
        public InputDouble m_K;
        [InputDouble(11, "L", "L", 0.999999)]
        public InputDouble m_L;
        [InputDouble(12, "OPER_LOAD", "OPER_LOAD", 0.999999)]
        public InputDouble m_OPER_LOAD;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BASE", "BASE")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("TOP", "TOP")]
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

                Double M = m_M.Value;
                Double actTravel = m_ACT_TRAVEL.Value;
                Double totalTrav = m_TOTAL_TRAV.Value;
                Double B = m_B.Value;
                Double C = m_C.Value;
                Double J = m_J.Value;
                Double K = m_K.Value;
                Double L = m_L.Value;
                int dir = (int)m_DIR.Value;
                Double operLoad = m_OPER_LOAD.Value;
                string size = m_SIZE.Value;
                double A = 0;

                PropertyValueCodelist dirCodelist = ((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrAnvil_FIG81F", "DIR"));

                Port port1 = new Port(OccurrenceConnection, part, "Other", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                CatalogBaseHelper cataloghelper = new CatalogBaseHelper();
                PartClass anvil81F_A = (PartClass)cataloghelper.GetPartClass("Anvil_FIG81F_A");
                ReadOnlyCollection<BusinessObject> anvil81F_AclassItems = anvil81F_A.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject classItem in anvil81F_AclassItems)
                {
                    if (((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81F_A", "SIZE_MAX")).PropValue >= int.Parse(size)) && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81F_A", "SIZE_MIN")).PropValue <= int.Parse(size)) && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81F_A", "TRAV_MAX")).PropValue >= totalTrav) && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81F_A", "TRAV_MIN")).PropValue <= totalTrav))
                    {
                        A = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81F_A", "A")).PropValue;
                        A = A + 0.5 * actTravel;
                        break;
                    }
                }
                string direction = dirCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(dir).DisplayName;
                if (direction == "Up")
                    A = A - 0.5 * actTravel;

                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, A), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                //Validating Inputs
                if (M <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidZeroMvalue, "M value should be greater than zero"));
                    return;
                }
                if (J <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidZeroJvalue, "J value should be greater than zero"));
                    return;
                }                
                if (B <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidZeroBvalue, "B value should be greater than zero"));
                    return;
                }
                if (C <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidZeroCvalue, "C value should be greater than zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(L - K, 0)==true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidZeroLandKvalue, "L-K value cannot be zero"));
                    return;
                }
                if (K <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidZeroKvalue, "K value should be greater than zero"));
                    return;
                }

                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                matrix.Rotate((Math.PI) / 2, new Vector(0, 1, 0));
                matrix.Translate(new Vector(K / 2, 0, A / 2));
                matrix.Rotate((3 * Math.PI) / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d baseCylinder = symbolGeometryHelper.CreateCylinder(null, M / 2, L - K);
                baseCylinder.Transform(matrix);
                m_Symbolic.Outputs["BASE"] = baseCylinder;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(-K / 2, -J / 2, A / 2 - M / 2));
                matrix.Rotate((3 * Math.PI) / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d bodybox = symbolGeometryHelper.CreateBox(null, K, J, A / 2 + M / 2, 9);
                bodybox.Transform(matrix);
                m_Symbolic.Outputs["BODY"] = bodybox;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(-B / 2, -C / 2, 0));
                matrix.Rotate((3 * Math.PI) / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d topBox = symbolGeometryHelper.CreateBox(null, B, C, A / 2 - M / 2, 9);
                topBox.Transform(matrix);
                m_Symbolic.Outputs["TOP"] = topBox;

            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG81F"));
                    return;
                }
            }
        }
        #endregion

    }

}
