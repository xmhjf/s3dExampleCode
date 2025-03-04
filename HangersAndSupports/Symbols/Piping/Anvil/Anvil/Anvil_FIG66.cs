//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG66.cs
//   Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG66
//   Author       :  Hema
//   Creation Date:  06-05-2013
//   Description:    Converted Anvil_Parts VB Project to C#.Net Project

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------ 
//   07-05-2013     Chethan  HF-CP-249460  [TR] Some Anvil_FIG66 parts converted to .Net fail to upgrade correctly.  
//   06-05-2013     Hema     CR-CP-222292 Convert HS_Anvil VB Project to C# .Net
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
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
    [SymbolVersion("1.0.0.0")]
    [CacheOption(CacheOptionType.Cached)]

    public class Anvil_FIG66 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG66"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputString(2, "PLACE_MODE_USER", "PLACE_MODE_USER", "No Value")]
        public InputString m_oPLACE_MODE_USER;
        [InputDouble(3, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble m_dROD_DIA;
        [InputDouble(4, "E", "E", 0.999999)]
        public InputDouble m_dE;
        [InputDouble(5, "E_PRIME", "E_PRIME", 0.999999)]
        public InputDouble m_dE_PRIME;
        [InputDouble(6, "H", "H", 0.999999)]
        public InputDouble m_dH;
        [InputDouble(7, "R", "R", 0.999999)]
        public InputDouble m_dR;
        [InputDouble(8, "S", "S", 0.999999)]
        public InputDouble m_dS;
        [InputDouble(9, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(10, "T", "T", 0.999999)]
        public InputDouble m_dT;
        [InputDouble(11, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("TOP", "TOP")]
        [SymbolOutput("SIDE1", "SIDE1")]
        [SymbolOutput("SIDE2", "SIDE2")]
        [SymbolOutput("PIN", "PIN")]
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

                String placeModeUser = m_oPLACE_MODE_USER.Value;
                Double rodDiameter = m_dROD_DIA.Value;
                Double E = m_dE.Value;
                Double ePrime = m_dE_PRIME.Value;
                Double H = m_dH.Value;
                Double R = m_dR.Value;
                Double S = m_dS.Value;
                Double B = m_dB.Value;
                Double T = m_dT.Value;

                //Initializing SymbolGeometryHelper
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                //Validating Inputs
                if (S <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidSGTZero, "S value should be greater than zero"));
                    return;
                }
                if (B <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidBGTZero, "B value should be greater than zero"));
                    return;
                }
                if (T <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTGTZero, "T value should be greater than zero"));
                    return;
                }
                if (H <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidHGTZero, "H value should be greater than zero"));
                    return;
                }
                if (ePrime== 0 && R==0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidePrimeAndRNZero, "E_PRIME and R values cannot be zero"));
                    return;
                }

                if (placeModeUser == "Connect with Hanger Rod")
                {
                    symbolGeometryHelper.ActivePosition = new Position(-S / 2, -B / 2, -(R + ePrime));
                    Projection3d top = symbolGeometryHelper.CreateBox(null, S, B, T, 9);
                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                    top.Transform(matrix);
                    m_Symbolic.Outputs["TOP"] = top;

                    Port port2 = new Port(OccurrenceConnection, part, "InThdRH", new Position(0, 0, -E), new Vector(1, 0, 0), new Vector(0, 0, -1));
                    m_Symbolic.Outputs["Port2"] = port2;
                }
                else
                {
                    symbolGeometryHelper.ActivePosition = new Position(-(S / 2), -B / 2, -T);
                    Projection3d top = symbolGeometryHelper.CreateBox(null, S, B, T, 9);
                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                    top.Transform(matrix);
                    m_Symbolic.Outputs["TOP"] = top;

                    Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, -(ePrime - H / 2)), new Vector(1, 0, 0), new Vector(0, 0, -1));
                    m_Symbolic.Outputs["Port2"] = port2;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    Vector normal = new Position(0, -(S / 2 + T + H / 2), -ePrime).Subtract(new Position(0, (S / 2 + T + H / 2), -ePrime));
                    symbolGeometryHelper.ActivePosition = new Position(0, (S / 2 + T + H / 2), -ePrime);
                    symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                    Projection3d pin = symbolGeometryHelper.CreateCylinder(null, H / 2, normal.Length);
                    m_Symbolic.Outputs["PIN"] = pin;
                }

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-(S / 2 + T), -B / 2, -(ePrime + R));
                Projection3d side1 = symbolGeometryHelper.CreateBox(null, T, B, (ePrime + R), 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                side1.Transform(matrix);
                m_Symbolic.Outputs["SIDE1"] = side1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(S / 2, -B / 2, -(ePrime + R));
                Projection3d side2 = symbolGeometryHelper.CreateBox(null, T, B, (ePrime + R), 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                side2.Transform(matrix);
                m_Symbolic.Outputs["SIDE2"] = side2;

            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG66"));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                string addDescription = " WITHOUT Pin or Bolt ";

                //To get FINISH
                string placeModeUser = (string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJUAHgrAnvil_fig66", "PLACE_MODE_USER")).PropValue;
                double rodDiameter = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJUAHgrRod_Dia", "ROD_DIA")).PropValue;

                if (placeModeUser.Equals("Connect to Bolt") && rodDiameter <= 1)
                    addDescription = " with bolt and nut";
                if (placeModeUser.Equals("Connect to Bolt") && rodDiameter >= 1.25)
                    addDescription = " with pin and cotter pins";

                bomDescription = part.PartDescription + addDescription;

                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrBOMDescription, "Error in BOM description of Anvil_FIG66"));
                return ""; 
            }
        }
        #endregion

    }

}
