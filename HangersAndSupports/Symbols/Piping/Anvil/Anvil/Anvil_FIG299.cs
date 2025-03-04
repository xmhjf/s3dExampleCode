//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG299.cs
//   Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG299
//   Author       :  Hema
//   Creation Date:  2-05-2013
//   Description:    Converted Anvil_Parts VB Project to C#.Net Project

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   2-05-2013     Hema      CR-CP-222292 Convert HS_Anvil VB Project to C# .Net
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

    public class Anvil_FIG299 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG299"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputString(2, "CONFIG", "CONFIG", "No Value")]
        public InputString m_oCONFIG;
        [InputDouble(3, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble m_dROD_DIA;
        [InputDouble(4, "A", "A", 0.999999)]
        public InputDouble m_dA;
        [InputDouble(5, "P", "P", 0.999999)]
        public InputDouble m_dP;
        [InputDouble(6, "N", "N", 0.999999)]
        public InputDouble m_dN;
        [InputDouble(7, "GRIP", "GRIP", 0.999999)]
        public InputDouble m_dGRIP;
        [InputDouble(8, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(9, "T", "T", 0.999999)]
        public InputDouble m_dT;
        [InputDouble(10, "W", "W", 0.999999)]
        public InputDouble m_dW;
        [InputDouble(11, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("PIN", "PIN")]
        [SymbolOutput("TOP", "TOP")]
        [SymbolOutput("LEFT", "LEFT")]
        [SymbolOutput("RIGHT", "RIGHT")]
        [SymbolOutput("LEFT_CYL", "LEFT_CYL")]
        [SymbolOutput("RIGHT_CYL", "RIGHT_CYL")]
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

                string config = m_oCONFIG.Value;
                Double rodDiameter = m_dROD_DIA.Value;
                Double A = m_dA.Value;
                Double P = m_dP.Value;
                Double N = m_dN.Value;
                Double grip = m_dGRIP.Value;
                Double D = m_dD.Value;
                Double T = m_dT.Value;
                Double W = m_dW.Value;
                double pinSize = m_dP.Value;

                //Initializing SymbolGeometryHelper
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
               
                Port port1 = new Port(OccurrenceConnection, part, "InThdRH", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_Symbolic.Outputs["Port1"] = port1;

                if (config == "Without Pin")
                {
                    Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, rodDiameter / 2 - (A + pinSize / 2)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_Symbolic.Outputs["Port2"] = port2;
                }
                else
                {
                    Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, -A + pinSize / 2 + rodDiameter / 2), new Vector(1, 0, 0), new Vector(0, 0, -1));
                    m_Symbolic.Outputs["Port2"] = port2;

                    //Validating Inputs
                    if (pinSize <= 0)
                    {
                        ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidPinSizeGTZero, "Pinsize should be greater than zero"));
                        return;
                    }
                    Vector normal = new Position(0, -(grip / 2 + 2 * T), -A + rodDiameter / 2).Subtract(new Position(0, grip / 2 + 2 * T, -A + rodDiameter / 2));
                    symbolGeometryHelper.ActivePosition = new Position(0, grip / 2 + 2 * T, -A + rodDiameter / 2);
                    symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                    Projection3d pin = symbolGeometryHelper.CreateCylinder(null, pinSize / 2, normal.Length);
                    m_Symbolic.Outputs["PIN"] = pin;
                }

                //Validating Inputs
                if (grip <= 0 && T <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidGripnTGTZero, "Grip and T value should be greater than zero"));
                    return;
                }
                if (N == 0 && A==0 && D==0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidNValueAvalueAndDNZero, "N,A and D values cannot be zero"));
                    return;
                }
                if (T <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTGTZero, "T value should be greater than zero"));
                    return;
                }
                if (W <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidWGTZero, "W value should be greater than zero"));
                    return;
                }
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, rodDiameter / 2);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d top = symbolGeometryHelper.CreateCylinder(null, (grip + T) / 2, N);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                top.Transform(matrix);
                m_Symbolic.Outputs["TOP"] = top;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-grip / 2 - T, -W / 2, -A + 0.4 * D + rodDiameter / 2);
                Projection3d left = symbolGeometryHelper.CreateBox(null, T, W, A + N / 2 - 0.4 * D, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                left.Transform(matrix);
                m_Symbolic.Outputs["LEFT"] = left;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(grip / 2, -W / 2, -A + 0.4 * D + rodDiameter / 2);
                Projection3d right = symbolGeometryHelper.CreateBox(null, T, W, A + N / 2 - 0.4 * D, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                right.Transform(matrix);
                m_Symbolic.Outputs["RIGHT"] = right;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d leftCylinder = symbolGeometryHelper.CreateCylinder(null, D / 2, T);
                matrix = new Matrix4X4();
                matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0));
                matrix.Translate(new Vector(grip / 2, 0, -A + rodDiameter / 2));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1),new Position(0,0,0));
                leftCylinder.Transform(matrix);
                m_Symbolic.Outputs["LEFT_CYL"] = leftCylinder;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d rightCylinder = symbolGeometryHelper.CreateCylinder(null, D / 2, T);
                matrix = new Matrix4X4();
                matrix.Rotate(-Math.PI / 2, new Vector(0, 1, 0));
                matrix.Translate(new Vector(-grip / 2, 0, -A + rodDiameter / 2));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1),new Position(0,0,0));
                rightCylinder.Transform(matrix);
                m_Symbolic.Outputs["RIGHT_CYL"] = rightCylinder;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG299"));
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
                PropertyValueCodelist finishCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrFinish", "FINISH");

                if (finishCodelist.PropValue == 0)
                    finishCodelist.PropValue = 1;
                if (finishCodelist.PropValue != 1 && finishCodelist.PropValue != 2)
                {
                    ToDoListMessage ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidFinishvalue, "Finish Value Should be 1 or 2"));
                    return "";
                } 
                string finish = finishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(finishCodelist.PropValue).DisplayName;
                string config = (string)((PropertyValueString)part.GetPropertyValue("IJUAHgrAnvil_fig299", "CONFIG")).PropValue;

                bomDescription = part.PartDescription + ", Config: " + config + ", Finish: " + finish;

                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrBOMDescription, "Error in BOM description of Anvil_FIG299"));
                return ""; 
            }
        }
        #endregion
    }
}
