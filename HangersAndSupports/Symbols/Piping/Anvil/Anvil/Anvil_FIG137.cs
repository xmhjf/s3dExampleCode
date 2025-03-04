//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG137.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG137
//   Author       :  Vijay
//   Creation Date:  30-04-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   30-04-2013     Vijay    CR-CP-222292  Convert HS_Anvil VB Project to C# .Net 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
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
    public class Anvil_FIG137 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG137"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "EXTRA_TANGENT", "EXTRA_TANGENT", 0.999999)]
        public InputDouble m_dEXTRA_TANGENT;
        [InputDouble(3, "GAGE", "GAGE", 0.999999)]
        public InputDouble m_dGAGE;
        [InputDouble(4, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;
        [InputDouble(5, "A", "A", 0.999999)]
        public InputDouble m_dA;
        [InputDouble(6, "C", "C", 0.999999)]
        public InputDouble m_dC;
        [InputDouble(7, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(8, "NUTS", "NUTS", 1)]
        public InputDouble m_oNUTS;
        [InputDouble(9, "EXTRA_THREAD", "EXTRA_THREAD", 0.999999)]
        public InputDouble m_dEXTRA_THREAD;
        [InputDouble(10, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BEND", "BEND")]
        [SymbolOutput("R", "R")]
        [SymbolOutput("L", "L")]
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

                Double extraTangent = m_dEXTRA_TANGENT.Value;
                Double gage = m_dGAGE.Value;
                Double pipeDiameter = m_dPIPE_DIA.Value;
                Double A = m_dA.Value;
                Double C = m_dC.Value;
                Double D = m_dD.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, pipeDiameter / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (D == 0 && extraTangent == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidDandExtraTangentNZero, "D and ExtraTanget values cannot be zero."));
                    return;
                }

                if (A <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidAGTZero, "A value should be greater than zero."));
                    return;
                }

                Matrix4X4 matrix = new Matrix4X4();
                matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                matrix.Translate(new Vector(0, gage, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Revolution3d bend = new Revolution3d((new Circle3d(new Position(C / 2.0, 0, 0), new Vector(0, 0, 1), A / 2.0)), new Vector(0, -1, 0), new Position(0, 0, 0), Math.PI, true);
                bend.Transform(matrix);
                m_Symbolic.Outputs["BEND"] = bend;

                matrix = new Matrix4X4();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                matrix.Translate(new Vector(C / 2, gage, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d rightCylinder = symbolGeometryHelper.CreateCylinder(null, A / 2, D + extraTangent);
                rightCylinder.Transform(matrix);
                m_Symbolic.Outputs["R"] = rightCylinder;

                matrix = new Matrix4X4();
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                matrix.Translate(new Vector(-C / 2, gage, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d leftCylinder = symbolGeometryHelper.CreateCylinder(null, A / 2, D + extraTangent);
                leftCylinder.Transform(matrix);
                m_Symbolic.Outputs["L"] = leftCylinder;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG137"));
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

                string finish = finishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(finishCodelist.PropValue).DisplayName;

                int nuts = ((PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG137", "NUTS")).PropValue;
                double pipeDiameter = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPipe_Dia", "PIPE_DIA")).PropValue;
                double extraTangent = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG137", "EXTRA_TANGENT")).PropValue;
                double extraThread = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG137", "EXTRA_THREAD")).PropValue;

                string nutText, extraTan, extraThr;

                if (nuts > 0)
                    nutText = "with hex nuts";
                else
                    nutText = "without hex nuts";

                if (extraTangent > 0)
                    extraTan = ", Extra Tan: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, extraTangent, UnitName.DISTANCE_INCH);
                else
                    extraTan = "";

                if (extraThread > 0)
                    extraThr = ", Extra Thr: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, extraThread, UnitName.DISTANCE_INCH);
                else
                    extraThr = "";

                bomDescription = part.PartDescription + ", Finish: " + finish + " " + nutText + " " + extraTan + " " + extraThr;

                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrBOMDescription, "Error in BOM description of Anvil_FIG137"));
                return "";
            }
        }
        #endregion
    }
}
