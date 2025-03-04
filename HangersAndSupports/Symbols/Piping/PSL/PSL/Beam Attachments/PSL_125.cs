//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_125.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_125
//   Author       :  Vijay
//   Creation Date:  21-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   21-08-2013     Vijay    CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net    
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
    public class PSL_125 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_122A"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputString(2, "SIZE", "SIZE", "No Value")]
        public InputString SIZE;
        [InputDouble(3, "A", "A", 0.999999)]
        public InputDouble A;
        [InputDouble(4, "C", "C", 0.999999)]
        public InputDouble C;
        [InputDouble(5, "E", "E", 0.999999)]
        public InputDouble E;
        [InputDouble(6, "B", "B", 0.999999)]
        public InputDouble B;
        [InputDouble(7, "F", "F", 0.999999)]
        public InputDouble F;
        [InputDouble(8, "G", "G", 0.999999)]
        public InputDouble G;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("PIN", "PIN")]
        [SymbolOutput("TOP", "TOP")]
        [SymbolOutput("SIDE1", "SIDE1")]
        [SymbolOutput("SIDE2", "SIDE2")]
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

                string size = SIZE.Value;
                Double a = A.Value;
                Double b = B.Value;
                Double c = C.Value;
                Double f = F.Value;
                Double g = G.Value;
                Double e = E.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, Convert.ToDouble(size) / 2000 - a), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (Convert.ToDouble(size) <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidSizeGTZ, "SIZE should be greater than zero"));
                    return;
                }
                if (e <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidEGTZ, "E value should be greater than zero"));
                    return;
                }
                if (f <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidFGTZ, "F value should be greater than zero"));
                    return;
                }

                Vector normal = new Position(0, -(c / 2 + f + g), -a).Subtract(new Position(0, (c / 2 + f + g), -a));
                symbolGeometryHelper.ActivePosition = new Position(0, (c / 2 + f + g), -a);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d pin = symbolGeometryHelper.CreateCylinder(null, Convert.ToDouble(size) / 2000, normal.Length);
                m_Symbolic.Outputs["PIN"] = pin;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.Translate(new Vector(-(c / 2 + f + g), -e / 2, -f));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d topBox = symbolGeometryHelper.CreateBox(null, c + f * 2 + g * 2, e, f, 9);
                topBox.Transform(matrix);
                m_Symbolic.Outputs["TOP"] = topBox;

                matrix = new Matrix4X4();
                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.Translate(new Vector(-(c / 2 + f), -e / 2, -(a + b)));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d side1Box = symbolGeometryHelper.CreateBox(null, f, e, (a + b - f), 9);
                side1Box.Transform(matrix);
                m_Symbolic.Outputs["SIDE1"] = side1Box;

                matrix = new Matrix4X4();
                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.Translate(new Vector(c / 2, -e / 2, -(a + b)));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d side2Box = symbolGeometryHelper.CreateBox(null, f, e, (a + b - f), 9);
                side2Box.Transform(matrix);
                m_Symbolic.Outputs["SIDE2"] = side2Box;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_125."));
                    return;
                }
            }
        }

        #endregion


        #region ICustomHgrBOMDescription Members

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                bomDescription = "PSL " + (string)((PropertyValueString)part.GetPropertyValue("IJUAHgrPSL_PART_NUMBER", "PART_NUMBER")).PropValue + "  Inverted Beam Welding Attachment";
                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of PSL_125.cs."));
                return "";
            }
        }

        #endregion
    }
}
