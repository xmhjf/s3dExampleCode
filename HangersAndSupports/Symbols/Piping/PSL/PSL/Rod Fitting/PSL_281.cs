//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   PSL_281.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_281
//   Author       :  Manikanth
//   Creation Date:  21-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   21-08-2013    Manikanth CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net 
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
    [SymbolVersion("1.0.0.0")]
    [CacheOption(CacheOptionType.Cached)]
    public class PSL_281 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_281"
        //----------------------------------------------------------------------------------
        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "E", "E", 0.999999)]
        public InputDouble E1;
        [InputDouble(3, "B", "B", 0.999999)]
        public InputDouble B1;
        [InputDouble(4, "K", "K", 0.999999)]
        public InputDouble K1;
        [InputDouble(5, "H", "H", 0.999999)]
        public InputDouble H1;
        #endregion

        #region "Definitions of Aspects and their Outputs"
        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("CYL", "CYL")]
        [SymbolOutput("CYL2", "CYL2")]
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
                Part part = (Part)PartInput.Value;

                double B = B1.Value;
                double E = E1.Value;
                double K = K1.Value;
                double H = H1.Value;

                Port port1 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, E), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();


                if (B <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidBGTZ, "B value should be greater than zero"));
                    return;
                }
                if (E == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidENEZ, "E value cannot be zero"));
                    return;
                }
                if (H <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidHGTZ, "H value should be greater than zero"));
                    return;
                }
                if (K == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidKNEZ, "K value cannot be zero"));
                    return;
                }
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), (new Vector(0, 0, 1)).GetOrthogonalVector());
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d cylinder = (Projection3d)symbolGeometryHelper.CreateCylinder(null, B / 2, E);
                cylinder.Transform(matrix);
                m_Symbolic.Outputs["CYL"] = cylinder;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), (new Vector(0, 0, 1)).GetOrthogonalVector());
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(0, 0, E));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d cylinder2 = (Projection3d)symbolGeometryHelper.CreateCylinder(null, H / 2, K);
                cylinder2.Transform(matrix);
                m_Symbolic.Outputs["CYL2"] = cylinder2;

            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_281"));
                    return;
                }
            }
        }

        #endregion
    }
}
