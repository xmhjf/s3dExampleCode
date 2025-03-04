//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_488.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_488
//   Author       :  Hema
//   Creation Date:  21-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   21-08-2013     Hema    CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net    
//   30-Dec-2014    PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Middle;
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
    public class PSL_488 : HangerComponentSymbolDefinition
    {
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_488"

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "C_A", "C_A", 0.999999)]
        public InputDouble C_A;
        [InputDouble(3, "D_A", "D_A", 0.999999)]
        public InputDouble D_A;
        [InputDouble(4, "E_A", "E_A", 0.999999)]
        public InputDouble E_A;
        [InputDouble(5, "B_A", "B_A", 0.999999)]
        public InputDouble B_A;
        [InputDouble(6, "C_B", "C_B", 0.999999)]
        public InputDouble C_B;
        [InputDouble(7, "D_B", "D_B", 0.999999)]
        public InputDouble D_B;
        [InputDouble(8, "E_B", "E_B", 0.999999)]
        public InputDouble E_B;
        [InputDouble(9, "B_B", "B_B", 0.999999)]
        public InputDouble B_B;
        [InputDouble(10, "C_C", "C_C", 0.999999)]
        public InputDouble C_C;
        [InputDouble(11, "D_C", "D_C", 0.999999)]
        public InputDouble D_C;
        [InputDouble(12, "E_C", "E_C", 0.999999)]
        public InputDouble E_C;
        [InputDouble(13, "B_C", "B_C", 0.999999)]
        public InputDouble B_C;
        [InputDouble(14, "BEAM_TYPE", "BEAM_TYPE", 1)]
        public InputDouble BEAM_TYPE;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("HOLE", "HOLE")]
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
                //-----------------------------------------------------------------------
                // Construction of Physical Aspect 
                //----------------------------------------------------------------------
                Part part = (Part)PartInput.Value;
                Double c, d, e, b;
                int beamType = (int)BEAM_TYPE.Value;
                if (beamType == 1) //code list calue e.g. "1" --> "A"
                {
                    c = C_A.Value;
                    d = D_A.Value;
                    e = E_A.Value;
                    b = B_A.Value;
                }
                else
                {
                    if (beamType == 2)
                    {
                        c = C_B.Value;
                        d = D_B.Value;
                        e = E_B.Value;
                        b = B_B.Value;
                    }
                    else
                    {
                        c = C_C.Value;
                        d = D_C.Value;
                        e = E_C.Value;
                        b = B_C.Value;
                    }
                }
                //error checking
                if (HgrCompareDoubleService.cmpdbl(c , 0)==true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidBeamType, "BEAM_TYPE is not available for this part"));
                    return;
                }
                //Intializing SymbolGeomHelper
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                //Ports
                Port port1 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, e), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (c <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidCGTZ, "C value should be greater than zero"));
                    return;
                }
                if (e <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidEGTZ, "E value should be greater than zero"));
                    return;
                }
                if (b <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidBGTZ, "B value should be greater than zero"));
                    return;
                }
                if (d <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidDGTZ, "D value should be greater than zero"));
                    return;
                }               
                symbolGeometryHelper.ActivePosition = new Position(-b / 2.0, -c / 2.0, 0);
                Projection3d body = symbolGeometryHelper.CreateBox(null, b, c, e, 9);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                body.Transform(matrix);
                m_Symbolic.Outputs["BODY"] = body;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d hole = symbolGeometryHelper.CreateCylinder(null, d / 2.0, e);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                hole.Transform(matrix);
                m_Symbolic.Outputs["HOLE"] = hole;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_488."));
                    return;
                }
            }
        }
        #endregion
    }
}