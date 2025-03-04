//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   CVerticalPipe.cs
//   HgrSupPad,Ingr.SP3D.Content.Support.Symbols.CVerticalPipe
//   Author       :  Rajeswari
//   Creation Date:  18-Dec-2012 
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   18-Dec-2012   Rajeswari CR-CP-222301 .Net HgrSupPad project creation
//  18/03/2013     Rajeswari DI-CP-228142  Modify the error handling for delivered H&S symbols
//  6/11/2013      Vijaya    CR-CP-242533  Provide the ability to store GType outputs as a single blob for H&S .net symbols  
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
    [VariableOutputs]
    [CacheOption(CacheOptionType.NonCached)]
    [SymbolVersion("1.0.0.0")]
    public class CVerticalPipe : HangerComponentSymbolDefinition
    {
        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Length", "Length", 0.999999)]
        public InputDouble m_dLength;
        [InputDouble(3, "Height", "Height", 0.999999)]
        public InputDouble m_dHeight;
        [InputDouble(4, "Thickness", "Thickness", 0.999999)]
        public InputDouble m_dThickness;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("HgrPort_1", "Description for HgrPort")]
        [SymbolOutput("HgrPort_2", "Description for HgrPort")]
        [SymbolOutput("VerticalPipe", "VerticalPipe")]

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
                Double length = m_dLength.Value;
                Double height = m_dHeight.Value;
                Double thickness = m_dThickness.Value;

                Double diagonalOffset = 0.05;
                Double topOffset = 0.025;
                Double verticalOffset = 0.03;
                Double xOffset = verticalOffset * diagonalOffset / height;

                if (thickness == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HgrSupPadLocalizer.GetString(HgrSupPadSymbolsResourceIDs.ErrInvalidThickness, "Thickness cannot be zero"));
                    return;
                }
                //--------------------------------------------------
                //Vertical Pipe Pad and the Projection
                //--------------------------------------------------
                Collection<ICurve> curveCollection = new Collection<ICurve>();
                Line3d bottomLengthLine = new Line3d(new Position(0, 0, 0), new Position(length + 2 * (diagonalOffset + topOffset), 0, 0));
                curveCollection.Add(bottomLengthLine);
                Line3d topLengthline = new Line3d(new Position(diagonalOffset, 0, height), new Position(length + 2 * topOffset + diagonalOffset, 0, height));
                curveCollection.Add(topLengthline);
                Line3d leftVerticalLine = new Line3d(new Position(0, 0, 0), new Position(0, 0, verticalOffset));
                curveCollection.Add(leftVerticalLine);
                Line3d rightVerticalLine = new Line3d(new Position(length + 2 * (diagonalOffset + topOffset), 0, 0), new Position(length + 2 * (diagonalOffset + topOffset), 0, verticalOffset));
                curveCollection.Add(rightVerticalLine);
                Line3d leftDiagonalLine = new Line3d(new Position(diagonalOffset, 0, height), new Position(2*xOffset,0,2*verticalOffset));
                curveCollection.Add(leftDiagonalLine);
                Line3d rightDiagonalLine = new Line3d(new Position(length + diagonalOffset + 2 * topOffset, 0, height), new Position(length + 2 * (diagonalOffset + topOffset)-2*xOffset, 0,2 * verticalOffset));
                curveCollection.Add(rightDiagonalLine);
                Arc3d lowerArc1 = new Arc3d(new Position(0,0,verticalOffset), new Position(xOffset,0,4/3.0*verticalOffset), new Position(2*xOffset,0,2*verticalOffset));
                curveCollection.Add(lowerArc1);
                Arc3d lowerArc2 = new Arc3d(new Position(length+2*(diagonalOffset+topOffset),0,verticalOffset), new Position(length+2*(diagonalOffset+topOffset)-xOffset,0,4/3.0*verticalOffset), new Position(length+2*(diagonalOffset+topOffset)-2*xOffset,0,2*verticalOffset));
                curveCollection.Add(lowerArc2);

                ComplexString3d verticalPipeLine = new ComplexString3d(curveCollection);
                Projection3d verticalPipe = new Projection3d( verticalPipeLine, new Vector(0, -1, 0), thickness, true);
                m_Symbolic.Outputs["VerticalPipe"] = verticalPipe;

                //CREATE HANGER PORTS
                Port port1 = new Port(OccurrenceConnection, part, "HgrPort_1", new Position(length / 2 + diagonalOffset + topOffset, -thickness / 2, 0), new Vector(1.0, 0, 0.0), new Vector(0, 0.0, 1.0));
                m_Symbolic.Outputs["HgrPort_1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "HgrPort_2", new Position(length / 2 + diagonalOffset + topOffset, -thickness / 2, height/2), new Vector(1.0, 0, 0.0), new Vector(0, 0.0, 1.0));
                m_Symbolic.Outputs["HgrPort_2"] = port2;
               
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HgrSupPadLocalizer.GetString(HgrSupPadSymbolsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of CVerticalPipe.cs."));
                    return;
                }
            }
        }

        #endregion
    }
}
