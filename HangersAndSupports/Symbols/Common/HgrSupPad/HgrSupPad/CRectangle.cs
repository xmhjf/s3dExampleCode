//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//  CRectangle.cs
//    HgrSupPad,Ingr.SP3D.Content.Support.Symbols.CRectangle
//   Author       :  Rajeswari
//   Creation Date:  17-Dec-2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   17-Dec-2012   Rajeswari CR-CP-222301 .Net HgrSupPad project creation
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
    public class CRectangle : HangerComponentSymbolDefinition
    {
        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Length", "Length", 0.999999)]
        public InputDouble m_dLength;
        [InputDouble(3, "Breadth", "Breadth", 0.999999)]
        public InputDouble m_dBreadth;
        [InputDouble(4, "Radius", "Radius", 0.999999)]
        public InputDouble m_dRadius;
        [InputDouble(5, "Thickness", "Thickness", 0.999999)]
        public InputDouble m_dThickness;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("HgrPort_1", "Description for HgrPort")]
        [SymbolOutput("HgrPort_2", "Description for HgrPort")]
        [SymbolOutput("Rectangle", "Rectangle")]

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
                Double breadth =m_dLength.Value;
                Double length = m_dBreadth.Value;
                Double radius = m_dRadius.Value;
                Double thickness = m_dThickness.Value;

                if (thickness == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HgrSupPadLocalizer.GetString(HgrSupPadSymbolsResourceIDs.ErrInvalidThickness, "Thickness cannot be zero"));
                    return;
                }
                               
                //--------------------------------------------------
                //Rectangle Pad and the Projection
                //--------------------------------------------------
                Collection<ICurve> curveCollection = new Collection<ICurve>();
                Line3d bottomLengthLine = new Line3d(new Position(radius, 0, 0), new Position(length - radius, 0, 0));
                curveCollection.Add(bottomLengthLine);
                Arc3d lowerArc2 = new Arc3d(new Position(length-radius,radius,0),null, new Position(length - radius, 0, 0), new Position(length, radius, 0));
                curveCollection.Add(lowerArc2);
                Line3d secondBreadthLine = new Line3d(new Position(length, radius, 0), new Position(length, breadth - radius, 0));
                curveCollection.Add(secondBreadthLine);
                Arc3d UpperArc2 = new Arc3d(new Position(length - radius, breadth - radius, 0), null, new Position(length, breadth - radius, 0), new Position(length - radius, breadth, 0));
                curveCollection.Add(UpperArc2);
                Line3d topLengthLine = new Line3d(new Position(length - radius, breadth, 0), new Position(radius, breadth, 0));
                curveCollection.Add(topLengthLine);
                Arc3d upperarc1 = new Arc3d(new Position(radius, breadth - radius, 0), null, new Position(radius, breadth, 0), new Position(0, breadth - radius, 0));
                curveCollection.Add(upperarc1);
                Line3d firstBreadthLine = new Line3d(new Position(0, breadth - radius, 0), new Position(0, radius, 0));
                curveCollection.Add(firstBreadthLine);
                 Arc3d lowerArc1 = new Arc3d(new Position(radius, radius, 0), null, new Position(0, radius, 0), new Position(radius, 0, 0));
                curveCollection.Add(lowerArc1);

                ComplexString3d rectangleLine = new ComplexString3d(curveCollection);
                Projection3d rectangle = new Projection3d(rectangleLine, new Vector(0, 0, 1), thickness, true);
                m_Symbolic.Outputs["Rectangle"] = rectangle;

                //CREATE HANGER PORTS
                Port port1 = new Port(OccurrenceConnection, part, "HgrPort_1", new Position(length/2,breadth/2,0), new Vector(1.0, 0, 0.0), new Vector(0, 0.0, 1.0));
                m_Symbolic.Outputs["HgrPort_1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "HgrPort_2", new Position(length/2,breadth/2,thickness), new Vector(1.0, 0, 0.0), new Vector(0, 0.0, 1.0));
                m_Symbolic.Outputs["HgrPort_2"] = port2;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HgrSupPadLocalizer.GetString(HgrSupPadSymbolsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of CRectangle.cs."));
                    return;
                }
            }
        }

        #endregion
    }
}
