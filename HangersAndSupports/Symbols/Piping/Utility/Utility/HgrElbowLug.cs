//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   HgrElbowLug.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.HgrElbowLug
//   Author       :  Rajeswari
//   Creation Date:  06/11/2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   06/11/2012   Rajeswari   CR-CP-222288  Converted HS_Utility VB Project to C# .Net  
//	 27/03/2013		Hema 	  DI-CP-228142  Modify the error handling for delivered H&S symbols
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

    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class HgrElbowLug : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.HgrElbowLug"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "PipeRadius", "PipeRadius", 1)]
        public InputDouble m_PipeRadius;
        [InputDouble(3, "Thickness", "Thickness", 0.375)]
        public InputDouble m_Thickness;
        [InputDouble(4, "BendRadius", "BendRadius", 2.5)]
        public InputDouble m_BendRadius;
        [InputDouble(5, "HoleRadius", "HoleRadius", 0.3125)]
        public InputDouble m_HoleRadius;
        [InputDouble(6, "DropValue", "Drop Value", 1)]
        public InputDouble m_DropValue;
        [InputDouble(7, "RodTakeOut", "Rod Take Out", 3.875)]
        public InputDouble m_RodTakeOut;
        [InputDouble(8, "ClevisDiam", "Clevis Ring Diameter", 1.4375)]
        public InputDouble m_ClevisDiam;
        [InputDouble(9, "ClevisWidth", "Clevis Leg Width", 1.0625)]
        public InputDouble m_ClevisWidth;
        [InputDouble(10, "RodNutHeight", "Clevis Nut Height", 0.625)]
        public InputDouble m_RodNutHeight;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Lug", "Lug Plate")]
        [SymbolOutput("LugPin", "Lug Pin")]
        [SymbolOutput("ClevisRing", "Clevis Ring")]
        [SymbolOutput("RightClevisLeg", "Right Clevis Leg")]
        [SymbolOutput("LeftClevisLeg", "Left Clevis Leg")]
        [SymbolOutput("SupportedPort", "Supported Port")]
        [SymbolOutput("ClevisNut", "Clevis Nut")]
        [SymbolOutput("FlexMembPort", "FlexMembPort")]
        [SymbolOutput("LugPin", "Lug Pin")]
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

                Double pipeRadius = m_PipeRadius.Value;
                Double thickness = m_Thickness.Value;
                Double bendRadius = m_BendRadius.Value;
                Double holeRadius = m_HoleRadius.Value;
                Double dropValue = m_DropValue.Value;
                dropValue = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, dropValue, UnitName.DISTANCE_INCH);
                Double rodTakeOut = m_RodTakeOut.Value;
                rodTakeOut = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, rodTakeOut, UnitName.DISTANCE_INCH);
                Double clevisDiam = m_ClevisDiam.Value;
                clevisDiam = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, clevisDiam, UnitName.DISTANCE_INCH);
                Double clevisWidth = m_ClevisWidth.Value;
                clevisWidth = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, clevisWidth, UnitName.DISTANCE_INCH);
                Double rodNutHeight = m_RodNutHeight.Value;
                rodNutHeight = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, rodNutHeight, UnitName.DISTANCE_INCH);

                if (thickness == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidThickness, "Thickness cannot be zero"));
                    return;
                }

                Double bottomWidthBase = 4;
                bottomWidthBase = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, bottomWidthBase, UnitName.DISTANCE_INCH);
                Double bottomWidth = ((int)(pipeRadius / bottomWidthBase) + 1) * bottomWidthBase;
                //Correct BendRadius if necessary
                bendRadius = (bendRadius < pipeRadius) ? pipeRadius : bendRadius;

                //Correct Bottom Width if necessary
                bottomWidth = (bottomWidth / 2 > bendRadius) ? 0.95 * bendRadius : bottomWidth;
                //Calculate the LugHeight
                Double LugHeight = 1.2 * bottomWidth;
                //Calculate the Top Raidus
                Double topRadius = 0.875 * (bottomWidth / 2);
                topRadius = (topRadius < 3.5 * holeRadius) ? 3.5 * holeRadius : topRadius;

                //Correct the Clevis Diameter if necessary
                clevisDiam = (clevisDiam < 2 * 2.5 * holeRadius) ? 2 * 2.5 * holeRadius : clevisDiam;

                //Correct the Rod Take Out if necessary
                rodTakeOut = (clevisDiam > rodTakeOut) ? clevisDiam : rodTakeOut;
                rodTakeOut = (topRadius * 1.5 > rodTakeOut) ? topRadius * 1.5 : rodTakeOut;

                //Correct the Clevis Width if necessary
                clevisWidth = (clevisDiam < clevisWidth) ? clevisDiam : clevisWidth;
                clevisWidth = (holeRadius * 2 > clevisWidth) ? holeRadius * 2 * 1.5 : clevisWidth;

                //Correct the Rod Nut Height if necessary
                rodNutHeight = (rodNutHeight < clevisDiam) ? 0.45 * clevisDiam : rodNutHeight;

                Double sinTheta1 = (bendRadius - (bottomWidth / 2)) / (bendRadius + pipeRadius);
                Double theta1 = Math.Asin(sinTheta1);

                Double bottomArcEndY = (pipeRadius + bendRadius) * Math.Cos(theta1);
                Double bottomArcEndX = -(bendRadius - (bottomWidth / 2));

                Double bottomArcStartX, bottomArcStartY;
                if (pipeRadius > (bottomWidth / 2))
                {
                    Double sinTheta2 = (bendRadius + (bottomWidth / 2)) / (bendRadius + pipeRadius);
                    Double theta2 = Math.Asin(sinTheta2);

                    bottomArcStartY = (pipeRadius + bendRadius) * Math.Cos(theta2);
                    bottomArcStartX = -(bendRadius + (bottomWidth / 2));
                }
                else
                {
                    Double sinTheta3 = (bendRadius + pipeRadius) / (bendRadius + bottomWidth / 2);
                    Double theta3 = Math.Asin(sinTheta3);

                    bottomArcStartX = -(bendRadius + pipeRadius);
                    bottomArcStartY = (bendRadius + pipeRadius) * Math.Cos(theta3);
                }

                // ocurve1, ocurve2, ocurve3;
                // Double xStart, yStart, zStart, xEnd, yEnd, zEnd;
                ICurve ocurve = new Arc3d(new Position(0, 0, -thickness / 2), new Vector(0, 0, 1), new Position(bottomArcEndX, bottomArcEndY, -thickness / 2), new Position(bottomArcStartX, bottomArcStartY, -thickness / 2));
                Position startPosition = new Position();
                Position endPosition = new Position();
                ocurve.EndPoints(out startPosition, out endPosition);

                bottomArcStartX = endPosition.X;
                bottomArcStartY = endPosition.Y;

                //Create the top arc.  It center is at the hole of the lug
                Double holeX = -bendRadius;
                Double holeY = bendRadius + pipeRadius + LugHeight;
                Double topArcStartX = holeX + topRadius;
                Double topArcStartY = holeY;
                Double topArcEndX = holeX - topRadius;
                Double topArcEndY = holeY;

                ocurve = new Arc3d(new Position(holeX, holeY, -thickness / 2), new Vector(0, 0, 1), new Position(topArcStartX, topArcStartY, -thickness / 2), new Position(topArcEndX, topArcEndY, -thickness / 2));
                Position startPositionone = new Position();
                Position endPositionone = new Position();
                ocurve.EndPoints(out startPositionone, out endPositionone);

                Collection<Position> collectionPoints = new Collection<Position>();

                collectionPoints.Add(new Position(topArcEndX, topArcEndY, -thickness / 2));
                collectionPoints.Add(new Position(holeX - (bottomWidth / 2), holeY - LugHeight, -thickness / 2));
                if (pipeRadius > (bottomWidth / 2))
                {
                    collectionPoints.Add(new Position(bottomArcStartX, bottomArcStartY, -thickness / 2));
                }
                else
                {
                    collectionPoints.Add(new Position(holeX - (bottomWidth / 2), bottomArcStartY + dropValue, -thickness / 2));
                    collectionPoints.Add(new Position(bottomArcStartX, bottomArcStartY, -thickness / 2));
                }

                ocurve = new LineString3d(collectionPoints);
                Position startPositionTwo = new Position();
                Position enPostwo = new Position();
                ocurve.EndPoints(out startPositionTwo, out enPostwo);

                //Create Line String Along the Front of the lug
                Collection<Position> collectionPoints1 = new Collection<Position>();
                collectionPoints1.Add(new Position(bottomArcEndX, bottomArcEndY, -thickness / 2));
                collectionPoints1.Add(new Position(holeX + (bottomWidth / 2), holeY - LugHeight, -thickness / 2));
                collectionPoints1.Add(new Position(topArcStartX, topArcStartY, -thickness / 2));

                ocurve = new LineString3d(collectionPoints1);
                Position startPositionThree = new Position();
                Position endPositionThree = new Position();
                ocurve.EndPoints(out startPositionThree, out endPositionThree);

                //Create the Complex String Representing the outline (i.e. contuour) of the lug

                Collection<ICurve> curveElements = new Collection<ICurve>();
                curveElements.Add(new Arc3d(new Position(0, 0, -thickness / 2), new Vector(0, 0, 1), new Position(bottomArcEndX, bottomArcEndY, -thickness / 2), new Position(bottomArcStartX, bottomArcStartY, -thickness / 2)));
                curveElements.Add(new LineString3d(collectionPoints));
                curveElements.Add(new Arc3d(new Position(holeX, holeY, -thickness / 2), new Vector(0, 0, 1), new Position(topArcStartX, topArcStartY, -thickness / 2), new Position(topArcEndX, topArcEndY, -thickness / 2)));
                curveElements.Add(new LineString3d(collectionPoints1));

                //construct the lug
                Projection3d lug = new Projection3d(new ComplexString3d(curveElements), new Vector(0, 0, 1), thickness, true);

                //Add lug to output collection
                m_Symbolic.Outputs["Lug"] = lug;
                //========================================
                // Lug Pin projection
                //========================================
                //Create the lug pin
                Projection3d pin = new Projection3d(new Circle3d(new Position(holeX, holeY, -(3 * thickness)), new Vector(0, 0, 1), holeRadius), new Vector(0, 0, 1), (6 * thickness), true);

                //Clevis Ring projection
                Projection3d ring = new Projection3d(new Circle3d(new Position(holeX, holeY, -2 * thickness), new Vector(0, 0, 1), (clevisDiam / 2)), new Vector(0, 0, 1), 4 * thickness, true);

                //Clevis Nut projection
                Projection3d nut = new Projection3d(new Circle3d(new Position(holeX, (holeY + rodTakeOut), 0), new Vector(0, 1, 0), (clevisDiam / 2)), new Vector(0, 1, 0), rodNutHeight, true);

                //Add pin, ring, and clevis not to output collection
                m_Symbolic.Outputs["LugPin"] = pin;
                m_Symbolic.Outputs["ClevisRing"] = ring;
                m_Symbolic.Outputs["ClevisNut"] = nut;

                //========================================
                //  Right & Left Clevis Leg projection
                //========================================
                Collection<Position> collPoints2 = new Collection<Position>();
                collPoints2.Add(new Position(holeX - (clevisWidth / 2), holeY, -thickness / 2));
                collPoints2.Add(new Position(holeX + (clevisWidth / 2), holeY, -thickness / 2));
                collPoints2.Add(new Position(holeX + (clevisWidth / 2), holeY + rodTakeOut, -thickness / 2));
                collPoints2.Add(new Position(holeX - (clevisWidth / 2), holeY + rodTakeOut, -thickness / 2));
                collPoints2.Add(new Position(holeX - (clevisWidth / 2), holeY, -thickness / 2));

                Projection3d rightClevisLeg = new Projection3d(new LineString3d(collPoints2), new Vector(0, 0, -1), thickness, true);

                Collection<Position> collPoints3 = new Collection<Position>();
                collPoints3.Add(new Position(holeX - (clevisWidth / 2), holeY, thickness / 2));
                collPoints3.Add(new Position(holeX + (clevisWidth / 2), holeY, thickness / 2));
                collPoints3.Add(new Position(holeX + (clevisWidth / 2), holeY + rodTakeOut, thickness / 2));
                collPoints3.Add(new Position(holeX - (clevisWidth / 2), holeY + rodTakeOut, thickness / 2));
                collPoints3.Add(new Position(holeX - (clevisWidth / 2), holeY, thickness / 2));

                Projection3d leftClevisLeg = new Projection3d(new LineString3d(collPoints3), new Vector(0, 0, 1), thickness, true);
                //Add legs to output collection
                m_Symbolic.Outputs["RightClevisLeg"] = rightClevisLeg;
                m_Symbolic.Outputs["LeftClevisLeg"] = leftClevisLeg;

                // ports//
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(-bendRadius, 0, 0), new Vector(1, 0, 0), new Vector(0, -1, 0));
                m_Symbolic.Outputs["SupportedPort"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Rod", new Position(holeX, (holeY + rodTakeOut + rodNutHeight), 0), new Vector(1, 0, 0), new Vector(0, 1, 0));
                m_Symbolic.Outputs["FlexMembPort"] = port2;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of HgrElbowLug"));
                    return;
                }
            }
        }

        #endregion
    }
}
