//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PentrtPlate_1.cs
//   PipeHgrAssemblySymbols,Ingr.SP3D.Content.Support.Symbols.PentrtPlate_1
//   Author       :  Rajeswari
//   Creation Date:  03-OCT-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   03-OCT-2013  Rajeswari CR-CP-241285--Convert HgrAssemblySymbols to C# .Net
//   30/12/2014      PVK     TR-CP-264951	Resolve P3 coverity issues found in November 2014 report 
//   14/03/2016      PR      TR 288812	Graphics for penetration plate are missing in a specific case
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle.Services;

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
    [CacheOption(CacheOptionType.NonCached)]
    public class PentrtPlate_1 : HangerComponentSymbolDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PipeHgrAssemblySymbols,Ingr.SP3D.Content.Support.Symbols.PentrtPlate_1"
        //----------------------------------------------------------------------------------
        //bool firstTime = false;
        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "Thickness", "Thickness", 1)]
        public InputDouble Thickness;
        [InputDouble(3, "PipeRadius", "Pipe Radius", 0.25)]
        public InputDouble PipeRadius;
        [InputDouble(4, "Width", "Width", 2)]
        public InputDouble Width;
        [InputDouble(5, "Height", "Height", 2)]
        public InputDouble Height;
        [InputDouble(6, "PipePort_X", "Pipe Port X", 2)]
        public InputDouble PipePort_X;
        [InputDouble(7, "PipePort_Y", "Pipe Port Y", 2)]
        public InputDouble PipePort_Y;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("SimplePhysical", "SimplePhysical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("StructPort", "Structure Port")]
        [SymbolOutput("PipePort", "Pipe Port")]
        [SymbolOutput("Plate1", "Plate1")]
        [SymbolOutput("Plate2", "Plate2")]
        [SymbolOutput("Plate3", "Plate3")]
        [SymbolOutput("Plate4", "Plate4")]
        [SymbolOutput("Plate5", "Plate5")]
        [SymbolOutput("Plate6", "Plate6")]
        [SymbolOutput("Plate7", "Plate7")]
        [SymbolOutput("Plate8", "Plate8")]
        [SymbolOutput("Plate9", "Plate9")]
        [SymbolOutput("Plate10", "Plate10")]
        [SymbolOutput("Plate11", "Plate11")]
        [SymbolOutput("Plate12", "Plate12")]
        [SymbolOutput("Plate13", "Plate13")]
        [SymbolOutput("Plate14", "Plate14")]
        [SymbolOutput("Plate15", "Plate15")]
        [SymbolOutput("Plate16", "Plate16")]
        [SymbolOutput("Plate17", "Plate17")]
        [SymbolOutput("Plate18", "Plate18")]
        [SymbolOutput("Plate19", "Plate19")]
        [SymbolOutput("Plate20", "Plate20")]
        [SymbolOutput("Plate21", "Plate21")]
        [SymbolOutput("Plate22", "Plate22")]
        [SymbolOutput("Plate23", "Plate23")]
        [SymbolOutput("Plate24", "Plate24")]
        [SymbolOutput("Plate25", "Plate25")]
        [SymbolOutput("Plate26", "Plate26")]
        [SymbolOutput("Plate27", "Plate27")]
        [SymbolOutput("Plate28", "Plate28")]
        [SymbolOutput("Plate29", "Plate29")]
        [SymbolOutput("Plate30", "Plate30")]
        [SymbolOutput("Plate31", "Plate31")]
        [SymbolOutput("Plate32", "Plate32")]
        [SymbolOutput("Plate33", "Plate33")]
        [SymbolOutput("Plate34", "Plate34")]
        [SymbolOutput("Plate35", "Plate35")]
        [SymbolOutput("Plate36", "Plate36")]
        [SymbolOutput("Plate37", "Plate37")]
        [SymbolOutput("Plate38", "Plate38")]
        [SymbolOutput("Plate39", "Plate39")]
        [SymbolOutput("Plate40", "Plate40")]
        [SymbolOutput("Plate41", "Plate41")]
        [SymbolOutput("Plate42", "Plate42")]
        [SymbolOutput("Plate43", "Plate43")]
        [SymbolOutput("Plate44", "Plate44")]
        [SymbolOutput("Plate45", "Plate45")]
        [SymbolOutput("Plate46", "Plate46")]
        [SymbolOutput("Plate47", "Plate47")]
        [SymbolOutput("Plate48", "Plate48")]
        [SymbolOutput("Plate49", "Plate49")]
        [SymbolOutput("Plate50", "Plate50")]
        [SymbolOutput("Plate51", "Plate51")]
        [SymbolOutput("Plate52", "Plate52")]
        [SymbolOutput("Plate53", "Plate53")]
        [SymbolOutput("Plate54", "Plate54")]
        [SymbolOutput("Plate55", "Plate55")]
        [SymbolOutput("Plate56", "Plate56")]
        [SymbolOutput("Plate57", "Plate57")]
        [SymbolOutput("Plate58", "Plate58")]
        [SymbolOutput("Plate59", "Plate59")]
        [SymbolOutput("Plate60", "Plate60")]
        [SymbolOutput("Plate61", "Plate61")]
        [SymbolOutput("Plate62", "Plate62")]
        [SymbolOutput("Plate63", "Plate63")]
        [SymbolOutput("Plate64", "Plate64")]
        [SymbolOutput("Plate65", "Plate65")]
        [SymbolOutput("Plate66", "Plate66")]
        [SymbolOutput("Plate67", "Plate67")]
        [SymbolOutput("Plate68", "Plate68")]
        [SymbolOutput("Plate69", "Plate69")]
        [SymbolOutput("Plate70", "Plate70")]
        [SymbolOutput("Plate71", "Plate71")]
        [SymbolOutput("Plate72", "Plate72")]
        [SymbolOutput("Plate73", "Plate73")]
        [SymbolOutput("Plate74", "Plate74")]
        [SymbolOutput("Plate75", "Plate75")]
        [SymbolOutput("Plate76", "Plate76")]
        [SymbolOutput("Plate77", "Plate77")]
        [SymbolOutput("Plate78", "Plate78")]
        [SymbolOutput("Plate79", "Plate79")]
        [SymbolOutput("Plate80", "Plate80")]
        [SymbolOutput("Plate81", "Plate81")]
        [SymbolOutput("Plate82", "Plate82")]
        [SymbolOutput("Plate83", "Plate83")]
        [SymbolOutput("Plate84", "Plate84")]
        [SymbolOutput("Plate85", "Plate85")]
        [SymbolOutput("Plate86", "Plate86")]
        [SymbolOutput("Plate87", "Plate87")]
        [SymbolOutput("Plate88", "Plate88")]
        [SymbolOutput("Plate89", "Plate89")]
        [SymbolOutput("Plate90", "Plate90")]
        [SymbolOutput("Plate91", "Plate91")]
        [SymbolOutput("Plate92", "Plate92")]
        [SymbolOutput("Plate93", "Plate93")]
        [SymbolOutput("Plate94", "Plate94")]
        [SymbolOutput("Plate95", "Plate95")]
        [SymbolOutput("Plate96", "Plate96")]
        [SymbolOutput("Plate97", "Plate97")]
        [SymbolOutput("Plate98", "Plate98")]
        [SymbolOutput("Plate99", "Plate99")]
        [SymbolOutput("Plate100", "Plate100")]
        public AspectDefinition simplePhysicalAspect;

        #endregion

        #region "Construct Outputs"

        /// <summary>
        /// Construct symbol outputs in aspects.
        /// </summary>
        /// <remarks></remarks>
        SupportHelper supportHelper = null; BusinessObject businessObject = null;
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)PartInput.Value;
                Double thickness = Thickness.Value;
                Double pipeRadius = PipeRadius.Value;
                Boolean isPlaceByStruct = false;
                double width = 0, height = 0, pipePortX = 0, pipePortY = 0;

                RelationCollection hgrRelation = Occurrence.GetRelationship("SupportHasComponents", "Support");
                businessObject = hgrRelation.TargetObjects[0];
                supportHelper = new SupportHelper((Ingr.SP3D.Support.Middle.Support)businessObject);
                if (supportHelper.PlacementType == PlacementType.PlaceByStruct)
                    isPlaceByStruct = true;

                Position position = new Position(0, 0, 0);
                string boundingBox = string.Empty, boundingBoxName = string.Empty;
                BoundingBoxHelper boundingBoxHelper = new BoundingBoxHelper((Ingr.SP3D.Support.Middle.Support)businessObject);
                try
                {
                    boundingBoxHelper.CreateStandardBoundingBoxes(false);

                    // Get required information about the Bounding Box Surrounding the Pipe.
                    // The Bounding Box used depends on the command.
                    
                    if (isPlaceByStruct)
                    {
                        // Use the Structure Bounding Box
                        boundingBox = "BBS";
                        boundingBoxName = "BBS_Low";
                    }
                    else
                    {
                        // Use the Route Bouding Box
                        boundingBox = "BBR";
                        boundingBoxName = "BBR_Low";
                    }

                    if (boundingBoxHelper.GetBoundingBox(boundingBox) != null)
                    {
                        width = boundingBoxHelper.GetBoundingBox(boundingBox).Width;
                        height = boundingBoxHelper.GetBoundingBox(boundingBox).Height;
                    }
                    else
                    {
                        width = 0.0;
                        height = 0.0;
                    }


                    position = boundingBoxHelper.GetBoundingBox(boundingBox).GetRelativeRouteCenterPosition(1);
                    if (position != null)
                    {
                        pipePortX = position.X;
                        pipePortY = position.Y;
                    }
                }
                catch 
                {
                    width = 0.1;
                    height = 0.1;
                    pipePortX = 0.0;
                    pipePortY = 0.0;
                    pipeRadius = 0.05;
                    thickness = 0.025;
                }
                
                // Correct Height if necessary
                if (height <= 0)
                    height = 1.0;
                // Correct Width if necessary
                if (width <= 0)
                    width = 1.0;
                // Correct Thickness if necessary
                if (thickness <= 0.0)
                    thickness = 0.0;
                // Correct PipeRadius if necessary
                if (pipeRadius <= 0.0)
                    pipeRadius = 1.0;

                if (supportHelper.SupportedObjects == null)
                {
                    pipeRadius = 0.05;
                    thickness = 0.025;
                    width = 0.1;
                    height = 0.1;
                }

                // Set the Height value on the Plate.
                Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)businessObject;
                Occurrence.SetPropertyValue(width, "IJUAHgrOccGeometry", "Width");
                // Set the Height value on the Plate.
                Occurrence.SetPropertyValue(height, "IJUAHgrOccGeometry", "Height");
                // Set the PipePort_X value on the Plate.
                Occurrence.SetPropertyValue(pipePortX, "IJUAHgrOccPenPlate", "PipePort_X");
                // Set the PipePort_Y value on the Plate.
                Occurrence.SetPropertyValue(pipePortY, "IJUAHgrOccPenPlate", "PipePort_Y");

                // Define the offset from the provided bounding box.
                double offset = pipeRadius;
                double chamfer = 0.5 * offset;
                // ========================================
                // Bracket contour and projection
                // ========================================

                // Create Line String Representing the outline of the plate
                Collection<Position> pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(-chamfer, -offset, 0.0));
                pointCollection.Add(new Position(width + chamfer, -offset, 0.0));
                pointCollection.Add(new Position(width + offset, -chamfer, 0.0));
                pointCollection.Add(new Position(width + offset, height + chamfer, 0.0));
                pointCollection.Add(new Position(width + chamfer, height + offset, 0.0));
                pointCollection.Add(new Position(-chamfer, height + offset, 0.0));
                pointCollection.Add(new Position(-offset, height + chamfer, 0.0));
                pointCollection.Add(new Position(-offset, -chamfer, 0.0));
                pointCollection.Add(new Position(-chamfer, -offset, 0.0));

                Collection<ICurve> curveCollection = new Collection<ICurve>();
                Plane3d plane = new Plane3d(pointCollection);
                SupportedHelper supportedHelper = new SupportedHelper((Ingr.SP3D.Support.Middle.Support)businessObject);
                if (supportHelper.SupportedObjects != null)
                {
                    for (int i = 1; i <= supportHelper.SupportedObjects.Count; i++)
                    {
                        SupportedObjectInfo supportedObjectInfo = supportedHelper.SupportedObjectInfo(i);
                        if (supportedObjectInfo.GetType() == typeof(PipeObjectInfo))
                            pipeRadius = ((PipeObjectInfo)supportedObjectInfo).OutsideDiameter / 2.0;
                        else if (supportedObjectInfo.GetType() == typeof(ConduitObjectInfo))
                            pipeRadius = ((ConduitObjectInfo)supportedObjectInfo).OutsideDiameter / 2.0;
                        try
                        {
                            position = boundingBoxHelper.GetBoundingBox(boundingBox).GetRelativeRouteCenterPosition(i);
                        }
                        catch
                        {
                            position = new Position(0, 0, 0);
                        }
                        Circle3d hole = new Circle3d(new Position(position.X, position.Y, 0.0), new Vector(0.0, 0.0, 1.0), pipeRadius);
                        curveCollection.Add(hole);
                        Collection<ComplexString3d> collection = new Collection<ComplexString3d>();
                        ComplexString3d holeComplexString = new ComplexString3d(curveCollection);
                        plane.AddHole(holeComplexString);
                        curveCollection.Clear();
                    }
                }
                Line3d zVector = new Line3d(new Position(0, 0, 0), new Position(0, 0, thickness));
                Collection<Surface3d> surface = Surface3d.GetSweepSurfacesFromPlane(plane, new Line3d(zVector), (SurfaceSweepOptions)1);
                int i1 = 1;
                foreach (Surface3d item in surface)
                {
                    simplePhysicalAspect.Outputs.Add("Plate" + i1, item);
                    i1++;
                }

                // Hanger Ports
                Port structPort = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0.0, 0.0), new Vector(1.0, 0.0, 0.0), new Vector(0.0, 0.0, 1.0));
                simplePhysicalAspect.Outputs["StructPort"] = structPort;
                Port pipePort = new Port(OccurrenceConnection, part, "Route", new Position(pipePortX, pipePortY, 0.0), new Vector(1.0, 0.0, 0.0), new Vector(0.0, 0.0, 1.0));
                simplePhysicalAspect.Outputs["PipePort"] = pipePort;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PipeHgrAssemblySymbolsLocalizer.GetString(PipeHgrAssemblySymbolsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PentrtPlate_1.cs"));
                    return;
                }
            }
        }

        #endregion

        #region "ICustomHgrWeightCG Members"
        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                double width = 0, height = 0;
                Part catalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                double thickness = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAHgrThickness", "Thickness")).PropValue;
                try
                {
                    width = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrOccGeometry", "Width")).PropValue;
                }
                catch { width = 0; }
                try
                {
                    height = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrOccGeometry", "Height")).PropValue;
                }
                catch { height = 0; }
                double pipeRadius = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAHgrRadius", "PipeRadius")).PropValue;
                double offset = pipeRadius;
                double chamfer = 0.5 * offset;
                RelationCollection hgrRelation = supportComponentBO.GetRelationship("SupportHasComponents", "Support");
                businessObject = hgrRelation.TargetObjects[0];
                supportHelper = new SupportHelper((Ingr.SP3D.Support.Middle.Support)businessObject);
                SupportedHelper supportedHelper = new SupportedHelper((Ingr.SP3D.Support.Middle.Support)businessObject);
                if (supportHelper.SupportedObjects != null)
                {
                    for (int i = 1; i <= supportHelper.SupportedObjects.Count; i++)
                    {
                        SupportedObjectInfo supportedObjectInfo = supportedHelper.SupportedObjectInfo(i);
                        if (supportedObjectInfo.GetType() == typeof(PipeObjectInfo))
                            pipeRadius = ((PipeObjectInfo)supportedObjectInfo).OutsideDiameter / 2.0;
                        else if (supportedObjectInfo.GetType() == typeof(ConduitObjectInfo))
                            pipeRadius = ((ConduitObjectInfo)supportedObjectInfo).OutsideDiameter / 2.0;
                    }
                }

                double volume = (width * height * thickness) - (4 * 1 / 2 * offset * chamfer * thickness) - (Math.PI * (pipeRadius * pipeRadius) * thickness);
                CatalogStructHelper catalogStructHelper = new CatalogStructHelper();
                string materialType = (string)((PropertyValueString)catalogPart.GetPropertyValue("IJHgrPartMaterial", "MaterialType")).PropValue;
                string materialGrade = (string)((PropertyValueString)catalogPart.GetPropertyValue("IJHgrPartMaterial", "MaterialGrade")).PropValue;
                Material material = catalogStructHelper.GetMaterial(materialType, materialGrade);

                double weight = volume * material.Density;
                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, 0, 0, 0);
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PipeHgrAssemblySymbolsLocalizer.GetString(PipeHgrAssemblySymbolsResourceIDs.ErrWeightCG, "Error in EvaluateWeightCG of PentrtPlate_1.cs."));
            }
        }
        #endregion
    }
}
