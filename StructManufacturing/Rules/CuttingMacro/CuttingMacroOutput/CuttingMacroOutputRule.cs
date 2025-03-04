//-----------------------------------------------------------------------------
//      Copyright (C) 2012, Intergraph Corporation. All rights reserved.
//
//      Manufacturing Annotation Symbols Rule : Entity Name.
//
//      History:
//      Apr 16, 2014    Ninad        Creation.
//-----------------------------------------------------------------------------

using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;

using Ingr.SP3D.ReferenceData.Middle.Xml;
using Ingr.SP3D.ReferenceData.Middle.Services;

using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Structure.Middle;
using System;



namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// Class CuttingMacroOutputRule.
    /// </summary>
    public class CuttingMacroOutputRule : CuttingMacroRuleBase
    {
        /// <summary>
        /// Given a Manufacturing Part, an identifier for a library of Cutting Macro Definitions, and the
        /// context in which the rule is called, return a collection of ManufacturingMacro business objects.
        /// 
        /// The rule takes responsibility to -
        /// 1. Evaluate which reference entities (typically feature cuts) of the Manufacturing part require Macros
        /// 2. Find the most appropriate macro definition for each reference entity (or entities)
        /// 3. Compute the macro parameters required by the definition for the part
        /// 4. Construct the ManufacturingMacro business object
        /// </summary>
        public override ReadOnlyCollection<ManufacturingMacro>
            GetMacroObjectsForManufacturingPart(ManufacturingOutputBase mfgPart, int macroLibraryIdentifier, ProcessInfo.MacroRuleExecutionContextType ruleContext)
        {
            try
            {
                ReadOnlyCollection<ManufacturingMacro> persistedMacroCollection;
                Collection<BusinessObject> featuresWithoutMacros;
                persistedMacroCollection = GetPersistedMacros(mfgPart, out featuresWithoutMacros);
               
                string macroStandard = null;
                string macroVersion = null;
                GetMacroStandardVersionForLibraryIdentifier( macroLibraryIdentifier, 
                                                             out macroStandard,
                                                             out macroVersion );

                CuttingMacroCatalogService macroCatalogService = new CuttingMacroCatalogService( macroStandard,
                                                                                                 macroVersion );

                UOMManager uomMgr = MiddleServiceProvider.UOMMgr;

                Collection<ManufacturingMacro> retColl = new Collection<ManufacturingMacro>();

                if (macroLibraryIdentifier == 101) // Tubes
                {
                    //ReadOnlyCollection<CuttingMacroDefinition> macroDefinitions =
                    //    CuttingMacroCatalogService.GetMacroDefinitions(macroLibraryIdentifier);

                    MemberPart memberObj = mfgPart.DetailedPart as MemberPart;
                    if (memberObj == null) throw new CmnException("Part type and macro library identifier do not match");
                    string sectionType = memberObj.SectionType;

                    ReadOnlyCollection<CuttingMacroDefinition> definitions = macroCatalogService.GetMacroDefinitions(MacroLibraryLocation.All, sectionType);
                    if (definitions.Count <= 0) throw new CmnException("No valid definitions for member section type");
                    
                    //int[] idxOfXML = new int[macroDefinitions.Count];
                    //for ( int i = 0; i < macroDefinitions.Count; ++i )
                    //{
                    //    idxOfXML[i] = -1;
                    //    for (int j = 0; j < macroDefinitions[i].OutputFormats.Length; ++j)
                    //        if (macroDefinitions[i].OutputFormats[j].OutputFormatType == "XML")
                    //            idxOfXML[i] = j;
                    //}

                    //add all persisted macros to return collection;
                    foreach(ManufacturingMacro persistedMacro in persistedMacroCollection) 
                    {
                        if (persistedMacro.Status != MacroStatus.Generated) retColl.Add(persistedMacro);
                        else
                        {
                            Collection<BusinessObject> macroRefEntities = persistedMacro.ReferenceEntities;
                            foreach (BusinessObject refEntity in macroRefEntities) featuresWithoutMacros.Add(refEntity);
                        }
                    }


                    ReadOnlyCollection<TubeMacroInformation> tubeMacroInfoColl =
                        GetMacroInformationForTubularMembers(mfgPart);

                    Vector refVecForClTurn = null;
                    Position refPosForClOrigin = null;
                    foreach (TubeMacroInformation tubeMacroInfo in tubeMacroInfoColl)
                    {
                        //only create macros on features lacking one
                        if (persistedMacroCollection.Count > 0)
                        {
                            BusinessObject feature = tubeMacroInfo.feature;
                            if (!featuresWithoutMacros.Contains(feature)) continue;
                        }

                        bool IsStart = tubeMacroInfo.myPort.ContextId.HasFlag(
                                        Ingr.SP3D.Structure.Middle.ContextTypes.Base);
                        bool IsEnd = tubeMacroInfo.myPort.ContextId.HasFlag(
                                        Ingr.SP3D.Structure.Middle.ContextTypes.Offset);

                        if (!IsStart && !IsEnd)
                        {
                            continue;
                        }

                        bool surfacePlanar;
                        try { surfacePlanar = CheckIsSurfacePlanar(tubeMacroInfo); }
                        catch { continue; }
                        
                        for (int i = 0; i < definitions.Count; ++i)
                        {
                            //if (idxOfXML[i] == -1)
                            //    continue;

                            //select the appropriate definition
                            CuttingMacroDefinition macroDefinition = definitions[i];
                            string geometryType = macroDefinition.Maps[0].SelectionCriteria[0].ConditionData[0].GeometryType;
                            if (string.IsNullOrEmpty(geometryType)) continue;
                            else geometryType = geometryType.ToLower();
                            if (geometryType != "planar" && geometryType != "nonplanar") continue;
                            if (surfacePlanar)
                            {
                                if (geometryType != "planar") continue;
                            }
                            else
                            {
                                if (geometryType != "nonplanar") continue;
                            }

                            CuttingMacroDefinition macroDefForTube = CloneCuttingMacroDefinition(macroDefinition);
                            if (macroDefForTube == null)
                                throw new System.Xml.XmlException("Failed to clone macro definition");

                            CuttingMacroXmlFormat macroOutput =
                                macroDefForTube.OutputFormats.First(
                                x => x.OutputFormatType.Equals("xml", System.StringComparison.CurrentCultureIgnoreCase)).XmlFormat;

                            macroOutput.FaceSide = (IsStart ? "Start" : "End");
                            macroOutput.VerticalOffsetAsDouble = 0.0;

                            for (int k = 0; k < macroOutput.Parameters.Length; ++k)
                            {
                                switch (macroOutput.Parameters[ k ].Name)
                                {
                                    // Parameters common to saddle and chamfer
                                    case "clturn":
                                        double clturn = 0.0;
                                        if (refVecForClTurn == null)
                                        {
                                            refVecForClTurn = tubeMacroInfo.myAxisTangentAtConnection *
                                                              tubeMacroInfo.connectedAxisTangentAtConnection;
                                        }
                                        else
                                        {
                                            Vector crossVec = tubeMacroInfo.myAxisTangentAtConnection *
                                                              tubeMacroInfo.connectedAxisTangentAtConnection;
                                            clturn = refVecForClTurn.Angle( crossVec,
                                                                            tubeMacroInfo.
                                                                            myAxisTangentAtConnection );
                                        }
                                        macroOutput.Parameters[ k ].ValueAsDouble = clturn;
                                        macroOutput.RotationAsDouble = clturn;
                                        break;

                                    case "clorigin":
                                        double clOrigin = 0.0;
                                        if (refPosForClOrigin == null)
                                            refPosForClOrigin = tubeMacroInfo.pointOnMyAxis;
                                        else
                                            // Assume linear for now.
                                            clOrigin = refPosForClOrigin.DistanceToPoint(tubeMacroInfo.pointOnMyAxis);
                                        macroOutput.Parameters[ k ].ValueAsDouble = clOrigin;
                                        macroOutput.HorizontalOffsetAsDouble = clOrigin;
                                        break;

                                    case "slope":
                                        macroOutput.Parameters[ k ].ValueAsDouble =
                                            System.Math.Acos(
                                            tubeMacroInfo.myAxisTangentAtConnection.Dot(
                                            tubeMacroInfo.connectedAxisTangentAtConnection) );
                                        break;

                                    case "bevel1":
                                    case "rootopening":
                                    case "shrinkage":
                                        macroOutput.Parameters[ k ].ValueAsDouble = 0.0;
                                        break;

                                    case "side":
                                        macroOutput.Parameters[ k ].Value = macroOutput.FaceSide;
                                        break;


                                    // Parameters specific to saddle
                                    case "fpdiameter":
                                        macroOutput.Parameters[ k ].ValueAsDouble = 0.0;
                                        break;

                                    case "through":
                                        macroOutput.Parameters[ k ].Value =
                                            (tubeMacroInfo.IsConnectedPartGoingThrough ? "Yes" : "No");
                                        break;

                                    case "eccentricity":
                                        macroOutput.Parameters[ k ].ValueAsDouble =
                                            tubeMacroInfo.eccentricity;
                                        macroOutput.VerticalOffsetAsDouble =
                                            tubeMacroInfo.eccentricity;
                                        break;
                                } // end switch

                                if ( macroOutput.Parameters[ k ].Name != "through" &&
                                     macroOutput.Parameters[ k ].Name != "side" )
                                {
                                    macroOutput.Parameters[ k ].ValueAsDouble =
                                        macroOutput.Parameters[ k ].ValueAsDouble != null ?
                                            uomMgr.ConvertDBUtoUnit( (UnitType) macroOutput.Parameters[ k ].UnitType,
                                                (double) macroOutput.Parameters[ k ].ValueAsDouble )
                                                    :
                                                0;
                                }

                            } // for-loop around Macro output parameters

                            BusinessObject[] macroRefs = { tubeMacroInfo.feature };

                            ManufacturingMacro macroForTube = new ManufacturingMacro( mfgPart, macroRefs, macroLibraryIdentifier,
                                IsStart ? MacroLocation.EndCutAtStart : MacroLocation.EndCutAtEnd);
                            //macroForTube.Location = (IsStart ? 9 : 10); // Remind Anand to update Rule code!
                            macroForTube.XmlOutput = macroOutput;

                            
                            retColl.Add(macroForTube);
                           
                        } // for-loop around Cutting macro definitions
                    } // for-loop around Tube Macro Information objects.
                } // Tube macro library
                else if (macroLibraryIdentifier == 100) // Profiles
                {
                    ReadOnlyCollection<ManufacturingProfileSectionInformation>
                        mfgProfileXnInfo = GetCrossSectionBreakdown(mfgPart);

                    Dictionary<string, ReadOnlyCollection<CuttingMacroDefinition>>
                        macroDefinitionsForSection = new Dictionary<string,ReadOnlyCollection<CuttingMacroDefinition>>
                            (mfgProfileXnInfo.Count);

                    foreach(ManufacturingProfileSectionInformation mfgProfile in mfgProfileXnInfo)
                    {
                        if ( ! mfgProfile.manufacturedAsProfile )
                            continue;

                        /*if ( ! macroDefinitionsForSection.ContainsKey(mfgProfile.sectionType) )
                        {
                            // NB: GetMacroDefintions will thrown an exception if the input section
                            // type does not correspond to section type in the macro definition.
                            macroDefinitionsForSection[ mfgProfile.sectionType ] =
                                macroCatalogService.GetMacroDefinitions( MacroLibraryLocation.All,
                                                                         mfgProfile.sectionType );
                        }*/
                    }
                }

                macroCatalogService = null;

                return new ReadOnlyCollection<ManufacturingMacro>(retColl);
            }
            catch (System.Exception)
            {
                // Handle "gracefully" and avoid not rethrow upstream
                return null;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="macroBO"></param>
        /// <param name="macroOutputFormat"></param>
        /// <returns></returns>
        public override string GetOutputData(ManufacturingMacro macroBO, MacroOutputFormat macroOutputFormat)
        {
            if (macroOutputFormat == MacroOutputFormat.Label) return GetMacroLabel(macroBO);
            else return macroBO.XmlOutput.GetXML();            
        }

        private string GetMacroLabel(ManufacturingMacro macro)
        {

            CuttingMacroCatalogService cuttingMacroCatalogService = new CuttingMacroCatalogService(macro.Standard, macro.Version);
            CuttingMacroDefinition macroDefinition = cuttingMacroCatalogService.GetMacroDefinitions(macro.Name).First();//1 to 1
            if (macroDefinition == null) return "";

            //parse the arguments and the macro string from the definition
            string macroString;
            string macroArguments;
            if (macroDefinition.OutputFormats.FirstOrDefault(x => x.OutputFormatType.Equals("String", StringComparison.CurrentCultureIgnoreCase)) != null)
            {
                macroString =
                    macroDefinition.OutputFormats.FirstOrDefault(
                        x => x.OutputFormatType.Equals("String", StringComparison.CurrentCultureIgnoreCase)).Label.Label;

                macroArguments =
                    macroDefinition.OutputFormats.FirstOrDefault(
                        x => x.OutputFormatType.Equals("String", StringComparison.CurrentCultureIgnoreCase)).Label.EvaluationArguments;
            }
            else return "";

            string returnMacroString;

            string[] arguments = macroArguments.Split(',');
            for (int i = 0; i < arguments.Length; i++)  arguments[i] = arguments[i].Trim();

            string[] macroParts = macroString.Split(',');
            for (int i = 0; i < macroParts.Length; i++) macroParts[i] = macroParts[i].Trim();

            for (int i = 1; i < macroParts.Length; i++)
            {
                string arg = arguments[i - 1];

                if (arg.Equals("Face_Side", StringComparison.CurrentCultureIgnoreCase)) macroParts[i] = 
                    macro.XmlOutput.FaceSide;
                else if (arg.Equals("Offset_X", StringComparison.CurrentCultureIgnoreCase)) macroParts[i] = 
                    Math.Round(double.Parse(macro.XmlOutput.HorizontalOffset), 3).ToString();
                else if (arg.Equals("Offset_Y", StringComparison.CurrentCultureIgnoreCase)) macroParts[i] =
                    Math.Round(double.Parse(macro.XmlOutput.VerticalOffset), 3).ToString();
                else if (arg.Equals("Rotation", StringComparison.CurrentCultureIgnoreCase)) macroParts[i] = 
                    Math.Round(double.Parse(macro.XmlOutput.Rotation), 3).ToString();
                else macroParts[i] = arg;
            }

            returnMacroString = macroParts[0];

            for (int i = 1; i < macroParts.Length; i++)
                returnMacroString += "," + macroParts[i];

            cuttingMacroCatalogService = null;

            return returnMacroString;
        }
        
    }
}
