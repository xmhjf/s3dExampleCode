//-----------------------------------------------------------------------------
//      Copyright (C) 2012, Intergraph Corporation. All rights reserved.
//
//      Profile Stock Nesting Report Query Interpreter implementation.
//
//      Author: Nautilus - HSV  
//
//      History:
//      10-15-2012    Brandon Affenzeller    Created
//
//      ToDo:
//          change object arrays to double/string arrays
//  	
//      Modified  on  25 OCT  2013  Praveen Babu  Oracle support
//      Modified  on  13 Mar  2015  Praveen Babu  TR-CP-269163	Structural Manufacturing - Several Catalog Reports Fail on Oracle
//-----------------------------------------------------------------------------

using System;
using System.Data;
using System.Data.OleDb;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Reports.Middle;
using Ingr.SP3D.Reports.Exceptions;

using Ingr.SP3D.Interop.StrMfgProfileStockNesting;

namespace StrMfgProfileStockQI
{
    /// <summary>
    /// Class ProfileStockQI.
    /// </summary>
    public class ProfileStockQI : QueryInterpreter
    {
        private DataTable m_DataTable;

        #region Constants

        private const string FIELD_NESTING_OID =        "NestingOid";
        private const string FIELD_OIDS =               "Oids";
        private const string FIELD_OID =                "Oid";
        private const string FIELD_REFERENCE_STANDARD = "ReferenceStandard";
        private const string FIELD_SECTION_TYPE =       "SectionType";
        private const string FIELD_SECTION_NAME =       "SectionName";
        private const string FIELD_MATERIAL_TYPE =      "MaterialType";
        private const string FIELD_MATERIAL_GRADE =     "MaterialGrade";
        private const string FIELD_PART_LENGTHS =       "PartLengths";
        private const string FIELD_QUANTITY =           "Quantity";
        private const string FIELD_LENGTH_USED =        "UsedLength";
        private const string FIELD_LENGTH_UNUSED =      "UnusedLength";
        private const string FIELD_WEIGHT_USED =        "WeightUsed";
        private const string FIELD_WEIGHT_UNUSED =      "WeightUnused";
        private const string FIELD_PERCENTAGE_USED =    "Percentageused";
        private const string FIELD_PERCENTAGE_UNUSED =  "PercentageUnused";
        private const string FIELD_PART_NAMES =         "PartNames";
        private const string FIELD_RANK =               "EntityRank";
        private const string FIELD_PART_NAME =          "PartName";
        private const string FIELD_DENSITY =            "MaterialDensity";
        private const string FIELD_AREA =               "SectionArea";
        //private const string FIELD_PROFILE_SKETCH_DWG = "ProfileSketchDwg";
        private const string FIELD_PART_LENGTH =        "PartLength";
        private const string FIELD_STOCK_LENGTH =       "StockLength";
        private const string FIELD_LENGTH =             "Length";

        private const string REPORT_FORMAT_ONSET = "ONSET";
        private const string REPORT_FORMAT_OFFSET = "OFFSET";

        private const string REPORT_MEMBER = "MEMBER";
        private const string REPORT_PROFILE = "PROFILE";
        private const string REPORT_MEMBER_PROFILE = "PROFILE_MEMBER";

        private const string REPORT_ALL = "ALL";
        private const string REPORT_EXCLUDE = "EXCLUDE";

        private const string REPORT_END_OFFSETS = "ENDOFFSETS";
        private const string REPORT_MIDDLE_OFFSETS = "MIDDLEOFFSETS";
        private const string REPORT_STOCKS = "STOCKS";

        #endregion

        #region Hard Coded Profile Stock Query

        private const string stockQuery =
                @"SELECT
                    JRS.Name As ReferenceStandard,
                    JSCS.SectionName,
                    JDM.MaterialType,
                    JDM.MaterialGrade,
                    JDPS.Length As StockLength,
                    DENSE_RANK() OVER(Order BY 
                                        JRS.Name,
                                        JSCS.SectionName,
                                        JDM.MaterialType,
                                        JDM.MaterialGrade)
                                        As EntityRank
                						
                FROM JDProfileStock JDPS

                --Profile Stock Material Type and Grade
                INNER JOIN XProfileStockMaterial XPSM ON XPSM.OidDestination = JDPS.Oid
                INNER JOIN JDMaterial JDM ON JDM.Oid = XPSM.OidOrigin

                -- Profile Cross Section
                INNER JOIN XStockCrossSection XSCS ON XSCS.OidDestination = JDPS.Oid
                INNER JOIN JSTRUCTCrossSection JSCS ON JSCS.Oid = XSCS.OidOrigin

                -- Reference Standard
                INNER JOIN JDCrossSection JDCS ON JDCS.Oid = JSCS.Oid
                INNER JOIN JDPartClass JDPC ON JDPC.Name = JDCS.Type
                INNER JOIN XReferenceStdHasPartClasses XRSHPC 
                    ON XRSHPC.OidDestination = JDPC.Oid
                INNER JOIN JReferenceStandard JRS ON JRS.Oid = XRSHPC.OidOrigin

                ORDER BY EntityRank, JDPS.Length";

        #endregion

        #region Hard Coded Profile Query

        private const string profileQueryTemplate =
                @"SET NOCOUNT ON
                SET QUOTED_IDENTIFIER ON
                SET ANSI_NULLS ON
                 
                if exists (select * from tempdb..sysobjects where ( id = OBJECT_ID(N'[tempdb]..[#TempProfileLengthTbl]') ))
                    DROP TABLE #TempProfileLengthTbl

                -- Create temporary table to contain the profile oids and lengths
                CREATE TABLE #TempProfileLengthTbl(oid Uniqueidentifier, PartLength FLOAT)

                CREATE INDEX i_LengthOid ON #TempProfileLengthTbl (oid)

                INSERT INTO #TempProfileLengthTbl
                SELECT 
                    PP.Oid,
                    PP.ProfileLength AS PartLength
                FROM JProfilePart PP
                @Assemblies

                if exists (select * from tempdb..sysobjects where ( id = OBJECT_ID(N'[tempdb]..[#TempProfileSectionTbl]') ))
                    DROP TABLE #TempProfileSectionTbl

                -- create temporaty table to contain the reference standard, section type,
                -- section name, and section area for each profile in #TempProfileLengthTbl
                CREATE TABLE #TempProfileSectionTbl(oid Uniqueidentifier, ReferenceStandard NVARCHAR(256), SectionType NVARCHAR(256), SectionName NVARCHAR(256), SectionArea FLOAT)

                CREATE INDEX i_SectionOid ON #TempProfileSectionTbl (oid)

                INSERT INTO #TempProfileSectionTbl
                SELECT 
                    PP.Oid,
                    JRS.Name As ReferenceStandard,
                    JDCS.Type As SectionType,
                    COALESCE(CADCS1.csSectName, 
                             CADCS2.csSectName, 
                             CADCS3.csSectName, 
                             CADCS4.csSectName)
                             As SectionName,
                    JSCSD.Area
                		
                FROM #TempProfileLengthTbl PP
                -- Profile Section Name
                INNER JOIN XShpStrDesignHierarchy XSSDH1 ON XSSDH1.OidDestination = PP.Oid
                INNER JOIN XShpStrDesignHierarchy XSSDH2 ON XSSDH2.OidDestination = XSSDH1.OidOrigin
                INNER JOIN XShpStrDesignHierarchy XSSDH3 ON XSSDH3.OidDestination = XSSDH2.OidOrigin
                LEFT JOIN GSCADGeom GEOM1 ON GEOM1.goidDest = PP.Oid
                LEFT JOIN GSCADGeom GEOM2 ON GEOM2.goidDest = XSSDH1.oidOrigin
                LEFT JOIN GSCADGeom GEOM3 ON GEOM3.goidDest = XSSDH2.oidOrigin
                LEFT JOIN GSCADGeom GEOM4 ON GEOM4.goidDest = XSSDH3.oidOrigin
                LEFT JOIN GSCADCrossSection CADCS1 ON CADCS1.oidOrigin = GEOM1.gOidOrigin
                LEFT JOIN GSCADCrossSection CADCS2 ON CADCS2.oidOrigin = GEOM2.gOidOrigin
                LEFT JOIN GSCADCrossSection CADCS3 ON CADCS3.oidOrigin = GEOM3.gOidOrigin
                LEFT JOIN GSCADCrossSection CADCS4 ON CADCS4.oidOrigin = GEOM4.gOidOrigin
                	                
                -- Profile Cross Section Area
                INNER JOIN JSTRUCTCrossSectionDimensions JSCSD ON JSCSD.Oid = COALESCE(CADCS1.csOid, 
												                                       CADCS2.csOid, 
												                                       CADCS3.csOid, 
												                                       CADCS4.csOid)

                -- Reference Standard
                INNER JOIN JDCrossSection JDCS ON JDCS.Oid = JSCSD.Oid                                       
                INNER JOIN JDPartClass JDPC ON JDPC.Name = JDCS.Type
                INNER JOIN XReferenceStdHasPartClasses XRSHPC 
                    ON XRSHPC.OidDestination = JDPC.Oid
                INNER JOIN JReferenceStandard JRS ON JRS.Oid = XRSHPC.OidOrigin

                if exists (select * from tempdb..sysobjects where ( id = OBJECT_ID(N'[tempdb]..[#TempProfileMaterialTbl]') ))
                    DROP TABLE #TempProfileMaterialTbl

                -- create temporaty table to contain the material type, material grade, and
                -- density for each profile in #TempProfileLengthTbl
                CREATE TABLE #TempProfileMaterialTbl(oid Uniqueidentifier, MaterialType NVARCHAR(256), MaterialGrade NVARCHAR(256), MaterialDensity FLOAT)

                CREATE INDEX i_MaterialOid ON #TempProfileMaterialTbl (oid)

                INSERT INTO #TempProfileMaterialTbl
                SELECT
                    PP.Oid,
                    JDM.MaterialType,
                    JDM.MaterialGrade,
                    JDM.Density As MaterialDensity
                FROM #TempProfileLengthTbl PP
                -- Profile Material Type, Grade, and Density
                INNER JOIN XSystemHasMaterial XSHM 
                    ON XSHM.OidDestination = dbo.REPORTGetParentRelationOid(PP.Oid, 'SystemHasMaterial')
                INNER JOIN JDMaterial JDM ON JDM.Oid = XSHM.OidOrigin

                if exists (select * from tempdb..sysobjects where ( id = OBJECT_ID(N'[tempdb]..[#TempProfileNameTbl]') ))
                    DROP TABLE #TempProfileNameTbl

                -- create temporaty table to contain the names for each profile in
                -- #TempProfileLengthTbl
                CREATE TABLE #TempProfileNameTbl(oid Uniqueidentifier, PartName NVARCHAR(256))

                CREATE INDEX i_NameOid ON #TempProfileNameTbl (oid)

                INSERT INTO #TempProfileNameTbl
                SELECT 
                    PP.Oid,
                    JNI.ItemName As PartName
                FROM #TempProfileLengthTbl PP
                -- ProfileName
                INNER JOIN JNamedItem JNI ON PP.Oid = JNI.Oid	

                -- use the temporary tables to get the desired query 
                SELECT
                    #TempProfileLengthTbl.oid,
                    #TempProfileLengthTbl.PartLength,
                    #TempProfileSectionTbl.ReferenceStandard,
                    #TempProfileSectionTbl.SectionArea,
                    #TempProfileSectionTbl.SectionName,
                    #TempProfileSectionTbl.SectionType,
                    #TempProfileMaterialTbl.MaterialDensity,
                    #TempProfileMaterialTbl.MaterialGrade,
                    #TempProfileMaterialTbl.MaterialType,
                    #TempProfileNameTbl.PartName,
                    DENSE_RANK() OVER (ORDER BY
		                                    #TempProfileSectionTbl.ReferenceStandard,
		                                    #TempProfileSectionTbl.SectionName,
		                                    #TempProfileMaterialTbl.MaterialGrade,
		                                    #TempProfileMaterialTbl.MaterialType) 
		                                    As EntityRank
                FROM #TempProfileLengthTbl, #TempProfileMaterialTbl, #TempProfileSectionTbl, #TempProfileNameTbl
                WHERE #TempProfileLengthTbl.oid = #TempProfileSectionTbl.oid AND
                      #TempProfileMaterialTbl.oid = #TempProfileLengthTbl.oid AND
                      #TempProfileNameTbl.oid = #TempProfileLengthTbl.oid
                	  
                ORDER BY EntityRank, PartLength

                -- delete temporary tables
                DROP TABLE #TempProfileLengthTbl
                DROP TABLE #TempProfileSectionTbl
                DROP TABLE #TempProfileMaterialTbl
                DROP TABLE #TempProfileNameTbl";

        private const string OraprofileQueryTemplate =

                @"WITH
                TempProfileLengthTbl(oid, PartLength)
                as
                (
	                SELECT 
	                PP.Oid,
	                PP.ProfileLength AS PartLength
	                FROM JProfilePart PP
	                @Assemblies
                ),

                TempProfileSectionTbl(oid , ReferenceStandard , SectionType , SectionName, SectionArea )
                as
                (
	                SELECT 
	                PP.Oid,
	                JRS.Name As ReferenceStandard,
	                JDCS.Type As SectionType,
	                COALESCE(CADCS1.csSectName, 
	                CADCS2.csSectName, 
	                CADCS3.csSectName, 
	                CADCS4.csSectName)
	                As SectionName,
	                JSCSD.Area
	                FROM TempProfileLengthTbl PP
	                INNER JOIN XShpStrDesignHierarchy XSSDH1 ON XSSDH1.OidDestination = PP.Oid
	                INNER JOIN XShpStrDesignHierarchy XSSDH2 ON XSSDH2.OidDestination = XSSDH1.OidOrigin
	                INNER JOIN XShpStrDesignHierarchy XSSDH3 ON XSSDH3.OidDestination = XSSDH2.OidOrigin
	                LEFT JOIN GSCADGeom GEOM1 ON GEOM1.goidDest = PP.Oid
	                LEFT JOIN GSCADGeom GEOM2 ON GEOM2.goidDest = XSSDH1.oidOrigin
	                LEFT JOIN GSCADGeom GEOM3 ON GEOM3.goidDest = XSSDH2.oidOrigin
	                LEFT JOIN GSCADGeom GEOM4 ON GEOM4.goidDest = XSSDH3.oidOrigin
	                LEFT JOIN GSCADCrossSection CADCS1 ON CADCS1.oidOrigin = GEOM1.gOidOrigin
	                LEFT JOIN GSCADCrossSection CADCS2 ON CADCS2.oidOrigin = GEOM2.gOidOrigin
	                LEFT JOIN GSCADCrossSection CADCS3 ON CADCS3.oidOrigin = GEOM3.gOidOrigin
	                LEFT JOIN GSCADCrossSection CADCS4 ON CADCS4.oidOrigin = GEOM4.gOidOrigin

	                INNER JOIN JSTRUCTCrossSectionDimensions JSCSD ON JSCSD.Oid = COALESCE(CADCS1.csOid, 
	                CADCS2.csOid, 
	                CADCS3.csOid, 
	                CADCS4.csOid)

	                INNER JOIN JDCrossSection JDCS ON JDCS.Oid = JSCSD.Oid                                       
	                INNER JOIN JDPartClass JDPC ON JDPC.Name = JDCS.Type
	                INNER JOIN XReferenceStdHasPartClasses XRSHPC 
	                ON XRSHPC.OidDestination = JDPC.Oid
	                INNER JOIN JReferenceStandard JRS ON JRS.Oid = XRSHPC.OidOrigin
                ),

                TempProfileMaterialTbl(oid , MaterialType , MaterialGrade , MaterialDensity)
                as
                (
	                SELECT
	                PP.Oid,
	                JDM.MaterialType,
	                JDM.MaterialGrade,
	                JDM.Density As MaterialDensity
	                FROM TempProfileLengthTbl PP
	                INNER JOIN XSystemHasMaterial XSHM 
	                ON XSHM.OidDestination = REPORTGetParentRelationOid(PP.Oid, 'SystemHasMaterial')
	                INNER JOIN JDMaterial JDM ON JDM.Oid = XSHM.OidOrigin
                ),

                TempProfileNameTbl(oid , PartName )
                as
                (
	                SELECT 
	                PP.Oid,
	                JNI.ItemName As PartName
	                FROM TempProfileLengthTbl PP
	                INNER JOIN JNamedItem JNI ON PP.Oid = JNI.Oid	
                )

                SELECT
                ProfileLength.oid,
                ProfileLength.PartLength,
                ProfileSection.ReferenceStandard,
                ProfileSection.SectionArea,
                ProfileSection.SectionName,
                ProfileSection.SectionType,
                ProfileMaterial.MaterialDensity,
                ProfileMaterial.MaterialGrade,
                ProfileMaterial.MaterialType,
                ProfileName.PartName,
                DENSE_RANK() OVER (ORDER BY
                ProfileSection.ReferenceStandard,
                ProfileSection.SectionName,
                ProfileMaterial.MaterialGrade,
                ProfileMaterial.MaterialType) 
                As EntityRank
                FROM TempProfileLengthTbl ProfileLength, TempProfileMaterialTbl ProfileMaterial, 
                TempProfileSectionTbl ProfileSection, TempProfileNameTbl ProfileName
                WHERE ProfileLength.oid = ProfileSection.oid AND
                ProfileMaterial.oid = ProfileLength.oid AND
                ProfileName.oid = ProfileLength.oid

                ORDER BY EntityRank, PartLength";


        private const string queryAssemblyStart = @"
                -- Profiles that are children of the specified assembly or block
                INNER JOIN (";

        private const string queryAssemblyMiddle = @"
                SELECT *
			    FROM REPORTGetAllChildrenInHierarchyByOID( ?, 'AssemblyHierarchy')";

        private const string OraqueryAssemblyMiddle = @"
                SELECT *
			    FROM TABLE(RPTAllChildrenInHierarchyByOID( ?, 'AssemblyHierarchy'))";

        private const string queryAssemblyEnd = @"
                ) AH ON PP.Oid = AH.oidChild";

        #endregion

        #region Hard Coded Member Query

        private const string memberQueryTemplate =
                @"SET NOCOUNT ON
                SET QUOTED_IDENTIFIER ON
                SET ANSI_NULLS ON

                if exists (select * from tempdb..sysobjects where ( id = OBJECT_ID(N'[tempdb]..[#TempLengthTbl]') ))
                    DROP TABLE #TempLengthTbl

                CREATE TABLE #TempLengthTbl(oid Uniqueidentifier, PartLength FLOAT)

                CREATE INDEX i_LengthOid ON #TempLengthTbl (oid)

                INSERT INTO #TempLengthTbl
                SELECT 
	                PP.Oid,
	                PP.CutLength AS PartLength
                FROM SPSMemberPartPrismatic PP
                @Assemblies

                if exists (select * from tempdb..sysobjects where ( id = OBJECT_ID(N'[tempdb]..[#TempSectionTbl]') ))
                    DROP TABLE #TempSectionTbl

                CREATE TABLE #TempSectionTbl(oid Uniqueidentifier, ReferenceStandard NVARCHAR(256), SectionType NVARCHAR(256), SectionName NVARCHAR(256), SectionArea FLOAT)

                CREATE INDEX i_SectionOid ON #TempSectionTbl (oid)

                INSERT INTO #TempSectionTbl
                SELECT
	                PP.Oid,
	                JRS.Name As ReferenceStandard,
	                JDCS.Type As SectionType,
	                JSCS.SectionName,
	                JSCSD.Area As SectionArea
                FROM #TempLengthTbl PP
                INNER JOIN YSPSMemberPartToCrossSectionEd YMCS ON YMCS.OidDestination = PP.Oid
                INNER JOIN JDCrossSection JDCS ON JDCS.Oid = YMCS.OidOrigin
                INNER JOIN JSTRUCTCrossSection JSCS ON JSCS.Oid = YMCS.OidOrigin
                INNER JOIN JSTRUCTCrossSectionDimensions JSCSD ON JSCSD.Oid = JSCS.Oid

                INNER JOIN YSPSMemberPartToCSectionStdEd YMCSS ON YMCSS.OidDestination = PP.Oid
                INNER JOIN JReferenceStandard JRS ON JRS.Oid = YMCSS.OidOrigin

                if exists (select * from tempdb..sysobjects where ( id = OBJECT_ID(N'[tempdb]..[#TempMaterialTbl]') ))
                    DROP TABLE #TempMaterialTbl

                CREATE TABLE #TempMaterialTbl(oid Uniqueidentifier, MaterialType NVARCHAR(256), MaterialGrade NVARCHAR(256), MaterialDensity FLOAT)

                CREATE INDEX i_MaterialOid ON #TempMaterialTbl (oid)

                INSERT INTO #TempMaterialTbl
                SELECT 
	                PP.Oid,
	                JDM.MaterialType,
	                JDM.MaterialGrade,
	                JDM.Density As MaterialDensity
                FROM #TempLengthTbl PP
                INNER JOIN YSPSMemberPartToMaterialEd MPM ON PP.Oid = MPM.OidDestination
                INNER JOIN JDMaterial JDM ON JDM.Oid = MPM.OidOrigin

                if exists (select * from tempdb..sysobjects where ( id = OBJECT_ID(N'[tempdb]..[#TempNameTbl]') ))
                    DROP TABLE #TempNameTbl

                CREATE TABLE #TempNameTbl(oid Uniqueidentifier, PartName NVARCHAR(256))

                CREATE INDEX i_NameOid ON #TempNameTbl (oid)

                INSERT INTO #TempNameTbl
                SELECT 
	                PP.Oid,
	                JNI.ItemName As PartName
                FROM #TempLengthTbl PP
                -- ProfileName
                INNER JOIN JNamedItem JNI ON PP.Oid = JNI.Oid	

                SELECT 
	                #TempLengthTbl.oid,
	                #TempLengthTbl.PartLength,
	                #TempSectionTbl.ReferenceStandard,
	                #TempSectionTbl.SectionArea,
	                #TempSectionTbl.SectionName,
	                #TempSectionTbl.SectionType,
	                #TempMaterialTbl.MaterialDensity,
	                #TempMaterialTbl.MaterialGrade,
	                #TempMaterialTbl.MaterialType,
	                #TempNameTbl.PartName,
	                DENSE_RANK() OVER (ORDER BY 
							                #TempSectionTbl.ReferenceStandard,
							                #TempSectionTbl.SectionName,
							                #TempMaterialTbl.MaterialGrade,
							                #TempMaterialTbl.MaterialType)
							                As EntityRank
                FROM #TempLengthTbl, #TempSectionTbl, #TempMaterialTbl, #TempNameTbl
                WHERE #TempLengthTbl.oid = #TempSectionTbl.oid AND
	                  #TempLengthTbl.oid = #TempMaterialTbl.oid AND
	                  #TempLengthTbl.oid = #TempNameTbl.oid
                	  
                ORDER BY EntityRank, PartLength

                DROP TABLE #TempLengthTbl
                DROP TABLE #TempSectionTbl
                DROP TABLE #TempMaterialTbl
                DROP TABLE #TempNameTbl";

        private const string OramemberQueryTemplate =

                @"WITH
                TempLengthTbl(oid , PartLength)
                AS
                (
                SELECT 
	                PP.Oid,
	                PP.CutLength AS PartLength
	                FROM SPSMemberPartPrismatic PP
	                @Assemblies
                ),

                TempSectionTbl(oid , ReferenceStandard , SectionType , SectionName , SectionArea )
                AS
                (
	                SELECT
	                PP.Oid,
	                JRS.Name As ReferenceStandard,
	                JDCS.Type As SectionType,
	                JSCS.SectionName,
	                JSCSD.Area As SectionArea
	                FROM TempLengthTbl PP
	                INNER JOIN YSPSMemberPartToCrossSectionEd YMCS ON YMCS.OidDestination = PP.Oid
	                INNER JOIN JDCrossSection JDCS ON JDCS.Oid = YMCS.OidOrigin
	                INNER JOIN JSTRUCTCrossSection JSCS ON JSCS.Oid = YMCS.OidOrigin
	                INNER JOIN JSTRUCTCrossSectionDimensions JSCSD ON JSCSD.Oid = JSCS.Oid

	                INNER JOIN YSPSMemberPartToCSectionStdEd YMCSS ON YMCSS.OidDestination = PP.Oid
	                INNER JOIN JReferenceStandard JRS ON JRS.Oid = YMCSS.OidOrigin
                ),

                TempMaterialTbl(oid , MaterialType , MaterialGrade, MaterialDensity )
                AS
                (
	                SELECT 
	                PP.Oid,
	                JDM.MaterialType,
	                JDM.MaterialGrade,
	                JDM.Density As MaterialDensity
	                FROM TempLengthTbl PP
	                INNER JOIN YSPSMemberPartToMaterialEd MPM ON PP.Oid = MPM.OidDestination
	                INNER JOIN JDMaterial JDM ON JDM.Oid = MPM.OidOrigin
                ),

                TempNameTbl(oid , PartName )
                AS
                (
	                SELECT 
	                PP.Oid,
	                JNI.ItemName As PartName
	                FROM TempLengthTbl PP
	                INNER JOIN JNamedItem JNI ON PP.Oid = JNI.Oid	
                )

                SELECT 
                Length.oid,
                Length.PartLength,
                Section.ReferenceStandard,
                Section.SectionArea,
                Section.SectionName,
                Section.SectionType,
                Material.MaterialDensity,
                Material.MaterialGrade,
                Material.MaterialType,
                Name.PartName,
                DENSE_RANK() OVER (ORDER BY 
                Section.ReferenceStandard,
                Section.SectionName,
                Material.MaterialGrade,
                Material.MaterialType
               )
                As EntityRank
                FROM TempLengthTbl Length, TempSectionTbl Section, TempMaterialTbl Material, TempNameTbl Name
                WHERE Length.oid = Section.oid AND
                Length.oid = Material.oid AND
                Length.oid = Name.oid

                ORDER BY EntityRank, PartLength";


        #endregion

        bool IsOracleDB = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.DBProvider.IsOracleProvider();

        /// <summary>
        /// Initializes a new instance of the <see cref="ProfileStockQI" /> class.
        /// </summary>
        public ProfileStockQI()
        {
        }

        /// <summary>
        /// Executes the specified action.
        /// </summary>
        /// <param name="action">The action.</param>
        /// <param name="argument">The argument.</param>
        /// <returns>DataTable.</returns>
        /// <exception cref="System.NotImplementedException"></exception>
        public override DataTable Execute(string action, string argument)
        {
            try
            {
                m_DataTable = InitializeDataTable(CreateReturnDataTable());

                if (EvaluateOnly)
                {
                    return m_DataTable;
                }

                string strReportType = String.Empty;
                string strReportFormat = String.Empty;
                string testNesting = String.Empty;
                double[] dEndOffsetArr = null;
                double[] dMiddleOffsetArr = null;
                double[] defaultStocks = null;
                ParseArguments(argument, ref strReportType, ref strReportFormat, ref testNesting, ref dEndOffsetArr, ref dMiddleOffsetArr, ref defaultStocks);

                SetQuery(stockQuery);

                DataTable partStockTable = ExecuteDelegatedQuery();

                string[] partQueries;

                if (strReportType == REPORT_MEMBER_PROFILE)
                {
                    partQueries = new string[2];
                    if (IsOracleDB)
                    {
                        partQueries[0] = GetFormattedPartsQuery(OraprofileQueryTemplate);
                        partQueries[1] = GetFormattedPartsQuery(OramemberQueryTemplate);
                    }
                    else
                    {
                        partQueries[0] = GetFormattedPartsQuery(profileQueryTemplate);
                        partQueries[1] = GetFormattedPartsQuery(memberQueryTemplate);
                    }
                    
                }
                else if (strReportType == REPORT_MEMBER)
                {
                    partQueries = new string[1];
                    if (IsOracleDB)
                    {
                        partQueries[0] = GetFormattedPartsQuery(OramemberQueryTemplate);
                    }
                    else
                    {
                        partQueries[0] = GetFormattedPartsQuery(memberQueryTemplate);
                    }
                }
                else
                {
                    partQueries = new string[1];
                    if (IsOracleDB)
                    {
                        partQueries[0] = GetFormattedPartsQuery(OraprofileQueryTemplate);
                    }
                    else
                    {
                        partQueries[0] = GetFormattedPartsQuery(profileQueryTemplate);
                    }
                    
                }

                for (int queryIndex = 0; queryIndex < partQueries.Length; queryIndex++)
                {
                    // The profile query needs to be able to query with respect to the
                    // entire model or the select set. The queries for each are slightly
                    // different, and the select set query  requires processing of
                    // the select set to build the query. GetProfilesQuery handles
                    // all of this.
                    //string profileQuery = GetProfilesQuery();
                    //string partQuery = GetFormattedPartsQuery(profileQueryTemplate);
                    string partQuery = partQueries[queryIndex];

                    SetQuery(partQuery);

                    DataTable partTable = ExecuteDelegatedQuery();

                    // ToDo: Use ProgID to create the object containing the nesting algorithm.
                    StrMfgProfileStockNesting profileStockNesting = new StrMfgProfileStockNesting();

                    // The profiles are sent to the nesting algorithm in groups based
                    // on the rank assigned to ea ch profile in the query. The relevant
                    // profile stocks are sent to the algorithm based on the profiles
                    // in the current group.
                    long maxPartRank = 0;

                    // If the query returned any profiles, get the largest rank value.
                    if (partTable.Rows.Count > 0)
                    {
                        // Since the query results are sorted by the rank, the rank
                        // from the last entry of the table is the highest rank.
                        maxPartRank = Convert.ToInt64(partTable.Rows[partTable.Rows.Count - 1][FIELD_RANK]);
                    }

                    // For each rank, get the profiles and the relevant profile stocks
                    // and pass them to the nesting algorithm.
                    for (long partRank = 1; partRank <= maxPartRank; partRank++)
                    {
                        // Create a new table with the same structure as the profile table.
                        // We want to get sub tables of profiles with the same
                        // refereence standard, section name, material grade, and material type.
                        DataTable subPartTable = partTable.Clone();

                        // get all rows that have the current rank
                        var partSubQuery =
                            from row in partTable.AsEnumerable()
                            where Convert.ToInt64(row[FIELD_RANK]) == partRank
                            select row;

                        // For each row that contains the current rank, copy it to
                        // the sub table.
                        foreach (DataRow row in partSubQuery.ToList<DataRow>())
                        {
                            subPartTable.ImportRow(row);
                        }

                        if (subPartTable.Rows[0][FIELD_SECTION_TYPE].ToString() == "FB")
                        {
                            continue;
                        }

                        object[] arrPartOIDs = new object[subPartTable.Rows.Count];
                        object[] arrPartNames = new object[subPartTable.Rows.Count];
                        object[] arrPartLengths = new object[subPartTable.Rows.Count];

                        for (int index = 0; index < subPartTable.Rows.Count; index++)
                        {
                            arrPartOIDs[index] = subPartTable.Rows[index][FIELD_OID].ToString();
                            arrPartNames[index] = subPartTable.Rows[index][FIELD_PART_NAME].ToString();
                            arrPartLengths[index] = subPartTable.Rows[index][FIELD_PART_LENGTH];
                        }

                        // get all profile stocks of the specified type
                        var partStockSubQuery =
                            from row in partStockTable.AsEnumerable()
                            where row[FIELD_REFERENCE_STANDARD].ToString() == subPartTable.Rows[0][FIELD_REFERENCE_STANDARD].ToString()
                                && row[FIELD_SECTION_NAME].ToString() == subPartTable.Rows[0][FIELD_SECTION_NAME].ToString()
                                && row[FIELD_MATERIAL_TYPE].ToString() == subPartTable.Rows[0][FIELD_MATERIAL_TYPE].ToString()
                                && row[FIELD_MATERIAL_GRADE].ToString() == subPartTable.Rows[0][FIELD_MATERIAL_GRADE].ToString()
                            select row;

                        // Create a new table with the same structure as the profile table.
                        DataTable subPartStockTable = partStockTable.Clone();

                        // For each row that contains the specified type, copy it to
                        // the sub table.
                        foreach (DataRow dataRow in partStockSubQuery.ToList<DataRow>())
                        {
                            subPartStockTable.ImportRow(dataRow);
                        }

                        object[] arrStockLengths = new object[subPartStockTable.Rows.Count];

                        for (int index = 0; index < subPartStockTable.Rows.Count; index++)
                        {
                            arrStockLengths[index] = subPartStockTable.Rows[index][FIELD_STOCK_LENGTH];
                        }

                        Array nestingSolution = null;

                        nestingSolution = (Array)profileStockNesting.GetNestingSolution(0,
                                                                    (Array)arrStockLengths,
                                                                    (Array)arrPartOIDs,
                                                                    (Array)arrPartNames,
                                                                    (Array)arrPartLengths,
                                                                    dEndOffsetArr[queryIndex],
                                                                    dMiddleOffsetArr[queryIndex],
                                                                    (Array)Array.ConvertAll<double, object>(defaultStocks,
                                                                    delegate(double d)
                                                                    {
                                                                        return (object)d;
                                                                    }));

                        if (nestingSolution != null)
                        {
                            m_DataTable.Merge(AddToDataTable(subPartTable, nestingSolution, strReportFormat));
                        }
                    }

                    // clean up profile tables
                    partTable.Clear();
                    partTable.Dispose();
                    partTable = null;
                }

                // clean up stock tables
                partStockTable.Clear();
                partStockTable.Dispose();
                partStockTable = null;

                return m_DataTable;
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log(0, typeof(ProfileStockQI).FullName, e.ToString(), string.Empty, (new StackTrace()).GetFrame(1).GetMethod().Name, string.Empty, string.Empty, -1);
                throw;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="queryTemplate"></param>
        /// <returns></returns>
        private string GetFormattedPartsQuery(string queryTemplate)
        {
            string returnQuery;

            // If the select set is empty, return profile query for the model.
            // Otherwise, build the query string for profiles using the select set
            // BO OIDs.
            if (InputObjects.Count == 0)
            {
                returnQuery = queryTemplate;
                returnQuery = returnQuery.Replace("@Assemblies", String.Empty);
            }
            else
            {
                // using select set for the queries
                returnQuery = queryTemplate;

                string selectSetAssemblies = queryAssemblyStart;

                int selectSetCount = InputObjects.Count;

                for (int index = 0; index < selectSetCount; index++)
                {
                    if (index != 0)
                    {
                        selectSetAssemblies = selectSetAssemblies +
                                              Environment.NewLine +
                                              "UNION" + Environment.NewLine;
                    }

                    if (IsOracleDB)
                    {
                        string OraOid = InputObjects[index].ObjectID.ToString();
                        OraOid = OraOid.Replace("{", "");
                        OraOid = OraOid.Replace("-", "");
                        OraOid = OraOid.Replace("}", "");
                        OraOid = OraOid.ToUpper();
                        selectSetAssemblies = selectSetAssemblies +
                                          OraqueryAssemblyMiddle.Replace("?", "'" + OraOid + "'");
                    }
                    else
                    {
                        selectSetAssemblies = selectSetAssemblies +
                                          queryAssemblyMiddle.Replace("?", "'" + InputObjects[index].ObjectID.ToString() + "'");
                    }
                    
                }

                selectSetAssemblies = selectSetAssemblies + queryAssemblyEnd;

                returnQuery = returnQuery.Replace("@Assemblies", selectSetAssemblies);
            }

            return returnQuery;
        }

        #region Report Formats

        private List<Column> CreateReturnDataTable()
        {
            // create datatable columns
            List<Column> FieldColumnList = new List<Column>();

            Column recFieldColumnNesdtingOid = new Column(FIELD_NESTING_OID, typeof(System.String));
            FieldColumnList.Add(recFieldColumnNesdtingOid);

            Column recFieldColumnOIDS = new Column(FIELD_OIDS, typeof(System.String));
            FieldColumnList.Add(recFieldColumnOIDS);

            Column recFieldColumnReferenceStandard = new Column(FIELD_REFERENCE_STANDARD, typeof(System.String));
            FieldColumnList.Add(recFieldColumnReferenceStandard);

            Column recFieldColumnSectionName = new Column(FIELD_SECTION_NAME, typeof(System.String));
            FieldColumnList.Add(recFieldColumnSectionName);

            Column recFieldColumnMaterialType = new Column(FIELD_MATERIAL_TYPE, typeof(System.String));
            FieldColumnList.Add(recFieldColumnMaterialType);

            Column recFieldColumnMaterialGrade = new Column(FIELD_MATERIAL_GRADE, typeof(System.String));
            FieldColumnList.Add(recFieldColumnMaterialGrade);

            Column recFieldColumnStockLength = new Column(FIELD_LENGTH, typeof(System.Double));
            FieldColumnList.Add(recFieldColumnStockLength);

            Column recFieldColumnQuantity = new Column(FIELD_QUANTITY, typeof(System.Int32));
            FieldColumnList.Add(recFieldColumnQuantity);

            Column recFieldColumnLengthUsed = new Column(FIELD_LENGTH_USED, typeof(System.Double));
            FieldColumnList.Add(recFieldColumnLengthUsed);

            Column recFieldColumnLengthUnused = new Column(FIELD_LENGTH_UNUSED, typeof(System.Double));
            FieldColumnList.Add(recFieldColumnLengthUnused);

            Column recFieldColumnWeightUsed = new Column(FIELD_WEIGHT_USED, typeof(System.Double));
            FieldColumnList.Add(recFieldColumnWeightUsed);

            Column recFieldColumnWeightUnused = new Column(FIELD_WEIGHT_UNUSED, typeof(System.Double));
            FieldColumnList.Add(recFieldColumnWeightUnused);

            Column recFieldColumnPercentageUsed = new Column(FIELD_PERCENTAGE_USED, typeof(System.Double));
            FieldColumnList.Add(recFieldColumnPercentageUsed);

            Column recFieldColumnPercentageUnused = new Column(FIELD_PERCENTAGE_UNUSED, typeof(System.Double));
            FieldColumnList.Add(recFieldColumnPercentageUnused);

            Column recFieldColumnPartNames = new Column(FIELD_PART_NAMES, typeof(System.String));
            FieldColumnList.Add(recFieldColumnPartNames);

            Column recFieldColumnPartLengths = new Column(FIELD_PART_LENGTH, typeof(System.String));
            FieldColumnList.Add(recFieldColumnPartLengths);

            DataTable returnDataTable = new DataTable();

            return FieldColumnList;
        }

        /// <summary>
        /// Adds to data table.
        /// </summary>
        /// <param name="profilesDataTable">The profiles data table.</param>
        /// <param name="nestingResult">The nesting result.</param>
        /// <param name="strReportFormat">The string report format.</param>
        /// <returns>DataTable.</returns>
        private DataTable AddToDataTable(DataTable profilesDataTable, Array nestingResult, string strReportFormat)
        {
            try
            {
                DataTable returnDataTable = m_DataTable.Clone();

                for (int stockIndex = 0; stockIndex < nestingResult.GetLength(0); stockIndex++)
                {
                    Array stock = (Array)nestingResult.GetValue(stockIndex);
                    Double stockLength = (Double)stock.GetValue(0);
                    Double remainingLength = (Double)stock.GetValue(1);

                    Array profileOIDs = (Array)stock.GetValue(2);
                    Array profileNames = (Array)stock.GetValue(3);
                    Array profileLengths = (Array)stock.GetValue(4);

                    DataRow row = returnDataTable.NewRow();

                    row.SetField(FIELD_NESTING_OID, "Not Implemented");

                    row.SetField(FIELD_REFERENCE_STANDARD, profilesDataTable.Rows[0][FIELD_REFERENCE_STANDARD]);
                    row.SetField(FIELD_SECTION_NAME, profilesDataTable.Rows[0][FIELD_SECTION_NAME]);
                    row.SetField(FIELD_MATERIAL_TYPE, profilesDataTable.Rows[0][FIELD_MATERIAL_TYPE]);
                    row.SetField(FIELD_MATERIAL_GRADE, profilesDataTable.Rows[0][FIELD_MATERIAL_GRADE]);

                    row.SetField(FIELD_LENGTH, stockLength);
                    row.SetField(FIELD_QUANTITY, 1);
                    row.SetField(FIELD_LENGTH_USED, stockLength - remainingLength);
                    row.SetField(FIELD_LENGTH_UNUSED, remainingLength);

                    Double percentageUsed = (stockLength - remainingLength) / stockLength;

                    row.SetField(FIELD_WEIGHT_USED, percentageUsed * 100);
                    row.SetField(FIELD_WEIGHT_UNUSED, (1 - percentageUsed) * 100);
                    row.SetField(FIELD_PERCENTAGE_USED, percentageUsed * 100);
                    row.SetField(FIELD_PERCENTAGE_UNUSED, (1 - percentageUsed) * 100);

                    int iTableIndexOffset = 0;

                    if (strReportFormat == REPORT_FORMAT_ONSET)
                    {
                        row.SetField(FIELD_PART_LENGTH, profileLengths.GetValue(0).ToString());
                        row.SetField(FIELD_PART_NAMES, profileNames.GetValue(0).ToString());
                        row.SetField(FIELD_OIDS, profileOIDs.GetValue(0).ToString());

                        iTableIndexOffset = 1;
                    }

                    returnDataTable.Rows.Add(row);

                    for (int profileArrayIndex = 0 + iTableIndexOffset; profileArrayIndex < profileLengths.GetLength(0); profileArrayIndex++)
                    {
                        DataRow partNameRow = returnDataTable.NewRow();

                        partNameRow.SetField(FIELD_PART_LENGTH, profileLengths.GetValue(profileArrayIndex).ToString());

                        partNameRow.SetField(FIELD_PART_NAMES, profileNames.GetValue(profileArrayIndex).ToString());

                        partNameRow.SetField(FIELD_OIDS, profileOIDs.GetValue(profileArrayIndex).ToString());

                        returnDataTable.Rows.Add(partNameRow);
                    }
                }

                return returnDataTable;
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log(0, typeof(ProfileStockQI).FullName, e.ToString(), string.Empty, (new StackTrace()).GetFrame(1).GetMethod().Name, string.Empty, string.Empty, -1);
                throw;
            }
        }

        private void ParseArguments(string arguments, ref string reportType, ref string reportFormat, ref string nesting, ref double[] endOffsets, ref double[] middleOffsets, ref double[] stocks)
        {
            try
            {
                string arg = arguments.Replace(" ", String.Empty).Replace("\t", String.Empty).Replace("\r", String.Empty).Replace("\n", String.Empty).ToUpper();

                int argLength = arg.Length;
                char[] delimiters = { ',', '(', ')' };

                // I am expecting that the arguments input contains 6 comma separated
                // substrings each of which corresponds to some information that is needed
                // to run the report, so I attempt to get the substring and store them in 
                // variables that reflect their purpose for later processing.

                // get entered report format flag
                int subStringStartIndex = 0;
                int delimeterLocation = arg.IndexOf(delimiters[0], 0);

                if (delimeterLocation < 0 || delimeterLocation <= subStringStartIndex)
                {
                    MessageBox.Show("First argument error");
                    throw new Exception(".rqe arguments is not valid");
                }

                string enteredType = arg.Substring(subStringStartIndex, delimeterLocation - subStringStartIndex);

                // get entered report type flag
                subStringStartIndex = delimeterLocation + 1;

                delimeterLocation = arg.IndexOf(delimiters[0], subStringStartIndex);

                if (delimeterLocation < 0 || delimeterLocation >= argLength || delimeterLocation <= subStringStartIndex)
                {
                    MessageBox.Show("Second argument error");
                    throw new Exception(".rqe arguments is not valid");
                }

                string enteredFormat = arg.Substring(subStringStartIndex, delimeterLocation - subStringStartIndex);

                // get entered nesting flag
                /*subStringStartIndex = delimeterLocation + 1;

                delimeterLocation = arg.IndexOf(delimiters[0], subStringStartIndex);

                if (delimeterLocation < 0 || delimeterLocation >= argLength || delimeterLocation <= subStringStartIndex)
                {
                    MessageBox.Show("Third argument error");
                    throw new Exception(".rqe arguments is not valid");
                }

                string enteredNesting = arg.Substring(subStringStartIndex, delimeterLocation - subStringStartIndex);*/

                // get entered end offsets
                subStringStartIndex = delimeterLocation + 1;

                delimeterLocation = arg.IndexOf(delimiters[2].ToString() + delimiters[0].ToString(), subStringStartIndex) + 1;

                if (delimeterLocation < 0 || delimeterLocation >= argLength || delimeterLocation <= subStringStartIndex)
                {
                    MessageBox.Show("Fourth argument error");
                    throw new Exception(".rqe arguments is not valid");
                }

                string enteredEndOffsets = arg.Substring(subStringStartIndex, delimeterLocation - subStringStartIndex);

                // get entered middle offsets
                subStringStartIndex = delimeterLocation + 1;

                delimeterLocation = arg.IndexOf(delimiters[2].ToString() + delimiters[0].ToString(), subStringStartIndex) + 1;

                if (delimeterLocation < 0 || delimeterLocation >= argLength || delimeterLocation <= subStringStartIndex)
                {
                    MessageBox.Show("Fifth argument error");
                    throw new Exception(".rqe arguments is not valid");
                }

                string enteredMiddleOffsets = arg.Substring(subStringStartIndex, delimeterLocation - subStringStartIndex);

                // get entered default stock
                subStringStartIndex = delimeterLocation + 1;

                string enteredStocks = arg.Substring(subStringStartIndex);

                // Now that we have the substrings, we need to check that the enteredFormat,
                // enteredType, and enteredNesting are valid values. We need to do a little
                // more work to get the desired array of doubles from the enteredEndOffsets, 
                // enteredMiddleoffsets, and enteredStocks strings.

                // Check enteredType
                if (enteredType != REPORT_PROFILE && enteredType != REPORT_MEMBER && enteredType != REPORT_MEMBER_PROFILE)
                {
                    MessageBox.Show("Invalid type");
                    throw new Exception(".rqe arguments is not valid");
                }

                reportType = enteredType;

                // Check enteredFormat
                // If enteredFormat does not equal one of the predefined formats, report
                // an error.
                if (enteredFormat != REPORT_FORMAT_ONSET && enteredFormat != REPORT_FORMAT_OFFSET)
                {
                    MessageBox.Show("Invalid format");
                    throw new Exception(".rqe arguments is not valid");
                }

                reportFormat = enteredFormat;

                int numberOfOffsets;
                numberOfOffsets = (enteredType == REPORT_MEMBER_PROFILE) ? 2 : 1;

                // Check enteredNesting
                /*if (enteredNesting != REPORT_ALL && enteredNesting != REPORT_EXCLUDE)
                {
                    MessageBox.Show("Invalid nesting");
                    throw new Exception(".rqe arguments is not valid");
                }

                nesting = enteredNesting;*/

                // Check enteredEndOffsets
                // I am expecting this to be of the form ENDOFFSETS(values), where
                // values is a list of comma separated doubles. I know that the string
                // ends with a ')' because of how this string was retrived from arguments.
                subStringStartIndex = 0;
                delimeterLocation = enteredEndOffsets.IndexOf(delimiters[1], subStringStartIndex);

                // Get the substring starting at the start of the string to the first '('.
                // Everything after '(', not including the final character, should
                // be the list of comma separated doubles.
                string endOffsetLabel = enteredEndOffsets.Substring(subStringStartIndex, delimeterLocation);
                string endOffsetsString = enteredEndOffsets.Substring(delimeterLocation + 1, enteredEndOffsets.Length - delimeterLocation - 2);

                if (endOffsetLabel != REPORT_END_OFFSETS)
                {
                    MessageBox.Show("End offset not found");
                    throw new Exception(".rqe arguments is not valid");
                }

                string[] endOffsetsAsStrings = endOffsetsString.Split(delimiters[0]);

                UOMManager oUOMManager = MiddleServiceProvider.UOMMgr;
                endOffsets = new double[endOffsetsAsStrings.Length];

                for (int index = 0; index < endOffsetsAsStrings.Length; index++)
                {
                    double tmpEndOffset = oUOMManager.ParseUnit(UnitType.Distance, endOffsetsAsStrings[index]);

                    if (tmpEndOffset >= 0)
                    {
                        endOffsets[index] = tmpEndOffset;
                    }
                    else
                    {
                        MessageBox.Show("Negative end offset");
                        throw new Exception("Negative end offset");
                    }
                }

                // This code is the same as for the endoffset code.
                delimeterLocation = enteredMiddleOffsets.IndexOf(delimiters[1], subStringStartIndex);

                string middleOffsetLabel = enteredMiddleOffsets.Substring(subStringStartIndex, delimeterLocation);
                string middleOffsetsString = enteredMiddleOffsets.Substring(delimeterLocation + 1, enteredMiddleOffsets.Length - delimeterLocation - 2);

                if (middleOffsetLabel != REPORT_MIDDLE_OFFSETS)
                {
                    MessageBox.Show("Middle offset not found");
                    throw new Exception(".rqe arguments is not valid");
                }

                string[] middleOffsetsAsStrings = middleOffsetsString.Split(delimiters[0]);

                middleOffsets = new double[middleOffsetsAsStrings.Length];

                for (int index = 0; index < middleOffsetsAsStrings.Length; index++)
                {
                    double tmpMiddleOffset = oUOMManager.ParseUnit(UnitType.Distance, middleOffsetsAsStrings[index]);

                    if (tmpMiddleOffset >= 0)
                    {
                        middleOffsets[index] = tmpMiddleOffset;
                    }
                    else
                    {
                        MessageBox.Show("Negative middle offset");
                        throw new Exception("Negative middle offset");
                    }
                }

                if (middleOffsetsAsStrings.Length != endOffsetsAsStrings.Length || endOffsetsAsStrings.Length != numberOfOffsets)
                {
                    MessageBox.Show("Number of end and middle offsets do not match or not of expected number");
                    throw new Exception(".rqe arguments is not valid");
                }
                
                // This code is the same as the middleoffset and endoffset code above.
                delimeterLocation = enteredStocks.IndexOf(delimiters[1], subStringStartIndex);

                string stocksLabel = enteredStocks.Substring(subStringStartIndex, delimeterLocation);
                string stocksString = enteredStocks.Substring(delimeterLocation + 1, enteredStocks.Length - delimeterLocation - 2);

                if (stocksLabel != REPORT_STOCKS)
                {
                    MessageBox.Show("Stocks expected");
                    throw new Exception(".rqe arguments is not valid");
                }

                string[] stocksAsStrings;

                if (stocksString.Length > 0)
                {
                    stocksAsStrings = stocksString.Split(delimiters[0]);
                }
                else
                {
                    stocksAsStrings = new string[0];
                }

                stocks = new double[stocksAsStrings.Length];

                for (int index = 0; index < stocksAsStrings.Length; index++)
                {
                    double tmpStock = oUOMManager.ParseUnit(UnitType.Distance, stocksAsStrings[index]);

                    if (tmpStock > 0)
                    {
                        stocks[index] = tmpStock;
                    }
                    else
                    {
                        MessageBox.Show("Negative stock length");
                        throw new Exception("Negative stock length");
                    }
                }

                stocks = stocks.Distinct().ToArray();
                Array.Sort(stocks);
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log(0, typeof(ProfileStockQI).FullName, e.ToString(), string.Empty, (new StackTrace()).GetFrame(1).GetMethod().Name, string.Empty, string.Empty, -1);
                throw;
            }
        }

        //private DataTable StockNesting(string strPartQuery, DataTable dtPartStock, string strFormat, Double dPartEndOffset, Double dPartMiddleOffset)
        //{
        //    SetQuery(strPartQuery);

        //    //DataTable profileTable = ExecuteDelegatedQuery();
        //    object rsTest = ExecuteDelegatedQuery();
        //    DataTable profileTable = (DataTable)rsTest;

        //    // ToDo: Use ProgID to create the object containing the nesting algorithm.
        //    StrMfgProfileStockNesting profileStockNesting = new StrMfgProfileStockNesting();

        //    // The profiles are sent to the nesting algorithm in groups based
        //    // on the rank assigned to ea ch profile in the query. The relevant
        //    // profile stocks are sent to the algorithm based on the profiles
        //    // in the current group.
        //    long maxProfileRank = 0;

        //    // If the query returned any profiles, get the largest rank value.
        //    if (profileTable.Rows.Count > 0)
        //    {
        //        // Since the query results are sorted by the rank, the rank
        //        // from the last entry of the table is the highest rank.
        //        maxProfileRank = (long)profileTable.Rows[profileTable.Rows.Count - 1][FIELD_RANK];
        //    }

        //    DataTable dtReturnTable = m_DataTable.Clone();

        //    // For each rank, get the profiles and the relevant profile stocks
        //    // and pass them to the nesting algorithm.
        //    for (long profileRank = 1; profileRank <= maxProfileRank; profileRank++)
        //    {
        //        // Create a new table with the same structure as the profile table.
        //        // We want to get sub tables of profiles with the same
        //        // refereence standard, section name, material grade, and material type.
        //        DataTable subProfileTable = profileTable.Clone();

        //        // get all rows that have the current rank
        //        var profileSubQuery =
        //            from row in profileTable.AsEnumerable()
        //            where row.Field<long>(FIELD_RANK) == profileRank
        //            select row;

        //        // For each row that contains the current rank, copy it to
        //        // the sub table.
        //        foreach (DataRow row in profileSubQuery.ToList<DataRow>())
        //        {
        //            subProfileTable.ImportRow(row);
        //        }

        //        if (subProfileTable.Rows[0][FIELD_SECTION_TYPE].ToString() == "FB")
        //        {
        //            continue;
        //        }

        //        object[] arrProfileOIDs = new object[subProfileTable.Rows.Count];
        //        object[] arrProfileNames = new object[subProfileTable.Rows.Count];
        //        object[] arrProfileLengths = new object[subProfileTable.Rows.Count];

        //        for (int index = 0; index < subProfileTable.Rows.Count; index++)
        //        {
        //            arrProfileOIDs[index] = subProfileTable.Rows[index][FIELD_OID].ToString();
        //            arrProfileNames[index] = subProfileTable.Rows[index]["ProfileName"].ToString();
        //            arrProfileLengths[index] = subProfileTable.Rows[index]["ProfileLength"];
        //        }

        //        // get all profile stocks of the specified type
        //        var profileStockSubQuery =
        //            from row in dtPartStock.AsEnumerable()
        //            where row.Field<string>(FIELD_REFERENCE_STANDARD) == subProfileTable.Rows[0][FIELD_REFERENCE_STANDARD].ToString()
        //                && row.Field<string>(FIELD_SECTION_NAME) == subProfileTable.Rows[0][FIELD_SECTION_NAME].ToString()
        //                && row.Field<string>(FIELD_MATERIAL_TYPE) == subProfileTable.Rows[0][FIELD_MATERIAL_TYPE].ToString()
        //                && row.Field<string>(FIELD_MATERIAL_GRADE) == subProfileTable.Rows[0][FIELD_MATERIAL_GRADE].ToString()
        //            select row;

        //        // Create a new table with the same structure as the profile table.
        //        DataTable subProfileStockTable = dtPartStock.Clone();

        //        // For each row that contains the specified type, copy it to
        //        // the sub table.
        //        foreach (DataRow dataRow in profileStockSubQuery.ToList<DataRow>())
        //        {
        //            subProfileStockTable.ImportRow(dataRow);
        //        }

        //        // ToDo: pass only stock length
        //        object[] arrStockLengths = new object[subProfileStockTable.Rows.Count];

        //        for (int index = 0; index < subProfileStockTable.Rows.Count; index++)
        //        {
        //            arrStockLengths[index] = subProfileStockTable.Rows[index][FIELD_STOCK_LENGTH];
        //        }

        //        Array nestingSolution = null;

        //        object[] defaultStocks = new object[] { 99.0 };

        //        nestingSolution = (Array)profileStockNesting.GetNestingSolution(0,
        //                                                    (Array)arrStockLengths,
        //                                                    (Array)arrProfileOIDs,
        //                                                    (Array)arrProfileNames,
        //                                                    (Array)arrProfileLengths,
        //                                                    dPartEndOffset,
        //                                                    dPartMiddleOffset,
        //                                                    (Array)defaultStocks);

        //        //dtReturnTable.Merge(AddToDataTable(strFormat, subProfileTable, nestingSolution));
        //        dtReturnTable.Merge(Format3(strFormat, subProfileTable, nestingSolution));
        //    }

        //    profileTable.Clear();
        //    profileTable.Dispose();
        //    profileTable = null;

        //    return dtReturnTable;
        //}

        #endregion

    }
}