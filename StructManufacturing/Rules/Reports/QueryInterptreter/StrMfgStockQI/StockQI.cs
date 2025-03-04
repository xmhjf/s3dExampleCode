//-----------------------------------------------------------------------------
//      Copyright (C) 2012, Intergraph Corporation. All rights reserved.
//
//      Profile Stock Nesting Report Query Interpreter implementation.
//
//      Author: Nautilus - HSV  
//
//      History:
//      1-15-2013    Brandon Affenzeller    Created
//  	
//		Modified  on  25 OCT  2013  Praveen Babu  Oracle support
//-----------------------------------------------------------------------------

using System;
using System.Data;
using System.Collections.Generic;
using System.Diagnostics;

using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Reports.Middle;
using Ingr.SP3D.Reports.Exceptions;

namespace SM3DStrMfgStockQI
{
    /// <summary>
    /// Class StockQI.
    /// </summary>
    public class StockQI : QueryInterpreter
    {
        private DataTable m_DataTable;

        #region Constants

        private const String FIELD_BULK_LOAD_ACTION = "BulkLoadAction";
        private const String FIELD_PART_NUMBER = "PartNumber";
        private const String FIELD_PART_DESCRIPTION = "PartDescription";
        private const String FIELD_MATERIAL_TYPE = "MaterialType";
        private const String FIELD_MATERIAL_GRADE = "MaterialGrade";
        private const String FIELD_REFERENCE_STANDARD = "ReferenceStandard";
        private const String FIELD_SECTION_NAME = "SectionName";
        private const String FIELD_LENGTH = "Length";

        private const String STOCK_START = "Start";
        private const String STOCK_END = "End";

        private const String REPORT_PROFILE = "PROFILE";
        private const String REPORT_MEMBER = "MEMBER";
        private const String REPORT_PROFILE_MEMBER = "PROFILE_MEMBER";

        #endregion

        #region Hard Coded Profile Query

        private const String profileStockQuery =
                @"SET NOCOUNT ON
                SET QUOTED_IDENTIFIER ON
                SET ANSI_NULLS ON
                 
                if exists (select * from tempdb..sysobjects where ( id = OBJECT_ID(N'[tempdb]..[#TempProfileOidTbl]') ))
	                DROP TABLE #TempProfileOidTbl

                -- Create temporary table to contain the profile oids
                CREATE TABLE #TempProfileOidTbl(oid Uniqueidentifier)

                CREATE INDEX i_LengthOid ON #TempProfileOidTbl (oid)

                INSERT INTO #TempProfileOidTbl
                SELECT 
	                JPP.Oid
                FROM JProfilePart JPP

                if exists (select * from tempdb..sysobjects where ( id = OBJECT_ID(N'[tempdb]..[#TempProfileSectionTbl]') ))
                    DROP TABLE #TempProfileSectionTbl

                -- create temporaty table to contain the reference standard and section name
                CREATE TABLE #TempProfileSectionTbl(oid Uniqueidentifier, ReferenceStandard NVARCHAR(256), SectionName NVARCHAR(256))

                CREATE INDEX i_SectionOid ON #TempProfileSectionTbl (oid)

                INSERT INTO #TempProfileSectionTbl
                SELECT 
	                JPP.Oid,
	                JRS.Name As ReferenceStandard,
	                COALESCE(CADCS1.csSectName, 
			                 CADCS2.csSectName, 
			                 CADCS3.csSectName, 
			                 CADCS4.csSectName)
			                 As SectionName
                FROM #TempProfileOidTbl JPP
                -- Profile Section Name
                INNER JOIN XShpStrDesignHierarchy XSSDH1 ON XSSDH1.OidDestination = JPP.Oid
                INNER JOIN XShpStrDesignHierarchy XSSDH2 ON XSSDH2.OidDestination = XSSDH1.OidOrigin
                INNER JOIN XShpStrDesignHierarchy XSSDH3 ON XSSDH3.OidDestination = XSSDH2.OidOrigin
                LEFT JOIN GSCADGeom GEOM1 ON GEOM1.goidDest = JPP.Oid
                LEFT JOIN GSCADGeom GEOM2 ON GEOM2.goidDest = XSSDH1.oidOrigin
                LEFT JOIN GSCADGeom GEOM3 ON GEOM3.goidDest = XSSDH2.oidOrigin
                LEFT JOIN GSCADGeom GEOM4 ON GEOM4.goidDest = XSSDH3.oidOrigin
                LEFT JOIN GSCADCrossSection CADCS1 ON CADCS1.oidOrigin = GEOM1.gOidOrigin
                LEFT JOIN GSCADCrossSection CADCS2 ON CADCS2.oidOrigin = GEOM2.gOidOrigin
                LEFT JOIN GSCADCrossSection CADCS3 ON CADCS3.oidOrigin = GEOM3.gOidOrigin
                LEFT JOIN GSCADCrossSection CADCS4 ON CADCS4.oidOrigin = GEOM4.gOidOrigin

                -- Reference Standard
                INNER JOIN JDCrossSection JDCS ON JDCS.Oid = COALESCE(CADCS1.csOid, 
													                  CADCS2.csOid, 
													                  CADCS3.csOid, 
													                  CADCS4.csOid)                                      
                INNER JOIN JDPartClass JDPC ON JDPC.Name = JDCS.Type
                INNER JOIN XReferenceStdHasPartClasses XRSHPC 
	                ON XRSHPC.OidDestination = JDPC.Oid
                INNER JOIN JReferenceStandard JRS ON JRS.Oid = XRSHPC.OidOrigin

                if exists (select * from tempdb..sysobjects where ( id = OBJECT_ID(N'[tempdb]..[#TempProfileMaterialTbl]') ))
                    DROP TABLE #TempProfileMaterialTbl

                -- create temporaty table to contain the profile material type and material grade
                CREATE TABLE #TempProfileMaterialTbl(oid Uniqueidentifier, MaterialType NVARCHAR(256), MaterialGrade NVARCHAR(256))

                CREATE INDEX i_MaterialOid ON #TempProfileMaterialTbl (oid)

                INSERT INTO #TempProfileMaterialTbl
                SELECT
	                JPP.Oid,
	                JDM.MaterialType,
	                JDM.MaterialGrade
                FROM #TempProfileOidTbl JPP
                -- Profile Material Type and Grade
                INNER JOIN XSystemHasMaterial XSHM 
	                ON XSHM.OidDestination = dbo.REPORTGetParentRelationOid(JPP.Oid, 'SystemHasMaterial')
                INNER JOIN JDMaterial JDM ON JDM.Oid = XSHM.OidOrigin

                -- use the temporary tables to get the desired query results
                SELECT DISTINCT
	                #TempProfileSectionTbl.ReferenceStandard,
	                #TempProfileSectionTbl.SectionName,
	                #TempProfileMaterialTbl.MaterialGrade,
	                #TempProfileMaterialTbl.MaterialType
                FROM #TempProfileOidTbl, #TempProfileMaterialTbl, #TempProfileSectionTbl
                WHERE #TempProfileOidTbl.oid = #TempProfileSectionTbl.oid AND
	                  #TempProfileMaterialTbl.oid = #TempProfileOidTbl.oid

                -- drop temporary tables
                DROP TABLE #TempProfileOidTbl
                DROP TABLE #TempProfileSectionTbl
                DROP TABLE #TempProfileMaterialTbl";
        private const String OraprofileStockQuery =
                 @"With 
                TempProfileOidTbl(oid) as
                (
	                SELECT JPP.Oid FROM JProfilePart JPP
                ),

                TempProfileSectionTbl(oid , ReferenceStandard , SectionName)
                as
                (
	                SELECT 
	                JPP.Oid,
	                JRS.Name As ReferenceStandard,
	                COALESCE(CADCS1.csSectName, 
	                CADCS2.csSectName, 
	                CADCS3.csSectName, 
	                CADCS4.csSectName)
	                As SectionName
	                FROM TempProfileOidTbl JPP
	                INNER JOIN XShpStrDesignHierarchy XSSDH1 ON XSSDH1.OidDestination = JPP.Oid
	                INNER JOIN XShpStrDesignHierarchy XSSDH2 ON XSSDH2.OidDestination = XSSDH1.OidOrigin
	                INNER JOIN XShpStrDesignHierarchy XSSDH3 ON XSSDH3.OidDestination = XSSDH2.OidOrigin
	                LEFT JOIN GSCADGeom GEOM1 ON GEOM1.goidDest = JPP.Oid
	                LEFT JOIN GSCADGeom GEOM2 ON GEOM2.goidDest = XSSDH1.oidOrigin
	                LEFT JOIN GSCADGeom GEOM3 ON GEOM3.goidDest = XSSDH2.oidOrigin
	                LEFT JOIN GSCADGeom GEOM4 ON GEOM4.goidDest = XSSDH3.oidOrigin
	                LEFT JOIN GSCADCrossSection CADCS1 ON CADCS1.oidOrigin = GEOM1.gOidOrigin
	                LEFT JOIN GSCADCrossSection CADCS2 ON CADCS2.oidOrigin = GEOM2.gOidOrigin
	                LEFT JOIN GSCADCrossSection CADCS3 ON CADCS3.oidOrigin = GEOM3.gOidOrigin
	                LEFT JOIN GSCADCrossSection CADCS4 ON CADCS4.oidOrigin = GEOM4.gOidOrigin

	                INNER JOIN JDCrossSection JDCS ON JDCS.Oid = COALESCE(CADCS1.csOid, 
	                CADCS2.csOid, 
	                CADCS3.csOid, 
	                CADCS4.csOid)                                      
	                INNER JOIN JDPartClass JDPC ON JDPC.Name = JDCS.Type
	                INNER JOIN XReferenceStdHasPartClasses XRSHPC 
	                ON XRSHPC.OidDestination = JDPC.Oid
	                INNER JOIN JReferenceStandard JRS ON JRS.Oid = XRSHPC.OidOrigin
                ),

                TempProfileMaterialTbl(oid , MaterialType , MaterialGrade)
                as
                (
	                SELECT
	                JPP.Oid,
	                JDM.MaterialType,
	                JDM.MaterialGrade
	                FROM TempProfileOidTbl JPP
	                INNER JOIN XSystemHasMaterial XSHM 
	                ON XSHM.OidDestination = REPORTGetParentRelationOid(JPP.Oid, 'SystemHasMaterial')
	                INNER JOIN JDMaterial JDM ON JDM.Oid = XSHM.OidOrigin
                )

                SELECT DISTINCT
                ProfileSection.ReferenceStandard ,
                ProfileSection.SectionName ,
                ProfileMaterial.MaterialGrade ,
                ProfileMaterial.MaterialType 
                FROM TempProfileOidTbl ProfileOid, TempProfileMaterialTbl ProfileMaterial, TempProfileSectionTbl ProfileSection
                WHERE ProfileOid.oid = ProfileSection.oid AND
                ProfileMaterial.oid = ProfileOid.oid";
        #endregion

        #region Hard Coded Member Query

        private const String memberStockQuery =
                @"SET NOCOUNT ON
                SET QUOTED_IDENTIFIER ON
                SET ANSI_NULLS ON

                if exists (select * from tempdb..sysobjects where ( id = OBJECT_ID(N'[tempdb]..[#TempMemberTbl]') ))
                    DROP TABLE #TempMemberTbl

                CREATE TABLE #TempMemberTbl(oid Uniqueidentifier)

                CREATE INDEX i_LengthOid ON #TempMemberTbl (oid)

                INSERT INTO #TempMemberTbl
                SELECT 
	                MPP.Oid
                FROM SPSMemberPartPrismatic MPP

                if exists (select * from tempdb..sysobjects where ( id = OBJECT_ID(N'[tempdb]..[#TempSectionTbl]') ))
                    DROP TABLE #TempSectionTbl

                -- Create temporary table to contain the member oids
                CREATE TABLE #TempSectionTbl(oid Uniqueidentifier, ReferenceStandard NVARCHAR(256), SectionName NVARCHAR(256))

                CREATE INDEX i_SectionOid ON #TempSectionTbl (oid)

                INSERT INTO #TempSectionTbl
                SELECT
	                MPP.Oid,
	                JRS.Name As ReferenceStandard,
	                JSCS.SectionName
                FROM SPSMemberPartPrismatic MPP
                INNER JOIN YSPSMemberPartToCrossSectionEd YMCS ON YMCS.OidDestination = MPP.Oid
                INNER JOIN JDCrossSection JDCS ON JDCS.Oid = YMCS.OidOrigin
                -- Member Section Name
                INNER JOIN JSTRUCTCrossSection JSCS ON JSCS.Oid = YMCS.OidOrigin
                -- Member Reference Standard
                INNER JOIN YSPSMemberPartToCSectionStdEd YMCSS ON YMCSS.OidDestination = MPP.Oid
                INNER JOIN JReferenceStandard JRS ON JRS.Oid = YMCSS.OidOrigin

                if exists (select * from tempdb..sysobjects where ( id = OBJECT_ID(N'[tempdb]..[#TempMaterialTbl]') ))
                    DROP TABLE #TempMaterialTbl

                -- Create temporary table to contain the member material type and material grade
                CREATE TABLE #TempMaterialTbl(oid Uniqueidentifier, MaterialType NVARCHAR(256), MaterialGrade NVARCHAR(256))

                CREATE INDEX i_MaterialOid ON #TempMaterialTbl (oid)

                INSERT INTO #TempMaterialTbl
                SELECT 
	                MPP.Oid,
	                JDM.MaterialType,
	                JDM.MaterialGrade
                FROM SPSMemberPartPrismatic MPP
                INNER JOIN YSPSMemberPartToMaterialEd MPM ON MPP.Oid = MPM.OidDestination
                INNER JOIN JDMaterial JDM ON JDM.Oid = MPM.OidOrigin

                -- use the temporary tables to get the desired query results
                SELECT DISTINCT
	                #TempSectionTbl.ReferenceStandard,
	                #TempSectionTbl.SectionName,
	                #TempMaterialTbl.MaterialGrade,
	                #TempMaterialTbl.MaterialType
                FROM #TempMemberTbl, #TempSectionTbl, #TempMaterialTbl
                WHERE #TempMemberTbl.oid = #TempSectionTbl.oid AND
	                  #TempMemberTbl.oid = #TempMaterialTbl.oid  
                ORDER BY ReferenceStandard, SectionName

                -- drop temporary tables
                DROP TABLE #TempMemberTbl
                DROP TABLE #TempMaterialTbl
                DROP TABLE #TempSectionTbl";

        private const String OramemberStockQuery =
                @"With 
                TempMemberTbl(oid )
                as
                (
	                SELECT 
	                MPP.Oid
	                FROM SPSMemberPartPrismatic MPP
                ),

                TempSectionTbl(oid , ReferenceStandard , SectionName )
                as
                (
	                SELECT
	                MPP.Oid,
	                JRS.Name  ReferenceStandard,
	                JSCS.SectionName
	                FROM SPSMemberPartPrismatic MPP
	                INNER JOIN YSPSMemberPartToCrossSectionEd YMCS ON YMCS.OidDestination = MPP.Oid
	                INNER JOIN JDCrossSection JDCS ON JDCS.Oid = YMCS.OidOrigin
	                INNER JOIN JSTRUCTCrossSection JSCS ON JSCS.Oid = YMCS.OidOrigin
	                INNER JOIN YSPSMemberPartToCSectionStdEd YMCSS ON YMCSS.OidDestination = MPP.Oid
	                INNER JOIN JReferenceStandard JRS ON JRS.Oid = YMCSS.OidOrigin
                ),

                TempMaterialTbl(oid , MaterialType , MaterialGrade )
                as
                (
	                SELECT 
	                MPP.Oid,
	                JDM.MaterialType,
	                JDM.MaterialGrade
	                FROM SPSMemberPartPrismatic MPP
	                INNER JOIN YSPSMemberPartToMaterialEd MPM ON MPP.Oid = MPM.OidDestination
	                INNER JOIN JDMaterial JDM ON JDM.Oid = MPM.OidOrigin
                )
                SELECT DISTINCT
                Section.ReferenceStandard ,
                Section.SectionName ,
                Material.MaterialGrade ,
                Material.MaterialType 
                FROM TempMemberTbl Member, TempSectionTbl Section, TempMaterialTbl Material
                WHERE Member.oid = Section.oid AND
                Member.oid = Material.oid  
                ORDER BY ReferenceStandard, SectionName";

        #endregion

        bool IsOracleDb = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.DBProvider.IsOracleProvider();

        /// <summary>
        /// Initializes a new instance of the <see cref="StockQI" /> class.
        /// </summary>
        public StockQI()
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

                // parse the action input to get the report type (profile, memeber, ot both),
                // and check if a default bulkload value (A, M, or D) has been provided
                String formattedAction = action.Replace(" ", String.Empty);
                formattedAction = formattedAction.ToUpper();

                Char[] delimiter = { ',' };

                String[] actionArr = formattedAction.Split(delimiter);

                String strReportType = REPORT_PROFILE_MEMBER;
                String strDefaultBulkLoadOption = String.Empty;

                for (int index = 0; index < actionArr.Length; index++)
                {
                    if (index == 0)
                    {
                        if (actionArr[index] == REPORT_PROFILE ||
                            actionArr[index] == REPORT_MEMBER)
                        {
                            strReportType = actionArr[index];
                        }
                    }
                    else if (index == 1)
                    {
                        if (actionArr[index] == "A" || actionArr[index] == "M" || actionArr[index] == "D")
                        {
                            strDefaultBulkLoadOption = actionArr[index];
                        }
                    }
                }

                // parse the argument to get the lengths to be used for the stocks
                String[] strDefaultLengths = argument.Replace(" ", String.Empty).Split(delimiter);

                // The profile and member queries are handled seperately. Each query that needs
                // to be run, determined by the report type, is added to the stockQueries array.
                String[] stockQueries;

                if (strReportType == REPORT_PROFILE)
                {
                    stockQueries = new String[1];
                    if (IsOracleDb)
                    {
                        stockQueries[0] = OraprofileStockQuery;
                    }
                    else
                    {
                        stockQueries[0] = profileStockQuery;
                    }
                    
                }
                else if (strReportType == REPORT_MEMBER)
                {
                    stockQueries = new String[1];

                    if (IsOracleDb)
                    {
                        stockQueries[0] = OramemberStockQuery;
                    }
                    else
                    {
                        stockQueries[0] = memberStockQuery;
                    }

                }
                else
                {
                    stockQueries = new String[2];

                    if (IsOracleDb)
                    {
                        stockQueries[0] = OraprofileStockQuery;
                        stockQueries[1] = OramemberStockQuery;
                    }
                    else
                    {
                        stockQueries[0] = profileStockQuery;
                        stockQueries[1] = memberStockQuery;
                    }
                }

                // The report output format is identical to the ProfileStock spreadsheet
                // used in bulkloading profile and member stocks. The ProfileStock spreadsheet
                // requires a start and an end identifier. This handles the start identifier.
                DataRow startRow = m_DataTable.NewRow();

                startRow[FIELD_BULK_LOAD_ACTION] = STOCK_START;

                m_DataTable.Rows.Add(startRow);

                // Create a stock entries
                for (int queryIndex = 0; queryIndex < stockQueries.Length; queryIndex++)
                {
                    String stockQuery = stockQueries[queryIndex];

                    SetQuery(stockQuery);

                    DataTable stockTable = ExecuteDelegatedQuery();

                    for (int rowIndex = 0; rowIndex < stockTable.Rows.Count; rowIndex++)
                    {
                        String strMaterialType = stockTable.Rows[rowIndex][FIELD_MATERIAL_TYPE].ToString();
                        String strMaterialGrade = stockTable.Rows[rowIndex][FIELD_MATERIAL_GRADE].ToString();
                        String strReferenceStandard = stockTable.Rows[rowIndex][FIELD_REFERENCE_STANDARD].ToString();
                        String strSectionName = stockTable.Rows[rowIndex][FIELD_SECTION_NAME].ToString();

                        for (int defaultLengthsIndex = 0; defaultLengthsIndex < strDefaultLengths.Length; defaultLengthsIndex++)
                        {
                            DataRow dataRow = m_DataTable.NewRow();

                            String strLength = strDefaultLengths[defaultLengthsIndex];

                            dataRow[FIELD_BULK_LOAD_ACTION] = strDefaultBulkLoadOption;
                            dataRow[FIELD_PART_NUMBER] = strMaterialType + strMaterialGrade +
                                                           strReferenceStandard + strSectionName + strLength;
                            dataRow[FIELD_PART_DESCRIPTION] = String.Empty;
                            dataRow[FIELD_MATERIAL_TYPE] = strMaterialType;
                            dataRow[FIELD_MATERIAL_GRADE] = strMaterialGrade;
                            dataRow[FIELD_REFERENCE_STANDARD] = strReferenceStandard;
                            dataRow[FIELD_SECTION_NAME] = strSectionName;
                            dataRow[FIELD_LENGTH] = strLength;

                            m_DataTable.Rows.Add(dataRow);
                        }
                    }
                }

                // Adds the end row to signal to the bulkloading process that there are
                // no more stock to load.
                DataRow endRow = m_DataTable.NewRow();

                endRow[FIELD_BULK_LOAD_ACTION] = STOCK_END;

                m_DataTable.Rows.Add(endRow);

                return m_DataTable;
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log(0, typeof(StockQI).FullName, e.ToString(), string.Empty, (new StackTrace()).GetFrame(1).GetMethod().Name, string.Empty, string.Empty, -1);
                throw;
            }
        }

        private List<Column> CreateReturnDataTable()
        {
            // create datatable columns
            List<Column> FieldColumnList = new List<Column>();

            Column recFieldColumnBulkloadAction = new Column(FIELD_BULK_LOAD_ACTION, typeof(System.String));
            FieldColumnList.Add(recFieldColumnBulkloadAction);

            Column recFieldColumnPartnumber = new Column(FIELD_PART_NUMBER, typeof(System.String));
            FieldColumnList.Add(recFieldColumnPartnumber);

            Column recFieldColumnPartDescription = new Column(FIELD_PART_DESCRIPTION, typeof(System.String));
            FieldColumnList.Add(recFieldColumnPartDescription);

            Column recFieldColumnMaterialType = new Column(FIELD_MATERIAL_TYPE, typeof(System.String));
            FieldColumnList.Add(recFieldColumnMaterialType);

            Column recFieldColumnMaterialGrade = new Column(FIELD_MATERIAL_GRADE, typeof(System.String));
            FieldColumnList.Add(recFieldColumnMaterialGrade);

            Column recFieldColumnReferenceStandard = new Column(FIELD_REFERENCE_STANDARD, typeof(System.String));
            FieldColumnList.Add(recFieldColumnReferenceStandard);

            Column recFieldColumnSectionName = new Column(FIELD_SECTION_NAME, typeof(System.String));
            FieldColumnList.Add(recFieldColumnSectionName);

            Column recFieldColumnStockLength = new Column(FIELD_LENGTH, typeof(System.String));
            FieldColumnList.Add(recFieldColumnStockLength);

            //DataTable returnDataTable = new DataTable();

            return FieldColumnList;
        }

    }
}
