//--------------------------------------------------------------------------------------------------
//	Copyright (C) 2016, Intergraph Corporation.  All rights reserved.
//
//	FILE: StructuralBOMQI.cs
//
//	DESCRIPTION:
//		the query interpreter for the structural BOM report
//
//	NOTES:
//		this query interpreter only supports plates and profiles at this time
//
//		nesting data may result in a single part having multiple rows in the report - a
//		profile manufactured as a plate may have 1, 2, or 3 entries based on how many
//		flanges it has (FB = 1 row, T = 2 rows, I = 3 rows)
//
//	HISTORY:
//		Feb-09-2015		Pam Livingston
//			TR-CP-280643 Improve "Structural BOM" report creation performance on ORACLE database
//
//--------------------------------------------------------------------------------------------------

using System;
using System.Data;
using System.Collections.Generic;
using System.Diagnostics;

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Common.Middle.Services.Hidden;
using Ingr.SP3D.Reports.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Manufacturing.Middle;

namespace StructuralBOM
{
	/// <summary>
	/// QueryInterpreter for the Structural BOM Report
	/// </summary>
	public class StructuralBOMQI : QueryInterpreter
	{
		#region Private Data
		private DataTable m_DataTable;	// the DataTable that will be returned by this query interpreter
		#endregion

		#region Constants
		// these column field names must match exactly what is in the .xls template for the report
		private const String FIELD_PART_ID = "OID";
		private const String FIELD_PART_NAME = "PartName";
		private const String FIELD_SYMMETRY = "Symmetry";
		private const String FIELD_OBJECT = "Object";
		private const String FIELD_OBJECT_TYPE = "ObjectType";
		private const String FIELD_PROFILE_TYPE = "ProfileType";
		private const String FIELD_STATUS = "Status";
		private const String FIELD_STAGE = "Stage";
		private const String FIELD_MATERIAL_TYPE = "MaterialType";
		private const String FIELD_GRADE = "Grade";
		private const String FIELD_PLATE_LENGTH = "PlateLength";
		private const String FIELD_PLATE_WIDTH = "PlateWidth";
		private const String FIELD_PLATE_THICKNESS = "PlateThickness";
		private const String FIELD_PROFILE_LENGTH = "ProfileLength";
		private const String FIELD_WEB_WIDTH = "WebWidth";
		private const String FIELD_FLANGE_WIDTH = "FlangeWidth";
		private const String FIELD_WEB_THICKNESS = "WebThickness";
		private const String FIELD_FLANGE_THICKNESS = "FlangeThickness";
		private const String FIELD_IS_TWISTED = "IsTwisted";
		private const String FIELD_CURVED = "Curved";
		private const String FIELD_PART_DWG_NUMBER = "PartDwgNumber";
		private const String FIELD_ARRG_DWG_NUMBER = "ArrgDwgNumber";
		private const String FIELD_MFG_PART_ID = "Mfg Part Id";
		private const String FIELD_MFG_PART_NAME = "Mfg Part Name";
		private const String FIELD_NEST_TYPE = "NestType";
		private const String FIELD_NEST_ID = "NestID";
		private const String FIELD_NESTING_STAGE = "Nesting Stage";
		private const String FIELD_COG_X = "COGx";
		private const String FIELD_COG_Y = "COGy";
		private const String FIELD_COG_Z = "COGz";
		private const String FIELD_WEIGHT = "Weight";
		private const String FIELD_AREA = "Area";
		#endregion

		#region Private Classes
		/// <summary>
		/// using an internal structure makes it easier as the DataRow doesn't have
		/// a copy constructor for when a part needs multiple rows
		/// </summary>
		internal class BOMInfo
		{
			#region Data
			internal string FieldPartID;
			internal string FieldPartName;
			internal string FieldSymmetry;
			internal string FieldObject;
			internal string FieldObjectType;
			internal string FieldProfileType;
			internal string FieldStatus;
			internal string FieldStage;
			internal string FieldMaterialType;
			internal string FieldGrade;
			internal double FieldPlateLength;
			internal double FieldPlateWidth;
			internal double FieldPlateThickness;
			internal double FieldProfileLength;
			internal double FieldWebWidth;
			internal double FieldFlangeWidth;
			internal double FieldWebThickness;
			internal double FieldFlangeThickness;
			internal string FieldIsTwisted;
			internal string FieldCurved;
			internal string FieldPartDwgNumber;		// SQL for report does not retrieve data for this field ...
			internal string FieldArrgDwgNumber;		// SQL for report does not retrieve data for this field ...
			internal string FieldMfgPartId;
			internal string FieldMfgPartName;
			internal string FieldNestType;
			internal string FieldNestID;
			internal string FieldNestingStage;		// SQL for report does not retrieve data for this field ...
			internal double FieldCOGx;
			internal double FieldCOGy;
			internal double FieldCOGz;
			internal double FieldWeight;
			internal double FieldArea;
			#endregion

			/// <summary>
			/// default constructor
			/// </summary>
			public BOMInfo()
			{
				FieldPartID = string.Empty;
				FieldPartName = string.Empty;
				FieldSymmetry = string.Empty;
				FieldObject = string.Empty;
				FieldObjectType = string.Empty;
				FieldProfileType = string.Empty;
				FieldStatus = string.Empty;
				FieldStage = string.Empty;
				FieldMaterialType = string.Empty;
				FieldGrade = string.Empty;
				FieldPlateLength = 0.0;
				FieldPlateWidth = 0.0;
				FieldPlateThickness = 0.0;
				FieldProfileLength = 0.0;
				FieldWebWidth = 0.0;
				FieldFlangeWidth = 0.0;
				FieldWebThickness = 0.0;
				FieldFlangeThickness = 0.0;
				FieldIsTwisted = string.Empty;
				FieldCurved = string.Empty;
				FieldPartDwgNumber = string.Empty;
				FieldArrgDwgNumber = string.Empty;
				FieldMfgPartId = string.Empty;
				FieldMfgPartName = string.Empty;
				FieldNestType = string.Empty;
				FieldNestID = string.Empty;
				FieldNestingStage = string.Empty;
				FieldCOGx = 0.0;
				FieldCOGy = 0.0;
				FieldCOGz = 0.0;
				FieldWeight = 0.0;
				FieldArea = 0.0;
			}

			/// <summary>
			/// copy constructor
			/// </summary>
			/// <param name="oFrom">struct to copy from</param>
			public BOMInfo(BOMInfo oFrom)
			{
				FieldPartID = oFrom.FieldPartID;
				FieldPartName = oFrom.FieldPartName;
				FieldSymmetry = oFrom.FieldSymmetry;
				FieldObject = oFrom.FieldObject;
				FieldObjectType = oFrom.FieldObjectType;
				FieldProfileType = oFrom.FieldProfileType;
				FieldStatus = oFrom.FieldStatus;
				FieldStage = oFrom.FieldStage;
				FieldMaterialType = oFrom.FieldMaterialType;
				FieldGrade = oFrom.FieldGrade;
				FieldPlateLength = oFrom.FieldPlateLength;
				FieldPlateWidth = oFrom.FieldPlateWidth;
				FieldPlateThickness = oFrom.FieldPlateThickness;
				FieldProfileLength = oFrom.FieldProfileLength;
				FieldWebWidth = oFrom.FieldWebWidth;
				FieldFlangeWidth = oFrom.FieldFlangeWidth;
				FieldWebThickness = oFrom.FieldWebThickness;
				FieldFlangeThickness = oFrom.FieldFlangeThickness;
				FieldIsTwisted = oFrom.FieldIsTwisted;
				FieldCurved = oFrom.FieldCurved;
				FieldPartDwgNumber = oFrom.FieldPartDwgNumber;
				FieldArrgDwgNumber = oFrom.FieldArrgDwgNumber;
				FieldMfgPartId = oFrom.FieldMfgPartId;
				FieldMfgPartName = oFrom.FieldMfgPartName;
				FieldNestType = oFrom.FieldNestType;
				FieldNestID = oFrom.FieldNestID;
				FieldNestingStage = oFrom.FieldNestingStage;
				FieldCOGx = oFrom.FieldCOGx;
				FieldCOGy = oFrom.FieldCOGy;
				FieldCOGz = oFrom.FieldCOGz;
				FieldWeight = oFrom.FieldWeight;
				FieldArea = oFrom.FieldArea;
			}

			/// <summary>
			/// fill a DataRow with this info
			/// </summary>
			/// <param name="currentRow">the DataRow class to fill</param>
			public DataRow Data(DataTable dataTable)
			{
				DataRow currentRow = dataTable.NewRow();

				currentRow.SetField(FIELD_PART_ID, FieldPartID);
				currentRow.SetField(FIELD_PART_NAME, FieldPartName);
				currentRow.SetField(FIELD_SYMMETRY, FieldSymmetry);
				currentRow.SetField(FIELD_OBJECT, FieldObject);
				currentRow.SetField(FIELD_OBJECT_TYPE, FieldObjectType);
				currentRow.SetField(FIELD_PROFILE_TYPE, FieldProfileType);
				currentRow.SetField(FIELD_STATUS, FieldStatus);
				currentRow.SetField(FIELD_STAGE, FieldStage);
				currentRow.SetField(FIELD_MATERIAL_TYPE, FieldMaterialType);
				currentRow.SetField(FIELD_GRADE, FieldGrade);
				currentRow.SetField(FIELD_PLATE_LENGTH, FieldPlateLength);
				currentRow.SetField(FIELD_PLATE_WIDTH, FieldPlateWidth);
				currentRow.SetField(FIELD_PLATE_THICKNESS, FieldPlateThickness);
				currentRow.SetField(FIELD_PROFILE_LENGTH, FieldProfileLength);
				currentRow.SetField(FIELD_WEB_WIDTH, FieldWebWidth);
				currentRow.SetField(FIELD_FLANGE_WIDTH, FieldFlangeWidth);
				currentRow.SetField(FIELD_WEB_THICKNESS, FieldWebThickness);
				currentRow.SetField(FIELD_FLANGE_THICKNESS, FieldFlangeThickness);
				currentRow.SetField(FIELD_IS_TWISTED, FieldIsTwisted);
				currentRow.SetField(FIELD_CURVED, FieldCurved);
				currentRow.SetField(FIELD_PART_DWG_NUMBER, FieldPartDwgNumber);
				currentRow.SetField(FIELD_ARRG_DWG_NUMBER, FieldArrgDwgNumber);
				currentRow.SetField(FIELD_MFG_PART_ID, FieldMfgPartId);
				currentRow.SetField(FIELD_MFG_PART_NAME, FieldMfgPartName);
				currentRow.SetField(FIELD_NEST_TYPE, FieldNestType);
				currentRow.SetField(FIELD_NEST_ID, FieldNestID);
				currentRow.SetField(FIELD_NESTING_STAGE, FieldNestingStage);
				currentRow.SetField(FIELD_COG_X, FieldCOGx);
				currentRow.SetField(FIELD_COG_Y, FieldCOGy);
				currentRow.SetField(FIELD_COG_Z, FieldCOGz);
				currentRow.SetField(FIELD_WEIGHT, FieldWeight);
				currentRow.SetField(FIELD_AREA, FieldArea);

				return currentRow;
			}
		}
		#endregion

		/// <summary>
		/// default constructor
		/// </summary>
		public StructuralBOMQI()
		{

		}

		/// <summary>
		/// public method called by the Reports environment framework to retrieve the data
		/// </summary>
		/// <param name="action">not used by this QI</param>
		/// <param name="argument">not used by this QI</param>
		/// <returns>the resulting DataTable</returns>
		public override System.Data.DataTable Execute(string action, string argument)
		{
			try
			{
				m_DataTable = InitializeDataTable(CreateColumnList());

				if (EvaluateOnly)
				{
					return m_DataTable;
				}

				//----------------------------------------------------------------------------------

				List<BusinessObject> SortedObjects = new List<BusinessObject>();
				foreach (BusinessObject oBO in InputObjects)
				{
					// only process plates and profiles for now
					if (oBO is PlatePartBase || oBO is ProfilePart)
						SortedObjects.Add(oBO);
				}

				SortedObjects.Sort((BoA, BoB) => BoA.ToString().CompareTo(BoB.ToString()));

				//----------------------------------------------------------------------------------

				foreach (BusinessObject oBO in SortedObjects)
				{
					try
					{
						if (oBO is PlatePartBase)
							ProcessPlatePart(oBO);

						else if (oBO is ProfilePart)
							ProcessProfilePart(oBO);
					}
					catch (Exception e)
					{
						// just log and move on to next row if an error occurs ...
						MiddleServiceProvider.ErrorLogger.Log(
							0,
							typeof(StructuralBOMQI).FullName,
							e.ToString(),
							oBO.ToString(),
							(new StackTrace()).GetFrame(1).GetMethod().Name,
							string.Empty,
							string.Empty,
							-1);
					}
				}

				//----------------------------------------------------------------------------------

				return m_DataTable;
			}
			catch (Exception e)
			{
				MiddleServiceProvider.ErrorLogger.Log(
					0,
					typeof(StructuralBOMQI).FullName,
					e.ToString(),
					string.Empty,
					(new StackTrace()).GetFrame(1).GetMethod().Name,
					string.Empty,
					string.Empty,
					-1);
				throw;
			}
		}

		/// <summary>
		/// create the columns expected by the report
		/// </summary>
		/// <returns>List of created Column objects</returns>
		private List<Column> CreateColumnList()
		{
			List<Column> FieldColumnList = new List<Column>();

			Column recFieldColumnPartID = new Column(FIELD_PART_ID, typeof(System.String));
			FieldColumnList.Add(recFieldColumnPartID);

			Column recFieldColumnPartName = new Column(FIELD_PART_NAME, typeof(System.String));
			FieldColumnList.Add(recFieldColumnPartName);

			Column recFieldColumnSymmetry = new Column(FIELD_SYMMETRY, typeof(System.String));
			FieldColumnList.Add(recFieldColumnSymmetry);

			Column recFieldColumnObject = new Column(FIELD_OBJECT, typeof(System.String));
			FieldColumnList.Add(recFieldColumnObject);

			Column recFieldColumnObjectType = new Column(FIELD_OBJECT_TYPE, typeof(System.String));
			FieldColumnList.Add(recFieldColumnObjectType);

			Column recFieldColumnProfileType = new Column(FIELD_PROFILE_TYPE, typeof(System.String));
			FieldColumnList.Add(recFieldColumnProfileType);

			Column recFieldColumnStatus = new Column(FIELD_STATUS, typeof(System.String));
			FieldColumnList.Add(recFieldColumnStatus);

			Column recFieldColumnStage = new Column(FIELD_STAGE, typeof(System.String));
			FieldColumnList.Add(recFieldColumnStage);

			Column recFieldColumnMaterialType = new Column(FIELD_MATERIAL_TYPE, typeof(System.String));
			FieldColumnList.Add(recFieldColumnMaterialType);

			Column recFieldColumnGrade = new Column(FIELD_GRADE, typeof(System.String));
			FieldColumnList.Add(recFieldColumnGrade);

			Column recFieldColumnPlateLength = new Column(FIELD_PLATE_LENGTH, typeof(System.Double));
			FieldColumnList.Add(recFieldColumnPlateLength);

			Column recFieldColumnPlateWidth = new Column(FIELD_PLATE_WIDTH, typeof(System.Double));
			FieldColumnList.Add(recFieldColumnPlateWidth);

			Column recFieldColumnPlateThickness = new Column(FIELD_PLATE_THICKNESS, typeof(System.Double));
			FieldColumnList.Add(recFieldColumnPlateThickness);

			Column recFieldColumnProfileLength = new Column(FIELD_PROFILE_LENGTH, typeof(System.Double));
			FieldColumnList.Add(recFieldColumnProfileLength);

			Column recFieldColumnWebWidth = new Column(FIELD_WEB_WIDTH, typeof(System.Double));
			FieldColumnList.Add(recFieldColumnWebWidth);

			Column recFieldColumnFlangeWidth = new Column(FIELD_FLANGE_WIDTH, typeof(System.Double));
			FieldColumnList.Add(recFieldColumnFlangeWidth);

			Column recFieldColumnWebThickness = new Column(FIELD_WEB_THICKNESS, typeof(System.Double));
			FieldColumnList.Add(recFieldColumnWebThickness);

			Column recFieldColumnFlangeThickness = new Column(FIELD_FLANGE_THICKNESS, typeof(System.Double));
			FieldColumnList.Add(recFieldColumnFlangeThickness);

			Column recFieldColumnIsTwisted = new Column(FIELD_IS_TWISTED, typeof(System.String));
			FieldColumnList.Add(recFieldColumnIsTwisted);

			Column recFieldColumnCurved = new Column(FIELD_CURVED, typeof(System.String));
			FieldColumnList.Add(recFieldColumnCurved);

			Column recFieldColumnPartDwgNumber = new Column(FIELD_PART_DWG_NUMBER, typeof(System.String));
			FieldColumnList.Add(recFieldColumnPartDwgNumber);

			Column recFieldColumnArrgDwgNumber = new Column(FIELD_ARRG_DWG_NUMBER, typeof(System.String));
			FieldColumnList.Add(recFieldColumnArrgDwgNumber);

			Column recFieldColumnMfgPartId = new Column(FIELD_MFG_PART_ID, typeof(System.String));
			FieldColumnList.Add(recFieldColumnMfgPartId);

			Column recFieldColumnMfgPartName = new Column(FIELD_MFG_PART_NAME, typeof(System.String));
			FieldColumnList.Add(recFieldColumnMfgPartName);

			Column recFieldColumnNestType = new Column(FIELD_NEST_TYPE, typeof(System.String));
			FieldColumnList.Add(recFieldColumnNestType);

			Column recFieldColumnNestID = new Column(FIELD_NEST_ID, typeof(System.String));
			FieldColumnList.Add(recFieldColumnNestID);

			Column recFieldColumnNestingStage = new Column(FIELD_NESTING_STAGE, typeof(System.String));
			FieldColumnList.Add(recFieldColumnNestingStage);

			Column recFieldColumnCOGx = new Column(FIELD_COG_X, typeof(System.Double));
			FieldColumnList.Add(recFieldColumnCOGx);

			Column recFieldColumnCOGy = new Column(FIELD_COG_Y, typeof(System.Double));
			FieldColumnList.Add(recFieldColumnCOGy);

			Column recFieldColumnCOGz = new Column(FIELD_COG_Z, typeof(System.Double));
			FieldColumnList.Add(recFieldColumnCOGz);

			Column recFieldColumnWeight = new Column(FIELD_WEIGHT, typeof(System.Double));
			FieldColumnList.Add(recFieldColumnWeight);

			Column recFieldColumnArea = new Column(FIELD_AREA, typeof(System.Double));
			FieldColumnList.Add(recFieldColumnArea);

			return FieldColumnList;
		}

		/// <summary>
		/// process a PlatePartBase class object
		/// </summary>
		/// <param name="oBO">the structural part (assume it is a PlatePartBase)</param>
		private void ProcessPlatePart(BusinessObject oBO)
		{
			try
			{
				BOMInfo oBaseData = new BOMInfo();

				PlatePartBase oPlate = oBO as PlatePartBase;
				ManufacturingOutputBase oMfgPart = GetMfgPart(oBO);

				//----------------------------------------------------------------------------------

				List<NestData> oNestingData = GetNestData(oMfgPart);

				//----------------------------------------------------------------------------------

				oBaseData.FieldPartID = oBO.ObjectID.ToString();
				oBaseData.FieldPartName = oBO.ToString();
				oBaseData.FieldObject = oBO.ClassInfo.DisplayName;
				oBaseData.FieldObjectType = oBO.ClassInfo.Name;
				oBaseData.FieldStatus = oBO.ApprovalStatus.ToString();

				//----------------------------------------------------------------------------------

				oBaseData.FieldStage = GetPartState(oPlate, oMfgPart);

				//----------------------------------------------------------------------------------

				oBaseData.FieldSymmetry = oPlate.Symmetry.ToString();
				oBaseData.FieldPlateLength = GetPlateLength(oPlate, oMfgPart);
				oBaseData.FieldPlateWidth = GetPlateWidth(oPlate, oMfgPart);
				oBaseData.FieldPlateThickness = GetPlateThickness(oPlate);
				oBaseData.FieldCurved = GetPlateCurvature(oPlate);
				oBaseData.FieldMaterialType = oPlate.MaterialType;
				oBaseData.FieldGrade = oPlate.MaterialGrade;
				oBaseData.FieldArea = GetPlateArea(oPlate);

				//----------------------------------------------------------------------------------

				ProcessMfgPart(oMfgPart, oBaseData);
				ProcessWeightAndCG(oPlate, oBaseData);

				//----------------------------------------------------------------------------------

				if (oNestingData.Count > 0)
				{
					foreach (NestData oNestData in oNestingData)
					{
						BOMInfo oNewRow = new BOMInfo(oBaseData);
						oNewRow.FieldNestType = oNestData.PartType;
						oNewRow.FieldNestID = oNestData.LotNumber;
						m_DataTable.Rows.Add(oNewRow.Data(m_DataTable));
					}
				}
				else
					m_DataTable.Rows.Add(oBaseData.Data(m_DataTable));
			}
			catch (Exception e)
			{
				MiddleServiceProvider.ErrorLogger.Log(
					0,
					typeof(StructuralBOMQI).FullName,
					e.ToString(),
					string.Empty,
					(new StackTrace()).GetFrame(1).GetMethod().Name,
					string.Empty,
					string.Empty,
					-1);
				throw;
			}
		}

		/// <summary>
		/// process a ProfilePart class object
		/// </summary>
		/// <param name="oBO">the structural part (assume it is a ProfilePart)</param>
		private void ProcessProfilePart(BusinessObject oBO)
		{
			try
			{
				BOMInfo oBaseData = new BOMInfo();

				ProfilePart oProfile = oBO as ProfilePart;
				ManufacturingOutputBase oMfgPart = GetMfgPart(oBO);

				//----------------------------------------------------------------------------------

				List<NestData> oNestingData = GetNestData(oMfgPart);

				//----------------------------------------------------------------------------------

				oBaseData.FieldPartID = oBO.ObjectID.ToString();
				oBaseData.FieldPartName = oBO.ToString();
				oBaseData.FieldObject = oBO.ClassInfo.DisplayName;
				oBaseData.FieldObjectType = oBO.ClassInfo.Name;
				oBaseData.FieldStatus = oBO.ApprovalStatus.ToString();

				//----------------------------------------------------------------------------------

				oBaseData.FieldStage = GetPartState(oProfile, oMfgPart);

				//----------------------------------------------------------------------------------

				oBaseData.FieldSymmetry = GetProfileSymmetry(oProfile);
				oBaseData.FieldProfileType = oProfile.CrossSection.CrossSectionClass.ToString();
				oBaseData.FieldPlateLength = GetProfileLength(oProfile, oMfgPart);
				oBaseData.FieldIsTwisted = GetProfileIsTwisted(oProfile);
				oBaseData.FieldCurved = GetProfileCurvature(oProfile);
				oBaseData.FieldMaterialType = oProfile.MaterialType;
				oBaseData.FieldGrade = oProfile.MaterialGrade;
				oBaseData.FieldArea = GetProfileArea(oProfile);

				//----------------------------------------------------------------------------------

				ProcessMfgPart(oMfgPart, oBaseData);
				ProcessWeightAndCG(oProfile, oBaseData);
				ProcessCrossSection(oProfile.CrossSection, oBaseData);

				//----------------------------------------------------------------------------------

				if (oNestingData.Count > 0)
				{
					foreach (NestData oNestData in oNestingData)
					{
						BOMInfo oNewRow = new BOMInfo(oBaseData);
						oNewRow.FieldNestType = oNestData.PartType;
						oNewRow.FieldNestID = oNestData.LotNumber;
						m_DataTable.Rows.Add(oNewRow.Data(m_DataTable));
					}
				}
				else
					m_DataTable.Rows.Add(oBaseData.Data(m_DataTable));
			}
			catch (Exception e)
			{
				MiddleServiceProvider.ErrorLogger.Log(
					0,
					typeof(StructuralBOMQI).FullName,
					e.ToString(),
					string.Empty,
					(new StackTrace()).GetFrame(1).GetMethod().Name,
					string.Empty,
					string.Empty,
					-1);
				throw;
			}
		}

		/// <summary>
		/// process the manufacturing data
		/// </summary>
		/// <param name="oMfgPart">the manufacturing part</param>
		/// <param name="oBaseData">the structure to fill</param>
		private void ProcessMfgPart(ManufacturingOutputBase oMfgPart, BOMInfo oBaseData)
		{
			if (oMfgPart != null)
			{
				oBaseData.FieldMfgPartId = oMfgPart.ObjectID.ToString();
				oBaseData.FieldMfgPartName = oMfgPart.ToString();
			}
		}

		/// <summary>
		/// process the weight & cg data
		/// </summary>
		/// <param name="oBO">the structural part</param>
		/// <param name="oBaseData">the structure to fill</param>
		private void ProcessWeightAndCG(BusinessObject oBO, BOMInfo oBaseData)
		{
			const string IJWeightCG = "IJWeightCG";

			double dDryCGX = StructHelper.GetDoubleProperty(oBO, IJWeightCG, "DryCGX");
			double dDryCGY = StructHelper.GetDoubleProperty(oBO, IJWeightCG, "DryCGY");
			double dDryCGZ = StructHelper.GetDoubleProperty(oBO, IJWeightCG, "DryCGZ");
			double dDryWeight = StructHelper.GetDoubleProperty(oBO, IJWeightCG, "DryWeight");

			oBaseData.FieldCOGx = dDryCGX;
			oBaseData.FieldCOGy = dDryCGY;
			oBaseData.FieldCOGz = dDryCGZ;
			oBaseData.FieldWeight = dDryWeight;
		}

		/// <summary>
		/// process the cross section data
		/// </summary>
		/// <param name="oCrossSection">the cross section</param>
		/// <param name="oBaseData">the structure to fill</param>
		private void ProcessCrossSection(CrossSection oCrossSection, BOMInfo oBaseData)
		{
			const string IJUAXSectionWeb = "IJUAXSectionWeb";
			const string IJUAXSectionFlange = "IJUAXSectionFlange";

			double dWebWidth = StructHelper.GetDoubleProperty(oCrossSection, IJUAXSectionWeb, "WebLength");
			double dWebThickness = StructHelper.GetDoubleProperty(oCrossSection, IJUAXSectionWeb, "WebThickness");

			// not all cross sections have flanges ...
			double dFlangeWidth = 0.0, dFlangeThickness = 0.0;
			if (oCrossSection.SupportsInterface(IJUAXSectionFlange))
			{
				dFlangeWidth = StructHelper.GetDoubleProperty(oCrossSection, IJUAXSectionFlange, "FlangeLength");
				dFlangeThickness = StructHelper.GetDoubleProperty(oCrossSection, IJUAXSectionFlange, "FlangeThickness");
			}

			oBaseData.FieldWebWidth = dWebWidth;
			oBaseData.FieldFlangeWidth = dFlangeWidth;
			oBaseData.FieldWebThickness = dWebThickness;
			oBaseData.FieldFlangeThickness = dFlangeThickness;
		}

		/// <summary>
		/// get the manufacturing part if it exists
		/// </summary>
		/// <param name="oBO">the structural part</param>
		/// <returns>the manufacturing part (if it exists)</returns>
		private ManufacturingOutputBase GetMfgPart(BusinessObject oBO)
		{
			ManufacturingOutputBase oMfgPart = null;

			RelationCollection oRelationships = oBO.GetRelationship("StrMfgHierarchy", "IsStrMfgChildOf");
			if (oRelationships != null)
			{
				if (oRelationships.TargetObjects.Count > 0)
				{
					foreach (BusinessObject oMfgObj in oRelationships.TargetObjects)
					{
						if (oMfgObj is ManufacturingPlate || oMfgObj is ManufacturingProfile)
						{
							oMfgPart = oMfgObj as ManufacturingOutputBase;
							break;	// there should only be one ...
						}
					}
				}
			}

			return oMfgPart;
		}

		/// <summary>
		/// get the nesting data for the manufacturing part if it exists
		/// </summary>
		/// <param name="oMfgPart">the manufacturing part</param>
		/// <returns>List of nesting data (if there is any)</returns>
		private List<NestData> GetNestData(ManufacturingOutputBase oMfgPart)
		{
			List<NestData> oNestData = new List<NestData>();

			if (oMfgPart != null)
			{
				RelationCollection oRelationships = oMfgPart.GetRelationship("MfgPartNestData", "MfgPartNestData_ORIG");
				if (oRelationships != null)
				{
					if (oRelationships.TargetObjects.Count > 0)
					{
						foreach (BusinessObject oMfgObj in oRelationships.TargetObjects)
						{
							if (oMfgObj is NestData)
								oNestData.Add(oMfgObj as NestData);
						}
					}
				}
			}

			return oNestData;
		}

		/// <summary>
		/// return the part state (detailed, light, manufactured)
		/// </summary>
		/// <param name="oBO">the structural part</param>
		/// <param name="oMfgPart">its manufacturing part</param>
		/// <returns>string designating the part state</returns>
		private string GetPartState(BusinessObject oBO, ManufacturingOutputBase oMfgPart)
		{
			string sPartState = String.Empty;

			if (oMfgPart != null)
			{
				sPartState = "Manufacturing Part";
			}
			else
			{
				if (oBO is PlatePart)
				{
					PlatePart oPlatePart = oBO as PlatePart;
					sPartState = oPlatePart.PartGeometryState.ToString();
				}
				else if (oBO is CollarPart || oBO is StandAlonePlatePart)
				{
					// collar and stand-alone plate parts are always detailed ...
					sPartState = PartGeometryStateType.DetailedPart.ToString();
				}
				else if (oBO is StiffenerPart)
				{
					StiffenerPart oStiffenerPart = oBO as StiffenerPart;
					sPartState = oStiffenerPart.PartGeometryState.ToString();
				}
				else if (oBO is EdgeReinforcementPart)
				{
					EdgeReinforcementPart oERPart = oBO as EdgeReinforcementPart;
					sPartState = oERPart.PartGeometryState.ToString();
				}
				else if (oBO.SupportsInterface("IJPartGeometryState"))
				{
					PropertyValue vState = oBO.GetPropertyValue("IJPartGeometryState", "PartGeometryState");
					PropertyValueCodelist vclState = (PropertyValueCodelist)vState;
					sPartState = vclState.ToString();
				}
				else
				{
					// what is it ...
				}
			}

			return sPartState;
		}

		/// <summary>
		/// get the part symmetry value
		/// </summary>
		/// <param name="oBO">the structural part</param>
		/// <returns>string indicating the part symmetry</returns>
		private string GetPartSymmetry(BusinessObject oBO)
		{
			string sSymmetry = String.Empty;

			PropertyValue vSymmetry = oBO.GetPropertyValue("IJBoardMgt", "Symmetry");
			PropertyValueCodelist vclSymmetry = (PropertyValueCodelist)vSymmetry;

			// to maintain constistency, translate it to BoardManagementSymmetry enumeration
			BoardManagementSymmetry eSymmetry = (BoardManagementSymmetry)vclSymmetry.PropValue;
			sSymmetry = eSymmetry.ToString();

			return sSymmetry;
		}

		/// <summary>
		/// get the plate length - manufacturing length overrides part length
		/// </summary>
		/// <param name="oPlate">the plate part</param>
		/// <param name="oMfgPart">the manufacturing part (if there is one)</param>
		/// <returns>double value - plate length</returns>
		private double GetPlateLength(PlatePartBase oPlate, ManufacturingOutputBase oMfgPart)
		{
			double dLength = oPlate.Length;

			if (oMfgPart != null)
			{
				// have to get it from the COM interface - ManufacturingPlate doesn't expose it ...
				dLength = StructHelper.GetDoubleProperty(oMfgPart, "IJMfgPlatePart", "Length");
			}

			return dLength;
		}

		/// <summary>
		/// get the plate width - manufacturing width overrides part width
		/// </summary>
		/// <param name="oPlate">the plate part</param>
		/// <param name="oMfgPart">the manufacturing part (if there is one)</param>
		/// <returns>double value - plate width</returns>
		private double GetPlateWidth(PlatePartBase oPlate, ManufacturingOutputBase oMfgPart)
		{
			double dWidth = oPlate.Width;

			if (oMfgPart != null)
			{
				// have to get it from the COM interface - ManufacturingPlate doesn't expose it ...
				dWidth = StructHelper.GetDoubleProperty(oMfgPart, "IJMfgPlatePart", "Width");
			}

			return dWidth;
		}

		/// <summary>
		/// get plate thickness
		/// </summary>
		/// <param name="oPlate">the plate part</param>
		/// <returns>double value - plate thickness</returns>
		private double GetPlateThickness(PlatePartBase oPlate)
		{
			double dThickness = 0.0;

			if (oPlate is PlatePart)
			{
				PlatePart oPlatePart = oPlate as PlatePart;
				dThickness = oPlatePart.Thickness;
			}
			else if (oPlate is CollarPart)
			{
				CollarPart oPlatePart = oPlate as CollarPart;
				dThickness = oPlatePart.Thickness;
			}
			else if (oPlate is CustomPlatePart)
			{
				CustomPlatePart oPlatePart = oPlate as CustomPlatePart;
				dThickness = oPlatePart.Thickness;
			}
			else if (oPlate is StandAlonePlatePart)
			{
				StandAlonePlatePart oPlatePart = oPlate as StandAlonePlatePart;
				dThickness = oPlatePart.Thickness;
			}

			return dThickness;
		}

		/// <summary>
		/// get plate curavture - return empty string if plate is flat
		/// </summary>
		/// <param name="oPlate">the plate part</param>
		/// <returns>string indicating plate curvature</returns>
		private string GetPlateCurvature(PlatePartBase oPlate)
		{
			string sCurvature = String.Empty;

			PlatePartBase oPlatePart = oPlate as PlatePartBase;
			if (oPlatePart.Curvature != Curvature.Flat)
				sCurvature = oPlatePart.Curvature.ToString();

			return sCurvature;
		}

		/// <summary>
		/// get plate area
		/// </summary>
		/// <param name="oPlate">the plate part</param>
		/// <returns>double value - plate area</returns>
		private double GetPlateArea(PlatePartBase oPlate)
		{
			double dArea = 0.0;

			/*
			 * CAUTION - TotalSurfaceArea is calling the wrong COM method and is only returning
			 *				the surface area of one side of the plate instead of both sides
			 *				skip that call until TR289084 is fixed
			if (oPlate is PlatePart)
			{
				PlatePart oPlatePart = oPlate as PlatePart;
				dArea = oPlatePart.TotalSurfaceArea;
			}
			else
			{
				// SupportsInterface is costly - assuming calling method knows the BO is a IJPlatePart ...
				dArea = StructHelper.GetDoubleProperty(oPlate, "IJPlatePart", "Area");
			}
			*/

			// SupportsInterface is costly - assuming calling method knows the BO is a IJPlatePart ...
			dArea = StructHelper.GetDoubleProperty(oPlate, "IJPlatePart", "Area");

			return dArea;
		}

		/// <summary>
		/// get profile symmetry
		/// </summary>
		/// <param name="oProfile">the profile part</param>
		/// <returns>string indicating profile symmetry</returns>
		private string GetProfileSymmetry(ProfilePart oProfile)
		{
			string sSymmetry = String.Empty;

			if (oProfile is StiffenerPartBase)
			{
				StiffenerPartBase oStiffener = oProfile as StiffenerPartBase;
				sSymmetry = oStiffener.Symmetry.ToString();
			}
			else
			{
				sSymmetry = GetPartSymmetry(oProfile);
			}

			return sSymmetry;
		}

		/// <summary>
		/// get the profile length - manufacturing length overrides part length
		/// </summary>
		/// <param name="oProfile">the profile part</param>
		/// <param name="oMfgPart">the manufacturing part (if it exists)</param>
		/// <returns>double value - profile length</returns>
		private double GetProfileLength(ProfilePart oProfile, ManufacturingOutputBase oMfgPart)
		{
			double dLength = 0.0;

			if (oMfgPart == null)
			{
				// SupportsInterface is costly - assuming calling method knows the BO is a IJProfilePart ...
				dLength = StructHelper.GetDoubleProperty(oProfile, "IJProfilePart", "ProfileLength");
			}
			else
			{
				// SupportsInterface is costly - assuming calling method knows the mfg object is a IJMfgProfileLengths ...
				dLength = StructHelper.GetDoubleProperty(oMfgPart, "IJMfgProfileLengths", "BeforeFeaturesTotal");
			}

			return dLength;
		}

		/// <summary>
		/// get if the profile is twisted
		/// </summary>
		/// <param name="oBO">the structural part - assumed to be IJProfilePart</param>
		/// <returns>string indicating if the profile is twisted</returns>
		private string GetProfileIsTwisted(BusinessObject oBO)
		{
			string sIsTwisted = String.Empty;

			// SupportsInterface is costly - assuming calling method knows the BO is a IJProfilePart ...
			bool bIsTwisted = StructHelper.GetBoolProperty(oBO, "IJProfilePart", "IsTwisted");
			sIsTwisted = bIsTwisted.ToString();

			return sIsTwisted;
		}

		/// <summary>
		/// get profile curvature - return empty string if profile is straight
		/// </summary>
		/// <param name="oBO">the structural part - assumed to be IJProfilePart</param>
		/// <returns>string indicating profile curvature</returns>
		private string GetProfileCurvature(BusinessObject oBO)
		{
			string sCurvature = String.Empty;

			// SupportsInterface is costly - assuming calling method knows the BO is a IJProfilePart ...
			PropertyValue vCurved = oBO.GetPropertyValue("IJProfilePart", "Curved");
			if (vCurved.ToString() != "Straight")
				sCurvature = vCurved.ToString();

			return sCurvature;
		}

		/// <summary>
		/// get profile area
		/// </summary>
		/// <param name="oBO">the structural part - assumed to be IJProfilePart</param>
		/// <returns>double value - profile area</returns>
		private double GetProfileArea(BusinessObject oBO)
		{
			double dArea = 0.0;

			// SupportsInterface is costly - assuming calling method knows the BO is a IJProfilePart ...
			dArea = StructHelper.GetDoubleProperty(oBO, "IJProfilePart", "Area");

			return dArea;
		}
	}
}
