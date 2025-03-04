//--------------------------------------------------------------------------------------------------
//	Copyright (C) 2015-2016, Intergraph Corporation.  All rights reserved.
//
//	FILE: DBREndCutRule.cs
//
//	DESCRIPTION:
//		DBR rule implementation for determining what end cut symbol file to use for drawing
//		end cuts on landing curves
//
//	HISTORY:
//		Oct-08-2015		Pam Livingston
//			CR-CP-273727 Resymbolization of the Detailing endcuts from the Profile properties
//
//      Nov-12-2015     Emeralds
//           Updated the code to API standards. Added error handling where missing. Use Cmn version
//           exceptions so the errors will get logged on to S3DErrors.log file.
//
//      Jan-22-2016     Squids
//           DI-CP-287359 	Fix DwgEndCutRules project for DBRATP_Generated ATP failure
//
//--------------------------------------------------------------------------------------------------

using System;
using System.Linq;
using System.Xml;

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Drawings.Middle;
using Ingr.SP3D.Drawings.Middle.Services;
using Ingr.SP3D.Drawings.Exceptions;

using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Common.Exceptions;
using Ingr.SP3D.Common.Middle.Services;
using System.Text;
using System.IO;

namespace DwgEndCutRules
{
	public class DBREndCutRule : EndCutRule
	{
		#region Private Data

		private struct Codelist
		{
			/// <summary>
			/// the tablename for the codelist
			/// </summary>
			public string sTableName;

			/// <summary>
			/// the namespace for the codelist
			/// </summary>
			public string sNamespace;

			/// <summary>
			/// constructor
			/// </summary>
			/// <param name="table"></param>
			/// <param name="space"></param>
			public Codelist(string table, string space)
			{
				sTableName = table;
				sNamespace = space;
			}
		}

		private string m_sSymbolDirectory = "";

		private XmlNode m_xCustomNode = null;

		private Codelist[] m_sCodelists = { new Codelist("", ""),		// ConnectionType
											new Codelist("", ""),		// PrimaryWebCutInfo
											new Codelist("", ""),		// TopFlangeCutInfo
											new Codelist("", ""),		// BottomFlangeCutInfo
											new Codelist("", "") };		// SecondaryWebCutInfo
		#endregion

		#region Public Interface

		public DBREndCutRule()
			: base()
		{
		}

		public override void Initialize(string endcutRulePath)
		{
            if (string.IsNullOrEmpty(endcutRulePath))
                throw new CmnArgumentNullException("Argument 'End Cut Rule Path' is either NULL or Empty");

            if (System.IO.File.Exists(endcutRulePath) == false)
            {
                MiddleServiceProvider.ErrorLogger.Log("End Cut Rule file provided is not found. End Cut Rule Path: " + endcutRulePath, 1);
                throw new FileNotFoundException("End Cut Rule file provided is not found. End Cut Rule Path: " + endcutRulePath);
            }

			try
			{
                m_sSymbolDirectory = System.IO.Directory.GetParent(endcutRulePath).FullName;

				//--------------------------------------------------------------------------------------
				//	load the XML rule and get the <RULE/CUSTOM> node
				XmlDocument xDocument = new XmlDocument();
                xDocument.Load(endcutRulePath);

				m_xCustomNode = xDocument.SelectSingleNode("RULE/CUSTOM");
				if (m_xCustomNode == null)
                    throw new CmnUnexpectedException("Invalid End Cut Rule xml document, does not have minimum mandatory RULE/CUSTOM node. End Cut File Path: " + endcutRulePath);
                
				//----------------------------------------------------------------------------------
				//	read in the codelist tables to use (these are optional)
				GetCodelists();
			}
			catch (Exception ex)
			{
                throw new CmnUnexpectedException("Unexpected failure while Initializing the DBREndCut rule. End Cut Rule Path: " + endcutRulePath + ", Exception Message: " + ex.Message);
			}
		}

		public override string GetEndCutSymbolFilePath(BusinessObject bo, PositionOnObject location)
		{
            if (bo == null)
            {
                throw new CmnArgumentNullException("Input object is null");
            }

			string symbolFilePath = "";

			try
			{
				//----------------------------------------------------------------------------------
				// make sure Initialize() succeeded

				if (m_xCustomNode == null)
					throw new CmnUnexpectedException("Initialize did not succeed");

				//----------------------------------------------------------------------------------
				//	retrieve the profile info (see GetProfileInfo() for list of data)
				string[] sProfileInfo = GetProfileInfo(bo, location);

				//----------------------------------------------------------------------------------
				//	build the XML path to use to match the node in the XML file

				string sXMLSearchPath = GetXMLSearchPath(sProfileInfo);
				if (sXMLSearchPath.Length == 0)
                    throw new CmnUnexpectedException("Invalid XML path");

				//----------------------------------------------------------------------------------
				//	select the node from the XML and retrieve the symbol file name

				XmlNode xInfoNode = m_xCustomNode.SelectSingleNode(sXMLSearchPath);
				if (xInfoNode != null)
				{
					XmlNode xSymbolNode = xInfoNode.SelectSingleNode("Symbol");
					if (xSymbolNode != null)
					{
						symbolFilePath = System.IO.Path.Combine(m_sSymbolDirectory, xSymbolNode.InnerText);
					}
				}
			}
			catch (Exception ex)
			{
                throw new CmnUnexpectedException("Unexpected failure while fetching the End Cut Symbol Path. Error Message: " + ex.Message);
			}

			return symbolFilePath;
		}
		#endregion

		#region Private Methods

		private void GetCodelists()
		{
			//	<Codelists>
			//		<ConnectionType namespace="STRUCT">StiffenerConnectionType</ConnectionType>
			//		<PrimaryWebCutInfo namespace="UDP">WebCutDrawingTypeCodeList</PrimaryWebCutInfo>
			//		<TopFlangeCutInfo namespace="UDP">FlangeCutDrawingTypeCodeList</TopFlangeCutInfo>
			//		<BottomFlangeCutInfo namespace="UDP">FlangeCutDrawingTypeCodeList</BottomFlangeCutInfo>
			//		<SecondaryWebCutInfo namespace="UDP">WebCutDrawingTypeCodeList</SecondaryWebCutInfo>
			//	</Codelists>

			// the codelist nodes are optional - if a codelist table is not specified for a value - then the
			// actual value will be used
			try
			{
				XmlNode xCodelists = m_xCustomNode.SelectSingleNode("Codelists");
				if (xCodelists != null)
				{
					string[] sNodes = { "ConnectionType",
										"PrimaryWebCutInfo",
										"TopFlangeCutInfo",
										"BottomFlangeCutInfo",
										"SecondaryWebCutInfo" };
					for (int ii = 0; ii < 5; ii++)
					{
						XmlElement xElement = (XmlElement)xCodelists.SelectSingleNode(sNodes[ii]);
						if (xElement != null)
						{
							m_sCodelists[ii].sTableName = xElement.InnerText;
							m_sCodelists[ii].sNamespace = xElement.GetAttribute("namespace");
						}
					}
				}
			}
			catch (Exception ex)
			{
                throw new CmnUnexpectedException("Unexpected failure in GetCodeLists. Message: " + ex.Message);
			}
		}

		private string[] GetProfileInfo(BusinessObject bo, PositionOnObject location)
		{
			string[] sProfileInfo = { "None", "None", "None",
									  "None", "None", "None",
									  "None", "None", "None",
									  "None", "None", "None",
									  "None", "None", "None" };

			try
			{
				// sProfileInfo will hold the values to check against the XML
				//		 0	=	ProfileSide
				//		 1	=	SectionType
				//		 2	=	ConnectionType
				//
				//		 3	=	PrimaryWebCut - First Cut
				//		 4	=	PrimaryWebCut - Top Cut
				//		 5	=	PrimaryWebCut - Bottom Cut
				//
				//		 6	=	TopFlangeCut - First Cut
				//		 7	=	TopFlangeCut - Left Cut
				//		 8	=	TopFlangeCutInfo - Bottom Cut
				//
				//		 9	=	BottomFlangeCut - First Cut
				//		10	=	BottomFlangeCut - Left Cut
				//		11	=	BottomFlangeCut - Right Cut
				//
				//		12	=	SecondaryWebCut - First Cut
				//		13	=	SecondaryWebCut - Top Cut
				//		14	=	SecondaryWebCut - Bottom Cut

                InterfaceInformation interfaceInfo = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr.GetInterfaceInfo("IJSDXSectionType", "STRUCT");
                PropertyInformation propInfo = interfaceInfo.GetPropertyInfo("XSectionType");
                PropertyValueString stringProperty = (PropertyValueString) bo.GetPropertyValue(propInfo);

                sProfileInfo[1] = stringProperty.PropValue;

                string connectionTypePropertyName, endCutDataPropertyName;
				if (location == PositionOnObject.AtFirstEnd)
				{
					sProfileInfo[0] = "Start";
                    connectionTypePropertyName = "StartConnectionType";
                    endCutDataPropertyName = "StartEndCutData";
				}
				else
				{
					sProfileInfo[0] = "End";
                    connectionTypePropertyName = "EndConnectionType";
                    endCutDataPropertyName = "EndEndCutData";
				}

                InterfaceInformation endcutAttrsInterfaceInfo = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr.GetInterfaceInfo("IJEndCutAttributes", "STRUCT");
                PropertyInformation connectionTypePropInfo = endcutAttrsInterfaceInfo.GetPropertyInfo(connectionTypePropertyName);
                PropertyValueCodelist codeList = (PropertyValueCodelist)bo.GetPropertyValue(connectionTypePropInfo);
                sProfileInfo[2] = GetCodelistName(m_sCodelists[0], codeList.PropValue);

                InterfaceInformation endcutInterfaceInfo = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr.GetInterfaceInfo("IJSDEndCutData", "STRUCT");
                PropertyInformation endcutPropInfo = endcutInterfaceInfo.GetPropertyInfo(endCutDataPropertyName);
                PropertyValueString endCutPropValue = (PropertyValueString)bo.GetPropertyValue(endcutPropInfo);
                string sEndCutData = String.Empty;
                if (endCutPropValue.PropValue != null)
                    sEndCutData = endCutPropValue.PropValue;

                //--------------------------------------------------------------------------------------
				if (sEndCutData.Length > 0)
				{
					int index = 3;

					string[] sSplit = sEndCutData.Split('|');
					for (int ii = 0; ii < sSplit.Count(); ii++)
					{
						string[] sSubSplit = sSplit[ii].Split(';');
						for (int jj = 0; jj < sSubSplit.Count(); jj++, index++)
						{
							if (sSubSplit[jj].Length > 0)
								sProfileInfo[index] = GetCodelistName(m_sCodelists[ii+1], sSubSplit[jj]);
							else
								sProfileInfo[index] = "Any";
						}
						for (; index < (ii + 1) * 3 + 3; index++)
						{
							sProfileInfo[index] = "Any";
						}
					}
					for (; index < 15; index++)
					{
						sProfileInfo[index] = "Any";
					}
				}
			}
			catch (Exception excp)
			{
                throw new CmnUnexpectedException("Unexpected failure in fetching the Profile Information. Error Message: " + excp.Message);
			}

			return sProfileInfo;
		}

		private string GetXMLSearchPath(string[] sProfileInfo)
		{
			try
			{
				// sample XML path
				//	SymbolInfo[ProfileInfo
				//		[ProfileSide='Start'] <<<<< NOT USED
				//		[SectionType='EA' or not(SectionType) or SectionType='Any']
				//		[ConnectionType='ByRule' or not(ConnectionType) or ConnectionType='Any']
				//		[PrimaryWebCutInfo[First='W' or not(First) or First='Any'][Top='2' or not(Top) or Top='Any'][Bottom='3' or not(Bottom) or Bottom='Any']]
				//		[TopFlangeCutInfo[First='S' or not(First) or First='Any'][Left='4' or not(Left) or Left='Any'][Right='5' or not(Right) or Right='Any']]
				//		[BottomFlangCutInfo[First='C' or not(First) or First='Any'][Left='6' or not(Left) or Left='Any'][Right='7' or not(Right) or Right='Any']]
				//		[SecondaryWebCutInfo[First='B' or not(First) or First='Any'][Top='8' or not(Top) or Top='Any'][Bottom='9' or not(Bottom) or Bottom='Any']]]

                StringBuilder xmlSearchStringBuilder = new StringBuilder();
                xmlSearchStringBuilder.Append("SymbolInfo[ProfileInfo");
                xmlSearchStringBuilder.Append("[SectionType='");
                xmlSearchStringBuilder.Append(sProfileInfo[1]);
                xmlSearchStringBuilder.Append("' or not(SectionType) or SectionType='Any']");

                xmlSearchStringBuilder.Append("[ConnectionType='");
                xmlSearchStringBuilder.Append(sProfileInfo[2]);
                xmlSearchStringBuilder.Append("' or not(ConnectionType) or ConnectionType='Any']");

                xmlSearchStringBuilder.Append("[PrimaryWebCutInfo[First='");
                xmlSearchStringBuilder.Append(sProfileInfo[3]);
                xmlSearchStringBuilder.Append("' or not(First) or First='Any']");
                xmlSearchStringBuilder.Append("[Top='");
                xmlSearchStringBuilder.Append(sProfileInfo[4]);
                xmlSearchStringBuilder.Append("' or not(Top) or Top='Any']");
                xmlSearchStringBuilder.Append("[Bottom='");
                xmlSearchStringBuilder.Append(sProfileInfo[5]);
                xmlSearchStringBuilder.Append("' or not(Bottom) or Bottom='Any']]");

                xmlSearchStringBuilder.Append("[TopFlangeCutInfo[First='");
                xmlSearchStringBuilder.Append(sProfileInfo[6]);
                xmlSearchStringBuilder.Append("' or not(First) or First='Any']");
                xmlSearchStringBuilder.Append("[Left='");
                xmlSearchStringBuilder.Append(sProfileInfo[7]);
                xmlSearchStringBuilder.Append("' or not(Left) or Left='Any']");
                xmlSearchStringBuilder.Append("[Right='");
                xmlSearchStringBuilder.Append(sProfileInfo[8]);
                xmlSearchStringBuilder.Append("' or not(Right) or Right='Any']]");

                xmlSearchStringBuilder.Append("[BottomFlangeCutInfo[First='");
                xmlSearchStringBuilder.Append(sProfileInfo[9]);
                xmlSearchStringBuilder.Append("' or not(First) or First='Any']");
                xmlSearchStringBuilder.Append("[Left='");
                xmlSearchStringBuilder.Append(sProfileInfo[10]);
                xmlSearchStringBuilder.Append("' or not(Left) or Left='Any']");
                xmlSearchStringBuilder.Append("[Right='");
                xmlSearchStringBuilder.Append(sProfileInfo[11]);
                xmlSearchStringBuilder.Append("' or not(Right) or Right='Any']]");

                xmlSearchStringBuilder.Append("[SecondaryWebCutInfo[First='");
                xmlSearchStringBuilder.Append(sProfileInfo[12]);
                xmlSearchStringBuilder.Append("' or not(First) or First='Any']");
                xmlSearchStringBuilder.Append("[Top='");
                xmlSearchStringBuilder.Append(sProfileInfo[13]);
                xmlSearchStringBuilder.Append("' or not(Top) or Top='Any']");
                xmlSearchStringBuilder.Append("[Bottom='");
                xmlSearchStringBuilder.Append(sProfileInfo[14]);
                xmlSearchStringBuilder.Append("' or not(Bottom) or Bottom='Any']]");

                xmlSearchStringBuilder.Append("]");
                return xmlSearchStringBuilder.ToString();
			}
            catch (Exception excp)
            {
                throw new CmnUnexpectedException("Unexpected failure in preparing XML Search string. Error Message: " + excp.Message);
            }
		}

		private string GetCodelistName(Codelist sCodelist, string sValue)
		{
			string sReturn = sValue;	// return sValue if not a valid codelist value

			try
			{
				int listValue = 0;
				if (int.TryParse(sValue, out listValue) && sCodelist.sTableName.Length > 0)
				{
					if (StructHelper.IsValidCodeListValue(sCodelist.sTableName, sCodelist.sNamespace, listValue))
						sReturn = StructHelper.GetCodeListInfo(sCodelist.sTableName, sCodelist.sNamespace).GetCodelistItem(listValue).ShortDisplayName;
				}
			}
            catch (Exception excp)
            {
                throw new CmnUnexpectedException("Unexpected failure in obtaining code list short description. Error Message: " + excp.Message);
            }

			return sReturn;
		}

		private string GetCodelistName(Codelist sCodelist, int listValue)
		{
			string sReturn = "None";	// return "None" if not a valid codelist value

			try
			{
				if (StructHelper.IsValidCodeListValue(sCodelist.sTableName, sCodelist.sNamespace, listValue))
					sReturn = StructHelper.GetCodeListInfo(sCodelist.sTableName, sCodelist.sNamespace).GetCodelistItem(listValue).ShortDisplayName;
			}
            catch (Exception excp)
            {
                throw new CmnUnexpectedException("Unexpected failure in obtaining code list short description. Error Message: " + excp.Message);
            }

			return sReturn;
		}
		#endregion
	}
}
