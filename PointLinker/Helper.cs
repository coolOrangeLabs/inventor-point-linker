                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    using System;

//''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
// Copyright (c) Autodesk, Inc. All rights reserved 
// Written by Xiaodong Liang 2011 - ADN/Developer Technical Services
//
// This software is provided as is, without any warranty that 
// it will work. You choose to use this tool at your own risk.
// Neither Autodesk nor the author Xiaodong Liang can be taken 
// as responsible for any damage this tool can cause to your data. 
// Please always make a back up of your data prior to use this tool,
// as it will modify the documents involved in the feature transformation.
//''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

using System.IO;
using System.Diagnostics;
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    using System.Runtime.InteropServices;
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    using System.Windows.Forms;
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    using Inventor;
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    using Attribute = Inventor.Attribute;
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    using Excel = Microsoft.Office.Interop.Excel;


namespace PointLinker
{


	public class Helper
	{

		[System.Runtime.InteropServices.DllImport("user32", EntryPoint="FindWindowA", ExactSpelling=true, CharSet=System.Runtime.InteropServices.CharSet.Ansi, SetLastError=true)]
		private static extern int FindWindow(string lpClassName, string lpWindowName);

		[System.Runtime.InteropServices.DllImport("user32", EntryPoint="GetWindowThreadProcessId", ExactSpelling=true, CharSet=System.Runtime.InteropServices.CharSet.Ansi, SetLastError=true)]
		private static extern int GetWindowThreadProcessId(int hwnd, ref int lpdwProcessId);


		public enum PointOptionEnum: int
		{
			ePoints = 1,
			eLines = 2,
			eSpline = 3
		}

		public static Inventor.Application _InvApp;
		public static DataGridView DataGV;
		public static bool _isDrawingDoc;

		public static bool _isFromSketch;
		public static _Document _CurEditDoc;
		public static AttributeManager _AttributeMgr;
		public static TransientGeometry _TG;
		public static TransientObjects _TO;

		public static TransactionManager _TxnMgr;

		public static PointOptionEnum _ptOption = PointOptionEnum.ePoints;
		public static int _CurrentIndex;

		public static dynamic _curSketch;
		public static bool _is3DSketch;
		public static bool _isCenterPt;
		public static bool _isCenterLine;
		public static bool _isConstruction;



		public static void Initialize(Inventor.Application oInvApp, bool isFromSketch, ref DataGridView oDataGV)
		{

			_InvApp = oInvApp;


			_isFromSketch = isFromSketch;

			_CurEditDoc = _InvApp.ActiveEditDocument;

			_isDrawingDoc = _CurEditDoc is DrawingDocument;

			_curSketch = _InvApp.ActiveEditObject;

			_is3DSketch = (_curSketch) is Sketch3D;

			if (_isDrawingDoc)
			{
                _isCenterPt = ((ButtonDefinition)(_InvApp.CommandManager.ControlDefinitions["SketchHoleCenterCmd"])).Pressed;
				_isCenterLine = ((ButtonDefinition)(_InvApp.CommandManager.ControlDefinitions["SketchCenterlineCmd"])).Pressed;
			}
			else
			{
				if (_is3DSketch)
				{
					_isCenterPt = ((ButtonDefinition)(_InvApp.CommandManager.ControlDefinitions["Sketch3DHoleCenterCmd"])).Pressed;
					_isConstruction = ((ButtonDefinition)(_InvApp.CommandManager.ControlDefinitions["Sketch3DConstructionToggleCmd"])).Pressed;
				}
				else
				{
					_isCenterPt = ((ButtonDefinition)(_InvApp.CommandManager.ControlDefinitions["SketchHoleCenterCmd"])).Pressed;
					_isCenterLine = ((ButtonDefinition)(_InvApp.CommandManager.ControlDefinitions["SketchCenterlineCmd"])).Pressed;
					_isConstruction = ((ButtonDefinition)(_InvApp.CommandManager.ControlDefinitions["SketchConstructionCmd"])).Pressed;
				}
			}


			_AttributeMgr = _CurEditDoc.AttributeManager;

			_TG = _InvApp.TransientGeometry;

			_TO = _InvApp.TransientObjects;

			_TxnMgr = _InvApp.TransactionManager;



			DataGV = oDataGV;

		}


		/// <summary>
		/// get sketch from guid
		/// </summary>
		/// <param name="guid"></param>
		/// <returns></returns>
		/// <remarks></remarks>

		public static object GetSketchFromGuid(string guid)
		{
			object tempGetSketchFromGuid = null;

			tempGetSketchFromGuid = null;

			ObjectCollection oTempCollection = null;

			try
			{
				// _ADNPlugin_LinkExcelPoints_Sketch
				oTempCollection = _AttributeMgr.FindObjects(Properties.Settings.Default.sketchAttSetName, guid);

				if (oTempCollection.Count != 0)
				{

					object oSketch = null;

					if (IsDef(oTempCollection[1], ref oSketch))
					{

						// to get sketch from titleblockdefinition
						// borderdefinition and  _
						// SketchedSymbolDefinition()

						tempGetSketchFromGuid = oSketch;
					}
					else
					{
						// a common sketch
						tempGetSketchFromGuid = oTempCollection[1];
					}

				}
				else
				{
					// error!
					return tempGetSketchFromGuid;
				}

			}
			catch (Exception ex)
			{
				return tempGetSketchFromGuid;
			}


			return tempGetSketchFromGuid;
		}



		/// <summary>
		/// update or resolve the selected row
		/// </summary>
		/// <param name="oNewFileName"></param>
		/// <remarks></remarks>

		public static bool UpdateOrResolve()
		{
			return UpdateOrResolve("");
		}

//INSTANT C# NOTE: Overloaded method(s) are created above to convert the following method having optional parameters:
//ORIGINAL LINE: Public Shared Function UpdateOrResolve(Optional ByVal oNewFileName As String = "") As Boolean
		public static bool UpdateOrResolve(string oNewFileName)
		{
			bool tempUpdateOrResolve = false;

			tempUpdateOrResolve = false;

			DataGridViewRow oCurRow = DataGV.CurrentRow;

			// get guid
			string oGuid = oCurRow.Cells["Guid"].Value.ToString();

			// get point type
			PointOptionEnum oPtOp = 0;

			switch (oCurRow.Cells["PointType"].Value.ToString())
			{
				case "Points":
					oPtOp = PointOptionEnum.ePoints;
					break;
				case "Lines":
					oPtOp = PointOptionEnum.eLines;
					break;
				case "Spline":
					oPtOp = PointOptionEnum.eSpline;
					break;
			}

			// get the sketch from guid
			dynamic oCurSketch = GetSketchFromGuid(oGuid);

			if (oCurSketch == null)
			{

				MessageBox.Show(Properties.Settings.Default.generalError, Properties.Settings.Default.errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);

				return tempUpdateOrResolve;
			}

			bool otempIs3dSketch = false;
			//activate the sketch
			bool oFlagEditSketch = false;
			// this is a part or assembly document
			dynamic oCurrentEditObj = _InvApp.ActiveEditObject;
			if (oCurrentEditObj != oCurSketch)
			{
				if (oCurrentEditObj is Sketch | oCurrentEditObj is Sketch3D)
				{
					oCurrentEditObj.ExitEdit();
				}

				oCurSketch.Edit();
				oFlagEditSketch = true;


			}

			if (oCurSketch is Sketch3D)
			{
				otempIs3dSketch = true;
			}


			// file name used in this function

			string oFileName = null;

			if (oNewFileName == "")
			{
				// update, get from column string
				oFileName = oCurRow.Cells["FileName"].Value.ToString();
			}
			else
			{
				// resolve, new file
				oFileName = oNewFileName;
			}



			// previous points count
			int oPreviousPtCount = Convert.ToInt16(oCurRow.Cells["PointsCount"].Value);

			// new points from excel file
			ObjectCollection oNewPointDataColl = _TO.CreateObjectCollection();

			// read the file data
			if (ReadExcel(oFileName, ref oNewPointDataColl, oPtOp, otempIs3dSketch) == false)
			{
				return tempUpdateOrResolve;
			}

			// current points count
			// _ADNPlugin_LinkExcelPoints_Entity_Points_ID
			ObjectCollection oCurSketchObjColl = _AttributeMgr.FindObjects(oGuid, Properties.Settings.Default.sketchPtAttName);

			// messages to warn the user. 
			if (oPreviousPtCount == oCurSketchObjColl.Count && oPreviousPtCount == oNewPointDataColl.Count)
			{

				// The points will be updated one by one. 
				//  The lines and spline will be re-created.

				MessageBox.Show(Properties.Settings.Default.samePtCount, Properties.Settings.Default.messageCaption, MessageBoxButtons.OK, MessageBoxIcon.Information);

			}
			else
			{

				// Some previous points may have been missed, 
				// or the previous points count is different 
				// to the new file. 
				// Some points may be added and some will break
				//the link with the excel file. 
				// The lines or spline will be re-created.

				MessageBox.Show(Properties.Settings.Default.difPtCount, Properties.Settings.Default.messageCaption, MessageBoxButtons.OK, MessageBoxIcon.Information);


			}


			// Start a regular transaction
			Transaction oNewTrans = null;
			oNewTrans = _TxnMgr.StartTransaction(_CurEditDoc, "UpdatePoints");



			dynamic oSketchPoint = null;
			AttributeSet oPtAttSet = null;
			bool oSomeSpecialPoints = false;
			ObjectCollection oSketchPointColl = _InvApp.TransientObjects.CreateObjectCollection();

			ObjectCollection oTempColl = null;
			int oIndex = 0;

			for (oIndex = 1; oIndex <= oNewPointDataColl.Count; oIndex++)
			{

				try
				{
					oTempColl = _AttributeMgr.FindObjects(oGuid, Properties.Settings.Default.sketchPtAttName, oIndex);

					if (oTempColl.Count == 0)
					{
						//the point missed! add it
						if (_is3DSketch)
						{
							oSketchPoint = oCurSketch.SketchPoints3D.Add(oNewPointDataColl[oIndex]);
						}
						else
						{
							oSketchPoint = oCurSketch.SketchPoints.Add(oNewPointDataColl[oIndex]);
						}

						oSketchPoint.HoleCenter = _isCenterPt;

						//add attribute
						oPtAttSet = oSketchPoint.AttributeSets.Add(oGuid);

						oPtAttSet.Add(Properties.Settings.Default.sketchPtAttName, ValueTypeEnum.kIntegerType, oIndex);

						oSketchPointColl.Add(oSketchPoint);
					}
					else
					{

						//the point exists!
						oTempColl[1].MoveTo(oNewPointDataColl[oIndex]);
						oSketchPointColl.Add(oTempColl[1]);

					}

				}
				catch (Exception ex)
				{
					// error when update or resolve points

					MessageBox.Show(Properties.Settings.Default.errorUpdatePt, Properties.Settings.Default.errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);

					oNewTrans.Abort();
					return tempUpdateOrResolve;

				}
			}

			// break link of additional points with the file

			if (oPreviousPtCount > oNewPointDataColl.Count)
			{

				//break no use point
				for (oIndex = oNewPointDataColl.Count + 1; oIndex <= oPreviousPtCount; oIndex++)
				{

					try
					{
						oTempColl = _AttributeMgr.FindObjects(oGuid, Properties.Settings.Default.sketchPtAttName, oIndex);

						if (oTempColl.Count == 0)
						{

							// the point does not exist. do nothing

						}
						else
						{
							//the point exists! remove its attribute
							oSketchPoint = oTempColl[1];

							if (oSketchPoint.AttributeSets.NameIsUsed(oGuid))
							{
								oPtAttSet = oSketchPoint.AttributeSets(oGuid);
								oPtAttSet.Delete();
							}

						}
					}
					catch (Exception ex)
					{

						// error when update or resolve points!
						MessageBox.Show(Properties.Settings.Default.errorUpdatePt, Properties.Settings.Default.errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);

						oNewTrans.Abort();
						return tempUpdateOrResolve;
					}
				}

			}



			dynamic oSketchLine = null;
			dynamic oSketchSpline = null;

//INSTANT C# NOTE: Commented this declaration since looping variables in 'foreach' loops are declared in the 'foreach' header in C#:
//			dynamic oTempObj = null;

			//recreate the lines or spline

			try
			{
				switch (oPtOp)
				{
					case PointOptionEnum.ePoints:
					break;
					case PointOptionEnum.eLines:

						// _ADNPlugin_LinkExcelPoints_Entity_Lines_ID
						oTempColl = _AttributeMgr.FindObjects(oGuid, Properties.Settings.Default.sketchLineAttName);

						// delete the old lines
						foreach (dynamic oTempObj in oTempColl)
						{
							oTempObj.Delete();
						}

						// create the new lines
						for (var oPtIndex = 1; oPtIndex < oSketchPointColl.Count; oPtIndex++)
						{

							if (otempIs3dSketch)
							{
								// 3d lines
								oSketchLine = oCurSketch.SketchLines3D.AddByTwoPoints(oSketchPointColl[oPtIndex], oSketchPointColl[oPtIndex + 1]);
							}
							else
							{
								// 2d lines
								oSketchLine = oCurSketch.SketchLines.AddByTwoPoints(oSketchPointColl[oPtIndex], oSketchPointColl[oPtIndex + 1]);
								oSketchLine.Centerline = _isCenterLine;
							}

							oSketchLine.Construction = _isConstruction;
							// add attribute to the lines
							oPtAttSet = oSketchLine.AttributeSets.Add(oGuid);
							// _ADNPlugin_LinkExcelPoints_Entity_Lines_ID
							oPtAttSet.Add(Properties.Settings.Default.sketchLineAttName, ValueTypeEnum.kIntegerType, oPtIndex);

						}

						break;
					case PointOptionEnum.eSpline:

						// _ADNPlugin_LinkExcelPoints_Entity_Spline_ID
						oTempColl = _AttributeMgr.FindObjects(oGuid, Properties.Settings.Default.sketchSLAttName);

						if (oTempColl.Count != 0)
						{
							oTempColl[1].Delete();
						}

						if (otempIs3dSketch)
						{
							// 3d spline
							oSketchSpline = oCurSketch.SketchSplines3D.Add(oSketchPointColl, SplineFitMethodEnum.kSmoothSplineFit);
						}
						else
						{
							// 2d spline
							oSketchSpline = oCurSketch.SketchSplines.Add(oSketchPointColl, SplineFitMethodEnum.kSmoothSplineFit);


						}

						oSketchSpline.Construction = _isConstruction;

						// add attribute to the spline
						oPtAttSet = oSketchSpline.AttributeSets.Add(oGuid);
						// _ADNPlugin_LinkExcelPoints_Entity_Spline_ID
						oPtAttSet.Add(Properties.Settings.Default.sketchSLAttName, ValueTypeEnum.kIntegerType, 1); //always 1

						break;
				}
			}
			catch (Exception ex)
			{
				// error when update or resolve points!

				MessageBox.Show(Properties.Settings.Default.errorUpdatePt, Properties.Settings.Default.errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);

				oNewTrans.Abort();
				return tempUpdateOrResolve;
			}



			//update attribute
			string oValueString = "";
			CreateOrUpdateSketchAtt(ref oCurSketch, oFileName, oPtOp, oNewPointDataColl.Count, ref oGuid, ref oValueString);

			// submit transaction

			oNewTrans.End();

			_CurrentIndex = oCurRow.Index;

			_CurEditDoc.Update();

			// update strings of row
			DataGV.CurrentRow.Cells["FileName"].Value = oFileName;
			DataGV.CurrentRow.Cells["PointsCount"].Value = oNewPointDataColl.Count.ToString();
			DataGV.CurrentRow.Cells["Status"].Value = "Up-to-date";

			DataGV.Refresh();

			try
			{
				if (oFlagEditSketch)
				{
					oCurSketch.ExitEdit();
				}
			}
			catch (Exception ex)
			{

			}


			return true;

		}

		public static void WhenClose()
		{
			bool oAskIfDelete = false;

			foreach (DataGridViewRow oEachRow in DataGV.Rows)
			{

				ObjectCollection oTempColl = null;
				string oGuid = oEachRow.Cells["Guid"].Value.ToString();

				// _ADNPlugin_LinkExcelPoints_Entity*
				oTempColl = _AttributeMgr.FindObjects(oGuid, Properties.Settings.Default.findSketchEntity);

				if (oTempColl.Count == 0)
				{

					if (oAskIfDelete == false)
					{

						if (MessageBox.Show(Properties.Settings.Default.whenClose, Properties.Settings.Default.messageCaption, MessageBoxButtons.YesNo) == DialogResult.Yes)
						{

							oAskIfDelete = true;
						}
						else
						{
							break;
						}
					}


					if (oAskIfDelete)
					{
						AttributesEnumerator oAtts = _AttributeMgr.FindAttributes(Properties.Settings.Default.sketchAttSetName, oGuid);

						if (oAtts.Count != 0)
						{
							oAtts[1].Delete();
						}

					}
				}
			}

		}

		public static void ZoomWindowAndHilight(string guid)
		{
            
			// _ADNPlugin_LinkExcelPoints_Entity*
			ObjectCollection oObjCollection = _AttributeMgr.FindObjects(guid, Properties.Settings.Default.findSketchEntity);


			_CurEditDoc.SelectSet.Clear();
			foreach (object oEachEnt in oObjCollection)
			{
				_CurEditDoc.SelectSet.Select(oEachEnt);
			}

			Camera oC = _InvApp.ActiveView.Camera;
			oC.Fit();
			oC.Apply();
		}

		/// <summary>
		/// produce strings for row
		/// </summary>
		/// <param name="oSketch"></param>
		/// <param name="attName"></param>
		/// <param name="attValue"></param>
		/// <param name="rowStrs"></param>
		/// <remarks></remarks>

		public static void produceRowStrings(dynamic oSketch, string attName, string attValue, ref string[] rowStrs)
		{

			// parse the strings from attribute
			// format of att: "c:\temp\book1.xls|1|100|2011-3-21:14:30:24"
			//  FileName | Point Type  | Points Count | Last Saved Time
			string[] attParts = attValue.Split(new char[] {'|'});

			// file name

			rowStrs[0] = attParts[0];

			// point type
		    var pointOptionsEnum = (PointOptionEnum)Enum.Parse(typeof (PointOptionEnum), attParts[1]);
            switch (pointOptionsEnum)
			{
				case PointOptionEnum.ePoints:
					rowStrs[1] = "Points";
					break;
				case PointOptionEnum.eLines:
					rowStrs[1] = "Lines";
					break;
				case PointOptionEnum.eSpline:
					rowStrs[1] = "Spline";
					break;
			}

			// Status

			rowStrs[2] = "";
			if (System.IO.File.Exists(attParts[0]) == true)
			{
				DateTime oCurrentDate = System.IO.File.GetLastWriteTimeUtc(attParts[0]);
				if (attParts[3] != oCurrentDate.ToString())
				{

					// file out of date
					rowStrs[2] = "Out-of-date";

				}
			}
			else
			{

				// file not exist
				rowStrs[2] = "File not exist";
			}



			var oPointStatusStr = "";
			// check the points status
			long oCurrentPointsCount = _AttributeMgr.FindObjects(attName, Properties.Settings.Default.sketchPtAttName).Count;

			if (oCurrentPointsCount == 0)
			{

				// no points exist with this row
				oPointStatusStr = "No points";
			}
			else
			{
				if (oCurrentPointsCount != Convert.ToInt64(attParts[2]))
				{

					// some points missed
					oPointStatusStr = "Points missed";

				}
			}

			if (oPointStatusStr != "")
			{

				if (rowStrs[2] != "")
				{
					rowStrs[2] += ", ";
				}
				rowStrs[2] += oPointStatusStr;

			}

			var oEntityStatusStr = "";
			// check if any lines missed
			long oCurrentEntityCount = 0;
            if (pointOptionsEnum == PointOptionEnum.eLines)
			{
				oCurrentEntityCount = _AttributeMgr.FindObjects(attName, Properties.Settings.Default.attSearchLines).Count;

				if (oCurrentEntityCount == 0)
				{
					oEntityStatusStr = "No lines";
				}
				else
				{
                    if (oCurrentEntityCount != Convert.ToInt32(attParts[2]) - 1)
					{
						// some lines missed
						oEntityStatusStr = "Lines missed";
					}
				}

				// check if spline missed
			}
			else if (attParts[1] == ((int)PointOptionEnum.eSpline).ToString())
			{
				oCurrentEntityCount = _AttributeMgr.FindObjects(attName, Properties.Settings.Default.attSeatchSpline).Count;

				if (oCurrentEntityCount == 0)
				{
					oEntityStatusStr = "No spline";
				}

			}

			if (oEntityStatusStr != "")
			{

				if (rowStrs[2] != "")
				{
					rowStrs[2] += ", ";
				}

				rowStrs[2] += oEntityStatusStr;

			}

			// all is normal 

			if (rowStrs[2] == "")
			{
				rowStrs[2] = "Up-to-date";
			}

			// from Tools Tab in Part or Assembly
			if (!_isFromSketch)
			{

				// Sketch Name
				rowStrs[3] = oSketch.Name;

				// Sketch Type

				if (oSketch is Sketch3D)
				{
					rowStrs[4] = "3D";
				}
				else
				{
					rowStrs[4] = "2D";
				}

			}

			// guid
			rowStrs[5] = attName;

			// Current point count
			rowStrs[6] = attParts[2];

			//last save time
			rowStrs[7] = attParts[3];

		}


		/// <summary>
		/// check if the object with "_ADNPlugin_LinkExcelPoints_Sketch"
		/// is TitleBlockDefinition, BorderDefinition, or
		/// SketchedSymbolDefinition
		/// </summary>
		/// <param name="oObj"></param>
		/// <param name="oSketch"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		private static bool IsDef(dynamic oObj, ref object oSketch)
		{
			bool tempIsDef = false;
			tempIsDef = false;

			if (oObj is TitleBlockDefinition | oObj is BorderDefinition | oObj is SketchedSymbolDefinition)
			{

				oSketch = oObj.Sketch;
				tempIsDef = true;
			}
			return tempIsDef;
		}

		/// <summary>
		/// produce rows for datagridview
		/// </summary>
		/// <remarks></remarks>
		public static void ProduceRows()
		{

			DataGV.Rows.Clear();
            
			ObjectCollection oObjCollection = _AttributeMgr.FindObjects(Properties.Settings.Default.sketchAttSetName);

			AttributeSet oEachAttSet = null;

			object oActiveEditObj = _InvApp.ActiveEditObject;

			foreach (dynamic oEachObj in oObjCollection)
			{

				bool oToAdd = true;
//INSTANT C# TODO TASK: There is no C# equivalent to VB's implicit 'once only' variable initialization within loops:
				dynamic oDrawingDefSketch = null;

				if (_isDrawingDoc)
				{

					if (IsDef(oEachObj, ref oDrawingDefSketch))
					{

						if (oDrawingDefSketch == oActiveEditObj)
						{

							oToAdd = true;

							if (oDrawingDefSketch.AttributeSets.NameIsUsed(Properties.Settings.Default.sketchAttSetName))
							{

								oEachAttSet = oDrawingDefSketch.AttributeSets(Properties.Settings.Default.sketchAttSetName);

							}
						}
					}
				}
				else
				{
					if ((oEachObj) is Inventor.Sketch | oEachObj is Inventor.Sketch3D)
					{

						if (_isFromSketch && !(oEachObj == _InvApp.ActiveEditObject))
						{
							oToAdd = false;
						}
						else
						{
							if (oEachObj.AttributeSets.NameIsUsed(Properties.Settings.Default.sketchAttSetName))
							{

								oEachAttSet = oEachObj.AttributeSets(Properties.Settings.Default.sketchAttSetName);
							}
						}
					}

				}

				if (oToAdd && oEachAttSet != null)
				{

					string[] oParamArray = new string[8];
					foreach (Attribute oEachAtt in oEachAttSet)
					{

						produceRowStrings(oEachObj, oEachAtt.Name, oEachAtt.Value, ref oParamArray);

						DataGV.Rows.Add(oParamArray);

					}
				}

			}

		}



		/// <summary>
		/// Create AttributeSet and Attribute for a sketch, or
		/// Update the attribute
		/// AttSet Name: _ADNPlugin_LinkExcelPoints_SketchAtt 
		/// each att is one set of points from one file
		/// att name: guid  _9245fe4a-d402-451c-b9ed-9c1a04247482 
		/// value: c:\temp\book1.xls|1|100|2011-3-21:14:30:24
		///  FileName | Points Type | Points Count | Last Saved Time 
		/// </summary>
		/// <param name="oSketch"></param>
		/// <param name="oFileName"></param>
		/// <param name="oPointType"></param>
		/// <param name="oPointsCount"></param>
		/// <param name="oGuid"></param>
		/// <param name="oValueString"></param>
		/// <remarks></remarks>

        private static void CreateOrUpdateSketchAtt(ref dynamic oSketch, string oFileName, PointOptionEnum oPointType, long oPointsCount, ref string oGuid, ref string oValueString)
		{

			bool isCreate = false;
			AttributeSets oAttSets = oSketch.AttributeSets;

			AttributeSet oAttSet = null;
			if (oAttSets.NameIsUsed[Properties.Settings.Default.sketchAttSetName])
			{

				oAttSet = oAttSets[Properties.Settings.Default.sketchAttSetName];

			}
			else
			{
				oAttSet = oAttSets.Add(Properties.Settings.Default.sketchAttSetName);
			}

			Attribute oAtt = null;


			oValueString = oFileName + "|";
			oValueString += Convert.ToInt16(oPointType).ToString();
			oValueString += "|";
			oValueString += oPointsCount.ToString();
			oValueString += "|";
			oValueString += System.IO.File.GetLastWriteTimeUtc(oFileName);

			if (oGuid == "")
			{

				oGuid = System.Guid.NewGuid().ToString();

				oGuid = "_" + oGuid;

				oAtt = oAttSet.Add(oGuid, ValueTypeEnum.kStringType, oValueString);

			}
			else
			{
				if (oAttSet.NameIsUsed[oGuid])
				{
					oAtt = oAttSet[oGuid];
					oAtt.Value = oValueString;
				}
				else
				{

					MessageBox.Show(Properties.Settings.Default.generalError, Properties.Settings.Default.errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);

				}

			}

		}


		/// <summary>
		/// Create New Points
		/// </summary>
		/// <param name="oFileName"></param>
		/// <remarks></remarks>
		public static bool ImportNewPoints(string oFileName)
		{
			bool tempImportNewPoints = false;


			tempImportNewPoints = false;


			// the points collection from file
			ObjectCollection oPtDataCollc = _InvApp.TransientObjects.CreateObjectCollection();

			// read data from excel

			if (ReadExcel(oFileName, ref oPtDataCollc, _ptOption, _is3DSketch) == false)
			{

				return tempImportNewPoints;

			}


			//add attribute
			//only one attset for one sketch:  
			// Name: "_ADNPlugin_LinkExcelPoints_Sketch"
			//value

			// Start a regular transaction
			Transaction oNewTrans = null;
			oNewTrans = _TxnMgr.StartTransaction(_CurEditDoc, "ImportPoints");

			try
			{

				// create new AttributeSet and Attribute for the sketch
				string oNewGuid = "";
				string oValueString = "";
				CreateOrUpdateSketchAtt(ref _curSketch, oFileName, _ptOption, oPtDataCollc.Count, ref oNewGuid, ref oValueString);


				// Sketch points collection, 
				//for create lines and spline

				ObjectCollection oSketchPointColl = _TO.CreateObjectCollection();

				AttributeSet oPtAttSet = null;
				dynamic oSketchPoint = null;
				dynamic oSketchLine = null;
				dynamic oSketchSpline = null;

				// add sketch points

				int oPtIndex = 0;
				for (oPtIndex = 1; oPtIndex <= oPtDataCollc.Count; oPtIndex++)
				{

					if (_is3DSketch)
					{
						// add sketch point 3d
						oSketchPoint = _curSketch.SketchPoints3D.Add(oPtDataCollc[oPtIndex]);
					}
					else
					{
						// add sketch point 2d
						oSketchPoint = _curSketch.SketchPoints.Add(oPtDataCollc[oPtIndex]);
					}

					oSketchPoint.HoleCenter = _isCenterPt;

					// add attribute to sketch point

					oPtAttSet = oSketchPoint.AttributeSets.Add(oNewGuid);
					oPtAttSet.Add(Properties.Settings.Default.sketchPtAttName, ValueTypeEnum.kIntegerType, oPtIndex);

					oSketchPointColl.Add(oSketchPoint);
				}

				// Create lines or spline
				
				switch (_ptOption)
				{
					case PointOptionEnum.ePoints:
						//has done
					break;
					case PointOptionEnum.eLines:
						var autoBendWithLineCreation = _InvApp.Sketch3DOptions.AutoBendWithLineCreation; //remember state of AutoBendWithLineCreation
						for (oPtIndex = 1; oPtIndex < oSketchPointColl.Count; oPtIndex++)
						{

							if (_is3DSketch)
							{
								// sketch line 3d
								oSketchLine = _curSketch.SketchLines3D.AddByTwoPoints(oSketchPointColl[oPtIndex], oSketchPointColl[oPtIndex + 1]);
							}
							else
							{
								// sketch line 2d
								oSketchLine = _curSketch.SketchLines.AddByTwoPoints(oSketchPointColl[oPtIndex], oSketchPointColl[oPtIndex + 1]);

								oSketchLine.Centerline = _isCenterLine;
							}

							oSketchLine.Construction = _isConstruction;

							// add attribute to sketch line

							oPtAttSet = oSketchLine.AttributeSets.Add(oNewGuid);
							oPtAttSet.Add(Properties.Settings.Default.sketchLineAttName, ValueTypeEnum.kIntegerType, oPtIndex);

						}
						_InvApp.Sketch3DOptions.AutoBendWithLineCreation = autoBendWithLineCreation; //apply back state of AutoBendWithLineCreation
						break;
					case PointOptionEnum.eSpline:

						if (_is3DSketch)
						{
							// add sketch spline 3d
							oSketchSpline = _curSketch.SketchSplines3D.Add(oSketchPointColl, SplineFitMethodEnum.kSmoothSplineFit);


						}
						else
						{
							// add sketch spline 2d
							oSketchSpline = _curSketch.SketchSplines.Add(oSketchPointColl, SplineFitMethodEnum.kSmoothSplineFit);

						}

						oSketchSpline.Construction = _isConstruction;

						// add attribute to sketch spline

						oPtAttSet = oSketchSpline.AttributeSets.Add(oNewGuid);
						oPtAttSet.Add(Properties.Settings.Default.sketchSLAttName, ValueTypeEnum.kIntegerType, 1); //always 1

						break;
				}


				// submit transaction
				oNewTrans.End();

				// update view
				_CurEditDoc.Update();


				// add a row

				string[] oParamArray = new string[8];
				produceRowStrings(_curSketch, oNewGuid, oValueString, ref oParamArray);

				var oNewRowIndex = DataGV.Rows.Add(oParamArray);

				// the new row is selected

				DataGV.CurrentCell = DataGV.Rows[oNewRowIndex].Cells[0];

				_CurrentIndex = oNewRowIndex;

				tempImportNewPoints = true;

				// zoom to the object or hightlight
				ZoomWindowAndHilight(oNewGuid);


			}
			catch (Exception ex)
			{

				MessageBox.Show(Properties.Settings.Default.generalError, Properties.Settings.Default.errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);


				// abort transaction 
				if (oNewTrans.State != TransactionStateEnum.kDoneState)
				{

					oNewTrans.Abort();

				}

			}


			return tempImportNewPoints;
		}

		/// <summary>
		/// Read data from excel file
		/// 
		/// Criteria :
		/// 1. only import the first sheet
		/// 2. ignore any row until the first cell is 
		///     number (valid row)
		/// 3. start read the continuous rows till the first cell 
		///    is nonnumeric/empty
		/// 4. read 2 columns for 2d sketch, 
		///    and 3 colums for 3d sketch
		/// 5. error scenarios:  
		///    1) in 2d sketch, more than 2 columns. 
		///       ask user if continue:
		///       yes: just read the first and second columns
		///       no:  stop reading
		///    2) count of columns is less than 2 in 2d sketch, or
		///       less than 3 in 3d sketch. pop out a dialog that the
		///       data does not meet requirement formatting 
		///    3) if any cell in valid row is nonnumeric,
		///       stop can cancel reading.
		///       pop out a dialog that the data does not meet
		///       requirement formatting
		///    
		/// </summary>
		/// <param name="fullFileName"></param>
		/// <param name="oPtDataCollc"></param>
		/// <param name="ptOp"></param>
		/// <param name="is3D"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		public static bool ReadExcel(string fullFileName, ref ObjectCollection oPtDataCollc, PointOptionEnum ptOp, bool is3D)
		{
			bool tempReadExcel = false;


			tempReadExcel = false;

			//'*************
			// get addin location
		    System.Reflection.Assembly Asm
		        = System.Reflection.Assembly.GetExecutingAssembly();

			//' create option file
            //string oOptionFileName = System.IO.Path.GetDirectoryName(Asm.Location) + @"\log.txt";

            //if (System.IO.File.Exists(oOptionFileName))
            //    System.IO.File.Delete(oOptionFileName);

            //StreamWriter sw = System.IO.File.CreateText(oOptionFileName);
			//'***************



			System.Globalization.CultureInfo oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;

			System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");


			UnitsOfMeasure _UOM = _InvApp.UnitsOfMeasure;

			// Excel application
			Excel.Application excelApp = null;

			// if the Excel instance is created by this code,
			// need to dispose after reading.

			bool oNewExcel = false;

			try
			{
				// get active Excel process
				excelApp = (Excel.Application) Marshal.GetActiveObject("Excel.Application");
                //sw.WriteLine("succeed to get active excel");

			}
			catch (Exception ex)
			{
                //sw.WriteLine("fail to get active excel");
                //sw.WriteLine(ex.ToString());

			}

			// create an Excel process
			string oMyExcelCaption = "ADNPlugin_LinkExcelPoint";

			if (excelApp == null)
			{
				try
				{
					excelApp = new Excel.Application();
					excelApp.Caption = oMyExcelCaption;
					oNewExcel = true;

                    //sw.WriteLine("succeed to create active excel");

				}
				catch (Exception ex1)
				{

                    //sw.WriteLine("fail to get create excel");
                    //sw.WriteLine(ex1.ToString());

					System.Threading.Thread.CurrentThread.CurrentCulture = oldCI;

                    //sw.Flush();
                    //sw.Close();

					MessageBox.Show(Properties.Settings.Default.failToGetExcel, Properties.Settings.Default.errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);


					return tempReadExcel;

				}
			}

			// check if the book has been opened
			bool hasOpened = false;

			Excel.Workbook excelBook = null;


			try
			{
				// open file
				foreach (Excel.Workbook oEachBook in excelApp.Workbooks)
				{
					if (oEachBook.FullName == fullFileName)
					{
						excelBook = oEachBook;
						hasOpened = true;
					}
				}

				if (excelBook == null)
				{
					excelBook = excelApp.Workbooks.Open(fullFileName);
				}

                //sw.WriteLine("succeed to  open  book");

			}
			catch (Exception ex)
			{

                //sw.WriteLine("fail to open  book");
                //sw.WriteLine(ex.ToString());
                //System.Threading.Thread.CurrentThread.CurrentCulture = oldCI;
                //sw.Flush();
                //sw.Close();

				MessageBox.Show(Properties.Settings.Default.failToOpenWorkbook, Properties.Settings.Default.errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);


				return tempReadExcel;
			}


			//ready to read
			try
			{

				//' read first sheet
				Excel.Worksheet excelSheet = excelBook.Sheets[1];

				// used range
				Excel.Range rng = excelSheet.UsedRange;


				// if no error with reading
				bool isSuccessfulRead = true;

				// if count of columns <2 in 2d sketch or
				// <3 in 3d sketch
				if ((is3D == false && rng.Columns.Count < 2) | (is3D == true && rng.Columns.Count < 3))
				{

					MessageBox.Show(Properties.Settings.Default.notMeetFormat);

					isSuccessfulRead = false;

				}

				// check unit

				UnitsTypeEnum oFromUnit = _CurEditDoc.UnitsOfMeasure.LengthUnits;

				UnitsTypeEnum oOutUnit = UnitsTypeEnum.kCentimeterLengthUnits;


				// first cell for unit
				string oFirstCell = rng[1, 1].Value;

				if (string.IsNullOrEmpty(oFirstCell))
				{

					if (MessageBox.Show(Properties.Settings.Default.noUnits, Properties.Settings.Default.errorCaption, MessageBoxButtons.YesNo, MessageBoxIcon.Error) == DialogResult.No)
					{

						isSuccessfulRead = false;
					}

				}
				else // has string
				{

					try
					{
                        oFromUnit = _UOM.GetTypeFromString(oFirstCell);
					}
					catch (Exception ex)
					{
                        //sw.Write(ex);
						// may fail unit specified by the string does not 
						//exist in the constant list. will use the 
						// document units
					}

				}


				if (isSuccessfulRead)
				{
					// if start read
					bool oStartRead = false;

					// user decide to read data for 
					// 2d sketch if count columns >2
					bool hasUserDecide = false;

					// for detect adjoining points
					double lastPtX = double.NaN;
					double lastPtY = double.NaN;
					double lastPtZ = double.NaN;


					// read each row
					for (var i = 1; i <= rng.Rows.Count; i++)
					{

						var oColumn_1 = rng[i, 1].Value;

						// check if start read
						if (oStartRead == false)
						{

							if ((oColumn_1!=null) && !(string.IsNullOrEmpty(oColumn_1.ToString())) && SimulateIsNumeric.IsNumeric(oColumn_1.ToString()))
							{

								//not empty and is numeric, start read

								oStartRead = true;

							}
						}


						if (oStartRead)
						{
                            if(oColumn_1 == null)
                                break;

							if (string.IsNullOrEmpty(oColumn_1.ToString()) || !(SimulateIsNumeric.IsNumeric(oColumn_1.ToString())))
							{
								//stop reading
								break;
							}

							// 3d sketch

							var oColumn_2 = rng[i, 2].Value;
							var oColumn_3 = rng[i, 3].Value;

							if (is3D)
							{

								// try to read

								if (string.IsNullOrEmpty(oColumn_2.ToString()) || !(SimulateIsNumeric.IsNumeric(oColumn_2)) | string.IsNullOrEmpty(oColumn_3.ToString()) || !(SimulateIsNumeric.IsNumeric(oColumn_3)))
								{

									if (MessageBox.Show(Properties.Settings.Default.notMeetFormat, Properties.Settings.Default.errorCaption, MessageBoxButtons.YesNo, MessageBoxIcon.Error) == DialogResult.No)
									{

									}
									isSuccessfulRead = false;
									break;
								}

								// add 3d point
								double oX = _UOM.ConvertUnits(Convert.ToDouble(oColumn_1), oFromUnit, oOutUnit);

								double oY = _UOM.ConvertUnits(Convert.ToDouble(oColumn_2), oFromUnit, oOutUnit);
								double oZ = _UOM.ConvertUnits(Convert.ToDouble(oColumn_3), oFromUnit, oOutUnit);

								oPtDataCollc.Add(_TG.CreatePoint(oX, oY, oZ));


								if (oX == lastPtX && oY == lastPtY && oZ == lastPtZ)
								{

									if (ptOp == PointOptionEnum.eLines || ptOp == PointOptionEnum.eSpline)
									{

										MessageBox.Show(Properties.Settings.Default.adjoiningPts, Properties.Settings.Default.errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);

										isSuccessfulRead = false;
										break;

									}
								}

								lastPtX = oX;
								lastPtY = oY;
								lastPtZ = oZ;

							}
							else // 2d sketch
							{

								// try to read
								if (string.IsNullOrEmpty(oColumn_2.ToString()) || !(SimulateIsNumeric.IsNumeric(oColumn_2)))
								{

									MessageBox.Show(Properties.Settings.Default.adjoiningPts, Properties.Settings.Default.errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);

									isSuccessfulRead = false;
									break;

								}

								//check if the 3rd column has value

								if (oColumn_3 != null && !string.IsNullOrEmpty(oColumn_3.ToString()) && !hasUserDecide)
								{

									if (MessageBox.Show(Properties.Settings.Default.hasMoreColumn, Properties.Settings.Default.messageCaption, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
									{

										//will just fouse the 1 and 2 column

										hasUserDecide = true;
									}
									else
									{

										isSuccessfulRead = false;
										break;

									}

								}

								// add 2d point
								double oX = _UOM.ConvertUnits(Convert.ToDouble(oColumn_1), oFromUnit, oOutUnit);
								double oY = _UOM.ConvertUnits(Convert.ToDouble(oColumn_2), oFromUnit, oOutUnit);
								oPtDataCollc.Add(_TG.CreatePoint2d(oX, oY));


								if (oX == lastPtX && oY == lastPtY)
								{

									if (ptOp == PointOptionEnum.eLines || ptOp == PointOptionEnum.eSpline)
									{

										MessageBox.Show(Properties.Settings.Default.adjoiningPts, Properties.Settings.Default.errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);

										isSuccessfulRead = false;
										break;

									}
								}

								lastPtX = oX;
								lastPtY = oY;


							}

						}

					}

					if (isSuccessfulRead)
					{


                        //sw.WriteLine("SuccessfulRead");


						// points count < 2 for Lines
						if (ptOp == PointOptionEnum.eLines && oPtDataCollc.Count < 2)
						{

							MessageBox.Show(Properties.Settings.Default.lessPtForLines, Properties.Settings.Default.errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);

							isSuccessfulRead = false;

						}

						if (ptOp == PointOptionEnum.eSpline && oPtDataCollc.Count < 2)
						{

							MessageBox.Show(Properties.Settings.Default.lessPtForSpline, Properties.Settings.Default.errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);

							isSuccessfulRead = false;

						}
					}
				}

				// if the file is opened by this plug-in
				if (!hasOpened)
				{
					excelBook.Close(false);
				}


				//release the objects 
				if (oNewExcel == true)
				{
					excelApp.Quit();
					excelApp = null;

					// Excel proces may be not killed.
					KillMyExcel(oMyExcelCaption);
				}

				System.Threading.Thread.CurrentThread.CurrentCulture = oldCI;

				tempReadExcel = isSuccessfulRead;

			}
			catch (Exception ex)
			{

                //sw.WriteLine(ex.ToString());

			}
			finally
			{
                //sw.Flush();
                //sw.Close();
			}



			return tempReadExcel;
		}


		private static void KillMyExcel(string caption)
		{

			try
			{
				int processId = 0;
				int xlHwnd = FindWindow("XLMAIN", caption);
				GetWindowThreadProcessId(xlHwnd, ref processId);
				Process xlProcess = Process.GetProcessById(processId);

				xlProcess.Kill();
				GC.Collect();
				GC.WaitForPendingFinalizers();
			}
			catch (Exception ex)
			{
			}
		}


		public static void DoDelete(string guid)
		{

			Transaction oNewTrans = null;
			oNewTrans = _TxnMgr.StartTransaction(_CurEditDoc, "DeletePoints");


			try
			{
//INSTANT C# NOTE: Commented this declaration since looping variables in 'foreach' loops are declared in the 'foreach' header in C#:
//				dynamic oEachEnt = null;

				// _ADNPlugin_LinkExcelPoints_Entity_Lines_ID
				ObjectCollection oObjCollection = _AttributeMgr.FindObjects(guid, Properties.Settings.Default.sketchLineAttName);


				foreach (dynamic oEachEnt in oObjCollection)
				{
					oEachEnt.Delete();
				}

				oObjCollection = _AttributeMgr.FindObjects(guid, Properties.Settings.Default.sketchSLAttName);


				foreach (dynamic oEachEnt in oObjCollection)
				{
					oEachEnt.Delete();
				}

				oObjCollection = _AttributeMgr.FindObjects(guid, Properties.Settings.Default.sketchPtAttName);


				foreach (dynamic oEachEnt in oObjCollection)
				{
					oEachEnt.Delete();
				}

				oNewTrans.End();

			}
			catch (Exception ex)
			{

				// error when update or resolve points!

				MessageBox.Show(Properties.Settings.Default.failToDelete, Properties.Settings.Default.errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);

				oNewTrans.Abort();

			}




		}
	}




}                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       