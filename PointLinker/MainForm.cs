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
//'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


//INSTANT C# NOTE: Formerly VB project-level imports:

using System.Windows.Forms;
using Inventor;
using PointLinker.Properties;

namespace PointLinker
{
	namespace PointLinker
	{


		public partial class MainForm
		{

			//variables for  this class

			internal MainForm()
			{
				InitializeComponent();
			}
#region variables


			private Inventor.Application _InvApp;
			public Inventor.Application InvApp
			{
				get
				{
					return _InvApp;
				}
				set
				{
					_InvApp = value;
				}
			}

			private bool _isFromSketch;
			public bool FromSketch
			{
				get
				{
					return _isFromSketch;
				}
				set
				{
					_isFromSketch = value;
				}
			}

			private const int sizeWidth = 850;
			private const int sizeHeight = 250;

			private int _undoSta = 0; // 0: new points 1: update/resolve 2:delete

#endregion


#region Main Form

			private void MainForm_Load(object sender, System.EventArgs e)
			{



				// Initialize the objects of Inventor
				Helper.Initialize(_InvApp, _isFromSketch, ref DataGV);


				BtnUndo.Enabled = false;

				if (_isFromSketch)
				{
					// from sketch
					DataGV.Columns["SketchName"].Visible = false;
					DataGV.Columns["SketchType"].Visible = false;
					DataGV.Columns["FileName"].Width = 300;
					DataGV.Columns["PointType"].Width = 150;
					DataGV.Columns["Status"].Width = 300;
				}
				else
				{
					// from Tools Tab in Part or Assembly         
					DataGV.Columns["FileName"].Width = 200;
					DataGV.Columns["PointType"].Width = 150;
					DataGV.Columns["Status"].Width = 200;
					DataGV.Columns["SketchName"].Visible = true;
					DataGV.Columns["SketchName"].Width = 120;
					DataGV.Columns["SketchType"].Visible = true;
					DataGV.Columns["SketchType"].Width = 120;
				}

				// produce the strings for the rows
				Helper.ProduceRows();

			}


			/// <summary>
			/// Import Points
			/// </summary>
			/// <param name="sender"></param>
			/// <param name="e"></param>
			/// <remarks></remarks>
			private void BtnImport_Click(object sender, System.EventArgs e)
			{

				// if the current object is a sketch

				if (!(_InvApp.ActiveEditObject is Sketch) && !(_InvApp.ActiveEditObject is Sketch3D))
				{

					MessageBox.Show(Properties.Settings.Default.notSketch, Properties.Settings.Default.messageCaption, MessageBoxButtons.OK, MessageBoxIcon.Information);

					return;

				}

				// open the excel file

				string oSelectExcelFile = null;
				OpenFileDlg.Title = Properties.Settings.Default.capForOpenDlg;
				OpenFileDlg.Filter = Properties.Settings.Default.filterForOpenDlg;
				OpenFileDlg.Multiselect = false;

				if (OpenFileDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
				{

					oSelectExcelFile = OpenFileDlg.FileName;

				}
				else
				{

					return;

				}

				// check if the file has been imported to
				// the current sketch
				DialogResult oResult = System.Windows.Forms.DialogResult.OK;
				foreach (DataGridViewRow oRow in DataGV.Rows)
				{

					string oFileName = oRow.Cells["FileName"].Value.ToString();

					string oGuid = oRow.Cells["Guid"].Value.ToString();

					object oCurSketch = Helper.GetSketchFromGuid(oGuid);

					if (oCurSketch != null)
					{

						if (oFileName == oSelectExcelFile && oCurSketch == _InvApp.ActiveEditObject)
						{

							oResult = MessageBox.Show(Properties.Settings.Default.hasImport, Properties.Settings.Default.errorCaption, MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
							break;

						}

					}
					else
					{

						MessageBox.Show(Properties.Settings.Default.generalError, Properties.Settings.Default.errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);

						return;

					}

				}


				if (oResult == System.Windows.Forms.DialogResult.OK)
				{
					
					if (Helper.ImportNewPoints(oSelectExcelFile))
					{
					
						_undoSta = 0;
						BtnUndo.Enabled = true;
					}
				}


			}

			private void BtnCancel_Click(object sender, System.EventArgs e)
			{
				this.Close();
			}

			private void BtnUndo_Click(object sender, System.EventArgs e)
			{

				ControlDefinition oUndoControl = _InvApp.CommandManager.ControlDefinitions["AppUndoCmd"];

				switch (_undoSta)
				{
					case 0:
					break;
					case 1:
						if (_InvApp.ActiveEditDocument is DrawingDocument)
						{
							oUndoControl.Execute2(true); // for edit
						}
						else
						{
							oUndoControl.Execute2(true); // for select
							oUndoControl.Execute2(true); // for edit
						}

						break;
					case 2:
						oUndoControl.Execute2(true);
						break;
				}

				Helper.ProduceRows();

				BtnUndo.Enabled = false;

			}

			private void radPoints_Click(object sender, System.EventArgs e)
			{
				Helper._ptOption = Helper.PointOptionEnum.ePoints;
			}

			private void radLines_Click(object sender, System.EventArgs e)
			{
				Helper._ptOption = Helper.PointOptionEnum.eLines;
			}

			private void radSpline_Click(object sender, System.EventArgs e)
			{
				Helper._ptOption = Helper.PointOptionEnum.eSpline;
			}

			/// <summary>
			/// add warning flag ahead of the row if it is not up to date
			/// </summary>
			/// <param name="sender"></param>
			/// <param name="e"></param>
			/// <remarks></remarks>
			private void DataGV_CellPainting(object sender, System.Windows.Forms.DataGridViewCellPaintingEventArgs e)
			{

				if (e.ColumnIndex < 0 && e.RowIndex >= 0 && e.RowIndex < DataGV.Rows.Count)
				{

					if (DataGV.Rows[e.RowIndex].Cells["Status"].Value != "Up-to-date")
					{

						e.PaintBackground(e.ClipBounds, true);
						e.Graphics.DrawImage(Resources.Light_Bulb.ToBitmap(), e.CellBounds);
						e.Handled = true;

					}

				}
			}

			/// <summary>
			/// resolve the item to another excel file
			/// </summary>
			/// <param name="sender"></param>
			/// <param name="e"></param>
			/// <remarks></remarks>
			private void ResolveToolStripMenuItem_Click(object sender, System.EventArgs e)
			{

				//open the excel file

				OpenFileDlg.Title = Properties.Settings.Default.capForOpenDlg;
				OpenFileDlg.Filter = Properties.Settings.Default.filterForOpenDlg;
				OpenFileDlg.Multiselect = false;

				if (OpenFileDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
				{

					DataGridViewRow oCurRow = DataGV.CurrentRow;
					string oFileName = oCurRow.Cells["FileName"].Value.ToString();
					string oStatus = oCurRow.Cells["Status"].Value.ToString();

					if (oFileName == OpenFileDlg.FileName && oStatus == "Up-to-date")
					{

						MessageBox.Show(Properties.Settings.Default.upToDate, Properties.Settings.Default.messageCaption, MessageBoxButtons.OK, MessageBoxIcon.Information);

						return;
					}
					else
					{
						if (Helper.UpdateOrResolve(OpenFileDlg.FileName) == true)
						{
							_undoSta = 1;
							BtnUndo.Enabled = true;
						}
					}
				}
			}

			/// <summary>
			/// update the row with newer data of the excel file
			/// </summary>
			/// <param name="sender"></param>
			/// <param name="e"></param>
			/// <remarks></remarks>
			private void UpdateToolStripMenuItem_Click(object sender, System.EventArgs e)
			{


				DataGridViewRow oCurRow = DataGV.CurrentRow;
				string oStatus = oCurRow.Cells["Status"].Value.ToString();

				// if current status is normal, exit sub

				if (oStatus == "Up-to-date")
				{

					MessageBox.Show(Properties.Settings.Default.upToDate, Properties.Settings.Default.messageCaption, MessageBoxButtons.OK, MessageBoxIcon.Information);

					return;
				}

				string oFileName = oCurRow.Cells["FileName"].Value.ToString();

				//  the file does not exist!

				if (!(System.IO.File.Exists(oFileName)))
				{

					MessageBox.Show(Properties.Settings.Default.fileNotExists, Properties.Settings.Default.errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);

					return;

				}
				// go to update

				if (Helper.UpdateOrResolve() == true)
				{
					_undoSta = 1;
					BtnUndo.Enabled = true;
				}

			}

			private void DataGV_MouseClick(object sender, System.Windows.Forms.MouseEventArgs e)
			{


                DataGridView.HitTestInfo hti = DataGV.HitTest(e.X, e.Y);

				if (hti.Type == DataGridViewHitTestType.RowHeader || hti.Type == DataGridViewHitTestType.Cell)
				{

					DataGV.CurrentCell = DataGV.Rows[hti.RowIndex].Cells[0];

					if (e.Button == System.Windows.Forms.MouseButtons.Right)
					{

						DataGV.CurrentCell = DataGV.Rows[hti.RowIndex].Cells[0];

						DataGVContextMenu.Show(this, e.Location);

					}

					string oGuid = DataGV.CurrentRow.Cells["Guid"].Value.ToString();
					Helper.ZoomWindowAndHilight(oGuid);

					if (Helper._CurrentIndex != hti.RowIndex)
					{

						BtnUndo.Enabled = false;
						Helper._CurrentIndex = hti.RowIndex;
					}

				}

			}

			private void DataGV_KeyUp(object sender, System.Windows.Forms.KeyEventArgs e)
			{

				switch (e.KeyCode)
				{
					case Keys.Escape:
					{
						this.Close();
						break;
					}
					case Keys.Up:
					{

						int oCurrentRowIndex = DataGV.CurrentRow.Index;


						BtnUndo.Enabled = false;

						string oGuid = DataGV.CurrentRow.Cells["Guid"].Value.ToString();
						Helper.ZoomWindowAndHilight(oGuid);

						break;
					}
					case Keys.Down:
					{

						int oCurrentRowIndex = DataGV.CurrentRow.Index;


						BtnUndo.Enabled = false;

						string oGuid = DataGV.CurrentRow.Cells["Guid"].Value.ToString();
						Helper.ZoomWindowAndHilight(oGuid);

						break;
					}
				}
			}

			private void MainForm_FormClosed(object sender, System.Windows.Forms.FormClosedEventArgs e)
			{

				Helper.WhenClose();

			}

			private void MainForm_SizeChanged(object sender, System.EventArgs e)
			{

				if (this.Width < sizeWidth)
				{
					this.Width = sizeWidth;
				}

				if (this.Height < sizeHeight)
				{
					this.Height = sizeHeight;
				}

			}

			private void DeleteToolStripMenuItem_Click(object sender, System.EventArgs e)
			{


				DataGridViewRow oCurRow = DataGV.CurrentRow;

				string oGuid = oCurRow.Cells["Guid"].Value.ToString();
				Helper.DoDelete(oGuid);

				_undoSta = 2;
				BtnUndo.Enabled = true;

				Helper.ProduceRows();


			}

#endregion




		}

	}

}