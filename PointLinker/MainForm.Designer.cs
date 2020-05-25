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


using System;

namespace PointLinker
{
	namespace PointLinker
	{


		[Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
		public partial class MainForm : System.Windows.Forms.Form
		{
			//Form overrides dispose to clean up the component list.
			[System.Diagnostics.DebuggerNonUserCode()]
			protected override void Dispose(bool disposing)
			{
				try
				{
					if (disposing && components != null)
					{
						components.Dispose();
					}
				}
				finally
				{
					base.Dispose(disposing);
				}
			}

			//Required by the Windows Form Designer
			private System.ComponentModel.IContainer components;

			//NOTE: The following procedure is required by the Windows Form Designer
			//It can be modified using the Windows Form Designer.  
			//Do not modify it using the code editor.
			[System.Diagnostics.DebuggerStepThrough()]
			private void InitializeComponent()
			{
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.DataGVContextMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.UpdateToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ResolveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.DeleteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.OpenFileDlg = new System.Windows.Forms.OpenFileDialog();
            this.DataGV = new System.Windows.Forms.DataGridView();
            this.FileName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PointType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Status = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SketchName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SketchType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Guid = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PointsCount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SaveTime = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.GroupBox1 = new System.Windows.Forms.GroupBox();
            this.radLines = new System.Windows.Forms.RadioButton();
            this.radSpline = new System.Windows.Forms.RadioButton();
            this.radPoints = new System.Windows.Forms.RadioButton();
            this.Panel3 = new System.Windows.Forms.Panel();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.BtnUndo = new System.Windows.Forms.Button();
            this.BtnImport = new System.Windows.Forms.Button();
            this.Panel2 = new System.Windows.Forms.Panel();
            this.DataGVContextMenu.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DataGV)).BeginInit();
            this.GroupBox1.SuspendLayout();
            this.Panel3.SuspendLayout();
            this.Panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // DataGVContextMenu
            // 
            this.DataGVContextMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.UpdateToolStripMenuItem,
            this.ResolveToolStripMenuItem,
            this.DeleteToolStripMenuItem});
            this.DataGVContextMenu.Name = "DataGVContextMenu";
            this.DataGVContextMenu.Size = new System.Drawing.Size(115, 70);
            // 
            // UpdateToolStripMenuItem
            // 
            this.UpdateToolStripMenuItem.Name = "UpdateToolStripMenuItem";
            this.UpdateToolStripMenuItem.Size = new System.Drawing.Size(114, 22);
            this.UpdateToolStripMenuItem.Text = "Update";
            this.UpdateToolStripMenuItem.Click += new System.EventHandler(this.UpdateToolStripMenuItem_Click);
            // 
            // ResolveToolStripMenuItem
            // 
            this.ResolveToolStripMenuItem.Name = "ResolveToolStripMenuItem";
            this.ResolveToolStripMenuItem.Size = new System.Drawing.Size(114, 22);
            this.ResolveToolStripMenuItem.Text = "Resolve";
            this.ResolveToolStripMenuItem.Click += new System.EventHandler(this.ResolveToolStripMenuItem_Click);
            // 
            // DeleteToolStripMenuItem
            // 
            this.DeleteToolStripMenuItem.Name = "DeleteToolStripMenuItem";
            this.DeleteToolStripMenuItem.Size = new System.Drawing.Size(114, 22);
            this.DeleteToolStripMenuItem.Text = "Delete";
            this.DeleteToolStripMenuItem.Click += new System.EventHandler(this.DeleteToolStripMenuItem_Click);
            // 
            // DataGV
            // 
            this.DataGV.AllowUserToAddRows = false;
            this.DataGV.AllowUserToDeleteRows = false;
            this.DataGV.AllowUserToOrderColumns = true;
            this.DataGV.BorderStyle = System.Windows.Forms.BorderStyle.None;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.DataGV.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.DataGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DataGV.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.FileName,
            this.PointType,
            this.Status,
            this.SketchName,
            this.SketchType,
            this.Guid,
            this.PointsCount,
            this.SaveTime});
            this.DataGV.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DataGV.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.DataGV.Location = new System.Drawing.Point(0, 0);
            this.DataGV.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.DataGV.MultiSelect = false;
            this.DataGV.Name = "DataGV";
            this.DataGV.RowTemplate.Height = 27;
            this.DataGV.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.DataGV.Size = new System.Drawing.Size(832, 140);
            this.DataGV.TabIndex = 4;
            this.DataGV.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.DataGV_CellPainting);
            this.DataGV.KeyUp += new System.Windows.Forms.KeyEventHandler(this.DataGV_KeyUp);
            this.DataGV.MouseClick += new System.Windows.Forms.MouseEventHandler(this.DataGV_MouseClick);
            // 
            // FileName
            // 
            this.FileName.HeaderText = "File Name";
            this.FileName.Name = "FileName";
            this.FileName.Width = 200;
            // 
            // PointType
            // 
            this.PointType.HeaderText = "Point Type";
            this.PointType.Name = "PointType";
            this.PointType.Width = 150;
            // 
            // Status
            // 
            this.Status.HeaderText = "Status";
            this.Status.Name = "Status";
            this.Status.Width = 200;
            // 
            // SketchName
            // 
            this.SketchName.HeaderText = "Sketch Name";
            this.SketchName.Name = "SketchName";
            this.SketchName.Width = 120;
            // 
            // SketchType
            // 
            this.SketchType.HeaderText = "Sketch Type";
            this.SketchType.Name = "SketchType";
            this.SketchType.Width = 120;
            // 
            // Guid
            // 
            this.Guid.HeaderText = "Guid";
            this.Guid.Name = "Guid";
            this.Guid.Visible = false;
            // 
            // PointsCount
            // 
            this.PointsCount.HeaderText = "PointsCount";
            this.PointsCount.Name = "PointsCount";
            this.PointsCount.Visible = false;
            // 
            // SaveTime
            // 
            this.SaveTime.HeaderText = "SaveTime";
            this.SaveTime.Name = "SaveTime";
            this.SaveTime.Visible = false;
            // 
            // GroupBox1
            // 
            this.GroupBox1.Controls.Add(this.radLines);
            this.GroupBox1.Controls.Add(this.radSpline);
            this.GroupBox1.Controls.Add(this.radPoints);
            this.GroupBox1.Dock = System.Windows.Forms.DockStyle.Left;
            this.GroupBox1.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GroupBox1.Location = new System.Drawing.Point(0, 0);
            this.GroupBox1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.GroupBox1.Name = "GroupBox1";
            this.GroupBox1.Padding = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.GroupBox1.Size = new System.Drawing.Size(234, 74);
            this.GroupBox1.TabIndex = 4;
            this.GroupBox1.TabStop = false;
            this.GroupBox1.Text = "Point Type";
            // 
            // radLines
            // 
            this.radLines.AutoSize = true;
            this.radLines.Location = new System.Drawing.Point(88, 20);
            this.radLines.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.radLines.Name = "radLines";
            this.radLines.Size = new System.Drawing.Size(56, 19);
            this.radLines.TabIndex = 3;
            this.radLines.Text = "Lines";
            this.radLines.UseVisualStyleBackColor = true;
            this.radLines.Click += new System.EventHandler(this.radLines_Click);
            // 
            // radSpline
            // 
            this.radSpline.AutoSize = true;
            this.radSpline.Location = new System.Drawing.Point(152, 20);
            this.radSpline.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.radSpline.Name = "radSpline";
            this.radSpline.Size = new System.Drawing.Size(60, 19);
            this.radSpline.TabIndex = 3;
            this.radSpline.Text = "Spline";
            this.radSpline.UseVisualStyleBackColor = true;
            this.radSpline.Click += new System.EventHandler(this.radSpline_Click);
            // 
            // radPoints
            // 
            this.radPoints.AutoSize = true;
            this.radPoints.Checked = true;
            this.radPoints.Location = new System.Drawing.Point(20, 20);
            this.radPoints.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.radPoints.Name = "radPoints";
            this.radPoints.Size = new System.Drawing.Size(60, 19);
            this.radPoints.TabIndex = 3;
            this.radPoints.TabStop = true;
            this.radPoints.Text = "Points";
            this.radPoints.UseVisualStyleBackColor = true;
            this.radPoints.Click += new System.EventHandler(this.radPoints_Click);
            // 
            // Panel3
            // 
            this.Panel3.Controls.Add(this.BtnCancel);
            this.Panel3.Dock = System.Windows.Forms.DockStyle.Right;
            this.Panel3.Location = new System.Drawing.Point(699, 0);
            this.Panel3.Name = "Panel3";
            this.Panel3.Size = new System.Drawing.Size(133, 74);
            this.Panel3.TabIndex = 5;
            // 
            // BtnCancel
            // 
            this.BtnCancel.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnCancel.Location = new System.Drawing.Point(27, 20);
            this.BtnCancel.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(84, 30);
            this.BtnCancel.TabIndex = 1;
            this.BtnCancel.Text = "OK";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // BtnUndo
            // 
            this.BtnUndo.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnUndo.Location = new System.Drawing.Point(507, 20);
            this.BtnUndo.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.BtnUndo.Name = "BtnUndo";
            this.BtnUndo.Size = new System.Drawing.Size(86, 30);
            this.BtnUndo.TabIndex = 1;
            this.BtnUndo.Text = "Undo";
            this.BtnUndo.UseVisualStyleBackColor = true;
            this.BtnUndo.Click += new System.EventHandler(this.BtnUndo_Click);
            // 
            // BtnImport
            // 
            this.BtnImport.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnImport.Location = new System.Drawing.Point(258, 20);
            this.BtnImport.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.BtnImport.Name = "BtnImport";
            this.BtnImport.Size = new System.Drawing.Size(84, 30);
            this.BtnImport.TabIndex = 0;
            this.BtnImport.Text = "Import";
            this.BtnImport.UseVisualStyleBackColor = true;
            this.BtnImport.Click += new System.EventHandler(this.BtnImport_Click);
            // 
            // Panel2
            // 
            this.Panel2.Controls.Add(this.BtnUndo);
            this.Panel2.Controls.Add(this.Panel3);
            this.Panel2.Controls.Add(this.GroupBox1);
            this.Panel2.Controls.Add(this.BtnImport);
            this.Panel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.Panel2.Location = new System.Drawing.Point(0, 140);
            this.Panel2.Name = "Panel2";
            this.Panel2.Size = new System.Drawing.Size(832, 74);
            this.Panel2.TabIndex = 3;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(832, 214);
            this.Controls.Add(this.DataGV);
            this.Controls.Add(this.Panel2);
            this.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Point Linker";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MainForm_FormClosed);
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.SizeChanged += new System.EventHandler(this.MainForm_SizeChanged);
            this.DataGVContextMenu.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.DataGV)).EndInit();
            this.GroupBox1.ResumeLayout(false);
            this.GroupBox1.PerformLayout();
            this.Panel3.ResumeLayout(false);
            this.Panel2.ResumeLayout(false);
            this.ResumeLayout(false);

			}
			internal System.Windows.Forms.ContextMenuStrip DataGVContextMenu;
			internal System.Windows.Forms.OpenFileDialog OpenFileDlg;
			internal System.Windows.Forms.ToolStripMenuItem UpdateToolStripMenuItem;
			internal System.Windows.Forms.ToolStripMenuItem ResolveToolStripMenuItem;
			internal System.Windows.Forms.DataGridView DataGV;
			internal System.Windows.Forms.GroupBox GroupBox1;
			internal System.Windows.Forms.RadioButton radLines;
			internal System.Windows.Forms.RadioButton radSpline;
			internal System.Windows.Forms.RadioButton radPoints;
			internal System.Windows.Forms.Panel Panel3;
			internal System.Windows.Forms.Button BtnUndo;
			internal System.Windows.Forms.Button BtnCancel;
			internal System.Windows.Forms.Button BtnImport;
			internal System.Windows.Forms.Panel Panel2;
			internal System.Windows.Forms.DataGridViewTextBoxColumn FileName;
			internal System.Windows.Forms.DataGridViewTextBoxColumn PointType;
			internal System.Windows.Forms.DataGridViewTextBoxColumn Status;
			internal System.Windows.Forms.DataGridViewTextBoxColumn SketchName;
			internal System.Windows.Forms.DataGridViewTextBoxColumn SketchType;
			internal System.Windows.Forms.DataGridViewTextBoxColumn Guid;
			internal System.Windows.Forms.DataGridViewTextBoxColumn PointsCount;
			internal System.Windows.Forms.DataGridViewTextBoxColumn SaveTime;
			internal System.Windows.Forms.ToolStripMenuItem DeleteToolStripMenuItem;
		}

	}
}