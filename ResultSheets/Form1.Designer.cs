namespace ResultSheets
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.newProjectToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openProjectToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveProjectToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveProjectAsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripSeparator();
            this.renderAsJPGToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exportToXLSToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.schemesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripComboBox1 = new System.Windows.Forms.ToolStripComboBox();
            this.colorSchemeListEditorToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.sheetToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.addColumnToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.deleteColumnToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.addRowToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.deleteRowToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem3 = new System.Windows.Forms.ToolStripSeparator();
            this.rowSelectionToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.sortToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.testToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.resultViewSheet1 = new ResultSheets.ResultViewSheet();
            this.saveAsTemplateToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.schemesToolStripMenuItem,
            this.sheetToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1025, 28);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.newProjectToolStripMenuItem,
            this.openProjectToolStripMenuItem,
            this.saveProjectToolStripMenuItem,
            this.saveProjectAsToolStripMenuItem,
            this.saveAsTemplateToolStripMenuItem,
            this.toolStripMenuItem1,
            this.renderAsJPGToolStripMenuItem,
            this.exportToXLSToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(44, 24);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // newProjectToolStripMenuItem
            // 
            this.newProjectToolStripMenuItem.Name = "newProjectToolStripMenuItem";
            this.newProjectToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.N)));
            this.newProjectToolStripMenuItem.Size = new System.Drawing.Size(275, 26);
            this.newProjectToolStripMenuItem.Text = "New Project";
            this.newProjectToolStripMenuItem.Click += new System.EventHandler(this.newProjectToolStripMenuItem_Click);
            // 
            // openProjectToolStripMenuItem
            // 
            this.openProjectToolStripMenuItem.Name = "openProjectToolStripMenuItem";
            this.openProjectToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.O)));
            this.openProjectToolStripMenuItem.Size = new System.Drawing.Size(275, 26);
            this.openProjectToolStripMenuItem.Text = "Open Project";
            this.openProjectToolStripMenuItem.Click += new System.EventHandler(this.openProjectToolStripMenuItem_Click);
            // 
            // saveProjectToolStripMenuItem
            // 
            this.saveProjectToolStripMenuItem.Name = "saveProjectToolStripMenuItem";
            this.saveProjectToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.S)));
            this.saveProjectToolStripMenuItem.Size = new System.Drawing.Size(275, 26);
            this.saveProjectToolStripMenuItem.Text = "Save Project";
            this.saveProjectToolStripMenuItem.Click += new System.EventHandler(this.saveProjectToolStripMenuItem_Click);
            // 
            // saveProjectAsToolStripMenuItem
            // 
            this.saveProjectAsToolStripMenuItem.Name = "saveProjectAsToolStripMenuItem";
            this.saveProjectAsToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)(((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift) 
            | System.Windows.Forms.Keys.S)));
            this.saveProjectAsToolStripMenuItem.Size = new System.Drawing.Size(275, 26);
            this.saveProjectAsToolStripMenuItem.Text = "Save Project As";
            this.saveProjectAsToolStripMenuItem.Click += new System.EventHandler(this.saveProjectAsToolStripMenuItem_Click);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(272, 6);
            // 
            // renderAsJPGToolStripMenuItem
            // 
            this.renderAsJPGToolStripMenuItem.Name = "renderAsJPGToolStripMenuItem";
            this.renderAsJPGToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.R)));
            this.renderAsJPGToolStripMenuItem.Size = new System.Drawing.Size(275, 26);
            this.renderAsJPGToolStripMenuItem.Text = "Render as JPG";
            this.renderAsJPGToolStripMenuItem.Click += new System.EventHandler(this.renderAsJPGToolStripMenuItem_Click);
            // 
            // exportToXLSToolStripMenuItem
            // 
            this.exportToXLSToolStripMenuItem.Name = "exportToXLSToolStripMenuItem";
            this.exportToXLSToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.E)));
            this.exportToXLSToolStripMenuItem.Size = new System.Drawing.Size(275, 26);
            this.exportToXLSToolStripMenuItem.Text = "Export to XLS";
            this.exportToXLSToolStripMenuItem.Click += new System.EventHandler(this.exportToXLSToolStripMenuItem_Click);
            // 
            // schemesToolStripMenuItem
            // 
            this.schemesToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripComboBox1,
            this.colorSchemeListEditorToolStripMenuItem});
            this.schemesToolStripMenuItem.Name = "schemesToolStripMenuItem";
            this.schemesToolStripMenuItem.Size = new System.Drawing.Size(79, 24);
            this.schemesToolStripMenuItem.Text = "Schemes";
            // 
            // toolStripComboBox1
            // 
            this.toolStripComboBox1.Name = "toolStripComboBox1";
            this.toolStripComboBox1.Size = new System.Drawing.Size(121, 28);
            this.toolStripComboBox1.SelectedIndexChanged += new System.EventHandler(this.toolStripComboBox1_SelectedIndexChanged);
            // 
            // colorSchemeListEditorToolStripMenuItem
            // 
            this.colorSchemeListEditorToolStripMenuItem.Name = "colorSchemeListEditorToolStripMenuItem";
            this.colorSchemeListEditorToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.F1)));
            this.colorSchemeListEditorToolStripMenuItem.Size = new System.Drawing.Size(348, 26);
            this.colorSchemeListEditorToolStripMenuItem.Text = "Color Scheme Collection Editor";
            this.colorSchemeListEditorToolStripMenuItem.Click += new System.EventHandler(this.colorSchemeListEditorToolStripMenuItem_Click);
            // 
            // sheetToolStripMenuItem
            // 
            this.sheetToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.addColumnToolStripMenuItem,
            this.deleteColumnToolStripMenuItem,
            this.addRowToolStripMenuItem,
            this.deleteRowToolStripMenuItem,
            this.toolStripMenuItem3,
            this.rowSelectionToolStripMenuItem,
            this.sortToolStripMenuItem,
            this.testToolStripMenuItem});
            this.sheetToolStripMenuItem.Name = "sheetToolStripMenuItem";
            this.sheetToolStripMenuItem.Size = new System.Drawing.Size(58, 24);
            this.sheetToolStripMenuItem.Text = "Sheet";
            // 
            // addColumnToolStripMenuItem
            // 
            this.addColumnToolStripMenuItem.Name = "addColumnToolStripMenuItem";
            this.addColumnToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Oemplus)));
            this.addColumnToolStripMenuItem.Size = new System.Drawing.Size(296, 26);
            this.addColumnToolStripMenuItem.Text = "Add Column";
            this.addColumnToolStripMenuItem.Click += new System.EventHandler(this.addColumnToolStripMenuItem_Click);
            // 
            // deleteColumnToolStripMenuItem
            // 
            this.deleteColumnToolStripMenuItem.Name = "deleteColumnToolStripMenuItem";
            this.deleteColumnToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.OemMinus)));
            this.deleteColumnToolStripMenuItem.Size = new System.Drawing.Size(296, 26);
            this.deleteColumnToolStripMenuItem.Text = "Delete Column";
            this.deleteColumnToolStripMenuItem.Click += new System.EventHandler(this.deleteColumnToolStripMenuItem_Click);
            // 
            // addRowToolStripMenuItem
            // 
            this.addRowToolStripMenuItem.Name = "addRowToolStripMenuItem";
            this.addRowToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.A)));
            this.addRowToolStripMenuItem.Size = new System.Drawing.Size(296, 26);
            this.addRowToolStripMenuItem.Text = "Add Row";
            this.addRowToolStripMenuItem.Click += new System.EventHandler(this.addRowToolStripMenuItem_Click);
            // 
            // deleteRowToolStripMenuItem
            // 
            this.deleteRowToolStripMenuItem.Name = "deleteRowToolStripMenuItem";
            this.deleteRowToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.D)));
            this.deleteRowToolStripMenuItem.Size = new System.Drawing.Size(296, 26);
            this.deleteRowToolStripMenuItem.Text = "Delete Row";
            this.deleteRowToolStripMenuItem.Click += new System.EventHandler(this.deleteRowToolStripMenuItem_Click);
            // 
            // toolStripMenuItem3
            // 
            this.toolStripMenuItem3.Name = "toolStripMenuItem3";
            this.toolStripMenuItem3.Size = new System.Drawing.Size(293, 6);
            // 
            // rowSelectionToolStripMenuItem
            // 
            this.rowSelectionToolStripMenuItem.Checked = true;
            this.rowSelectionToolStripMenuItem.CheckOnClick = true;
            this.rowSelectionToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked;
            this.rowSelectionToolStripMenuItem.Name = "rowSelectionToolStripMenuItem";
            this.rowSelectionToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.W)));
            this.rowSelectionToolStripMenuItem.Size = new System.Drawing.Size(296, 26);
            this.rowSelectionToolStripMenuItem.Text = "Row Selection";
            this.rowSelectionToolStripMenuItem.Click += new System.EventHandler(this.rowSelectionToolStripMenuItem_Click);
            // 
            // sortToolStripMenuItem
            // 
            this.sortToolStripMenuItem.Name = "sortToolStripMenuItem";
            this.sortToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Up)));
            this.sortToolStripMenuItem.Size = new System.Drawing.Size(296, 26);
            this.sortToolStripMenuItem.Text = "Sort";
            this.sortToolStripMenuItem.Click += new System.EventHandler(this.sortToolStripMenuItem_Click);
            // 
            // testToolStripMenuItem
            // 
            this.testToolStripMenuItem.Name = "testToolStripMenuItem";
            this.testToolStripMenuItem.Size = new System.Drawing.Size(296, 26);
            this.testToolStripMenuItem.Text = "Test";
            this.testToolStripMenuItem.Click += new System.EventHandler(this.testToolStripMenuItem_Click);
            // 
            // resultViewSheet1
            // 
            this.resultViewSheet1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.resultViewSheet1.Location = new System.Drawing.Point(0, 28);
            this.resultViewSheet1.Name = "resultViewSheet1";
            this.resultViewSheet1.RowSelectioin = true;
            this.resultViewSheet1.Size = new System.Drawing.Size(1025, 389);
            this.resultViewSheet1.TabIndex = 2;
            // 
            // saveAsTemplateToolStripMenuItem
            // 
            this.saveAsTemplateToolStripMenuItem.Name = "saveAsTemplateToolStripMenuItem";
            this.saveAsTemplateToolStripMenuItem.Size = new System.Drawing.Size(275, 26);
            this.saveAsTemplateToolStripMenuItem.Text = "Save as Template";
            this.saveAsTemplateToolStripMenuItem.Click += new System.EventHandler(this.saveAsTemplateToolStripMenuItem_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1025, 417);
            this.Controls.Add(this.resultViewSheet1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Result Sheets";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem newProjectToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openProjectToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveProjectToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveProjectAsToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem renderAsJPGToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exportToXLSToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem schemesToolStripMenuItem;
        private System.Windows.Forms.ToolStripComboBox toolStripComboBox1;
        private ResultViewSheet resultViewSheet1;
        private System.Windows.Forms.ToolStripMenuItem sheetToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem addColumnToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem deleteColumnToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem addRowToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem deleteRowToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem3;
        private System.Windows.Forms.ToolStripMenuItem testToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem rowSelectionToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem colorSchemeListEditorToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem sortToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveAsTemplateToolStripMenuItem;
    }
}

