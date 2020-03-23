namespace ManyToOne
{
    partial class FormDataSource
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
            this.btnDataSource = new System.Windows.Forms.Button();
            this.groupBoxPath = new System.Windows.Forms.GroupBox();
            this.labelPath = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.comboBoxSheets = new System.Windows.Forms.ComboBox();
            this.buttonSetDataSource = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.openFileDialogDataSource = new System.Windows.Forms.OpenFileDialog();
            this.dataGridViewDataSource = new System.Windows.Forms.DataGridView();
            this.groupBoxPath.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewDataSource)).BeginInit();
            this.SuspendLayout();
            // 
            // btnDataSource
            // 
            this.btnDataSource.Location = new System.Drawing.Point(248, 15);
            this.btnDataSource.Margin = new System.Windows.Forms.Padding(4);
            this.btnDataSource.Name = "btnDataSource";
            this.btnDataSource.Size = new System.Drawing.Size(195, 28);
            this.btnDataSource.TabIndex = 0;
            this.btnDataSource.Text = "Select Data Source";
            this.btnDataSource.UseVisualStyleBackColor = true;
            this.btnDataSource.Click += new System.EventHandler(this.btnDataSource_Click);
            // 
            // groupBoxPath
            // 
            this.groupBoxPath.Controls.Add(this.labelPath);
            this.groupBoxPath.Location = new System.Drawing.Point(16, 50);
            this.groupBoxPath.Margin = new System.Windows.Forms.Padding(4);
            this.groupBoxPath.Name = "groupBoxPath";
            this.groupBoxPath.Padding = new System.Windows.Forms.Padding(4);
            this.groupBoxPath.Size = new System.Drawing.Size(631, 87);
            this.groupBoxPath.TabIndex = 1;
            this.groupBoxPath.TabStop = false;
            this.groupBoxPath.Text = "Data Source Path";
            // 
            // labelPath
            // 
            this.labelPath.AutoSize = true;
            this.labelPath.ForeColor = System.Drawing.Color.ForestGreen;
            this.labelPath.Location = new System.Drawing.Point(9, 25);
            this.labelPath.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelPath.Name = "labelPath";
            this.labelPath.Size = new System.Drawing.Size(159, 16);
            this.labelPath.TabIndex = 0;
            this.labelPath.Text = "No Data Source selected";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(67, 149);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(82, 16);
            this.label1.TabIndex = 2;
            this.label1.Text = "Select sheet";
            // 
            // comboBoxSheets
            // 
            this.comboBoxSheets.FormattingEnabled = true;
            this.comboBoxSheets.Location = new System.Drawing.Point(163, 145);
            this.comboBoxSheets.Margin = new System.Windows.Forms.Padding(4);
            this.comboBoxSheets.Name = "comboBoxSheets";
            this.comboBoxSheets.Size = new System.Drawing.Size(348, 24);
            this.comboBoxSheets.TabIndex = 3;
            this.comboBoxSheets.SelectedIndexChanged += new System.EventHandler(this.comboBoxSheets_SelectedIndexChanged);
            // 
            // buttonSetDataSource
            // 
            this.buttonSetDataSource.Enabled = false;
            this.buttonSetDataSource.Location = new System.Drawing.Point(248, 180);
            this.buttonSetDataSource.Margin = new System.Windows.Forms.Padding(4);
            this.buttonSetDataSource.Name = "buttonSetDataSource";
            this.buttonSetDataSource.Size = new System.Drawing.Size(195, 28);
            this.buttonSetDataSource.TabIndex = 4;
            this.buttonSetDataSource.Text = "Set Data Source";
            this.buttonSetDataSource.UseVisualStyleBackColor = true;
            this.buttonSetDataSource.Click += new System.EventHandler(this.buttonSetDataSource_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.Location = new System.Drawing.Point(291, 215);
            this.buttonCancel.Margin = new System.Windows.Forms.Padding(4);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(100, 28);
            this.buttonCancel.TabIndex = 5;
            this.buttonCancel.Text = "Close";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // openFileDialogDataSource
            // 
            this.openFileDialogDataSource.FileName = "dataSource";
            this.openFileDialogDataSource.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            this.openFileDialogDataSource.ReadOnlyChecked = true;
            // 
            // dataGridViewDataSource
            // 
            this.dataGridViewDataSource.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewDataSource.Location = new System.Drawing.Point(16, 257);
            this.dataGridViewDataSource.Name = "dataGridViewDataSource";
            this.dataGridViewDataSource.RowTemplate.Height = 24;
            this.dataGridViewDataSource.Size = new System.Drawing.Size(631, 294);
            this.dataGridViewDataSource.TabIndex = 6;
            // 
            // FormDataSource
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(667, 563);
            this.Controls.Add(this.dataGridViewDataSource);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonSetDataSource);
            this.Controls.Add(this.comboBoxSheets);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.groupBoxPath);
            this.Controls.Add(this.btnDataSource);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormDataSource";
            this.Text = "Data Source";
            this.Load += new System.EventHandler(this.FormDataSource_Load);
            this.groupBoxPath.ResumeLayout(false);
            this.groupBoxPath.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewDataSource)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnDataSource;
        private System.Windows.Forms.GroupBox groupBoxPath;
        private System.Windows.Forms.Label labelPath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox comboBoxSheets;
        private System.Windows.Forms.Button buttonSetDataSource;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.OpenFileDialog openFileDialogDataSource;
        private System.Windows.Forms.DataGridView dataGridViewDataSource;
    }
}