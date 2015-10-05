namespace Laser_GPS_Converter_2
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
            this.list_Tracks = new System.Windows.Forms.ListBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.materialFlatButton3 = new MaterialSkin.Controls.MaterialFlatButton();
            this.materialFlatButton2 = new MaterialSkin.Controls.MaterialFlatButton();
            this.label1 = new System.Windows.Forms.Label();
            this.n_Offset = new System.Windows.Forms.NumericUpDown();
            this.txt_Details = new System.Windows.Forms.TextBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.materialFlatButton1 = new MaterialSkin.Controls.MaterialFlatButton();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.n_Offset)).BeginInit();
            this.SuspendLayout();
            // 
            // list_Tracks
            // 
            this.list_Tracks.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.list_Tracks.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.list_Tracks.FormattingEnabled = true;
            this.list_Tracks.IntegralHeight = false;
            this.list_Tracks.Location = new System.Drawing.Point(12, 83);
            this.list_Tracks.Name = "list_Tracks";
            this.list_Tracks.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.list_Tracks.Size = new System.Drawing.Size(236, 218);
            this.list_Tracks.TabIndex = 0;
            this.list_Tracks.SelectedIndexChanged += new System.EventHandler(this.list_Tracks_SelectedIndexChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.materialFlatButton3);
            this.groupBox1.Controls.Add(this.materialFlatButton2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.n_Offset);
            this.groupBox1.Controls.Add(this.txt_Details);
            this.groupBox1.Location = new System.Drawing.Point(254, 83);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(390, 285);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Track Details";
            // 
            // materialFlatButton3
            // 
            this.materialFlatButton3.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.materialFlatButton3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(33)))), ((int)(((byte)(150)))), ((int)(((byte)(243)))));
            this.materialFlatButton3.Depth = 0;
            this.materialFlatButton3.Location = new System.Drawing.Point(305, 251);
            this.materialFlatButton3.Margin = new System.Windows.Forms.Padding(4, 6, 4, 6);
            this.materialFlatButton3.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialFlatButton3.Name = "materialFlatButton3";
            this.materialFlatButton3.Primary = false;
            this.materialFlatButton3.Size = new System.Drawing.Size(75, 23);
            this.materialFlatButton3.TabIndex = 10;
            this.materialFlatButton3.Text = "Export";
            this.materialFlatButton3.UseVisualStyleBackColor = false;
            this.materialFlatButton3.Click += new System.EventHandler(this.btn_Export_Click);
            // 
            // materialFlatButton2
            // 
            this.materialFlatButton2.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.materialFlatButton2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(33)))), ((int)(((byte)(150)))), ((int)(((byte)(243)))));
            this.materialFlatButton2.Depth = 0;
            this.materialFlatButton2.Location = new System.Drawing.Point(222, 251);
            this.materialFlatButton2.Margin = new System.Windows.Forms.Padding(4, 6, 4, 6);
            this.materialFlatButton2.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialFlatButton2.Name = "materialFlatButton2";
            this.materialFlatButton2.Primary = false;
            this.materialFlatButton2.Size = new System.Drawing.Size(75, 23);
            this.materialFlatButton2.TabIndex = 9;
            this.materialFlatButton2.Text = "Export all";
            this.materialFlatButton2.UseVisualStyleBackColor = false;
            this.materialFlatButton2.Click += new System.EventHandler(this.btn_ExportAll_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(20, 258);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(127, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Timezone offset (vs UTC)";
            // 
            // n_Offset
            // 
            this.n_Offset.DecimalPlaces = 1;
            this.n_Offset.Increment = new decimal(new int[] {
            5,
            0,
            0,
            65536});
            this.n_Offset.Location = new System.Drawing.Point(153, 254);
            this.n_Offset.Maximum = new decimal(new int[] {
            12,
            0,
            0,
            0});
            this.n_Offset.Minimum = new decimal(new int[] {
            12,
            0,
            0,
            -2147483648});
            this.n_Offset.Name = "n_Offset";
            this.n_Offset.Size = new System.Drawing.Size(45, 20);
            this.n_Offset.TabIndex = 2;
            this.n_Offset.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // txt_Details
            // 
            this.txt_Details.BackColor = System.Drawing.SystemColors.Window;
            this.txt_Details.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_Details.Location = new System.Drawing.Point(9, 19);
            this.txt_Details.Multiline = true;
            this.txt_Details.Name = "txt_Details";
            this.txt_Details.ReadOnly = true;
            this.txt_Details.Size = new System.Drawing.Size(371, 223);
            this.txt_Details.TabIndex = 0;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.DefaultExt = "gpx";
            this.openFileDialog1.FileName = "BridgeMin.dll";
            this.openFileDialog1.Filter = "BridgeMin.dll|BridgeMin.dll|All files|*";
            this.openFileDialog1.ReadOnlyChecked = true;
            this.openFileDialog1.SupportMultiDottedExtensions = true;
            this.openFileDialog1.Title = "Laser GPS Database";
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.DefaultExt = "gpx";
            this.saveFileDialog1.Filter = "GPX files|*.gpx";
            this.saveFileDialog1.SupportMultiDottedExtensions = true;
            this.saveFileDialog1.Title = "Export GPX file...";
            // 
            // materialFlatButton1
            // 
            this.materialFlatButton1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.materialFlatButton1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(33)))), ((int)(((byte)(150)))), ((int)(((byte)(243)))));
            this.materialFlatButton1.Depth = 0;
            this.materialFlatButton1.Location = new System.Drawing.Point(12, 315);
            this.materialFlatButton1.Margin = new System.Windows.Forms.Padding(4, 6, 4, 6);
            this.materialFlatButton1.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialFlatButton1.Name = "materialFlatButton1";
            this.materialFlatButton1.Primary = false;
            this.materialFlatButton1.Size = new System.Drawing.Size(75, 23);
            this.materialFlatButton1.TabIndex = 4;
            this.materialFlatButton1.Text = "Load";
            this.materialFlatButton1.UseVisualStyleBackColor = false;
            this.materialFlatButton1.Click += new System.EventHandler(this.btn_Load_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 350);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(235, 18);
            this.progressBar1.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Location = new System.Drawing.Point(93, 320);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(154, 18);
            this.label2.TabIndex = 0;
            this.label2.Text = "label2";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(657, 385);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.materialFlatButton1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.list_Tracks);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(657, 385);
            this.MinimumSize = new System.Drawing.Size(657, 385);
            this.Name = "Form1";
            this.Sizable = false;
            this.Text = "GPS Converter V2";
            this.Load += new System.EventHandler(this.Form_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.n_Offset)).EndInit();
            this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.ListBox list_Tracks;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown n_Offset;
        private System.Windows.Forms.TextBox txt_Details;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private MaterialSkin.Controls.MaterialFlatButton materialFlatButton1;
        private MaterialSkin.Controls.MaterialFlatButton materialFlatButton3;
        private MaterialSkin.Controls.MaterialFlatButton materialFlatButton2;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label label2;
    }
}

