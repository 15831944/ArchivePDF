using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace SolidworksBMPWrapper {
	public partial class ErrForm : Form {
		private TableLayoutPanel tableLayoutPanel1;
		private Button button1;
		private TableLayoutPanel tableLayoutPanel2;
		private Button cpyToClip;
		private TextBox textBox1;
		private TextBox tbxErrMsg;
		private TextBox tbMsg;

		public ErrForm() {
			InitializeComponent();
		}

		public void fillErrMsg(string str) {
			this.tbMsg.AppendText(str);
		}

		private void button1_Click(object sender, EventArgs e) {
			this.Close();
		}

		private void InitializeComponent() {
			this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
			this.button1 = new System.Windows.Forms.Button();
			this.tbMsg = new System.Windows.Forms.TextBox();
			this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
			this.cpyToClip = new System.Windows.Forms.Button();
			this.textBox1 = new System.Windows.Forms.TextBox();
			this.tbxErrMsg = new System.Windows.Forms.TextBox();
			this.tableLayoutPanel1.SuspendLayout();
			this.tableLayoutPanel2.SuspendLayout();
			this.SuspendLayout();
			// 
			// tableLayoutPanel1
			// 
			this.tableLayoutPanel1.ColumnCount = 1;
			this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
			this.tableLayoutPanel1.Controls.Add(this.tbMsg, 0, 0);
			this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel2, 0, 1);
			this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
			//this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
			this.tableLayoutPanel1.Name = "tableLayoutPanel1";
			this.tableLayoutPanel1.RowCount = 2;
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 85F));
			this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 15F));
			//this.tableLayoutPanel1.Size = new System.Drawing.Size(667, 270);
			this.tableLayoutPanel1.TabIndex = 0;
			// 
			// button1
			// 
			this.button1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			//this.button1.Location = new System.Drawing.Point(3, 3);
			this.button1.Name = "button1";
			//this.button1.Size = new System.Drawing.Size(324, 29);
			this.button1.TabIndex = 0;
			this.button1.Text = "Close";
			this.button1.UseVisualStyleBackColor = true;
			// 
			// tbMsg
			// 
			this.tbMsg.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tbMsg.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.tbMsg.Location = new System.Drawing.Point(3, 3);
			this.tbMsg.Multiline = true;
			this.tbMsg.Name = "tbMsg";
			this.tbMsg.Size = new System.Drawing.Size(661, 223);
			this.tbMsg.TabIndex = 1;
			this.tbMsg.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
			// 
			// tableLayoutPanel2
			// 
			this.tableLayoutPanel2.ColumnCount = 2;
			this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
			this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
			this.tableLayoutPanel2.Controls.Add(this.button1, 0, 0);
			this.tableLayoutPanel2.Controls.Add(this.cpyToClip, 1, 0);
			this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tableLayoutPanel2.Location = new System.Drawing.Point(3, 232);
			this.tableLayoutPanel2.Name = "tableLayoutPanel2";
			this.tableLayoutPanel2.RowCount = 1;
			this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
			this.tableLayoutPanel2.Size = new System.Drawing.Size(661, 35);
			this.tableLayoutPanel2.TabIndex = 2;
			// 
			// cpyToClip
			// 
			this.cpyToClip.Dock = System.Windows.Forms.DockStyle.Fill;
			this.cpyToClip.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.cpyToClip.Location = new System.Drawing.Point(333, 3);
			this.cpyToClip.Name = "cpyToClip";
			this.cpyToClip.Size = new System.Drawing.Size(325, 29);
			this.cpyToClip.TabIndex = 1;
			this.cpyToClip.Text = "Copy to Clipboard";
			this.cpyToClip.UseVisualStyleBackColor = true;
			this.cpyToClip.Click += new System.EventHandler(this.cpyToClip_Click);
			// 
			// textBox1
			// 
			this.textBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.textBox1.Location = new System.Drawing.Point(0, 0);
			this.textBox1.Multiline = true;
			this.textBox1.Name = "textBox1";
			this.textBox1.Size = new System.Drawing.Size(100, 20);
			this.textBox1.TabIndex = 1;
			// 
			// tbxErrMsg
			// 
			this.tbxErrMsg.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tbxErrMsg.Location = new System.Drawing.Point(0, 0);
			this.tbxErrMsg.Multiline = true;
			this.tbxErrMsg.Name = "tbxErrMsg";
			this.tbxErrMsg.Size = new System.Drawing.Size(100, 20);
			this.tbxErrMsg.TabIndex = 1;
			this.tbxErrMsg.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
			// 
			// ErrForm
			// 
			this.ClientSize = new System.Drawing.Size(667, 270);
			this.Controls.Add(this.tableLayoutPanel1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
			this.Name = "ErrForm";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.tableLayoutPanel1.ResumeLayout(false);
			this.tableLayoutPanel1.PerformLayout();
			this.tableLayoutPanel2.ResumeLayout(false);
			this.ResumeLayout(false);

		}

		private void cpyToClip_Click(object sender, EventArgs e) {
			//Clipboard.SetText(textBox1.Text);
			string sMailToLink = @"mailto:kcjuntunen@amstore.com?subject=Error in PDF Archiver&body=" + textBox1.Text;
			System.Diagnostics.Process.Start(sMailToLink);
		}

		private void textBox1_TextChanged(object sender, EventArgs e) {

		}
	}
}
