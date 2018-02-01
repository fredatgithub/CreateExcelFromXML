namespace XmlToExcel
{
    partial class FormMain
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
      this.buttonCreateExcel = new System.Windows.Forms.Button();
      this.buttonReadXml = new System.Windows.Forms.Button();
      this.SuspendLayout();
      // 
      // buttonCreateExcel
      // 
      this.buttonCreateExcel.Location = new System.Drawing.Point(60, 57);
      this.buttonCreateExcel.Margin = new System.Windows.Forms.Padding(2);
      this.buttonCreateExcel.Name = "buttonCreateExcel";
      this.buttonCreateExcel.Size = new System.Drawing.Size(83, 28);
      this.buttonCreateExcel.TabIndex = 0;
      this.buttonCreateExcel.Text = "create excel";
      this.buttonCreateExcel.UseVisualStyleBackColor = true;
      this.buttonCreateExcel.Click += new System.EventHandler(this.button1_Click);
      // 
      // buttonReadXml
      // 
      this.buttonReadXml.Location = new System.Drawing.Point(180, 57);
      this.buttonReadXml.Margin = new System.Windows.Forms.Padding(2);
      this.buttonReadXml.Name = "buttonReadXml";
      this.buttonReadXml.Size = new System.Drawing.Size(83, 28);
      this.buttonReadXml.TabIndex = 1;
      this.buttonReadXml.Text = "Read Xml";
      this.buttonReadXml.UseVisualStyleBackColor = true;
      this.buttonReadXml.Click += new System.EventHandler(this.ButtonReadXml_Click);
      // 
      // FormMain
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(348, 129);
      this.Controls.Add(this.buttonReadXml);
      this.Controls.Add(this.buttonCreateExcel);
      this.Margin = new System.Windows.Forms.Padding(2);
      this.Name = "FormMain";
      this.Text = "XML to Excel";
      this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonCreateExcel;
    private System.Windows.Forms.Button buttonReadXml;
  }
}

