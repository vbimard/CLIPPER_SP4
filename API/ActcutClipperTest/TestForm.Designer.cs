namespace ActcutClipperTest
{
    partial class TestForm
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur Windows Form

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            this.BtnInitApi = new System.Windows.Forms.Button();
            this.BtnGetQuoteApi = new System.Windows.Forms.Button();
            this.BtnGetQuote = new System.Windows.Forms.Button();
            this.BtnInit = new System.Windows.Forms.Button();
            this.BtnExportQuote = new System.Windows.Forms.Button();
            this.BtnExit = new System.Windows.Forms.Button();
            this.BtnExportQuoteApi = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // BtnInitApi
            // 
            this.BtnInitApi.Location = new System.Drawing.Point(8, 24);
            this.BtnInitApi.Name = "BtnInitApi";
            this.BtnInitApi.Size = new System.Drawing.Size(75, 23);
            this.BtnInitApi.TabIndex = 0;
            this.BtnInitApi.Text = "Init";
            this.BtnInitApi.UseVisualStyleBackColor = true;
            this.BtnInitApi.Click += new System.EventHandler(this.BtnInitApi_Click);
            // 
            // BtnGetQuoteApi
            // 
            this.BtnGetQuoteApi.Location = new System.Drawing.Point(96, 24);
            this.BtnGetQuoteApi.Name = "BtnGetQuoteApi";
            this.BtnGetQuoteApi.Size = new System.Drawing.Size(75, 23);
            this.BtnGetQuoteApi.TabIndex = 1;
            this.BtnGetQuoteApi.Text = "Get Quote";
            this.BtnGetQuoteApi.UseVisualStyleBackColor = true;
            this.BtnGetQuoteApi.Click += new System.EventHandler(this.BtnGetQuoteApi_Click);
            // 
            // BtnGetQuote
            // 
            this.BtnGetQuote.Location = new System.Drawing.Point(96, 24);
            this.BtnGetQuote.Name = "BtnGetQuote";
            this.BtnGetQuote.Size = new System.Drawing.Size(75, 23);
            this.BtnGetQuote.TabIndex = 3;
            this.BtnGetQuote.Text = "Get Quote";
            this.BtnGetQuote.UseVisualStyleBackColor = true;
            this.BtnGetQuote.Click += new System.EventHandler(this.BtnGetQuote_Click);
            // 
            // BtnInit
            // 
            this.BtnInit.Location = new System.Drawing.Point(8, 24);
            this.BtnInit.Name = "BtnInit";
            this.BtnInit.Size = new System.Drawing.Size(75, 23);
            this.BtnInit.TabIndex = 2;
            this.BtnInit.Text = "Init";
            this.BtnInit.UseVisualStyleBackColor = true;
            this.BtnInit.Click += new System.EventHandler(this.BtnInit_Click);
            // 
            // BtnExportQuote
            // 
            this.BtnExportQuote.Location = new System.Drawing.Point(184, 24);
            this.BtnExportQuote.Name = "BtnExportQuote";
            this.BtnExportQuote.Size = new System.Drawing.Size(75, 23);
            this.BtnExportQuote.TabIndex = 4;
            this.BtnExportQuote.Text = "Export Quote";
            this.BtnExportQuote.UseVisualStyleBackColor = true;
            this.BtnExportQuote.Click += new System.EventHandler(this.BtnExportQuote_Click);
            // 
            // BtnExit
            // 
            this.BtnExit.Location = new System.Drawing.Point(8, 56);
            this.BtnExit.Name = "BtnExit";
            this.BtnExit.Size = new System.Drawing.Size(75, 23);
            this.BtnExit.TabIndex = 5;
            this.BtnExit.Text = "Exit";
            this.BtnExit.UseVisualStyleBackColor = true;
            this.BtnExit.Click += new System.EventHandler(this.BtnExit_Click);
            // 
            // BtnExportQuoteApi
            // 
            this.BtnExportQuoteApi.Location = new System.Drawing.Point(184, 24);
            this.BtnExportQuoteApi.Name = "BtnExportQuoteApi";
            this.BtnExportQuoteApi.Size = new System.Drawing.Size(75, 23);
            this.BtnExportQuoteApi.TabIndex = 6;
            this.BtnExportQuoteApi.Text = "Export Quote";
            this.BtnExportQuoteApi.UseVisualStyleBackColor = true;
            this.BtnExportQuoteApi.Click += new System.EventHandler(this.BtnExportQuoteApi_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.BtnInitApi);
            this.groupBox1.Controls.Add(this.BtnGetQuoteApi);
            this.groupBox1.Controls.Add(this.BtnExportQuoteApi);
            this.groupBox1.Location = new System.Drawing.Point(8, 8);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(312, 56);
            this.groupBox1.TabIndex = 7;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Api Test";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.BtnInit);
            this.groupBox2.Controls.Add(this.BtnGetQuote);
            this.groupBox2.Controls.Add(this.BtnExportQuote);
            this.groupBox2.Controls.Add(this.BtnExit);
            this.groupBox2.Location = new System.Drawing.Point(8, 104);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(312, 88);
            this.groupBox2.TabIndex = 8;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Exe Test";
            // 
            // TestForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(329, 201);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "TestForm";
            this.Text = "Test Form";
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button BtnInitApi;
        private System.Windows.Forms.Button BtnGetQuoteApi;
        private System.Windows.Forms.Button BtnGetQuote;
        private System.Windows.Forms.Button BtnInit;
        private System.Windows.Forms.Button BtnExportQuote;
        private System.Windows.Forms.Button BtnExit;
        private System.Windows.Forms.Button BtnExportQuoteApi;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
    }
}

