namespace HSTEFrmDef
{
    partial class AssemblyForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AssemblyForm));
            this.FrontEndBtn = new System.Windows.Forms.Button();
            this.RearEndBtn = new System.Windows.Forms.Button();
            this.sp3DUOmCtrl1 = new Ingr.SP3D.Common.Client.Sp3DUOmCtrl();
            this.sp3DUOmCtrl2 = new Ingr.SP3D.Common.Client.Sp3DUOmCtrl();
            this.sP3DOKApplyCtrl1 = new Ingr.SP3D.Common.Client.SP3DOKApplyCtrl();
            this.SuspendLayout();
            // 
            // FrontEndBtn
            // 
            this.FrontEndBtn.Location = new System.Drawing.Point(750, 12);
            this.FrontEndBtn.Name = "FrontEndBtn";
            this.FrontEndBtn.Size = new System.Drawing.Size(87, 35);
            this.FrontEndBtn.TabIndex = 0;
            this.FrontEndBtn.Text = "Front End";
            this.FrontEndBtn.UseMnemonic = false;
            this.FrontEndBtn.UseVisualStyleBackColor = true;
            this.FrontEndBtn.Click += new System.EventHandler(this.FrontEndBtn_Click);
            // 
            // RearEndBtn
            // 
            this.RearEndBtn.Location = new System.Drawing.Point(750, 65);
            this.RearEndBtn.Name = "RearEndBtn";
            this.RearEndBtn.Size = new System.Drawing.Size(87, 34);
            this.RearEndBtn.TabIndex = 1;
            this.RearEndBtn.Text = "Rear End";
            this.RearEndBtn.UseMnemonic = false;
            this.RearEndBtn.UseVisualStyleBackColor = true;
            this.RearEndBtn.Click += new System.EventHandler(this.RearEndBtn_Click);
            // 
            // sp3DUOmCtrl1
            // 
            this.sp3DUOmCtrl1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.sp3DUOmCtrl1.InterfaceName = "IJUAHeatExchanger";
            this.sp3DUOmCtrl1.Location = new System.Drawing.Point(289, 280);
            this.sp3DUOmCtrl1.Name = "sp3DUOmCtrl1";
            this.sp3DUOmCtrl1.PropertyName = "ExchangerLength";
            this.sp3DUOmCtrl1.Size = new System.Drawing.Size(87, 20);
            this.sp3DUOmCtrl1.TabIndex = 2;
            // 
            // sp3DUOmCtrl2
            // 
            this.sp3DUOmCtrl2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.sp3DUOmCtrl2.InterfaceName = "IJUAHeatExchanger";
            this.sp3DUOmCtrl2.Location = new System.Drawing.Point(400, 181);
            this.sp3DUOmCtrl2.Name = "sp3DUOmCtrl2";
            this.sp3DUOmCtrl2.PropertyName = "ExchangerDiameter";
            this.sp3DUOmCtrl2.Size = new System.Drawing.Size(94, 19);
            this.sp3DUOmCtrl2.TabIndex = 3;
            // 
            // sP3DOKApplyCtrl1
            // 
            this.sP3DOKApplyCtrl1.BackColor = System.Drawing.SystemColors.Control;
            this.sP3DOKApplyCtrl1.Location = new System.Drawing.Point(668, 294);
            this.sP3DOKApplyCtrl1.MaximumSize = new System.Drawing.Size(169, 29);
            this.sP3DOKApplyCtrl1.MinimumSize = new System.Drawing.Size(169, 29);
            this.sP3DOKApplyCtrl1.Name = "sP3DOKApplyCtrl1";
            this.sP3DOKApplyCtrl1.Size = new System.Drawing.Size(169, 29);
            this.sP3DOKApplyCtrl1.TabIndex = 4;
            // 
            // AssemblyForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(838, 335);
            this.Controls.Add(this.sP3DOKApplyCtrl1);
            this.Controls.Add(this.sp3DUOmCtrl2);
            this.Controls.Add(this.sp3DUOmCtrl1);
            this.Controls.Add(this.RearEndBtn);
            this.Controls.Add(this.FrontEndBtn);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AssemblyForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "AssemblyForm";
            this.Load += new System.EventHandler(this.AssemblyForm_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button FrontEndBtn;
        private System.Windows.Forms.Button RearEndBtn;
        private Ingr.SP3D.Common.Client.Sp3DUOmCtrl sp3DUOmCtrl1;
        private Ingr.SP3D.Common.Client.Sp3DUOmCtrl sp3DUOmCtrl2;
        private Ingr.SP3D.Common.Client.SP3DOKApplyCtrl sP3DOKApplyCtrl1;
    }
}