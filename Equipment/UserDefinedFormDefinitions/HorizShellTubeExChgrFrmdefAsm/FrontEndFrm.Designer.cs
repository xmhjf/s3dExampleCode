namespace HSTEFrmDef
{
    partial class FrontEndFrm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrontEndFrm));
            this.sp3DUOmCtrl1 = new Ingr.SP3D.Common.Client.Sp3DUOmCtrl();
            this.sP3DOKApplyCtrl1 = new Ingr.SP3D.Common.Client.SP3DOKApplyCtrl();
            this.sp3DUOmCtrl2 = new Ingr.SP3D.Common.Client.Sp3DUOmCtrl();
            this.SuspendLayout();
            // 
            // sp3DUOmCtrl1
            // 
            this.sp3DUOmCtrl1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.sp3DUOmCtrl1.InterfaceName = "IJUAHeatExchangerFrontEnd";
            this.sp3DUOmCtrl1.Location = new System.Drawing.Point(352, 63);
            this.sp3DUOmCtrl1.Name = "sp3DUOmCtrl1";
            this.sp3DUOmCtrl1.PropertyName = "FrontEndFlangeDia";
            this.sp3DUOmCtrl1.Size = new System.Drawing.Size(129, 20);
            this.sp3DUOmCtrl1.TabIndex = 0;
            // 
            // sP3DOKApplyCtrl1
            // 
            this.sP3DOKApplyCtrl1.BackColor = System.Drawing.SystemColors.Window;
            this.sP3DOKApplyCtrl1.Location = new System.Drawing.Point(352, 262);
            this.sP3DOKApplyCtrl1.MaximumSize = new System.Drawing.Size(169, 29);
            this.sP3DOKApplyCtrl1.MinimumSize = new System.Drawing.Size(169, 29);
            this.sP3DOKApplyCtrl1.Name = "sP3DOKApplyCtrl1";
            this.sP3DOKApplyCtrl1.Size = new System.Drawing.Size(169, 29);
            this.sP3DOKApplyCtrl1.TabIndex = 1;
            // 
            // sp3DUOmCtrl2
            // 
            this.sp3DUOmCtrl2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.sp3DUOmCtrl2.InterfaceName = "IJUAHeatExchangerFrontEnd";
            this.sp3DUOmCtrl2.Location = new System.Drawing.Point(215, 145);
            this.sp3DUOmCtrl2.Name = "sp3DUOmCtrl2";
            this.sp3DUOmCtrl2.PropertyName = "FrontEndLength1";
            this.sp3DUOmCtrl2.Size = new System.Drawing.Size(123, 17);
            this.sp3DUOmCtrl2.TabIndex = 2;
            // 
            // FrontEndFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(533, 303);
            this.Controls.Add(this.sp3DUOmCtrl2);
            this.Controls.Add(this.sP3DOKApplyCtrl1);
            this.Controls.Add(this.sp3DUOmCtrl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrontEndFrm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "FrontEndFrm";
            this.ResumeLayout(false);

        }

        #endregion

        private Ingr.SP3D.Common.Client.Sp3DUOmCtrl sp3DUOmCtrl1;
        private Ingr.SP3D.Common.Client.SP3DOKApplyCtrl sP3DOKApplyCtrl1;
        private Ingr.SP3D.Common.Client.Sp3DUOmCtrl sp3DUOmCtrl2;

    }
}