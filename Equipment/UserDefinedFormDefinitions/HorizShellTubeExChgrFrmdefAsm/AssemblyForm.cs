using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace HSTEFrmDef
{
    public partial class AssemblyForm : Form
    {
        public AssemblyForm()
        {
            InitializeComponent();
        }

        private void FrontEndBtn_Click(object sender, EventArgs e)
        {
            FrontEndFrm ofrontEndForm = null;
            ofrontEndForm = new FrontEndFrm();
            Form oFrm = (Form)ofrontEndForm;
            oFrm.ShowDialog();  
        }

        private void RearEndBtn_Click(object sender, EventArgs e)
        {
            RearEndFrm oRearEndfrmClass = null;
            oRearEndfrmClass = new RearEndFrm();
            Form oFrm = (Form)oRearEndfrmClass;
            oFrm.ShowDialog();
        }

        private void AssemblyForm_Load(object sender, EventArgs e)
        {

        }

        private void sp3DUOmCtrl2_ValidateProperty(object sender, Ingr.SP3D.CustomFormDefinition.ValidatepropertyArgs e)
        {
            //MessageBox.Show("Called the Validation"); 
        }
    }
}
