using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PumpAsmFormDefinition
{
    public partial class AssemblyForm : Form
    {
        public AssemblyForm()
        {
            InitializeComponent();
        }

        private void AssemblyForm_Load(object sender, EventArgs e)
        {

        }

        private void sp3DUOmCtrl2_ValidateProperty(object sender, Ingr.SP3D.Common.Client.ValidatepropertyArgs e)
        {
            //implement end user validation routine here
        }
    }
}
