using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SmartLink.EncryptTool
{
    public partial class Main : Form
    {
        private EncryptionService _encryptService;

        public Main()
        {
            _encryptService = new EncryptionService();
            InitializeComponent();
        }

        private void btnEncryptString_Click(object sender, EventArgs e)
        {
            try
            {
                txtTarget.Text = _encryptService.EncryptString(txtSource.Text);
            }
            catch (Exception ex)
            {

            }
        }

        private void btnDecryptString_Click(object sender, EventArgs e)
        {
            try
            {
                txtTarget.Text = _encryptService.DecryptString(txtSource.Text);
            }
            catch (Exception ex)
            {

            }
        }

    }
}
