using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MusicBeePlugin
{
    public partial class NewIPod : Form
    {
        public bool? proceed = null;

        public NewIPod()
        {
            InitializeComponent();
        }

        private void buttonContinue_Click(object sender, EventArgs e)
        {
            proceed = true;
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            proceed = false;
        }

        private void NewIPod_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (proceed == null)
                proceed = false;
        }
    }
}
