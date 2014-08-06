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
    public partial class WaitingForIPod : Form
    {
        public bool? proceed = null;

        public WaitingForIPod()
        {
            InitializeComponent();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            proceed = false;
        }

        private void Waiting_FormClosing(object sender, FormClosingEventArgs e)
        {
            proceed = false;
        }

        private void buttonOffline_Click(object sender, EventArgs e)
        {
            proceed = true;
        }
    }
}
