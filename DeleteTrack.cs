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
    public partial class DeleteTrack : Form
    {
        public bool? deleteTrackOrPlaylist = null;

        public DeleteTrack()
        {
            InitializeComponent();
        }

        private void buttonYes_Click(object sender, EventArgs e)
        {
            deleteTrackOrPlaylist = true;
        }

        private void buttonNo_Click(object sender, EventArgs e)
        {
            deleteTrackOrPlaylist = false;
        }
    }
}
