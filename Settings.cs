using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace MusicBeePlugin
{
    public partial class SettingsPlugin : Form
    {
        protected Plugin.PluginInfo about;
        private Plugin plugin;

        public SettingsPlugin()
        {
            InitializeComponent();
        }

        public SettingsPlugin(Plugin tagToolsPluginParam, Plugin.PluginInfo aboutParam)
        {
            InitializeComponent();

            plugin = tagToolsPluginParam;
            about = aboutParam;

            initializeForm();
        }

        protected void initializeForm()
        {
            versionLabel.Text += about.VersionMajor + "." + about.VersionMinor + "." + about.Revision;

            //deviceName.Checked = Plugin.SavedSettings.deviceName;
        }

        private void saveSettings()
        {
            //Plugin.SavedSettings.deviceName = deviceName.Checked;
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            saveSettings();
            Close();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void buttonReset_Click(object sender, EventArgs e)
        {
            var files = new string[0];

            if (Plugin.MbApiInterface.Library_QueryFiles("domain=Library"))
                files = Plugin.MbApiInterface.Library_QueryGetAllFiles().Split(Plugin.FilesSeparators, StringSplitOptions.RemoveEmptyEntries);

            if (files.Length != 0)
            {
                foreach (string file in files)
                {
                    Plugin.MbApiInterface.Library_SetDevicePersistentId(file, Plugin.DeviceIdType.AppleDevice, "");
                    Plugin.MbApiInterface.Library_SetDevicePersistentId(file, Plugin.DeviceIdType.AppleDevice2, "");
                    Plugin.MbApiInterface.Library_CommitTagsToFile(file);
                }
            }


            if (!System.IO.Directory.Exists(Plugin.ReencodedFilesStorage))
            {
                System.IO.Directory.Delete(Plugin.ReencodedFilesStorage, true);
                System.IO.Directory.CreateDirectory(Plugin.ReencodedFilesStorage);
            }
        }
    }
}
