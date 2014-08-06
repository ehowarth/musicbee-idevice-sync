using System;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using System.Text;
using System.Text.RegularExpressions;
using System.ComponentModel;
using iTunesLib;


namespace MusicBeePlugin
{
    public partial class Plugin
    {
        public static bool DeveloperMode = false;
        private bool uninstalled = false;

        public static MusicBeeApiInterface MbApiInterface;
        private PluginInfo about = new PluginInfo();

        public static Form MbForm;

        public static string Language;
        private string pluginSettingsFileName;
        public static string ReencodedFilesStorage;

        public static char[] FilesSeparators = { '\0' };

        public const char MultipleItemsSplitterId = (char)0;
        public const char GuestId = (char)1;
        public const char PerformerId = (char)2;
        public const char RemixerId = (char)4;
        public const char EndOfStringId = (char)8;

        public static int Zero = 0; //For generating exception by dividing by 0


        //Defaults for controls
        public struct SavedSettingsType
        {
            public string deviceName;

            public string multipleItemsSplitterChar1;
            public string multipleItemsSplitterChar2;
        }
        
        public static SavedSettingsType SavedSettings;


        //Localizable strings

        //Plugin localizable strings
        private string pluginName;
        private string description;

        private string commandName;
        private string commandDescription;

        private string msgOffline;
        public static string msgFailedToAddTrackToiTunesLibrary;
        public static string msgNoAppropriateCategoryFoundIniTunesLibrary;
        public static string msgDeletePlaylist;
        public static string msgDeleteTrack;
        public static string msgFromDevice;
        public static string msgSomeDuplicatedTracksWereSkipped;
        public static string msgTrackSourceFileNotFound;
        public static string msgRemovingTrack;
        public static string msgAddingTrackToPlaylist;
        public static string msgSyncingITunesWithiPodIPhone;
        public static string msgSyncingPlayCountsAndRatingsFROMiPod;
        public static string msgFillingSyncedPlaylists;

        //Displayed text
        public string getTagRepresentation(string tag)
        {
            return replaceMSIds(removeMSIdAtTheEndOfString(removeRoleIds(tag)));
        }

        public string removeMSIdAtTheEndOfString(string value)
        {
            value += EndOfStringId;
            value = value.Replace("" + MultipleItemsSplitterId + EndOfStringId, "");
            value = value.Replace("" + EndOfStringId, "");

            return value;
        }

        public string replaceMSChars(string value)
        {
            value = value.Replace(SavedSettings.multipleItemsSplitterChar2, "" + MultipleItemsSplitterId);
            value = value.Replace(SavedSettings.multipleItemsSplitterChar1, "" + MultipleItemsSplitterId);

            return value;
        }

        public string replaceMSIds(string value)
        {
            value = value.Replace("" + MultipleItemsSplitterId, SavedSettings.multipleItemsSplitterChar2);

            return value;
        }

        public string removeRoleIds(string value)
        {
            value = value.Replace("" + GuestId, "");
            value = value.Replace("" + PerformerId, "");
            value = value.Replace("" + RemixerId, "");

            return value;
        }

        public void turnOnEventHandler(object sender, EventArgs e)
        {
            while (ITunes.IsInitialised == null) ;

            if (ITunes.IsInitialised == false)
            {
                Thread initThread = new Thread(ITunes.Initialise);
                initThread.Start();
            }
            else
            {
                Close(PluginCloseReason.UserDisabled);
            }
                
        }

        public PluginInfo Initialise(IntPtr apiInterfacePtr)
        {
            MbApiInterface = new MusicBeeApiInterface();
            MbApiInterface.Initialise(apiInterfacePtr);

            ITunes.SendNotificationHandler = MbApiInterface.MB_SendNotification;


            //Localizable strings

            //Plugin localizable strings
            pluginName = "iPod & iPhone Driver";
            description = "Allows to use iPod/iPhone by the means of iTunes";

            commandName = "Turn iPod && iPhone Driver on";
            commandDescription = "iPod & iPhone Driver: Turn on/off";

            msgOffline = " - [Offline]";
            msgFailedToAddTrackToiTunesLibrary = "Failed to add track to iTunes library";
            msgNoAppropriateCategoryFoundIniTunesLibrary = "No appropriate category found in iTunes library";
            msgDeletePlaylist = "Delete playlist \"";
            msgDeleteTrack = "Delete track \"";
            msgFromDevice = "\" from device?";
            msgSomeDuplicatedTracksWereSkipped = "Some duplicated tracks were skipped.";
            msgTrackSourceFileNotFound = "Track source file not found!";
            msgRemovingTrack = "Removing track ";
            msgAddingTrackToPlaylist = "Adding track to playlist ";
            msgSyncingITunesWithiPodIPhone = "Syncing iTunes with iPod/iPhone...";
            msgSyncingPlayCountsAndRatingsFROMiPod = "Syncing play counts and/or ratings from iPod/iPhone...";
            msgFillingSyncedPlaylists = "Filling synced playlists...";

            //Defaults for controls
            SavedSettings = new SavedSettingsType();

            //Lets set initial defaults
            //SavedSettings.menuPlacement = 2;


            //Lets try to read defaults for controls from settings file
            pluginSettingsFileName = System.IO.Path.Combine(MbApiInterface.Setting_GetPersistentStoragePath(), "mb_iPod&iPhoneDriver.Settings.xml");

            Encoding unicode = Encoding.UTF8;
            System.IO.FileStream stream = System.IO.File.Open(pluginSettingsFileName, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.Read, System.IO.FileShare.None);
            System.IO.StreamReader file = new System.IO.StreamReader(stream, unicode);

            System.Xml.Serialization.XmlSerializer controlsDefaultsSerializer = null;
            try
            {
                controlsDefaultsSerializer = new System.Xml.Serialization.XmlSerializer(typeof(SavedSettingsType));
            }
            catch (Exception e)
            {
                MessageBox.Show("" + e.Message);
            }

            try 
            {
                SavedSettings = (SavedSettingsType)controlsDefaultsSerializer.Deserialize(file);
            }
            catch 
            {
                //Ignore...
            };

            file.Close();


            ReencodedFilesStorage = MbApiInterface.Setting_GetPersistentStoragePath() + "iPod & iPhone Driver";
            if (!System.IO.Directory.Exists(ReencodedFilesStorage))
            {
                System.IO.Directory.CreateDirectory(ReencodedFilesStorage);
            }


            Language = Thread.CurrentThread.CurrentCulture.TwoLetterISOLanguageName;

            if (!System.IO.File.Exists(Application.StartupPath + @"\Plugins\ru\mb_iPod&iPhoneDriver.resources.dll")) Language = "en"; //For testing

            if (Language == "ru")
            {
                //Plugin localizable strings
                pluginName = "Драйвер iPod & iPhone";
                description = "Плагин позволяет использовать iPod/iPhone посредством iTunes";

                commandName = "Включить драйвер iPod && iPhone";
                commandDescription = "Драйвер iPod & iPhone: Включить/выключить";

                msgOffline = " - [Автономно]";
                msgFailedToAddTrackToiTunesLibrary = "Не удалось добавить композицию в медиатеку iTunes";
                msgNoAppropriateCategoryFoundIniTunesLibrary = "Не найдено подходящей категории в медиатеке iTunes";
                msgDeletePlaylist = "Удалить плейлист \"";
                msgDeleteTrack = "Удалить композицию \"";
                msgFromDevice = "\" из устройства?";
                msgSomeDuplicatedTracksWereSkipped = "Некоторые повторяющиеся композиции были пропущены.";
                msgTrackSourceFileNotFound = "Исходный файл композиции не найден!";
                msgRemovingTrack = "Удаляется композиция ";
                msgAddingTrackToPlaylist = "В плейлист добавляется композиция ";
                msgSyncingITunesWithiPodIPhone = "Синхронизация iTunes с iPod/iPhone...";
                msgSyncingPlayCountsAndRatingsFROMiPod = "Синхронизация счетчиков воспроизведения и рейтингов из iPod/iPhone...";
                msgFillingSyncedPlaylists = "Заполнение синхронизированных плейлистов...";
            }
            else
            {
                Language = "en";
            }


            //Lets reset invalid defaults for controls
            if (SavedSettings.deviceName == null) SavedSettings.deviceName = "";

            //Again lets reset invalid defaults for controls
            if ("" + SavedSettings.multipleItemsSplitterChar2 == "")
            {
                SavedSettings.multipleItemsSplitterChar1 = ";";
                SavedSettings.multipleItemsSplitterChar2 = "; ";
            }

            //Final initialization
            MbForm = (Form)Form.FromHandle(MbApiInterface.MB_GetWindowHandle());

            MbApiInterface.MB_AddMenuItem("mnuTools/" + commandName, commandDescription, turnOnEventHandler);

            if (System.IO.File.Exists(Application.StartupPath + @"\Plugins\DevMode.txt"))
                DeveloperMode = true;


            about.PluginInfoVersion = PluginInfoVersion;
            about.Name = pluginName;
            about.Description = description;
            about.Author = "boroda74";
            about.TargetApplication = "iTunes";   // current only applies to artwork, lyrics or instant messenger name that appears in the provider drop down selector or target Instant Messenger
            about.Type = PluginType.Storage;
            about.VersionMajor = (short)System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Major;  // .net version
            about.VersionMinor = (short)System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Minor; // plugin version
            about.Revision = (short)System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Build; // number of days since 2000-01-01 at build time
            about.MinInterfaceVersion = 28;
            about.MinApiRevision = 32;
            about.ReceiveNotifications = ReceiveNotificationFlags.StartupOnly;
            about.ConfigurationPanelHeight = 0;   // height in pixels that musicbee should reserve in a panel for config settings. When set, a handle to an empty panel will be passed to the Configure function


            return about;
        }

        public bool Configure(IntPtr panelHandle)
        {
            // save any persistent settings in a sub-folder of this path
            //string dataPath = mbApiInterface.Setting_GetPersistentStoragePath();
            // panelHandle will only be set if you set about.ConfigurationPanelHeight to a non-zero value
            // keep in mind the panel width is scaled according to the font the user has selected
            // if about.ConfigurationPanelHeight is set to 0, you can display your own popup window

            SettingsPlugin tagToolsForm = new SettingsPlugin(this, about);
            tagToolsForm.ShowDialog();
            SaveSettings();

            return true;
        }
       
        // called by MusicBee when the user clicks Apply or Save in the MusicBee Preferences screen.
        // its up to you to figure out whether anything has changed and needs updating
        public void SaveSettings()
        {
            Encoding unicode = Encoding.UTF8;
            System.IO.FileStream stream = System.IO.File.Open(pluginSettingsFileName, System.IO.FileMode.Create, System.IO.FileAccess.Write, System.IO.FileShare.None);
            System.IO.StreamWriter file = new System.IO.StreamWriter(stream, unicode);

            System.Xml.Serialization.XmlSerializer controlsDefaultsSerializer = new System.Xml.Serialization.XmlSerializer(typeof(SavedSettingsType));

            controlsDefaultsSerializer.Serialize(file, SavedSettings);

            file.Close();
        }

        // MusicBee is closing the plugin (plugin is being disabled by user or MusicBee is shutting down)
        public void Close(PluginCloseReason reason)
        {
            Eject();
        }

        // uninstall this plugin - clean up any persisted files
        public void Uninstall()
        {
            //Delete settings file
            Eject();

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


            if (System.IO.File.Exists(pluginSettingsFileName))
                System.IO.File.Delete(pluginSettingsFileName);

            if (System.IO.Directory.Exists(ReencodedFilesStorage))
                System.IO.Directory.Delete(ReencodedFilesStorage, true);

            uninstalled = true;
        }

        // receive event notifications from MusicBee
        // you need to set about.ReceiveNotificationFlags = PlayerEvents to receive all notifications, and not just the startup event
        public void ReceiveNotification(string sourceFileUrl, NotificationType type)
        {
            // perform some action depending on the notification type
            switch (type)
            {
                case NotificationType.PluginStartup:
                    // perform startup initialisation
                    break;
            }
        }


        // return an array of lyric or artwork provider names this plugin supports
        // the providers will be iterated through one by one and passed to the RetrieveLyrics/ RetrieveArtwork function in order set by the user in the MusicBee Tags(2) preferences screen until a match is found
        public string[] GetProviders()
        {
            return null;
        }

        // return lyrics for the requested artist/title from the requested provider
        // only required if PluginType = LyricsRetrieval
        // return null if no lyrics are found
        public string RetrieveLyrics(string sourceFileUrl, string artist, string trackTitle, string album, bool synchronisedPreferred, string provider)
        {
            return null;
        }

        // return Base64 string representation of the artwork binary data from the requested provider
        // only required if PluginType = ArtworkRetrieval
        // return null if no artwork is found
        public string RetrieveArtwork(string sourceFileUrl, string albumArtist, string album, string provider)
        {
            //Return Convert.ToBase64String(artworkBinaryData)
            return null;
        }

        [Flags()]
        public enum SynchronisationSupportedFormats : uint
        {
            None = 0,
            SyncWavSupported = 0x100,
            SyncMp3Supported = 0x200,
            SyncAacSupported = 0x400,
            SyncWmaSupported = 0x800,
            SyncOggSupported = 0x1000,
            SyncOpusSupported = 0x2000,
            SyncMpcSupported = 0x4000,
            SyncFlacSupported = 0x8000,
            SyncAlacSupported = 0x10000,
            SyncWavPackSupported = 0x20000,
            SyncTakSupported = 0x40000
        }

        [Flags()]
        public enum SynchronisationSettings
        {
            None = 0,
            SyncRemoveMissingFiles = 0x1,
            SyncRating2Way = 0x2,
            SyncPlayCount2Way = 0x4,
            SyncConfirmRemoveFiles = 0x8,
            SyncPlaylists2Way = 0x10
        }

        public enum SynchronisationCategory
        {
            Music = 0,
            Audiobook = 1,
            Podcast = 2,
            Video = 3,
            Playlist = 10
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct DeviceProperties
        {
            private int version;
            public string DeviceName;
            // this should be a 64x64 bitmap
            public Bitmap DeviceIcon64;
            public string Manufacturer;
            public string Model;
            public string FirmwareVersion;
            public int BatteryLevel;
            public ulong FreeSpace;
            public ulong TotalSpace;
            // if yes, a Music node and if enabled Audiobook, Podcast, Video and playlist nodes are displayed in the Devices tree
            public bool ShowCategoryNodes;
            // device supports audio books as a category - if no, the files are synchronised to the music library
            public bool AudiobooksSupported;
            // device supports podcasts as a category - if no, the files are synchronised to the music library
            public bool PodcastsSupported;
            // device supports video as a category - if no, the files are synchronised to the music library
            public bool VideoSupported;
            // device supports a folder structure and allows the user to specify a naming template for the files
            public bool OrganisedFoldersSupported;
            // when enabled you need to query the MetaDataType.Artwork tag which will return null, "embeded" or the file location of the artwork. When not enabled, MusicBee forces the artwork to always be embeded (a temporary file is created if the file does not already have embeded artwork)
            public bool SyncExternalArtwork;
            // allows the user can choose whether files not on the sync list are to be removed from the device. If not enabled, its up to the plugin what to do
            public bool SyncAllowFileRemoval;
            // allows the user to tick 2-way rating sync in the device preferences
            public bool SyncAllowRating2Way;
            // allows the user to tick 2-way play count sync in the device preferences
            public bool SyncAllowPlayCount2Way;
            public SynchronisationSupportedFormats SupportedFormats;
            // allows the user to tick 2-way playlist sync in the device preferences
            public bool SyncAllowPlaylists2Way;
        }

        public bool GetDeviceProperties(IntPtr handle)
        {
            DeviceProperties properties = new DeviceProperties();
            if (!ITunes.IgnoreIPod)
            {
                properties.DeviceName = pluginName;
                properties.FirmwareVersion = ITunes.IPodSource.SoftwareVersion;
                properties.FreeSpace = (ulong)ITunes.IPodSource.FreeSpace;
                properties.TotalSpace = (ulong)ITunes.IPodSource.Capacity;
            }
            else
            {
                properties.DeviceName = pluginName;
            }

            properties.ShowCategoryNodes = true;
            properties.AudiobooksSupported = false;
            properties.PodcastsSupported = false;
            properties.VideoSupported = false;
            properties.Manufacturer = "Apple Inc.";
            properties.SyncExternalArtwork = false;
            properties.SyncAllowFileRemoval = true;
            properties.SyncAllowRating2Way = true;
            properties.SyncAllowPlayCount2Way = true;
            properties.SyncAllowPlaylists2Way = false;

            System.Resources.ResourceManager resourceManager = new System.Resources.ResourceManager("MusicBeePlugin.Images", System.Reflection.Assembly.GetExecutingAssembly());
            properties.DeviceIcon64 = (Bitmap)resourceManager.GetObject("iTunes64");

            properties.SupportedFormats = SynchronisationSupportedFormats.SyncMp3Supported | SynchronisationSupportedFormats.SyncAacSupported | SynchronisationSupportedFormats.SyncAlacSupported;
            Marshal.StructureToPtr(properties, handle, false);

            return true;
        }

        public bool Synchronise(SynchronisationSettings flags, KeyValuePair<int, string[]>[] files)
        {
            while (ITunes.IsInitialised == null) ;

            try
            {
                return ITunes.Synchronise(flags, files);
            }
            catch (Exception ex)
            {
                ITunes.SendNotificationHandler.Invoke(CallbackType.StorageFailed);
                ITunes.lastEx = ex;
                return false;
            }
        }

        public void SynchroniseAbort()
        {
            // called when user has requested the synchronisation aborts
            ITunes.AbortSynchronization = true;
        }

        public bool Eject()
        {
            // called when user has requested the device be ejected
            ITunes.Close();

            if (!uninstalled)
                SaveSettings();

            return true;
        }

        public void Refresh()
        {
            ITunes.Refresh();
        }

        public bool IsReady()
        {
            return ITunes.IsInitialised == true ? true : false;
        }

        public Bitmap GetIcon()
        {
            System.Resources.ResourceManager resourceManager = new System.Resources.ResourceManager("MusicBeePlugin.Images", System.Reflection.Assembly.GetExecutingAssembly());
            return (Bitmap)resourceManager.GetObject("iTunes");
        }

        public bool FolderExists(string path)
        {
            return ITunes.FolderExists(path);
        }

        public string[] GetFolders(string path)
        {
            return ITunes.GetFolders(path);
        }

        public KeyValuePair<byte, string>[][] GetFiles(string path)
        {
            return ITunes.GetFiles(path);
        }

        public bool FileExists(string url)
        {
            return ITunes.FileExists(url);
        }

        public KeyValuePair<byte, string>[] GetFile(string url)
        {
            return ITunes.GetFile(url);
        }

        public byte[] GetFileArtwork(string url)
        {
            return ITunes.GetFileArtwork(url);
        }

        public KeyValuePair<string, string>[] GetPlaylists()
        {
            return ITunes.GetPlaylists();
        }

        public KeyValuePair<byte, string>[][] GetPlaylistFiles(string id)
        {
            return ITunes.GetPlaylistFiles(id);
        }

        public System.IO.Stream GetStream(string url)
        {
            return ITunes.GetStream(url);
        }

        public Exception GetError()
        {
            return ITunes.GetError();
        }
    }

    class ITunes
    {
        public static bool? IsInitialised = false;
        public static Plugin.MB_SendNotificationDelegate SendNotificationHandler;

        private static iTunesApp iTunes;
        private const int TagCount = 23;
        private const int AddOperationTimeout = 5; //Timeout of adding track to iTunes (1 means 0.5 sec, 2 - 1.5 sec, 3 - 2.5 sec, 30 - 29.5 sec, etc.)
        public static Exception lastEx = null;

        public static bool IgnoreIPod = false; //For testing
        public static IITIPodSource IPodSource = null;
        private static IITUserPlaylist MusicPlaylist = null;
        private static IITUserPlaylist AudiobooksPlaylist = null;
        private static IITUserPlaylist PodcastsPlaylist = null;
        private static IITUserPlaylist VideoPlaylist = null;


        public static bool SynchronizationInProgress = false;
        public static bool AbortSynchronization = false;

        private delegate void MBInvokedFunction();
        private static MBInvokedFunction InvokedFunction;

        private static SortedDictionary<string,int[]> ITTracksCache = new SortedDictionary<string,int[]>();

        private static string GetTrackRepresentation(IITTrack track)
        {
            string trackRepresentation = "";
            string displayedArtist;
            string album;
            string title;
            string diskNo;
            string trackNo;

            displayedArtist = track.Artist;
            album = track.Album;
            title = track.Name;
            diskNo = track.DiscNumber.ToString();
            trackNo = track.TrackNumber.ToString();

            trackRepresentation += diskNo;
            trackRepresentation += (trackRepresentation == "") ? (trackNo) : ("-" + trackNo);
            trackRepresentation += (trackRepresentation == "") ? (displayedArtist) : (". " + displayedArtist);
            trackRepresentation += (trackRepresentation == "") ? (album) : (" - " + album);
            trackRepresentation += (trackRepresentation == "") ? (title) : (" - " + title);

            return trackRepresentation;
        }

        public static void Initialise()
        {
            try
            {
                IsInitialised = null;

                iTunes = new iTunesApp();
                iTunes.BrowserWindow.Minimized = true;


                WaitingForIPod waitingWindow = new WaitingForIPod();
                Plugin.MbForm.AddOwnedForm(waitingWindow);

                InvokedFunction = waitingWindow.Show;
                Plugin.MbForm.Invoke(InvokedFunction);

                while (IPodSource == null && waitingWindow.proceed == null)
                {
                    for (int i = 1; i <= iTunes.Sources.Count; i++)
                    {
                        if (iTunes.Sources[i].Kind == ITSourceKind.ITSourceKindIPod)
                            IPodSource = (IITIPodSource)iTunes.Sources[i];
                    }
                }

                if (waitingWindow.proceed == true)
                    IgnoreIPod = true;

                InvokedFunction = waitingWindow.Close;
                Plugin.MbForm.Invoke(InvokedFunction);

                if (IPodSource == null && !IgnoreIPod)
                {
                    iTunes.Quit();
                    IsInitialised = false;
                    //SendNotificationHandler.Invoke(Plugin.CallbackType.StorageFailed);
                    return;
                }

                //if (Plugin.SavedSettings.deviceName != "" && IPodSource != null && IPodSource.Name != Plugin.SavedSettings.deviceName)
                //{
                //    NewIPod newIPodWindow = new NewIPod();
                //    Plugin.MbForm.AddOwnedForm(newIPodWindow);

                //    InvokedFunction = newIPodWindow.Show;
                //    Plugin.MbForm.Invoke(InvokedFunction);

                //    while (newIPodWindow.proceed == null) ;

                //    InvokedFunction = newIPodWindow.Close;
                //    Plugin.MbForm.Invoke(InvokedFunction);


                //    if (newIPodWindow.proceed == false)
                //    {
                //        IsInitialised = false;
                //        iTunes.Quit();
                //        return;
                //    }
                //}


                if (IPodSource != null)
                {
                    Plugin.SavedSettings.deviceName = IPodSource.Name;
                    iTunes.UpdateIPod();
                }


                foreach (IITPlaylist playlist in iTunes.LibrarySource.Playlists)
                {
                    if (playlist.Kind == ITPlaylistKind.ITPlaylistKindUser)
                    {
                        if (((IITUserPlaylist)playlist).SpecialKind == ITUserPlaylistSpecialKind.ITUserPlaylistSpecialKindMusic)
                            MusicPlaylist = (IITUserPlaylist)playlist;
                        else if (((IITUserPlaylist)playlist).SpecialKind == ITUserPlaylistSpecialKind.ITUserPlaylistSpecialKindBooks)
                            AudiobooksPlaylist = (IITUserPlaylist)playlist;
                        else if (((IITUserPlaylist)playlist).SpecialKind == ITUserPlaylistSpecialKind.ITUserPlaylistSpecialKindPodcasts)
                            PodcastsPlaylist = (IITUserPlaylist)playlist;
                        else if (((IITUserPlaylist)playlist).SpecialKind == ITUserPlaylistSpecialKind.ITUserPlaylistSpecialKindMovies)
                            VideoPlaylist = (IITUserPlaylist)playlist;
                    }
                }


                IsInitialised = true;
            }
            catch (Exception ex)
            {
                IsInitialised = false;
                lastEx = ex;
                MessageBox.Show(ex.Message);

                iTunes.Quit();
            }
            finally
            {
                if (IsInitialised == true)
                    SendNotificationHandler.Invoke(Plugin.CallbackType.StorageReady);
                //else
                //    SendNotificationHandler.Invoke(Plugin.CallbackType.StorageFailed);
            }

            return;
        }

        public static void Close()
        {
            while (IsInitialised == null) ;

            if (IsInitialised == true)
            {
                if (!IgnoreIPod)
                    IPodSource.EjectIPod();

                iTunes.Quit();
                IsInitialised = false;
                ITunes.SendNotificationHandler.Invoke(Plugin.CallbackType.StorageEject);
            }

            IgnoreIPod = false;
        }

        public static bool Synchronise(Plugin.SynchronisationSettings flags, KeyValuePair<int, string[]>[] files)
        {
            // flags has bit indicators whether 2 way rating and/or playcount is requested (only set if enabled in the device properties)
            // for files():
            //   Key - the SynchronisationCategory the file should be sychronised to if appropriate for the device
            //   Value is 3 strings
            //   (0) - source file or playlist to be synchronised - use to query MusicBee for file tags, or files in a playlist
            //   (1) - the filename extension of the file to be synchronised - normally the same as the extension for (0) but if the file would need to be re-encoded to meet the user's synch preferences then the extension for the encoded file
            //   (2) - if SyncOrganisedFolders is enabled, filename as formatted by a naming template otherwise null
            // for each file that is synchronised, call Sync_FileStart(filename(0))
            //   MusicBee will determine if the file needs to be re-encoded depending on the user synch settings and if so the returned filename will be the temporary encoded filename
            //   call Sync_FileEnd(filename(0), success, errorMessage) when that file has been synched or not
            // return true if all files synchronised ok

            if (SynchronizationInProgress)
            {
                lastEx = null;
                return false;
            }

            SynchronizationInProgress = true;

            SortedDictionary<string, IITPlaylist> playlists = new SortedDictionary<string, IITPlaylist>();
            SortedDictionary<string, object> iTunesPlaylists = new SortedDictionary<string, object>();
            SortedDictionary<string, IITTrack> tracks = new SortedDictionary<string, IITTrack>();
            SortedDictionary<string, object> iTunesTracks = new SortedDictionary<string, object>();

            string iTunesIDs;
            int highPlaylisID;
            int lowPlaylisID;
            int highTrackID;
            int lowTrackID;

            bool someTracksWereSkipped = false;


            foreach (KeyValuePair<int, string[]> item in files)
            {
                // determine if the the file needs to be synched
                // ...
                // if yes, indicate to MusicBee that you want to synch the file - the returned filename is either the same as the supplied filename or if re-encoding/ forced embeding artwork, a temporary filename is returned
                // if filename is returned as null, that means MusicBee wasnt able to encode the file and it should be skipped from synchronisation
                string filename = null;
                bool success = false;
                string errorMessage = null;

                if (AbortSynchronization)
                {
                    SynchronizationInProgress = false;
                    AbortSynchronization = false;
                    lastEx = null;
                    return true;
                }


                try
                {
                    bool webFile = false;
                    string iTunesTrackFilename = null;


                    // sync the file to the category (Key) requested
                    if (item.Key == (int)Plugin.SynchronisationCategory.Playlist) //Playlist. Lets remember synced playlists for later processing
                    {
                        iTunesTrackFilename = Regex.Replace(item.Value[0], "^.*\\\\(.*)(\\..*)", "$1");
                        IITPlaylist libraryPlaylist = null;
                        foreach (IITPlaylist playlist in iTunes.LibrarySource.Playlists)
                        {
                            if (
                                playlist.Name == iTunesTrackFilename
                                && playlist.Kind == ITPlaylistKind.ITPlaylistKindUser
                                && ((IITUserPlaylist)playlist).SpecialKind == ITUserPlaylistSpecialKind.ITUserPlaylistSpecialKindNone
                                && !((IITUserPlaylist)playlist).Smart
                                )
                            {
                                libraryPlaylist = playlist;
                                break;
                            }
                        }

                        if (libraryPlaylist == null) //Playlist doesn't exist
                            libraryPlaylist = iTunes.CreatePlaylist(iTunesTrackFilename);


                        object playlistObj = libraryPlaylist;
                        iTunes.GetITObjectPersistentIDs(ref playlistObj, out highPlaylisID, out lowPlaylisID);

                        playlists.Add(item.Value[0], libraryPlaylist);
                        iTunesPlaylists.Add(highPlaylisID.ToString() + "@" + lowPlaylisID.ToString(), null);

                        success = true;
                        errorMessage = null;
                    }
                    else //Media file
                    {
                        filename = Plugin.MbApiInterface.Sync_FileStart(item.Value[0]);

                        bool sync = true;

                        if ((Plugin.SynchronisationCategory)item.Key == Plugin.SynchronisationCategory.Audiobook && AudiobooksPlaylist == null)
                            sync = false;
                        else if ((Plugin.SynchronisationCategory)item.Key == Plugin.SynchronisationCategory.Podcast && PodcastsPlaylist == null)
                            sync = false;
                        else if ((Plugin.SynchronisationCategory)item.Key == Plugin.SynchronisationCategory.Video && VideoPlaylist == null)
                            sync = false;


                        if (sync)
                        {
                            IITTrack libraryTrack = null;

                            if (!Regex.IsMatch(item.Value[0], "^http\\:.*") && !Regex.IsMatch(item.Value[0], "^https\\:.*")) //Local or local network file. Lets copy tracks to cache
                            {
                                if (!System.IO.File.Exists(filename))
                                {
                                    errorMessage = Plugin.msgTrackSourceFileNotFound;
                                    int x = 1/Plugin.Zero; //Lets generate exception
                                }

                                if (item.Value[0] == filename) //Track wasn't converted
                                {
                                    iTunesTrackFilename = filename;
                                }
                                else //Track was converted
                                {
                                    string trackGUID = Plugin.MbApiInterface.Library_GetDevicePersistentId(item.Value[0], Plugin.DeviceIdType.AppleDevice);

                                    if ("" + trackGUID == "")
                                    {
                                        trackGUID = Guid.NewGuid().ToString();

                                        Plugin.MbApiInterface.Library_SetDevicePersistentId(item.Value[0], Plugin.DeviceIdType.AppleDevice, trackGUID);
                                        Plugin.MbApiInterface.Library_CommitTagsToFile(item.Value[0]);
                                    }

                                    iTunesTrackFilename = Plugin.ReencodedFilesStorage + "\\" + trackGUID + item.Value[1];

                                    System.IO.File.Copy(filename, iTunesTrackFilename, true);
                                }
                            }
                            else
                            {
                                iTunesTrackFilename = item.Value[0];

                                webFile = true;
                            }


                            IITUserPlaylist currentPlaylist;
                            if ((Plugin.SynchronisationCategory)item.Key == Plugin.SynchronisationCategory.Audiobook)
                                currentPlaylist = AudiobooksPlaylist;
                            else if ((Plugin.SynchronisationCategory)item.Key == Plugin.SynchronisationCategory.Podcast)
                                currentPlaylist = PodcastsPlaylist;
                            else if ((Plugin.SynchronisationCategory)item.Key == Plugin.SynchronisationCategory.Video)
                                currentPlaylist = VideoPlaylist;
                            else
                                currentPlaylist = MusicPlaylist;


                            iTunesIDs = Plugin.MbApiInterface.Library_GetDevicePersistentId(item.Value[0], Plugin.DeviceIdType.AppleDevice2);
                            highPlaylisID = 0;
                            lowPlaylisID = 0;
                            highTrackID = 0;
                            lowTrackID = 0;

                            bool IDsAreCorrect = false;


                            if (iTunesIDs != "") //Track was synced before
                            {
                                try
                                {
                                    highPlaylisID = Convert.ToInt32(Regex.Replace(iTunesIDs, "^(.*)@(.*)@(.*)@(.*)", "$1"));
                                    lowPlaylisID = Convert.ToInt32(Regex.Replace(iTunesIDs, "^(.*)@(.*)@(.*)@(.*)", "$2"));
                                    highTrackID = Convert.ToInt32(Regex.Replace(iTunesIDs, "^(.*)@(.*)@(.*)@(.*)", "$3"));
                                    lowTrackID = Convert.ToInt32(Regex.Replace(iTunesIDs, "^(.*)@(.*)@(.*)@(.*)", "$4"));

                                    IITPlaylist playlist = (IITPlaylist)iTunes.LibrarySource.Playlists.get_ItemByPersistentID(highPlaylisID, lowPlaylisID);
                                    libraryTrack = (IITTrack)playlist.Tracks.get_ItemByPersistentID(highTrackID, lowTrackID);

                                    ITTrackKind temp = libraryTrack.Kind; //To generate exception if libraryTrack == null


                                    //Sync of general tags...
                                    if (!webFile) //Local or local network file
                                        ((IITFileOrCDTrack)libraryTrack).UpdateInfoFromFile();

                                    IDsAreCorrect = true;
                                }
                                catch { }
                            }


                            if (!IDsAreCorrect) //Track was never synced before or was deleted from iTunes library. Lets add track to iTunes library. 
                            {
                                if (!webFile) //Local or local network file
                                {
                                    libraryTrack = null;

                                    IITOperationStatus status = currentPlaylist.AddFile(iTunesTrackFilename);
                                    while (status.InProgress) ;

                                    DateTime startTime = DateTime.UtcNow;
                                    while (libraryTrack == null && (DateTime.UtcNow - startTime).Seconds < AddOperationTimeout)
                                    {
                                        IITTrackCollection foundTracks = currentPlaylist.Search(Plugin.MbApiInterface.Library_GetFileTag(item.Value[0], Plugin.MetaDataType.TrackTitle), ITPlaylistSearchField.ITPlaylistSearchFieldSongNames);

                                        if (foundTracks != null)
                                        {
                                            foreach (IITTrack track in foundTracks)
                                            {
                                                if (track.Kind == ITTrackKind.ITTrackKindFile)
                                                {
                                                    if (((IITFileOrCDTrack)track).Location == iTunesTrackFilename)
                                                    {
                                                        libraryTrack = track;
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    errorMessage = Plugin.msgFailedToAddTrackToiTunesLibrary;
                                    int temp = libraryTrack.TrackNumber; //Lets generate exception if libraryTrack == null


                                    errorMessage = null;
                                }
                                else //Web file
                                {
                                    libraryTrack = currentPlaylist.AddURL(iTunesTrackFilename);
                                }
                            }


                            //Sync of ratings & play counts...
                            int intValue;
                            bool retry;

                            DateTime dateTimeValue;
                            try
                            {
                                dateTimeValue = DateTime.Parse(Plugin.MbApiInterface.Library_GetFileProperty(item.Value[0], Plugin.FilePropertyType.LastPlayed)).ToUniversalTime();
                            }
                            catch
                            {
                                dateTimeValue = DateTime.MinValue;
                            }

                            if (dateTimeValue > libraryTrack.PlayedDate.ToUniversalTime())
                            {
                                try
                                {
                                    intValue = Convert.ToInt32(Plugin.MbApiInterface.Library_GetFileProperty(item.Value[0], Plugin.FilePropertyType.PlayCount));
                                }
                                catch
                                {
                                    intValue = 0;
                                }
                                retry = true;
                                while (retry)
                                {
                                    try
                                    {
                                        libraryTrack.PlayedCount = intValue;
                                        retry = false;
                                    }
                                    catch { }
                                }

                                retry = true;
                                while (retry)
                                {
                                    try
                                    {
                                        libraryTrack.PlayedDate = dateTimeValue;
                                        retry = false;
                                    }
                                    catch { }
                                }

                                try
                                {
                                    intValue = Convert.ToInt32(Plugin.MbApiInterface.Library_GetFileTag(item.Value[0], Plugin.MetaDataType.Rating));
                                }
                                catch
                                {
                                    intValue = 0;
                                }
                                retry = true;
                                while (retry)
                                {
                                    try
                                    {
                                        libraryTrack.Rating = intValue;
                                        retry = false;
                                    }
                                    catch { }
                                }


                                if (libraryTrack.Kind == ITTrackKind.ITTrackKindFile)
                                {
                                    try
                                    {
                                        intValue = Convert.ToInt32(Plugin.MbApiInterface.Library_GetFileProperty(item.Value[0], Plugin.FilePropertyType.SkipCount));
                                    }
                                    catch
                                    {
                                        intValue = 0;
                                    }
                                    retry = true;
                                    while (retry)
                                    {
                                        try
                                        {
                                            ((IITFileOrCDTrack)libraryTrack).SkippedCount = intValue;
                                            retry = false;
                                        }
                                        catch { }
                                    }

                                    try
                                    {
                                        intValue = Convert.ToInt32(Plugin.MbApiInterface.Library_GetFileTag(item.Value[0], Plugin.MetaDataType.RatingAlbum));
                                    }
                                    catch
                                    {
                                        intValue = 0;
                                    }
                                    retry = true;
                                    while (retry)
                                    {
                                        try
                                        {
                                            ((IITFileOrCDTrack)libraryTrack).AlbumRating = intValue;
                                            retry = false;
                                        }
                                        catch { }
                                    }
                                }
                            }


                            if (!webFile) //Local or local network file. Lets set last modified time of iTunes track to the same as MB track.
                            {
                                System.IO.FileInfo sourceFileInfo = new System.IO.FileInfo(item.Value[0]);
                                System.IO.FileInfo destinationFileInfo = new System.IO.FileInfo(iTunesTrackFilename);
                                destinationFileInfo.LastWriteTimeUtc = sourceFileInfo.LastWriteTimeUtc;
                            }


                            object playlistObj;
                            object trackObj;

                            if (!IDsAreCorrect) //Track was never synced before or was deleted from iTunes library. Lets write new IDs to MB library.
                            {

                                playlistObj = libraryTrack.Playlist;
                                iTunes.GetITObjectPersistentIDs(ref playlistObj, out highPlaylisID, out lowPlaylisID);
                                trackObj = libraryTrack;
                                iTunes.GetITObjectPersistentIDs(ref trackObj, out highTrackID, out lowTrackID);

                                iTunesIDs = highPlaylisID.ToString() + "@" + lowPlaylisID.ToString() + "@" + highTrackID.ToString() + "@" + lowTrackID.ToString();
                                Plugin.MbApiInterface.Library_SetDevicePersistentId(item.Value[0], Plugin.DeviceIdType.AppleDevice2, iTunesIDs);
                                Plugin.MbApiInterface.Library_CommitTagsToFile(item.Value[0]);
                            }


                            playlistObj = libraryTrack.Playlist;
                            iTunes.GetITObjectPersistentIDs(ref playlistObj, out highPlaylisID, out lowPlaylisID);
                            trackObj = libraryTrack;
                            iTunes.GetITObjectPersistentIDs(ref trackObj, out highTrackID, out lowTrackID);

                            tracks.Add(item.Value[0], libraryTrack);
                            iTunesTracks.Add(highPlaylisID.ToString() + "@" + lowPlaylisID.ToString() + "@" + highTrackID.ToString() + "@" + lowTrackID.ToString(), null);

                            success = true;
                            errorMessage = null;
                        }
                        else
                        {
                            errorMessage = Plugin.msgNoAppropriateCategoryFoundIniTunesLibrary;
                        }
                    }
                }
                catch (System.ArgumentException)
                {
                    success = true;
                    errorMessage = null;
                    someTracksWereSkipped = true;
                }
                catch (Exception ex)
                {
                    if (errorMessage == null)
                        errorMessage = ex.Message;
                }
                finally
                {
                    // when the file synch is done
                    if (filename != null)
                        Plugin.MbApiInterface.Sync_FileEnd(item.Value[0], success, errorMessage);
                }
            }


            try
            {
                //Lets remove tracks and playlists which were not synced if removing missing items is required
                if ((flags & Plugin.SynchronisationSettings.SyncRemoveMissingFiles) != 0)
                {
                    //Lets remove not synced tracks
                    for (int i = iTunes.LibrarySource.Playlists.Count; i > 0; i--)
                    {
                        IITPlaylist playlist = iTunes.LibrarySource.Playlists[i];

                        object playlistObj = playlist;
                        iTunes.GetITObjectPersistentIDs(ref playlistObj, out highPlaylisID, out lowPlaylisID);

                        if (
                            playlist.Kind == ITPlaylistKind.ITPlaylistKindUser
                            &&
                                (
                                    ((IITUserPlaylist)playlist).SpecialKind == ITUserPlaylistSpecialKind.ITUserPlaylistSpecialKindMusic
                            //|| ((IITUserPlaylist)playlist).SpecialKind == ITUserPlaylistSpecialKind.ITUserPlaylistSpecialKindBooks
                            //|| ((IITUserPlaylist)playlist).SpecialKind == ITUserPlaylistSpecialKind.ITUserPlaylistSpecialKindMovies
                            //|| ((IITUserPlaylist)playlist).SpecialKind == ITUserPlaylistSpecialKind.ITUserPlaylistSpecialKindPodcasts
                            //|| ((IITUserPlaylist)playlist).SpecialKind == ITUserPlaylistSpecialKind.ITUserPlaylistSpecialKindTVShows
                            //|| ((IITUserPlaylist)playlist).SpecialKind == ITUserPlaylistSpecialKind.ITUserPlaylistSpecialKindVideos
                                )
                            )
                        {
                            int tracksCount = playlist.Tracks.Count;

                            for (int j = playlist.Tracks.Count; j > 0; j--)
                            {
                                int curTrackNumber = tracksCount - j + 1;

                                Plugin.MbApiInterface.MB_SetBackgroundTaskMessage(Plugin.msgRemovingTrack + curTrackNumber + "/" + tracksCount);

                                IITTrack track = playlist.Tracks[j];

                                object trackObj = track;
                                iTunes.GetITObjectPersistentIDs(ref trackObj, out highTrackID, out lowTrackID);

                                iTunesIDs = highPlaylisID.ToString() + "@" + lowPlaylisID.ToString() + "@" + highTrackID.ToString() + "@" + lowTrackID.ToString();

                                object nothing;
                                if (!iTunesTracks.TryGetValue(iTunesIDs, out nothing)) //Track wasn't synced and needs to be deleted.
                                {
                                    bool? deleteTrack = true;

                                    if ((flags & Plugin.SynchronisationSettings.SyncConfirmRemoveFiles) != 0)
                                    {
                                        string trackRepresentation = GetTrackRepresentation(track);
                                        DeleteTrack deleteTrackWindow = new DeleteTrack();
                                        deleteTrackWindow.questionLabel.Text = Plugin.msgDeleteTrack + trackRepresentation + Plugin.msgFromDevice;
                                        Plugin.MbForm.AddOwnedForm(deleteTrackWindow);

                                        InvokedFunction = deleteTrackWindow.Show;
                                        Plugin.MbForm.Invoke(InvokedFunction);

                                        while (deleteTrackWindow.deleteTrackOrPlaylist == null) ;

                                        deleteTrack = deleteTrackWindow.deleteTrackOrPlaylist;

                                        InvokedFunction = deleteTrackWindow.Close;
                                        Plugin.MbForm.Invoke(InvokedFunction);
                                    }

                                    if (deleteTrack == true)
                                    {
                                        string[] mbFiles = null;
                                        Plugin.MbApiInterface.Library_FindDevicePersistentId(Plugin.DeviceIdType.AppleDevice2, new string[] { iTunesIDs }, ref mbFiles);

                                        foreach (string mbFile in mbFiles)
                                        {
                                            if (mbFile != null)
                                            {
                                                try
                                                {
                                                    System.IO.FileInfo fileInfo = new System.IO.FileInfo(((IITFileOrCDTrack)track).Location);

                                                    if (fileInfo.Directory.FullName == Plugin.ReencodedFilesStorage)
                                                        System.IO.File.Delete(((IITFileOrCDTrack)track).Location);
                                                }
                                                catch (Exception) { }

                                                Plugin.MbApiInterface.Library_SetDevicePersistentId(mbFile, Plugin.DeviceIdType.AppleDevice, "");
                                                Plugin.MbApiInterface.Library_SetDevicePersistentId(mbFile, Plugin.DeviceIdType.AppleDevice2, "");
                                                Plugin.MbApiInterface.Library_CommitTagsToFile(mbFile);
                                            }
                                        }

                                        track.Delete();
                                    }
                                }
                            }
                        }
                        else if ( //Not smart, user playlist
                            playlist.Kind == ITPlaylistKind.ITPlaylistKindUser
                            &&
                            ((IITUserPlaylist)playlist).SpecialKind == ITUserPlaylistSpecialKind.ITUserPlaylistSpecialKindNone
                            &&
                            !((IITUserPlaylist)playlist).Smart
                                )
                        {
                            iTunesIDs = highPlaylisID.ToString() + "@" + lowPlaylisID.ToString();

                            object nothing;
                            if (!iTunesPlaylists.TryGetValue(iTunesIDs, out nothing)) //Track wasn't synced and needs to be deleted.
                            {
                                bool? deletePlaylist = true;

                                if ((flags & Plugin.SynchronisationSettings.SyncConfirmRemoveFiles) != 0)
                                {
                                    string trackRepresentation = playlist.Name;
                                    DeleteTrack deleteTrackWindow = new DeleteTrack();
                                    deleteTrackWindow.questionLabel.Text = Plugin.msgDeletePlaylist + trackRepresentation + Plugin.msgFromDevice;
                                    Plugin.MbForm.AddOwnedForm(deleteTrackWindow);

                                    InvokedFunction = deleteTrackWindow.Show;
                                    Plugin.MbForm.Invoke(InvokedFunction);

                                    while (deleteTrackWindow.deleteTrackOrPlaylist == null) ;

                                    deletePlaylist = deleteTrackWindow.deleteTrackOrPlaylist;

                                    InvokedFunction = deleteTrackWindow.Close;
                                    Plugin.MbForm.Invoke(InvokedFunction);
                                }

                                if (deletePlaylist == true)
                                {
                                    playlist.Delete();
                                }
                            }
                        }
                    }
                }



                //Lets fill synced playlists
                //lastEx = null;
                foreach (KeyValuePair<string, IITPlaylist> playlistPair in playlists)
                {
                    Plugin.MbApiInterface.MB_SetBackgroundTaskMessage(Plugin.msgFillingSyncedPlaylists);

                    foreach (IITTrack track in playlistPair.Value.Tracks)
                        track.Delete();

                    if (Plugin.MbApiInterface.Playlist_QueryFiles(playlistPair.Key))
                    {
                        int m = 0;
                        string[] playlistFiles = Plugin.MbApiInterface.Playlist_QueryGetAllFiles().Split(Plugin.FilesSeparators, StringSplitOptions.RemoveEmptyEntries);
                        foreach (string playlistFile in playlistFiles)
                        {
                            m++;

                            Plugin.MbApiInterface.MB_SetBackgroundTaskMessage(Plugin.msgAddingTrackToPlaylist + m + "/" + playlistFiles.Length);

                            try
                            {
                                ((IITUserPlaylist)playlistPair.Value).AddTrack(tracks[playlistFile]);
                            }
                            catch(Exception ex)
                            {
                                //lastEx = ex;
                            }
                        }
                    }

                    Plugin.MbApiInterface.MB_SetBackgroundTaskMessage("");
                }


                if (!IgnoreIPod)
                {
                    Plugin.MbApiInterface.MB_SetBackgroundTaskMessage(Plugin.msgSyncingITunesWithiPodIPhone);
                    iTunes.UpdateIPod();
                    Plugin.MbApiInterface.MB_SetBackgroundTaskMessage("");
                }


                if ((flags & (Plugin.SynchronisationSettings.SyncPlayCount2Way | Plugin.SynchronisationSettings.SyncRating2Way)) != 0) //Sync play counts and/or ratings FROM iPod
                {
                    Plugin.MbApiInterface.MB_SetBackgroundTaskMessage(Plugin.msgSyncingPlayCountsAndRatingsFROMiPod);

                    foreach (KeyValuePair<string, IITTrack> trackPair in tracks)
                    {
                        string mbTrackUrl = trackPair.Key;
                        IITTrack libraryTrack = trackPair.Value;

                        DateTime dateTimeValue;
                        try
                        {
                            dateTimeValue = DateTime.Parse(Plugin.MbApiInterface.Library_GetFileProperty(mbTrackUrl, Plugin.FilePropertyType.LastPlayed)).ToUniversalTime();
                        }
                        catch
                        {
                            dateTimeValue = DateTime.MinValue;
                        }

                        if ((flags & Plugin.SynchronisationSettings.SyncPlayCount2Way) != 0 && dateTimeValue < libraryTrack.PlayedDate.ToUniversalTime()) //Sync play count and last played date FROM iPod
                        {
                            Plugin.MbApiInterface.Library_SetFileTag(mbTrackUrl, (Plugin.MetaDataType)Plugin.FilePropertyType.PlayCount, libraryTrack.PlayedCount.ToString());

                            Plugin.MbApiInterface.Library_SetFileTag(mbTrackUrl, (Plugin.MetaDataType)Plugin.FilePropertyType.LastPlayed, libraryTrack.PlayedDate.ToUniversalTime().ToString());

                            if (libraryTrack.Kind == ITTrackKind.ITTrackKindFile)
                                Plugin.MbApiInterface.Library_SetFileTag(mbTrackUrl, (Plugin.MetaDataType)Plugin.FilePropertyType.SkipCount, (((IITFileOrCDTrack)libraryTrack).SkippedCount).ToString());
                        }

                        if ((flags & Plugin.SynchronisationSettings.SyncRating2Way) != 0 && dateTimeValue < libraryTrack.PlayedDate.ToUniversalTime()) //Sync rating FROM iPod
                        {
                            Plugin.MbApiInterface.Library_SetFileTag(mbTrackUrl, Plugin.MetaDataType.Rating, libraryTrack.Rating.ToString());

                            if (libraryTrack.Kind == ITTrackKind.ITTrackKindFile)
                                Plugin.MbApiInterface.Library_SetFileTag(mbTrackUrl, Plugin.MetaDataType.RatingAlbum, (((IITFileOrCDTrack)libraryTrack).AlbumRating).ToString());
                        }

                        Plugin.MbApiInterface.Library_CommitTagsToFile(mbTrackUrl);
                    }

                    Plugin.MbApiInterface.MB_SetBackgroundTaskMessage("");
                }
            }
            catch (Exception ex)
            {
                Plugin.MbApiInterface.MB_RefreshPanels();

                SynchronizationInProgress = false;

                lastEx = ex;
                return false;
            }



            Plugin.MbApiInterface.MB_RefreshPanels();

            if (someTracksWereSkipped)
                MessageBox.Show(Plugin.msgSomeDuplicatedTracksWereSkipped);

            SynchronizationInProgress = false;


            if (lastEx != null)
                return false;


            return true;
        }

        public static void Refresh()
        {
        }

        public static bool FolderExists(string path)
        {
            return (string.IsNullOrEmpty(path) || path == @"\");
        }

        public static string[] GetFolders(string path)
        {
            string[] folders;
            folders = new string[1];
            folders[0] = "";

            return folders;
        }

        private static KeyValuePair<byte, string>[][] GetPlaylistTracks(IITPlaylist playlist)
        {
            ITTracksCache.Clear();

            List<KeyValuePair<byte, string>[]> files = new List<KeyValuePair<byte,string>[]>();

            foreach (IITTrack currTrack in playlist.Tracks)
            {
                KeyValuePair<byte, string>[] file = new KeyValuePair<byte, string>[TagCount];

                if (currTrack.Kind == ITTrackKind.ITTrackKindFile)
                {
                    IITFileOrCDTrack fileTrack = (IITFileOrCDTrack)currTrack;
                    if (fileTrack.Location != null)
                    {
                        object fileTrackObject = fileTrack;
                        int highID, lowID;
                        iTunes.GetITObjectPersistentIDs(ref fileTrackObject, out highID, out lowID);

                        ITTracksCache.Add(fileTrack.Location, new int[] { highID, lowID });


                        file[0] = new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.Url, fileTrack.Location);
                        file[1] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Artist, fileTrack.Artist);
                        file[2] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.TrackTitle, fileTrack.Name);
                        file[3] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Album, fileTrack.Album);
                        file[4] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Genre, fileTrack.Genre);
                        file[5] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Comment, fileTrack.Comment);
                        file[6] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.AlbumArtistRaw, fileTrack.AlbumArtist);
                        file[7] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.RatingAlbum, fileTrack.AlbumRating.ToString());
                        file[8] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Rating, fileTrack.Rating.ToString());
                        file[9] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Year, fileTrack.Year.ToString());
                        file[10] = new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.Bitrate, fileTrack.BitRate.ToString());
                        file[11] = new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.Size, fileTrack.Size.ToString());
                        file[12] = new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.Duration, (fileTrack.Duration * 1000).ToString());
                        file[13] = new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.PlayCount, fileTrack.PlayedCount.ToString());
                        file[14] = new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.SkipCount, fileTrack.SkippedCount.ToString());

                        file[15] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Composer, fileTrack.Composer);
                        file[16] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.DiscCount, fileTrack.DiscCount.ToString());
                        file[17] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.DiscNo, fileTrack.DiscNumber.ToString());
                        file[18] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Grouping, fileTrack.Grouping);
                        file[19] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.TrackCount, fileTrack.TrackCount.ToString());
                        file[20] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.TrackNo, fileTrack.TrackNumber.ToString());

                        file[21] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Artwork, fileTrack.Artwork.Count == 0 ? "" : "Y");

                        file[22] = new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.LastPlayed, fileTrack.PlayedDate.ToUniversalTime().ToString());

                        files.Add(file);
                    }
                }
            }

            return files.ToArray();
        }

        public static KeyValuePair<byte, string>[][] GetFiles(string path)
        {
            lastEx = null;
            KeyValuePair<byte, string>[][] files = null;

            if (IsInitialised != true)
            {
                files = new KeyValuePair<byte, string>[0][];
            }
            else
            {
                files = GetPlaylistTracks(iTunes.LibraryPlaylist);
            }

            return files;
        }

        public static bool FileExists(string url)
        {
            int[] ids = null;
            IITTrack foundTrack = null;
            if (ITTracksCache.TryGetValue(url, out ids))
                foundTrack = iTunes.LibraryPlaylist.Tracks.get_ItemByPersistentID(url[0], url[1]);

            return foundTrack != null;
        }

        public static KeyValuePair<byte, string>[] GetFile(string url)
        {
            KeyValuePair<byte, string>[] file = new KeyValuePair<byte, string>[TagCount];


            int[] ids = null;
            IITTrack foundTrack = null;
            if (ITTracksCache.TryGetValue(url, out ids))
                foundTrack = iTunes.LibraryPlaylist.Tracks.get_ItemByPersistentID(url[0], url[1]);


            if (foundTrack != null && foundTrack.Kind == ITTrackKind.ITTrackKindFile)
            {
                IITFileOrCDTrack fileTrack = (IITFileOrCDTrack)foundTrack;
                if (fileTrack.Location == url)
                {
                    file[0] = new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.Url, fileTrack.Location);
                    file[1] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Artist, fileTrack.Artist);
                    file[2] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.TrackTitle, fileTrack.Name);
                    file[3] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Album, fileTrack.Album);
                    file[4] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Genre, fileTrack.Genre);
                    file[5] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Comment, fileTrack.Comment);
                    file[6] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.AlbumArtistRaw, fileTrack.AlbumArtist);
                    file[7] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.RatingAlbum, fileTrack.AlbumRating.ToString());
                    file[8] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Rating, fileTrack.Rating.ToString());
                    file[9] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Year, fileTrack.Year.ToString());
                    file[10] = new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.Bitrate, fileTrack.BitRate.ToString());
                    file[11] = new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.Size, fileTrack.Size.ToString());
                    file[12] = new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.Duration, (fileTrack.Duration * 1000).ToString());
                    file[13] = new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.PlayCount, fileTrack.PlayedCount.ToString());
                    file[14] = new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.SkipCount, fileTrack.SkippedCount.ToString());

                    file[15] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Composer, fileTrack.Composer);
                    file[16] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.DiscCount, fileTrack.DiscCount.ToString());
                    file[17] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.DiscNo, fileTrack.DiscNumber.ToString());
                    file[18] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Grouping, fileTrack.Grouping);
                    file[19] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.TrackCount, fileTrack.TrackCount.ToString());
                    file[20] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.TrackNo, fileTrack.TrackNumber.ToString());

                    file[21] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Artwork, fileTrack.Artwork.Count == 0 ? "" : "Y");

                    file[22] = new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.LastPlayed, fileTrack.PlayedDate.ToUniversalTime().ToString());
                }
            }

            return file;
        }

        public static byte[] GetFileArtwork(string url)
        {
            lastEx = null;

            try
            {
                int[] ids = null;
                IITTrack foundTrack = null;
                if (ITTracksCache.TryGetValue(url, out ids))
                    foundTrack = iTunes.LibraryPlaylist.Tracks.get_ItemByPersistentID(url[0], url[1]);


                if (foundTrack != null && foundTrack.Kind == ITTrackKind.ITTrackKindFile)
                {
                    IITFileOrCDTrack fileTrack = (IITFileOrCDTrack)foundTrack;
                    if (fileTrack.Location == url)
                    {
                        if (fileTrack.Artwork.Count == 0)
                        {
                            return null;
                        }
                        else
                        {
                            fileTrack.Artwork[1].SaveArtworkToFile(Plugin.MbApiInterface.Setting_GetPersistentStoragePath() + "iPod & iPhone Driver.jpg");

                            Bitmap artwork = new Bitmap(Plugin.MbApiInterface.Setting_GetPersistentStoragePath() + "iPod & iPhone Driver.jpg");
                            TypeConverter tc = TypeDescriptor.GetConverter(typeof(Bitmap));
                            byte[] artworkBytes = (byte[])tc.ConvertTo(artwork, typeof(byte[]));
                            artwork.Dispose();
                            artwork = null;

                            return artworkBytes;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                lastEx = ex;
            }

            return null;
        }

        public static KeyValuePair<string, string>[] GetPlaylists()
        {
            List<KeyValuePair<string, string>> playlistsList = new List<KeyValuePair<string,string>>();

            foreach (IITPlaylist playlist in iTunes.LibrarySource.Playlists)
            {
                if ( //Not smart, user playlist
                    playlist.Kind == ITPlaylistKind.ITPlaylistKindUser
                    &&
                    ((IITUserPlaylist)playlist).SpecialKind == ITUserPlaylistSpecialKind.ITUserPlaylistSpecialKindNone
                    &&
                    !((IITUserPlaylist)playlist).Smart
                   )
                {
                    playlistsList.Add(new KeyValuePair<string, string>(playlist.Name, playlist.Name));
                }
            }

            return playlistsList.ToArray();
        }

        public static KeyValuePair<byte, string>[][] GetPlaylistFiles(string id)
        {
            List<KeyValuePair<byte, string>[]> playlistFiles = new List<KeyValuePair<byte, string>[]>();

            IITPlaylist libraryPlaylist = null;
            foreach (IITPlaylist playlist in iTunes.LibrarySource.Playlists)
            {
                if ( //Not smart, user playlist with given name/id
                    playlist.Kind == ITPlaylistKind.ITPlaylistKindUser
                    &&
                    ((IITUserPlaylist)playlist).SpecialKind == ITUserPlaylistSpecialKind.ITUserPlaylistSpecialKindNone
                    &&
                    !((IITUserPlaylist)playlist).Smart
                    &&
                    playlist.Name == id
                   )
                {
                    libraryPlaylist = playlist;

                    ITTracksCache.Clear();

                    foreach (IITTrack currTrack in libraryPlaylist.Tracks)
                    {
                        if (currTrack.Kind == ITTrackKind.ITTrackKindFile)
                        {
                            KeyValuePair<byte, string>[] file = new KeyValuePair<byte, string>[TagCount];

                            IITFileOrCDTrack fileTrack = (IITFileOrCDTrack)currTrack;

                            object fileTrackObject = fileTrack;
                            int highID, lowID;
                            iTunes.GetITObjectPersistentIDs(ref fileTrackObject, out highID, out lowID);

                            ITTracksCache.Add(fileTrack.Location, new int[] { highID, lowID });


                            file[0] = new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.Url, fileTrack.Location);
                            file[1] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Artist, fileTrack.Artist);
                            file[2] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.TrackTitle, fileTrack.Name);
                            file[3] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Album, fileTrack.Album);
                            file[4] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Genre, fileTrack.Genre);
                            file[5] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Comment, fileTrack.Comment);
                            file[6] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.AlbumArtistRaw, fileTrack.AlbumArtist);
                            file[7] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.RatingAlbum, fileTrack.AlbumRating.ToString());
                            file[8] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Rating, fileTrack.Rating.ToString());
                            file[9] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Year, fileTrack.Year.ToString());
                            file[10] = new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.Bitrate, fileTrack.BitRate.ToString());
                            file[11] = new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.Size, fileTrack.Size.ToString());
                            file[12] = new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.Duration, (fileTrack.Duration * 1000).ToString());
                            file[13] = new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.PlayCount, fileTrack.PlayedCount.ToString());
                            file[14] = new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.SkipCount, fileTrack.SkippedCount.ToString());

                            file[15] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Composer, fileTrack.Composer);
                            file[16] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.DiscCount, fileTrack.DiscCount.ToString());
                            file[17] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.DiscNo, fileTrack.DiscNumber.ToString());
                            file[18] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Grouping, fileTrack.Grouping);
                            file[19] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.TrackCount, fileTrack.TrackCount.ToString());
                            file[20] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.TrackNo, fileTrack.TrackNumber.ToString());

                            file[21] = new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Artwork, fileTrack.Artwork.Count == 0 ? "" : "Y");

                            file[22] = new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.LastPlayed, fileTrack.PlayedDate.ToUniversalTime().ToString());

                            playlistFiles.Add(file);
                        }
                    }

                    break;
                }
            }

            if (libraryPlaylist == null)
            {
                KeyValuePair<byte, string>[] file = new KeyValuePair<byte, string>[TagCount];
                playlistFiles.Add(file);
                return playlistFiles.ToArray();
            }
            else
            {
                return playlistFiles.ToArray();
            }
        }

        public static string CreatePlaylist(string name)
        {
            IITPlaylist libraryPlaylist = null;
            foreach (IITPlaylist playlist in iTunes.LibrarySource.Playlists)
            {
                if (
                    playlist.Name == name
                    && playlist.Kind == ITPlaylistKind.ITPlaylistKindUser
                    && ((IITUserPlaylist)playlist).SpecialKind == ITUserPlaylistSpecialKind.ITUserPlaylistSpecialKindNone
                    && !((IITUserPlaylist)playlist).Smart
                    )
                {
                    libraryPlaylist = playlist;
                    break;
                }
            }

            if (libraryPlaylist == null) //Playlist doesn't exist
            {
                libraryPlaylist = iTunes.CreatePlaylist(name);

                lastEx = null;
                return name;
            }
            else
            {
                MessageBox.Show("Playlist \"" + name + "\" already exists!");

                lastEx = null;
                return "";
            }
        }

        public static bool UpdatePlaylist(string id, KeyValuePair<byte, string>[][] files)
        {
            IITPlaylist libraryPlaylist = null;
            foreach (IITPlaylist playlist in iTunes.LibrarySource.Playlists)
            {
                if (
                    playlist.Name == id
                    && playlist.Kind == ITPlaylistKind.ITPlaylistKindUser
                    && ((IITUserPlaylist)playlist).SpecialKind == ITUserPlaylistSpecialKind.ITUserPlaylistSpecialKindNone
                    && !((IITUserPlaylist)playlist).Smart
                    )
                {
                    libraryPlaylist = playlist;
                    break;
                }
            }

            if (libraryPlaylist == null) //Playlist doesn't exist
            {
                MessageBox.Show("Playlist \"" + id + "\" doesn't exist!");

                lastEx = null;
                return false;
            }


            foreach (KeyValuePair<byte, string>[] file in files)
            {
                foreach (IITTrack currTrack in libraryPlaylist.Tracks)
                {
                    if (currTrack.Kind == ITTrackKind.ITTrackKindFile)
                    {
                        IITFileOrCDTrack fileTrack = (IITFileOrCDTrack)currTrack;

                        if //Found track
                            (
                                file[0].Value == fileTrack.Location
                                && file[1].Value == fileTrack.Artist
                                && file[2].Value == fileTrack.Name
                                && file[3].Value == fileTrack.Album
                                && file[4].Value == fileTrack.Genre
                                && file[5].Value == fileTrack.Comment
                                && file[6].Value == fileTrack.AlbumArtist
                                && file[7].Value == fileTrack.AlbumRating.ToString()
                                && file[8].Value == fileTrack.Rating.ToString()
                                && file[9].Value == fileTrack.Year.ToString()
                                && file[10].Value == fileTrack.BitRate.ToString()
                                && file[11].Value == fileTrack.Size.ToString()
                                && file[12].Value == (fileTrack.Duration * 1000).ToString()
                            )
                        {
                            ((IITUserPlaylist)libraryPlaylist).AddTrack(fileTrack);
                        }
                    }
                }
            }

            return true;
        }

        public static bool DeletePlaylist(string id)
        {
            IITPlaylist libraryPlaylist = null;
            foreach (IITPlaylist playlist in iTunes.LibrarySource.Playlists)
            {
                if (
                    playlist.Name == id
                    && playlist.Kind == ITPlaylistKind.ITPlaylistKindUser
                    && ((IITUserPlaylist)playlist).SpecialKind == ITUserPlaylistSpecialKind.ITUserPlaylistSpecialKindNone
                    && !((IITUserPlaylist)playlist).Smart
                    )
                {
                    libraryPlaylist = playlist;
                    break;
                }
            }

            if (libraryPlaylist == null) //Playlist doesn't exist
            {
                MessageBox.Show("Playlist \"" + id + "\" doesn't exist!");

                lastEx = null;
                return false;
            }

            libraryPlaylist.Delete();

            lastEx = null;
            return true;
        }

        public static System.IO.Stream GetStream(string url)
        {
            System.IO.FileStream stream = new System.IO.FileStream(url, System.IO.FileMode.Open);
            return stream;
        }

        public static Exception GetError()
        {
            return lastEx;
        }

        private sealed class FileSorter : Comparer<KeyValuePair<byte, string>[]>
        {
            public override int Compare(KeyValuePair<byte, string>[] tags1, KeyValuePair<byte, string>[] tags2)
            {
                return String.Compare(tags1[0].Value, tags2[0].Value, StringComparison.OrdinalIgnoreCase);
            }
        }
    }
}