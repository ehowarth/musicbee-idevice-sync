using iTunesLib;
using MusicBeeDeviceSyncPlugin;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace MusicBeePlugin
{
	public partial class Plugin
	{
		public static bool DeveloperMode = false;
		private bool uninstalled = false;

		public static MusicBeeApiInterface MbApiInterface;

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

		public void ToggleItunesOpenedAndClosed(object sender, EventArgs e)
		{
			while (ITunes.IsInitialised == null) ;

			if (ITunes.IsInitialised == false)
			{
				Backgrounding.RunInBackground(ITunes.Initialise);
			}
			else
			{
				Backgrounding.RunInBackground(() => Eject());
			}
		}

		public PluginInfo Initialise(IntPtr apiInterfacePtr)
		{
			MbApiInterface = new MusicBeeApiInterface();
			MbApiInterface.Initialise(apiInterfacePtr);

			//Lets try to read defaults for controls from settings file
			pluginSettingsFileName = Path.Combine(MbApiInterface.Setting_GetPersistentStoragePath(), Info.AssemblyName.Name + ".Settings.xml");

			using (var stream = File.Open(pluginSettingsFileName, FileMode.OpenOrCreate, FileAccess.Read, FileShare.None))
			using (var file = new StreamReader(stream, Encoding.UTF8))
			{
				XmlSerializer controlsDefaultsSerializer = null;
				try
				{
					controlsDefaultsSerializer = new XmlSerializer(typeof(SavedSettingsType));
				}
				catch (Exception e)
				{
					Trace.WriteLine(e); 
					MessageBox.Show("" + e.Message);
				}

				try
				{
					SavedSettings = (SavedSettingsType)controlsDefaultsSerializer.Deserialize(file);
				}
				catch
				{
					SavedSettings = new SavedSettingsType();
				};
			}

			ReencodedFilesStorage = MbApiInterface.Setting_GetPersistentStoragePath() + "iPod & iPhone Driver";
			if (!Directory.Exists(ReencodedFilesStorage))
			{
				Directory.CreateDirectory(ReencodedFilesStorage);
			}

			Language = Thread.CurrentThread.CurrentCulture.TwoLetterISOLanguageName;

			if (!File.Exists(Path.Combine(Application.StartupPath, "Plugins", "ru", Info.AssemblyName.Name + ".resources.dll")))
				Language = "en"; //For testing

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

			MbApiInterface.MB_AddMenuItem(
				"mnuTools/" + Text.L("Open iPod && iPhone Sync"),
				Text.L("Turns sync plugin on and off"),
				ToggleItunesOpenedAndClosed);

			DeveloperMode = File.Exists(Application.StartupPath + @"\Plugins\DevMode.txt");

			return Info.Current;
		}

		public bool Configure(IntPtr panelHandle)
		{
			// save any persistent settings in a sub-folder of this path
			//string dataPath = mbApiInterface.Setting_GetPersistentStoragePath();
			// panelHandle will only be set if you set about.ConfigurationPanelHeight to a non-zero value
			// keep in mind the panel width is scaled according to the font the user has selected
			// if about.ConfigurationPanelHeight is set to 0, you can display your own popup window

			new ConfigurationDialog().ShowDialog();
			SaveSettings();

			return true;
		}

		// called by MusicBee when the user clicks Apply or Save in the MusicBee Preferences screen.
		// its up to you to figure out whether anything has changed and needs updating
		public void SaveSettings()
		{
			using (var stream = File.Open(pluginSettingsFileName, FileMode.Create, FileAccess.Write, FileShare.None))
			using (var file = new StreamWriter(stream, Encoding.UTF8))
			{
				new XmlSerializer(typeof(SavedSettingsType)).Serialize(file, SavedSettings);
			}
		}

		// MusicBee is closing the plugin (plugin is being disabled by user or MusicBee is shutting down)
		public void Close(PluginCloseReason reason)
		{
			Eject();
		}

		// uninstall this plugin - clean up any persisted files
		public void Uninstall()
		{
			Eject();

			foreach (var file in MusicBeeFile.AllFiles())
			{
				file.ClearPluginTags();
			}

			if (File.Exists(pluginSettingsFileName))
				File.Delete(pluginSettingsFileName);

			if (Directory.Exists(ReencodedFilesStorage))
				Directory.Delete(ReencodedFilesStorage, true);

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

		public bool GetDeviceProperties(IntPtr handle)
		{
			var properties = new DeviceProperties
			{
				ShowCategoryNodes = true,
				AudiobooksSupported = false,
				PodcastsSupported = false,
				VideoSupported = false,
				SyncExternalArtwork = false,
				SyncAllowFileRemoval = true,
				SyncAllowRating2Way = true,
				SyncAllowPlayCount2Way = true,
				SyncAllowPlaylists2Way = false,
				SupportedFormats =
					SynchronisationSupportedFormats.SyncMp3Supported |
					SynchronisationSupportedFormats.SyncAacSupported |
					SynchronisationSupportedFormats.SyncAlacSupported |
					SynchronisationSupportedFormats.SyncWavSupported,
				DeviceIcon64 = (Bitmap)MusicBeeDeviceSyncPlugin.Properties.Resources.ResourceManager.GetObject("iTunes64"),
			};

			if (ITunes.SyncItunesButNotDevice)
			{
				properties.DeviceName = Text.L("iPod & iPhone Sync");
				properties.FirmwareVersion = "--";
				properties.Model = Text.L("No device");
				properties.Manufacturer = Text.L("No device");
			}
			else
			{
				properties.DeviceName = ITunes.IPodSource.Name;
				properties.FirmwareVersion = ITunes.IPodSource.SoftwareVersion;
				properties.FreeSpace = (ulong)ITunes.IPodSource.FreeSpace;
				properties.TotalSpace = (ulong)ITunes.IPodSource.Capacity;
				properties.Model = Text.L("<undisclosed>");
				properties.Manufacturer = "Apple Inc.";
			}

			Marshal.StructureToPtr(properties, handle, false);

			// Sync play history and ratings from iTunes right now.
			// Cannot wait for the Synchronize command from MB before collecting ratings and history
			// because this data may affect the file list passed to Synchronize.
			//		if (flags.HasFlag(Plugin.SynchronisationSettings.SyncPlayCount2Way) || flags.HasFlag(Plugin.SynchronisationSettings.SyncRating2Way))
			Backgrounding.RunInBackground(ITunes.SynchronizeRatingsAndPlaysFromItunesToMB);

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
				Trace.WriteLine(ex); 
				MbApiInterface.MB_SendNotification(CallbackType.StorageFailed);
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

			Trace.Flush();

			return true;
		}

		public void Refresh()
		{
			// TODO: what to refresh
		}

		public bool IsReady()
		{
			return ITunes.IsInitialised == true;
		}

		public Bitmap GetIcon()
		{
			return (Bitmap)MusicBeeDeviceSyncPlugin.Properties.Resources.ResourceManager.GetObject("iTunes");
			//(Bitmap)Info.Resources.GetObject("iTunes");
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
			return ITunes.iTunes.GetPlaylists().Select(playlist => new KeyValuePair<string, string>(playlist.Name, playlist.Name)).ToArray();
		}

		public KeyValuePair<byte, string>[][] GetPlaylistFiles(string id)
		{
			return ITunes.GetPlaylistFiles(id);
		}

		public Stream GetStream(string url)
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

		public static iTunesApp iTunes;
		private const int TagCount = 23;
		private const int AddOperationTimeout = 5; //Timeout of adding track to iTunes (1 means 0.5 sec, 2 - 1.5 sec, 3 - 2.5 sec, 30 - 29.5 sec, etc.)
		public static Exception lastEx = null;

		public static bool SyncItunesButNotDevice = false; //For testing
		public static IITIPodSource IPodSource = null;
		private static IITUserPlaylist MusicPlaylist = null;
		private static IITUserPlaylist AudiobooksPlaylist = null;
		private static IITUserPlaylist PodcastsPlaylist = null;
		private static IITUserPlaylist VideoPlaylist = null;


		public static bool SynchronizationInProgress = false;
		public static bool AbortSynchronization = false;

		private delegate void MBInvokedFunction();
		private static MBInvokedFunction InvokedFunction;

		private static SortedDictionary<string, int[]> ITTracksCache = new SortedDictionary<string, int[]>();

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
						else
							Marshal.ReleaseComObject(iTunes.Sources[i]);
					}
				}

				if (waitingWindow.proceed == true)
					SyncItunesButNotDevice = true;

				InvokedFunction = waitingWindow.Close;
				Plugin.MbForm.Invoke(InvokedFunction);

				if (IPodSource == null && !SyncItunesButNotDevice)
				{
					iTunes.Quit();
					IsInitialised = false;
					//SendNotificationHandler.Invoke(Plugin.CallbackType.StorageFailed);
					return;
				}

				if (IPodSource != null)
				{
					Plugin.SavedSettings.deviceName = IPodSource.Name;

					var originalSize = IPodSource.FreeSpace;
					IPodSource.UpdateIPod();
					while (true)
					{
						Thread.Sleep(1000);
						if (Math.Abs(originalSize - IPodSource.FreeSpace) > 1)
						{
							break;
						}
					}
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
				Trace.WriteLine(ex); 
				IsInitialised = false;
				lastEx = ex;
				MessageBox.Show(ex.Message);

				iTunes.Quit();
			}
			finally
			{
				Plugin.MbApiInterface.MB_SetBackgroundTaskMessage("");
				if (IsInitialised == true)
					Plugin.MbApiInterface.MB_SendNotification(Plugin.CallbackType.StorageReady);
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
				if (!SyncItunesButNotDevice)
				{
					IPodSource.EjectIPod();
					Marshal.ReleaseComObject(IPodSource);
					IPodSource = null;
				}

				iTunes.Quit();
				if (MusicPlaylist != null)
				{
					Marshal.ReleaseComObject(MusicPlaylist);
					MusicPlaylist = null;
				}
				if (AudiobooksPlaylist != null)
				{
					Marshal.ReleaseComObject(AudiobooksPlaylist);
					AudiobooksPlaylist = null;
				}
				if (PodcastsPlaylist != null)
				{
					Marshal.ReleaseComObject(PodcastsPlaylist);
					PodcastsPlaylist = null;
				}
				if (VideoPlaylist != null)
				{
					Marshal.ReleaseComObject(VideoPlaylist);
					VideoPlaylist = null;
				}
				Marshal.ReleaseComObject(iTunes);
				iTunes = null;

				IsInitialised = false;
				Plugin.MbApiInterface.MB_SendNotification(Plugin.CallbackType.StorageEject);
			}

			SyncItunesButNotDevice = false;
		}

		private static readonly DateTime MinITunesDateTime = new DateTime(1899, 12, 30);

		public static void SynchronizeRatingsAndPlaysFromItunesToMB()
		{
			try
			{
				foreach (var track in MusicPlaylist.GetAllTracks())
				{
					Plugin.MbApiInterface.MB_SetBackgroundTaskMessage(Text.L("Syncing play counts and/or ratings from iTunes: {0}", track.Name));

					var file = new MusicBeeFile(track.Location);
					if (!file.Exists) continue;

					var mbLastPlayed = file.LastPlayed;
					var itLastPlayed = track.MusicBeeLastPlayed();

					if (mbLastPlayed < itLastPlayed)
					{
						//if (flags.HasFlag(Plugin.SynchronisationSettings.SyncPlayCount2Way))
						{
							file.PlayCount = track.PlayedCount;
							file.LastPlayed = itLastPlayed;
							file.SkipCount = track.SkippedCount;
						}
						//if (flags.HasFlag(Plugin.SynchronisationSettings.SyncRating2Way))
						{
							file.Rating = track.MusicBeeRating();
							file.RatingAlbum = track.MusicBeeAlbumRating();
						}
						file.CommitChanges();
					}

					Marshal.ReleaseComObject(track);
				}
			}
			finally
			{
				Plugin.MbApiInterface.MB_SetBackgroundTaskMessage("");
			}
		}

		public static bool Synchronise(SynchronisationSettings flags, KeyValuePair<int, string[]>[] files)
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

			var playlistKeys = new Dictionary<long, string>();
			var trackKeys = new Dictionary<long, MusicBeeFile>();

			var someTracksWereSkipped = false;

			foreach (KeyValuePair<int, string[]> item in files)
			{
				// determine if the the file needs to be synched
				// ...
				// if yes, indicate to MusicBee that you want to synch the file - the returned filename is either the same as the supplied filename or if re-encoding/ forced embeding artwork, a temporary filename is returned
				// if filename is returned as null, that means MusicBee wasnt able to encode the file and it should be skipped from synchronisation
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
					// sync the file to the category (Key) requested
					if (item.Key == (int)SynchronisationCategory.Playlist) //Playlist. Lets remember synced playlists for later processing
					{
						var name = Regex.Replace(item.Value[0], "^.*\\\\(.*)(\\..*)", "$1");
						var playlist = iTunes.GetPlaylist(name) ?? iTunes.CreatePlaylist(name);
						var key = iTunes.GetPersistentId(playlist);
						Marshal.ReleaseComObject(playlist);
						playlistKeys[key] = item.Value[0];
						success = true;
						errorMessage = null;
					}
					else //Media file
					{
						var currentPlaylist = MusicPlaylist;
						if ((SynchronisationCategory)item.Key == SynchronisationCategory.Audiobook)
							currentPlaylist = AudiobooksPlaylist;
						else if ((SynchronisationCategory)item.Key == SynchronisationCategory.Podcast)
							currentPlaylist = PodcastsPlaylist;
						else if ((SynchronisationCategory)item.Key == SynchronisationCategory.Video)
							currentPlaylist = VideoPlaylist;

						if (currentPlaylist == null)
						{
							errorMessage = Text.L("No appropriate category found in iTunes library");
						}
						else
						{
							var filename = Plugin.MbApiInterface.Sync_FileStart(item.Value[0]);

							try
							{
								var mbFile = new MusicBeeFile(item.Value[0]);
								IITTrack libraryTrack = null;
								string iTunesTrackFilename = null;

								if (mbFile.WebFile)
								{
									iTunesTrackFilename = mbFile.Url;
								}
								else if (!File.Exists(filename))
								{
									throw new IOException(Text.L("Track source file not found: {0}", filename));
								}
								else if (mbFile.Url == filename)
								{
									iTunesTrackFilename = filename;
								}
								else
								{
									// Track was converted to a format that iTunes accepts...
									// Create a unique name for the converted file and store it in the MB file record
									var trackGUID = mbFile.ReencodingFileName;

									if (string.IsNullOrEmpty(trackGUID))
									{
										trackGUID = Guid.NewGuid().ToString();
										mbFile.ReencodingFileName = trackGUID;
										mbFile.CommitChanges();
									}

									iTunesTrackFilename = Path.Combine(Plugin.ReencodedFilesStorage, trackGUID + item.Value[1]);
									File.Copy(filename, iTunesTrackFilename, true);
								}

								var syncKey = mbFile.ITunesKey;

								if (syncKey != 0) //Track was synced before
								{
									libraryTrack = iTunes.GetTrackByPersistentId(syncKey);

									if (libraryTrack == null)
									{
										Trace.WriteLine("A file in MusicBee appears to have been sync'ed to iTunes before but is not found in iTunes: " + mbFile.Url);
										syncKey = 0;
									}
									else if (!mbFile.WebFile)
									{
										//Local or local network file
										((IITFileOrCDTrack)libraryTrack).UpdateInfoFromFile();
									}
								}

								var addingFile = false;

								if (libraryTrack == null) //Track was never synced before or was deleted from iTunes library. Lets add track to iTunes library. 
								{
									if (!mbFile.WebFile) //Local or local network file
									{
										currentPlaylist.AddFile(iTunesTrackFilename).Await();
										addingFile = true;

										// Verify the file was added to iTunes
										var endTime = DateTime.Now.AddSeconds(AddOperationTimeout);
										while (libraryTrack == null && DateTime.Now < endTime)
										{
											var foundTracks = currentPlaylist.Search(mbFile.Title, ITPlaylistSearchField.ITPlaylistSearchFieldSongNames);
											if (foundTracks == null) continue;
											foreach (IITTrack track in foundTracks)
											{
												if (track.Kind == ITTrackKind.ITTrackKindFile)
												{
													if (((IITFileOrCDTrack)track).Location == iTunesTrackFilename)
													{
														libraryTrack = track;
														syncKey = iTunes.GetPersistentId(libraryTrack);
														mbFile.ITunesKey = syncKey;
														mbFile.CommitChanges();
														break;
													}
												}
											}
										}

										if (libraryTrack == null)
										{
											throw new Exception(Text.L("Failed to add track to iTunes library"));
										}
									}
									else //Web file
									{
										libraryTrack = currentPlaylist.AddURL(iTunesTrackFilename);
									}
								}

								// Sync ratings & play counts...
								var mbLastPlayed = mbFile.LastPlayed.MusicBeeToITunes();
								if (addingFile || mbLastPlayed > libraryTrack.PlayedDate)
								{
									libraryTrack.RepeatTrackOperationUntilNoConflicts(t => t.PlayedDate = mbLastPlayed);
									libraryTrack.RepeatTrackOperationUntilNoConflicts(t => t.PlayedCount = mbFile.PlayCount);
									libraryTrack.RepeatTrackOperationUntilNoConflicts(t => t.SetMusicBeeRating(mbFile.Rating));
									if (libraryTrack.Kind == ITTrackKind.ITTrackKindFile)
									{
										libraryTrack.RepeatTrackOperationUntilNoConflicts(t => ((IITFileOrCDTrack)t).SkippedCount = mbFile.SkipCount);
										libraryTrack.RepeatTrackOperationUntilNoConflicts(t => ((IITFileOrCDTrack)t).SetMusicBeeAlbumRating(mbFile.RatingAlbum));
									}
								}

								if (!mbFile.WebFile) //Local or local network file. Lets set last modified time of iTunes track to the same as MB track.
								{
									var sourceFileInfo = new FileInfo(mbFile.Url);
									var destinationFileInfo = new FileInfo(iTunesTrackFilename);
									destinationFileInfo.LastWriteTimeUtc = sourceFileInfo.LastWriteTimeUtc;
								}

								Marshal.ReleaseComObject(libraryTrack);

								trackKeys[syncKey] = mbFile;

								success = true;
								errorMessage = null;
							}
							finally
							{
								// when the file synch is done
								if (filename != null)
									Plugin.MbApiInterface.Sync_FileEnd(item.Value[0], success, errorMessage);
							}
						}
					}
				}
				catch (ArgumentException ex)
				{
					Trace.WriteLine(ex); 
					success = true;
					errorMessage = null;
					someTracksWereSkipped = true;
				}
				catch (Exception ex)
				{
					Trace.WriteLine(ex); 
					if (errorMessage == null)
						errorMessage = ex.Message;
				}
			}

			try
			{
				// Remove non-sync'ed tracks
				foreach (var track in MusicPlaylist.GetAllTracks())
				{
					var key = iTunes.GetPersistentId(track);
					if (!trackKeys.ContainsKey(key))
					{
						Plugin.MbApiInterface.MB_SetBackgroundTaskMessage(Text.L("Removing track: \"{0}\"", track.Name));
						track.Delete();
					}
					Marshal.ReleaseComObject(track);
				}

				// Remove non-sync'ed playlists and populate the remaing ones
				foreach (var playlist in iTunes.GetPlaylists())
				{
					var key = iTunes.GetPersistentId(playlist);
					if (playlistKeys.ContainsKey(key))
					{
						Plugin.MbApiInterface.MB_SetBackgroundTaskMessage(Text.L("Clearing playlist: \"{0}\"", playlist.Name));
						while (playlist.Tracks.Count > 0)
						{
							var track = playlist.Tracks[1];
							track.Delete();
							Marshal.ReleaseComObject(track);
						}
						if (Plugin.MbApiInterface.Playlist_QueryFiles(playlistKeys[key]))
						{
							var m = 0;
							var playlistFiles = Plugin.MbApiInterface.Playlist_QueryGetAllFiles().Split(Plugin.FilesSeparators, StringSplitOptions.RemoveEmptyEntries);
							foreach (var playlistFile in playlistFiles)
							{
								m++;
								Plugin.MbApiInterface.MB_SetBackgroundTaskMessage(Text.L("Adding track {0} of {1} to playlist: \"{2}\"", m, playlistFiles.Length, playlist.Name));

								var track = iTunes.GetTrackByPersistentId(new MusicBeeFile(playlistFile).ITunesKey);
								try
								{
									playlist.AddTrack(track);
								}
								catch (Exception ex)
								{
									Trace.WriteLine(ex);
								}
								Marshal.ReleaseComObject(track);
							}
						}
					}
					else
					{
						Plugin.MbApiInterface.MB_SetBackgroundTaskMessage(Text.L("Removing playlist: \"{0}\"", playlist.Name));
						playlist.Delete();
					}
					Marshal.ReleaseComObject(playlist);
				}

				if (!SyncItunesButNotDevice)
				{
					Plugin.MbApiInterface.MB_SetBackgroundTaskMessage(Text.L("Syncing iTunes with iPod/iPhone..."));
					iTunes.UpdateIPod();
					// TODO: Wait for sync to finish?
				}
			}
			catch (Exception ex)
			{
				Trace.WriteLine(ex);
				lastEx = ex;
				return false;
			}
			finally
			{
				Plugin.MbApiInterface.MB_SetBackgroundTaskMessage("");
				Plugin.MbApiInterface.MB_RefreshPanels();
				SynchronizationInProgress = false;
			}

			if (someTracksWereSkipped)
				MessageBox.Show(Text.L("Some duplicated tracks were skipped."));

			return lastEx == null;
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

			List<KeyValuePair<byte, string>[]> files = new List<KeyValuePair<byte, string>[]>();

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
				//foundTrack = iTunes.LibraryPlaylist.Tracks.get_ItemByPersistentID(url[0], url[1]);
				foundTrack = iTunes.GetTrackByPersistentId(((long)url[0] << 32) | url[1]);

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
				Trace.WriteLine(ex); 
				lastEx = ex;
			}

			return null;
		}

		public static KeyValuePair<byte, string>[][] GetPlaylistFiles(string id)
		{
			List<KeyValuePair<byte, string>[]> playlistFiles = new List<KeyValuePair<byte, string>[]>();

			var libraryPlaylist = iTunes.GetPlaylist(id);
			if (libraryPlaylist == null)
			{
				KeyValuePair<byte, string>[] file = new KeyValuePair<byte, string>[TagCount];
				playlistFiles.Add(file);
				return playlistFiles.ToArray();
			}

			ITTracksCache.Clear();

			foreach (IITTrack currTrack in libraryPlaylist.Tracks)
			{
				if (currTrack.Kind == ITTrackKind.ITTrackKindFile)
				{
					IITFileOrCDTrack fileTrack = (IITFileOrCDTrack)currTrack;

					object fileTrackObject = fileTrack;
					int highID, lowID;
					iTunes.GetITObjectPersistentIDs(ref fileTrackObject, out highID, out lowID);
					int a, b, c, d;
					fileTrack.GetITObjectIDs(out a, out b, out c, out d);
					ITTracksCache.Add(fileTrack.Location, new int[] { highID, lowID });

					playlistFiles.Add(fileTrack.ToMusicBeeFileProperties());
				}
			}

			return playlistFiles.ToArray();
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

		public static Stream GetStream(string url)
		{
			FileStream stream = new FileStream(url, FileMode.Open);
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