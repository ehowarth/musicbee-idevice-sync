using iTunesLib;
using MusicBeeDeviceSyncPlugin;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace MusicBeePlugin
{
	public partial class Plugin
	{
		private const int TagCount = 23;
		public static readonly Assembly Assembly = Assembly.GetExecutingAssembly();
		public static readonly AssemblyName AssemblyName = Assembly.GetName();

		public static readonly PluginInfo Info = new PluginInfo
		{
			PluginInfoVersion = Plugin.PluginInfoVersion,
			Name = Text.L("iPod & iPhone Sync"),
			Description = Text.L("Takes control of the iTunes application to synchronize an iPod, iPhone, or iPad"),
			Author = "boroda74",
			TargetApplication = "iTunes",   // current only applies to artwork, lyrics or instant messenger name that appears in the provider drop down selector or target Instant Messenger
			Type = Plugin.PluginType.Storage,
			VersionMajor = (short)AssemblyName.Version.Major,  // .net version
			VersionMinor = (short)AssemblyName.Version.Minor, // plugin version
			Revision = (short)AssemblyName.Version.Build, // number of days since 2000-01-01 at build time
			MinInterfaceVersion = 28,
			MinApiRevision = 32,
			ReceiveNotifications = Plugin.ReceiveNotificationFlags.StartupOnly,
			ConfigurationPanelHeight = 0,   // height in pixels that musicbee should reserve in a panel for config settings. When set, a handle to an empty panel will be passed to the Configure function
		};

		public static MusicBeeApiInterface MbApiInterface;

		private DeviceProperties Device = new DeviceProperties
		{
			ShowCategoryNodes = true,
			AudiobooksSupported = false,
			PodcastsSupported = false,
			VideoSupported = false,
			SyncExternalArtwork = false,
			SyncAllowFileRemoval = false, // This plugin always removes non-sync-ed files from iTunes. It doesn't ask for permission.
			SyncAllowRating2Way = false, // This plugin always syncs plays from iTunes. It doesn't ask for permission.
			SyncAllowPlayCount2Way = false, // This plugin always syncs plays from iTunes. It doesn't ask for permission.
			SyncAllowPlaylists2Way = false,
			SupportedFormats =
				SynchronisationSupportedFormats.SyncMp3Supported |
				SynchronisationSupportedFormats.SyncAacSupported |
				SynchronisationSupportedFormats.SyncAlacSupported |
				SynchronisationSupportedFormats.SyncWavSupported,
			DeviceIcon64 = (Bitmap)MusicBeeDeviceSyncPlugin.Properties.Resources.ResourceManager.GetObject("iTunes64"),
		};

		private bool uninstalled = false;

		private Form MbForm;
		private ToolStripItem OpenMenu;

		private string ReencodedFilesStorage;

		private bool? ReadyForSync = false;

		private iTunesApp iTunes;
		private IITIPodSource IPodSource = null;

		private Exception lastEx = null;

		private bool SynchronizationInProgress = false;
		private bool AbortSynchronization = false;

		private readonly Dictionary<string, long> SelectedPlaylistLocationsToPersistentIds = new Dictionary<string, long>();

		/// <summary>
		/// Called by MusicBee whenever it is loading the plugin whether at startup or when enabling the plugin.
		/// The pointer parameter should be copied to a <c>MusicBeeApiInterface</c> structure.
		/// </summary>
		/// <param name="apiInterfacePtr">Pointer to a <c>MusicBeeApiInterface</c> object</param>
		/// <returns></returns>
		public PluginInfo Initialise(IntPtr apiInterfacePtr)
		{
			MbApiInterface = new MusicBeeApiInterface();
			MbApiInterface.Initialise(apiInterfacePtr);

			ReencodedFilesStorage = MbApiInterface.Setting_GetPersistentStoragePath() + "iPod & iPhone Driver";
			if (!Directory.Exists(ReencodedFilesStorage))
			{
				Directory.CreateDirectory(ReencodedFilesStorage);
			}

			MbForm = (Form)Form.FromHandle(MbApiInterface.MB_GetWindowHandle());

			OpenMenu = MbApiInterface.MB_AddMenuItem(
				"mnuTools/" + Text.L("Open iPod && iPhone Sync"),
				Text.L("Opens iTunes to begin sync"),
				ToggleItunesOpenedAndClosed);

			return Info;
		}

		private void ToggleItunesOpenedAndClosed(object sender, EventArgs e)
		{
			while (ReadyForSync == null) ; // Wait if already being toggled open

			if (ReadyForSync == false)
			{
				Backgrounding.RunInBackground(StartITunes);
			}

			OpenMenu.Enabled = false;
		}

		// save any persistent settings in a sub-folder of this path
		// string dataPath = mbApiInterface.Setting_GetPersistentStoragePath();
		// panelHandle will only be set if you set about.ConfigurationPanelHeight to a non-zero value
		// keep in mind the panel width is scaled according to the font the user has selected
		// if about.ConfigurationPanelHeight is set to 0, you can display your own popup window
		public bool Configure(IntPtr panelHandle)
		{
			var choice = MessageBox.Show(
				MbForm,
				Text.L("Clear All Sync Data?"),
				Info.Name + ' ' + AssemblyName.Version,
				MessageBoxButtons.YesNo
				);

			if (choice == DialogResult.Yes)
			{
				RemoveSyncTagsFrommAllLibraryFiles();
				DeleteAllReencodedFiles();
			}

			return true;
		}

		private static void RemoveSyncTagsFrommAllLibraryFiles()
		{
			foreach (var file in MusicBeeFile.AllFiles())
			{
				file.ClearPluginTags();
			}
		}

		private void DeleteAllReencodedFiles()
		{
			if (!Directory.Exists(ReencodedFilesStorage))
			{
				Directory.Delete(ReencodedFilesStorage, true);
				Directory.CreateDirectory(ReencodedFilesStorage);
			}
		}

		// called by MusicBee when the user clicks Apply or Save in the MusicBee Preferences screen.
		// its up to you to figure out whether anything has changed and needs updating
		public void SaveSettings()
		{
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

			OpenMenu.Visible = false;

			RemoveSyncTagsFrommAllLibraryFiles();

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

		/// <summary>
		/// Called by MusicBee in response to <c>MB_SendNotification(Plugin.CallbackType.StorageReady)</c>
		/// </summary>
		/// <param name="handle">Pointer to buffer in which to copy a <c>DeviceProperties</c> object</param>
		/// <returns></returns>
		public bool GetDeviceProperties(IntPtr handle)
		{
			Marshal.StructureToPtr(Device, handle, false);
			Trace.WriteLine("Reporting device \"" + Device.DeviceName + "\" to MusicBee");
			return true;
		}

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
		public bool Synchronise(SynchronisationSettings flags, KeyValuePair<int, string[]>[] syncItems)
		{
			while (ReadyForSync == null) ; // Wait if still opening iTunes

			try
			{
				if (SynchronizationInProgress)
				{
					lastEx = null;
					return false;
				}

				SynchronizationInProgress = true;

				var playlistKeys = new Dictionary<long, string>();
				var trackKeys = new Dictionary<long, MusicBeeFile>();

				foreach (var item in syncItems)
				{
					if (AbortSynchronization)
					{
						SynchronizationInProgress = false;
						AbortSynchronization = false;
						lastEx = null;
						return true;
					}

					if (item.Key == (int)SynchronisationCategory.Playlist)
					{
						// Create or verify playlist. Populate after all files have been processed.
						var name = Regex.Replace(item.Value[0], "^.*\\\\(.*)(\\..*)", "$1");
						var playlist = iTunes.GetPlaylist(name) ?? iTunes.CreatePlaylist(name);
						var key = iTunes.GetPersistentId(playlist);
						Marshal.ReleaseComObject(playlist);
						playlistKeys[key] = item.Value[0];
					}
					else
					{
						// item.Value[0] is the URL to a MusicBee file to be sync'ed
						// item.Value[1] is the extension of the file

						// indicate to MusicBee that you want to synch the file
						// the returned filename is either the same as the supplied filename or
						// if re-encoding/forced embedding artwork, a temporary filename is returned
						var filename = Plugin.MbApiInterface.Sync_FileStart(item.Value[0]);

						// if filename is returned as null, that means MusicBee wasnt able to encode the file and it should be skipped from synchronisation
						if (filename == null) continue;

						bool success = false;
						string errorMessage = null;

						try
						{
							var mbFile = new MusicBeeFile(item.Value[0]);
							IITTrack itTrack = null;
							string itTrackPath = null;

							if (mbFile.WebFile)
							{
								itTrackPath = mbFile.Url;
							}
							else if (!File.Exists(filename))
							{
								throw new IOException(Text.L("Track source file not found: {0}", filename));
							}
							else if (mbFile.Url == filename)
							{
								itTrackPath = filename;
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

								itTrackPath = Path.Combine(ReencodedFilesStorage, trackGUID + item.Value[1]);
								File.Copy(filename, itTrackPath, true);
							}

							var itKey = mbFile.ITunesKey;

							// Track was synced before
							if (itKey != 0)
							{
								itTrack = iTunes.GetTrackByPersistentId(itKey);

								if (itTrack == null)
								{
									Trace.WriteLine("A file in MusicBee appears to have been sync'ed to iTunes before but is not found in iTunes: " + mbFile.Url);
									itKey = 0;
								}
								else if (!mbFile.WebFile)
								{
									//Local or local network file
									((IITFileOrCDTrack)itTrack).UpdateInfoFromFile();
								}
							}

							// Track was never synced before or was deleted from iTunes library
							if (itTrack == null)
							{
								if (mbFile.WebFile)
								{
									itTrack = iTunes.LibraryPlaylist.AddURL(itTrackPath);
								}
								else
								{
									var operation = iTunes.LibraryPlaylist.AddFile(itTrackPath).Await();
									var tracks = operation.Tracks;
									itTrack = tracks[1];
									Marshal.ReleaseComObject(tracks);
									Marshal.ReleaseComObject(operation);
									itKey = iTunes.GetPersistentId(itTrack);
									mbFile.ITunesKey = itKey;
									mbFile.CommitChanges();
								}
							}

							// Sync ratings & play counts to iTunes
							itTrack.SyncMusicBeeHistoryToITunes(mbFile);
							mbFile.SyncFileTimestamp(itTrackPath);

							Marshal.ReleaseComObject(itTrack);

							trackKeys[itKey] = mbFile;

							success = true;
							errorMessage = null;
						}
						catch (Exception ex)
						{
							Trace.WriteLine(ex);
							lastEx = ex;
							if (errorMessage == null)
								errorMessage = ex.Message;
						}
						finally
						{
							Plugin.MbApiInterface.Sync_FileEnd(item.Value[0], success, errorMessage);
						}
					}
				}

				// Remove non-sync'ed tracks
				foreach (var track in iTunes.GetAllTracks())
				{
					var key = iTunes.GetPersistentId(track);
					if (!trackKeys.ContainsKey(key))
					{
						Plugin.MbApiInterface.MB_SetBackgroundTaskMessage(Text.L("Removing track: \"{0}\"", track.Name));
						track.Delete();
					}
					Marshal.ReleaseComObject(track);
				}

				// Remove non-sync'ed playlists and populate the remaining ones
				foreach (var playlist in iTunes.GetPlaylists())
				{
					var key = iTunes.GetPersistentId(playlist);
					if (playlistKeys.ContainsKey(key))
					{
						Plugin.MbApiInterface.MB_SetBackgroundTaskMessage(Text.L("Clearing playlist: \"{0}\"", playlist.Name));
						playlist.DeleteAllTracks();
						var m = 0;
						var playlistFiles = MusicBeeFile.GetPlaylistFiles(playlistKeys[key]).ToArray();
						foreach (var playlistFile in playlistFiles)
						{
							m++;
							Plugin.MbApiInterface.MB_SetBackgroundTaskMessage(Text.L("Adding track {0} of {1} to playlist: \"{2}\"", m, playlistFiles.Length, playlist.Name));
							iTunes.AddTrackToPlaylistByPersistentId(playlist, playlistFile.ITunesKey);
						}
					}
					else
					{
						Plugin.MbApiInterface.MB_SetBackgroundTaskMessage(Text.L("Removing playlist: \"{0}\"", playlist.Name));
						playlist.Delete();
					}
					Marshal.ReleaseComObject(playlist);
				}

				if (IPodSource != null)
				{
					Plugin.MbApiInterface.MB_SetBackgroundTaskMessage(Text.L("Syncing iTunes with {0}...", IPodSource.Name));
					iTunes.UpdateIPod();
					// TODO: Wait for sync to finish?
				}

				return lastEx == null;
			}
			catch (Exception ex)
			{
				Trace.WriteLine(ex);
				MbApiInterface.MB_SendNotification(CallbackType.StorageFailed);
				lastEx = ex;
				return false;
			}
			finally
			{
				Plugin.MbApiInterface.MB_SetBackgroundTaskMessage("");
				SynchronizationInProgress = false;
			}
		}

		// called when user has requested the synchronisation aborts
		public void SynchroniseAbort()
		{
			AbortSynchronization = true;
		}

		// called when user has requested the device be ejected
		public bool Eject()
		{
			CloseITunes();

			if (!uninstalled)
				SaveSettings();

			Trace.Flush();

			return true;
		}

		public void Refresh()
		{
			// TODO: what to refresh
		}

		public Bitmap GetIcon()
		{
			return (Bitmap)MusicBeeDeviceSyncPlugin.Properties.Resources.ResourceManager.GetObject("iTunes");
		}

		public bool IsReady()
		{
			return ReadyForSync == true;
		}

		private void StartITunes()
		{
			try
			{
				ReadyForSync = null;

				Plugin.MbApiInterface.MB_SetBackgroundTaskMessage(Text.L("Starting iTunes..."));

				iTunes = new iTunesApp();

#if !DEBUG
				iTunes.BrowserWindow.Minimized = true;
#endif

				if (LoadITunesStroage())
				{
					// Sync play history and ratings from iTunes right now.
					// Cannot wait for the Synchronize command from MB before collecting ratings and history
					// because this data may affect the file list passed to Synchronize.
					//		if (flags.HasFlag(Plugin.SynchronisationSettings.SyncPlayCount2Way) || flags.HasFlag(Plugin.SynchronisationSettings.SyncRating2Way))
					SyncITunesHistoryToMusicBee();

					ReadyForSync = true;
				}
				else
				{
					ReadyForSync = false;
					CloseITunes();
				}
			}
			catch (Exception ex)
			{
				Trace.WriteLine(ex);
				lastEx = ex;
				MessageBox.Show(ex.Message);

				ReadyForSync = false;
				CloseITunes();
			}
			finally
			{
				Plugin.MbApiInterface.MB_SetBackgroundTaskMessage("");
			}
		}

		private void SyncITunesHistoryToMusicBee()
		{
			try
			{
				Plugin.MbApiInterface.MB_SetBackgroundTaskMessage(Text.L("Syncing play counts and ratings from iTunes..."));

				if (IPodSource != null)
				{
					iTunes.UpdateIPod();
					Thread.Sleep(20000);
					// TODO: Wait until sync is done
				}

				foreach (var track in iTunes.GetAllTracks())
				{
					Plugin.MbApiInterface.MB_SetBackgroundTaskMessage(Text.L("Syncing play counts and ratings from iTunes: \"{0}\"", track.Name));

					var file = new MusicBeeFile(track.Location);
					if (!file.Exists) continue;

					track.SyncITunesHistoryToMusicBee(file);

					Marshal.ReleaseComObject(track);
				}
			}
			finally
			{
				Plugin.MbApiInterface.MB_SetBackgroundTaskMessage("");
			}
		}

		/// <summary>
		/// Returns true if this loads a MusicBee storage device.
		/// Normally the device is the value stored in IPodSource.
		/// However IPodSource may be null if the user chooses to sync iTunes only.
		/// </summary>
		private bool LoadITunesStroage()
		{
			var waitingWindow = ModelessMessageBox.Show(
				MbForm,
				Text.L("Connect an Apple device to the computer \r\nor click Skip to sync only the iTunes application."),
				Text.L("Waiting for iTunes to detect device"),
				Text.L("Skip"),
				Text.L("Cancel"));

			try
			{
				while (IPodSource == null && waitingWindow.DialogResult == 0)
				{
					for (int i = 1; i <= iTunes.Sources.Count; i++)
					{
						var source = iTunes.Sources[i];
						if (source.Kind == ITSourceKind.ITSourceKindIPod)
						{
							IPodSource = (IITIPodSource)source;
							while (IPodSource == null || IPodSource.Name == null)
							{
								Trace.WriteLine("Waiting for IITIPodSource to be fully loaded by iTunes...");
								Thread.Sleep(1000);
							}
							Device.DeviceName = IPodSource.Name;
							Device.FirmwareVersion = IPodSource.SoftwareVersion;
							Device.FreeSpace = (ulong)IPodSource.FreeSpace;
							Device.TotalSpace = (ulong)IPodSource.Capacity;
							Device.Model = "--";
							Device.Manufacturer = "Apple Inc.";
							break;
						}
						Marshal.ReleaseComObject(source);
					}
					if (IPodSource == null)
					{
						Thread.Sleep(1000);
					}
				}
			}
			finally
			{
				MbForm.Invoke((Action)waitingWindow.Close);
			}

			if (IPodSource == null)
			{
				if (waitingWindow.DialogResult == DialogResult.Cancel )
				{
					return false;
				}

				Device.DeviceName = Text.L("iTunes Sync Only -- No device");
				Device.FirmwareVersion = iTunes.Version;
				Device.Model = Text.L("No device");
				Device.Manufacturer = "Apple Inc.";
				Device.FreeSpace = 0;
				Device.TotalSpace = 0;
			}

			// Tell MusicBee to add an item to the Devices panel for this.
			// Music Bee calls GetDeviceProperties for details
			Plugin.MbApiInterface.MB_SendNotification(Plugin.CallbackType.StorageReady);

			return true;
		}

		private void CloseITunes()
		{
			while (ReadyForSync == null) ; // Wait if still opening iTunes before trying to close

			ReadyForSync = false;

			if (Device.DeviceName != null)
			{
				Plugin.MbApiInterface.MB_SendNotification(Plugin.CallbackType.StorageEject);
				Device.DeviceName = null;
			}

			if (IPodSource != null)
			{
				IPodSource.EjectIPod();
				Marshal.ReleaseComObject(IPodSource);
				IPodSource = null;
			}

			if (iTunes != null)
			{
				iTunes.Quit();
				Marshal.ReleaseComObject(iTunes);
				iTunes = null;
			}

			OpenMenu.Enabled = true;
		}

		public bool FolderExists(string path)
		{
			return (string.IsNullOrEmpty(path) || path == @"\");
		}

		public string[] GetFolders(string path)
		{
			string[] folders;
			folders = new string[1];
			folders[0] = "";

			return folders;
		}

		public KeyValuePair<string, string>[] GetPlaylists()
		{
			return iTunes.GetPlaylists().Select(playlist =>
			{
				try
				{
					return new KeyValuePair<string, string>(playlist.Name, playlist.Name);
				}
				finally
				{
					Marshal.ReleaseComObject(playlist);
				}
			}).ToArray();
		}

		private KeyValuePair<byte, string>[][] GetPlaylistTracks(IITPlaylist playlist)
		{
			var files = new List<KeyValuePair<byte, string>[]>();

			SelectedPlaylistLocationsToPersistentIds.Clear();

			foreach (IITTrack currTrack in playlist.Tracks)
			{
				if (currTrack.Kind == ITTrackKind.ITTrackKindFile)
				{
					var fileTrack = (IITFileOrCDTrack)currTrack;
					if (fileTrack.Location != null)
					{
						SelectedPlaylistLocationsToPersistentIds.Add(fileTrack.Location, iTunes.GetPersistentId(fileTrack));
						files.Add(fileTrack.ToMusicBeeFileProperties());
					}
				}
				Marshal.ReleaseComObject(currTrack);
			}

			return files.ToArray();
		}

		public KeyValuePair<byte, string>[][] GetPlaylistFiles(string id)
		{
			var playlistFiles = new List<KeyValuePair<byte, string>[]>();

			var libraryPlaylist = iTunes.GetPlaylist(id);

			if (libraryPlaylist == null)
			{
				var file = new KeyValuePair<byte, string>[TagCount];
				playlistFiles.Add(file);
				return playlistFiles.ToArray();
			}

			SelectedPlaylistLocationsToPersistentIds.Clear();

			foreach (IITTrack currTrack in libraryPlaylist.Tracks)
			{
				if (currTrack.Kind == ITTrackKind.ITTrackKindFile)
				{
					IITFileOrCDTrack fileTrack = (IITFileOrCDTrack)currTrack;
					SelectedPlaylistLocationsToPersistentIds.Add(fileTrack.Location, iTunes.GetPersistentId(fileTrack));
					playlistFiles.Add(fileTrack.ToMusicBeeFileProperties());
				}
				Marshal.ReleaseComObject(currTrack);
			}

			Marshal.ReleaseComObject(libraryPlaylist);

			return playlistFiles.ToArray();
		}

		public KeyValuePair<byte, string>[][] GetFiles(string path)
		{
			lastEx = null;
			KeyValuePair<byte, string>[][] files = null;

			if (ReadyForSync != true)
			{
				files = new KeyValuePair<byte, string>[0][];
			}
			else
			{
				files = GetPlaylistTracks(iTunes.LibraryPlaylist);
			}

			return files;
		}

		public bool FileExists(string url)
		{
			long ids = 0;
			if (SelectedPlaylistLocationsToPersistentIds.TryGetValue(url, out ids))
			{
				var foundTrack = iTunes.GetTrackByPersistentId(ids);
				if (foundTrack != null)
				{
					Marshal.ReleaseComObject(foundTrack);
					return true;
				}
			}
			return false;
		}

		public KeyValuePair<byte, string>[] GetFile(string url)
		{
			long ids = 0;
			if (SelectedPlaylistLocationsToPersistentIds.TryGetValue(url, out ids))
			{
				var foundTrack = iTunes.GetTrackByPersistentId(ids);
				if (foundTrack != null)
				{
					try
					{
						if (foundTrack.Kind == ITTrackKind.ITTrackKindFile)
						{
							return ((IITFileOrCDTrack)foundTrack).ToMusicBeeFileProperties();
						}
					}
					finally
					{
						Marshal.ReleaseComObject(foundTrack);
					}
				}
			}

			return new KeyValuePair<byte, string>[TagCount];
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

		public byte[] GetFileArtwork(string url)
		{
			lastEx = null;

			try
			{
				long ids = 0;
				if (SelectedPlaylistLocationsToPersistentIds.TryGetValue(url, out ids))
				{
					var foundTrack = iTunes.GetTrackByPersistentId(ids);
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
			}
			catch (Exception ex)
			{
				Trace.WriteLine(ex);
				lastEx = ex;
			}

			return null;
		}

		public Stream GetStream(string url)
		{
			return new FileStream(url, FileMode.Open);
		}

		public Exception GetError()
		{
			return lastEx;
		}
	}
}