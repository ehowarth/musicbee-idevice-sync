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
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace MusicBeePlugin
{
	public partial class Plugin
	{
		public static bool DeveloperMode = false;
		private bool uninstalled = false;

		public static MusicBeeApiInterface MbApiInterface;

		public static Form MbForm;

		public static string ReencodedFilesStorage;

		public static readonly char[] FilesSeparators = { '\0' };

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

			ReencodedFilesStorage = MbApiInterface.Setting_GetPersistentStoragePath() + "iPod & iPhone Driver";
			if (!Directory.Exists(ReencodedFilesStorage))
			{
				Directory.CreateDirectory(ReencodedFilesStorage);
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

			new ConfigurationDialog(ResetPlugin).ShowDialog();

			return true;
		}

		private static void ResetPlugin()
		{
			RemoveSyncTagsFrommAllLibraryFiles();
			DeleteAllReencodedFiles();
		}

		// TODO: Business logic doesn't belong in a UI class
		private static void RemoveSyncTagsFrommAllLibraryFiles()
		{
			foreach (var file in MusicBeeFile.AllFiles())
			{
				file.ClearPluginTags();
			}
		}

		// TODO: Business logic doesn't belong in a UI class
		private static void DeleteAllReencodedFiles()
		{
			if (!System.IO.Directory.Exists(Plugin.ReencodedFilesStorage))
			{
				System.IO.Directory.Delete(Plugin.ReencodedFilesStorage, true);
				System.IO.Directory.CreateDirectory(Plugin.ReencodedFilesStorage);
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

			foreach (var file in MusicBeeFile.AllFiles())
			{
				file.ClearPluginTags();
			}

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
		public static IITIPodSource IPodSource = null;

		private const int TagCount = 23;
		private const int AddOperationTimeout = 5; //Timeout of adding track to iTunes (1 means 0.5 sec, 2 - 1.5 sec, 3 - 2.5 sec, 30 - 29.5 sec, etc.)
		public static Exception lastEx = null;

		public static bool SyncItunesButNotDevice = false; //For testing

		public static bool SynchronizationInProgress = false;
		public static bool AbortSynchronization = false;

		private static Dictionary<string, long> SelectedPlaylistLocationsToPersistentIds = new Dictionary<string, long>();

		public static void Initialise()
		{
			try
			{
				IsInitialised = null;

				iTunes = new iTunesApp();
				iTunes.BrowserWindow.Minimized = true;

				var waitingWindow = new WaitingForIPod();
				Plugin.MbForm.AddOwnedForm(waitingWindow);
				Plugin.MbForm.Invoke((Action)waitingWindow.Show);

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

				Plugin.MbForm.Invoke((Action)waitingWindow.Close);

				if (IPodSource == null && !SyncItunesButNotDevice)
				{
					iTunes.Quit();
					IsInitialised = false;
					//SendNotificationHandler.Invoke(Plugin.CallbackType.StorageFailed);
					return;
				}

				if (IPodSource != null)
				{
					//var originalSize = IPodSource.FreeSpace;
					IPodSource.UpdateIPod();
					//while (true)
					//{
					Thread.Sleep(1000);
					//	if (Math.Abs(originalSize - IPodSource.FreeSpace) > 1)
					//	{
					//		break;
					//	}
					//}
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
				Marshal.ReleaseComObject(iTunes);
				iTunes = null;

				IsInitialised = false;
				Plugin.MbApiInterface.MB_SendNotification(Plugin.CallbackType.StorageEject);
			}

			SyncItunesButNotDevice = false;
		}

		public static void SynchronizeRatingsAndPlaysFromItunesToMB()
		{
			try
			{
				foreach (var track in iTunes.GetAllTracks())
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
					if (item.Key == (int)SynchronisationCategory.Playlist)
					{
						// Create or verify playlist. Populate after all files have been processed.
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

							if (libraryTrack == null) //Track was never synced before or was deleted from iTunes library. Lets add track to iTunes library. 
							{
								if (!mbFile.WebFile) //Local or local network file
								{
									var operation = //currentPlaylist
									iTunes.LibraryPlaylist.AddFile(iTunesTrackFilename).Await();
									Debug.Assert(!operation.InProgress);
									var tracks = operation.Tracks;
									Debug.Assert(tracks != null);
									Debug.Assert(tracks.Count == 1);
									libraryTrack = tracks[1];
									Marshal.ReleaseComObject(tracks);
									Marshal.ReleaseComObject(operation);
									syncKey = iTunes.GetPersistentId(libraryTrack);
									mbFile.ITunesKey = syncKey;
									mbFile.CommitChanges();
								}
								else //Web file
								{
									libraryTrack = iTunes.LibraryPlaylist.AddURL(iTunesTrackFilename);
								}
							}

							// Sync ratings & play counts...
							libraryTrack.RepeatTrackOperationUntilNoConflicts(t => t.PlayedDate = mbFile.LastPlayed.MusicBeeToITunes());
							libraryTrack.RepeatTrackOperationUntilNoConflicts(t => t.PlayedCount = mbFile.PlayCount);
							libraryTrack.RepeatTrackOperationUntilNoConflicts(t => t.SetMusicBeeRating(mbFile.Rating));
							if (libraryTrack.Kind == ITTrackKind.ITTrackKindFile)
							{
								libraryTrack.RepeatTrackOperationUntilNoConflicts(t => ((IITFileOrCDTrack)t).SkippedCount = mbFile.SkipCount);
								libraryTrack.RepeatTrackOperationUntilNoConflicts(t => ((IITFileOrCDTrack)t).SetMusicBeeAlbumRating(mbFile.RatingAlbum));
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

		public static KeyValuePair<byte, string>[] GetFile(string url)
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

		public static byte[] GetFileArtwork(string url)
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

		public static KeyValuePair<byte, string>[][] GetPlaylistFiles(string id)
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

		public static Stream GetStream(string url)
		{
			return new FileStream(url, FileMode.Open);
		}

		public static Exception GetError()
		{
			return lastEx;
		}
	}
}