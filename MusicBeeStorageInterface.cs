using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace MusicBeePlugin
{
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
}
