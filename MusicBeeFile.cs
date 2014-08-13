using MusicBeePlugin;
using System;
using System.Collections.Generic;
using System.Linq;

namespace MusicBeeDeviceSyncPlugin
{
	sealed class MusicBeeFile
	{
		public MusicBeeFile(string url)
		{
			this.Url = url;
		}

		public string Url { get; private set; }

		public bool Exists
		{
			get { try { return Url == Plugin.MbApiInterface.Library_GetFileProperty(Url, Plugin.FilePropertyType.Url); } catch { return false; } }
		}

		public bool WebFile
		{
			get { return Url.StartsWith("http\\:") || Url.StartsWith("https\\:"); }
		}

		public string Title
		{
			get { return Plugin.MbApiInterface.Library_GetFileTag(Url, Plugin.MetaDataType.TrackTitle); }
		}

		public string ReencodingFileName
		{
			get { return Plugin.MbApiInterface.Library_GetDevicePersistentId(Url, Plugin.DeviceIdType.AppleDevice); }
			set { Plugin.MbApiInterface.Library_SetDevicePersistentId(Url, Plugin.DeviceIdType.AppleDevice, value); }
		}

		public long ITunesKey
		{
			get
			{
				var value = 0L;
				long.TryParse(Plugin.MbApiInterface.Library_GetDevicePersistentId(Url, Plugin.DeviceIdType.AppleDevice2), out value);
				return value;
			}
			set { Plugin.MbApiInterface.Library_SetDevicePersistentId(Url, Plugin.DeviceIdType.AppleDevice2, value.ToString()); }
		}

		public DateTime? LastPlayed
		{
			get
			{
				var dateString = Plugin.MbApiInterface.Library_GetFileProperty(Url, Plugin.FilePropertyType.LastPlayed);
				DateTime dateBuffer;
				return DateTime.TryParse(dateString, out dateBuffer)
					? dateBuffer
					: (DateTime?)null;
			}
			set { Plugin.MbApiInterface.Library_SetFileTag(Url, (Plugin.MetaDataType)Plugin.FilePropertyType.LastPlayed, value.HasValue ? value.ToString() : ""); }
		}

		public int PlayCount
		{
			get { return TryParse(Plugin.MbApiInterface.Library_GetFileProperty(Url, Plugin.FilePropertyType.PlayCount)); }
			set { Plugin.MbApiInterface.Library_SetFileTag(Url, (Plugin.MetaDataType)Plugin.FilePropertyType.PlayCount, value.ToString()); }
		}

		public int SkipCount
		{
			get { return TryParse(Plugin.MbApiInterface.Library_GetFileProperty(Url, Plugin.FilePropertyType.SkipCount)); }
			set { Plugin.MbApiInterface.Library_SetFileTag(Url, (Plugin.MetaDataType)Plugin.FilePropertyType.SkipCount, value.ToString()); }
		}

		public int Rating
		{
			get { return TryParse(Plugin.MbApiInterface.Library_GetFileTag(Url, Plugin.MetaDataType.Rating)); }
			set { Plugin.MbApiInterface.Library_SetFileTag(Url, Plugin.MetaDataType.Rating, value.ToString()); }
		}

		public int RatingAlbum
		{
			get { return TryParse(Plugin.MbApiInterface.Library_GetFileTag(Url, Plugin.MetaDataType.RatingAlbum)); }
			set { Plugin.MbApiInterface.Library_SetFileTag(Url, Plugin.MetaDataType.RatingAlbum, value.ToString()); }
		}

		private static int TryParse(string stringValue, int unknownValue = 0)
		{
			var value = unknownValue;
			int.TryParse(stringValue, out value);
			return value;
		}

		public void CommitChanges()
		{
			Plugin.MbApiInterface.Library_CommitTagsToFile(Url);
		}

		public void ClearPluginTags()
		{
			ReencodingFileName = "";
			ITunesKey = 0;
			CommitChanges();
		}

		public static IEnumerable<MusicBeeFile> AllFiles()
		{
			if (Plugin.MbApiInterface.Library_QueryFiles("domain=Library"))
			{
				return Plugin.MbApiInterface
					.Library_QueryGetAllFiles()
					.Split(Plugin.FilesSeparators, StringSplitOptions.RemoveEmptyEntries)
					.Select(url => new MusicBeeFile(url));
			}
			return new MusicBeeFile[0];
		}
	}
}
