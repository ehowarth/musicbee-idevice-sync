using iTunesLib;
using MusicBeePlugin;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace MusicBeeDeviceSyncPlugin
{
	static class iTunesExtensions
	{
		public static IITUserPlaylist GetPlaylist(this iTunesApp iTunes, string name)
		{
			foreach (var playlist in iTunes.GetPlaylists())
			{
				if (playlist.Name == name) return playlist;
				Marshal.ReleaseComObject(playlist);
			}
			return null;
		}

		public static IEnumerable<IITUserPlaylist> GetPlaylists(this iTunesApp iTunes)
		{
			var library = iTunes.LibrarySource;
			var playlists = library.Playlists;
			try
			{
				for (var i = playlists.Count; i > 0; --i)
				{
					var playlist = playlists[i];
					if (
						 playlist.Kind == ITPlaylistKind.ITPlaylistKindUser
						 && ((IITUserPlaylist)playlist).SpecialKind == ITUserPlaylistSpecialKind.ITUserPlaylistSpecialKindNone
						 && !((IITUserPlaylist)playlist).Smart
						 )
					{
						yield return (IITUserPlaylist)playlist;
					}
					else
					{
						Marshal.ReleaseComObject(playlist);
					}
				}
			}
			finally
			{
				Marshal.ReleaseComObject(playlists);
				Marshal.ReleaseComObject(library);
			}
		}

		public static IEnumerable<IITFileOrCDTrack> GetAllTracks(this iTunesApp iTunes)
		{
			var library = iTunes.LibraryPlaylist;
			var tracks = library.Tracks;
			try
			{
				for (var i = tracks.Count; i > 0; --i)
				{
					var track = tracks[i];
					if (track.Kind == ITTrackKind.ITTrackKindFile)
						yield return (IITFileOrCDTrack)track;
					else
						Marshal.ReleaseComObject(track);
				}
			}
			finally
			{
				Marshal.ReleaseComObject(tracks);
				Marshal.ReleaseComObject(library);
			}
		}

		public static long GetPersistentId(this iTunesApp iTunes, IITObject item)
		{
			object o = item;
			int high, low;
			iTunes.GetITObjectPersistentIDs(ref o, out high, out low);
			return Encode(high, low);
		}

		public static IITTrack GetTrackByPersistentId(this iTunesApp iTunes, long id)
		{
			var library = iTunes.LibraryPlaylist;
			var tracks = library.Tracks;
			try
			{
				return tracks.get_ItemByPersistentID(id.Hi(), id.Lo());
			}
			finally
			{
				Marshal.ReleaseComObject(tracks);
				Marshal.ReleaseComObject(library);
			}
		}

		//public static IITObject GetObjectByRuntimeId(this iTunesApp iTunes, BigInteger id)
		//{
		//	return iTunes.GetITObjectByID( id.SourceId(),  id.PlaylistId(),  id.TrackId(),  id.DatabaseId());
		//}

		//public static BigInteger GetRuntimeId(this IITObject item)
		//{
		//	int source, playlist, track, database;
		//	item.GetITObjectIDs(out source, out playlist, out track, out database);
		//	return Encode(source, playlist, track, database);
		//}

		private static long Encode(int high, int low)
		{
			var uHigh = (ulong)high;
			var uLow = (uint)low;
			var combination = (uHigh << 32) | uLow;
			var result = (long)combination;
			Debug.Assert(result.Hi() == high);
			Debug.Assert(result.Lo() == low);
			return result;
		}

		private static int Hi(this long encoded)
		{
			return unchecked((int)(encoded >> 32));
		}

		private static int Lo(this long encoded)
		{
			return unchecked((int)(encoded & 0xFFFFFFFF));
		}

		public static KeyValuePair<byte, string>[] ToMusicBeeFileProperties(this IITFileOrCDTrack track)
		{
			return new KeyValuePair<byte, string>[]
			{
				new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.Url, track.Location),
				new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Artist, track.Artist),
				new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.TrackTitle, track.Name),
				new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Album, track.Album),
				new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Genre, track.Genre),
				new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Comment, track.Comment),
				new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.AlbumArtistRaw, track.AlbumArtist),
				new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.RatingAlbum, track.MusicBeeAlbumRating().ToString()),
				new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Rating, track.MusicBeeRating().ToString()),
				new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Year, track.Year.ToString()),
				new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.Bitrate, track.BitRate.ToString()),
				new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.Size, track.Size.ToString()),
				new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.Duration, track.MusicBeeDuration().ToString()),
				new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.PlayCount, track.PlayedCount.ToString()),
				new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.SkipCount, track.SkippedCount.ToString()),
				new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Composer, track.Composer),
				new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.DiscCount, track.DiscCount.ToString()),
				new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.DiscNo, track.DiscNumber.ToString()),
				new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Grouping, track.Grouping),
				new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.TrackCount, track.TrackCount.ToString()),
				new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.TrackNo, track.TrackNumber.ToString()),
				new KeyValuePair<byte, string>((byte)Plugin.MetaDataType.Artwork, track.Artwork.Count == 0 ? "" : "Y"),
				new KeyValuePair<byte, string>((byte)Plugin.FilePropertyType.LastPlayed, "" + track.MusicBeeLastPlayed()),
			};
		}

		public static DateTime? MusicBeeLastPlayed(this IITTrack track)
		{
			if (track.PlayedDate <= MinPlayedDate) return null;
			return track.PlayedDate.AddSeconds(-track.PlayedDate.Second).ToUniversalTime();
		}

		private static readonly DateTime MinPlayedDate = new DateTime(1899, 12, 30);

		public static DateTime MusicBeeToITunes(this DateTime? date)
		{
			if (date == null) return MinPlayedDate;
			return date.Value.AddSeconds(59).ToLocalTime();
		}

		public static int MusicBeeDuration(this IITTrack track)
		{
			return track.Duration * 1000;
		}

		public static int MusicBeeRating(this IITTrack track)
		{
			return track.Rating / 20;
		}

		public static void SetMusicBeeRating(this IITTrack track, int value)
		{
			track.Rating = value * 20;
		}

		public static int MusicBeeAlbumRating(this IITFileOrCDTrack track)
		{
			return track.AlbumRating / 20;
		}

		public static void SetMusicBeeAlbumRating(this IITFileOrCDTrack track, int value)
		{
			track.AlbumRating = value * 20;
		}

		public static void RepeatTrackOperationUntilNoConflicts(this IITTrack track, Action<IITTrack> action)
		{
			while (true)
			{
				try
				{
					action(track);
					break;
				}
				catch (Exception x)
				{
					Trace.WriteLine(x);
				}
			}
		}

		public static IITOperationStatus Await(this IITOperationStatus operation)
		{
			while (operation.InProgress)
			{
				Thread.Sleep(100);
			}
			return operation;
		}
	}
}
