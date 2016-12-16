using iTunesLib;
using MusicBeePlugin;
using System;
using System.Collections.Generic;

namespace MusicBeeITunesSyncPlugin
{
	static class FilePropertyAndTagConversions
	{
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

		public static void SyncITunesHistoryToMusicBee(this IITFileOrCDTrack track, MusicBeeFile file)
		{
			var mbLastPlayed = file.LastPlayed;
			var itLastPlayed = track.PlayedDate.AddSeconds(-track.PlayedDate.Second);

			if (mbLastPlayed == null || mbLastPlayed < itLastPlayed)
			{
				file.PlayCount = track.PlayedCount;
				file.LastPlayed = track.PlayedDate.ToUniversalTime();
				file.SkipCount = track.SkippedCount;
				file.Rating = track.MusicBeeRating();
				file.RatingAlbum = track.MusicBeeAlbumRating();
				file.CommitChanges();
				Plugin.MbApiInterface.MB_RefreshPanels();
			}
		}

		public static void SyncMusicBeeHistoryToITunes(this IITTrack itTrack, MusicBeeFile mbFile)
		{
			itTrack.RepeatTrackOperationUntilNoConflicts(t => t.PlayedDate = mbFile.LastPlayed.MusicBeeToITunes());
			itTrack.RepeatTrackOperationUntilNoConflicts(t => t.PlayedCount = mbFile.PlayCount);
			itTrack.RepeatTrackOperationUntilNoConflicts(t => t.SetMusicBeeRating(mbFile.Rating));
			if (itTrack.Kind == ITTrackKind.ITTrackKindFile)
			{
				itTrack.RepeatTrackOperationUntilNoConflicts(t => ((IITFileOrCDTrack)t).SkippedCount = mbFile.SkipCount);
				itTrack.RepeatTrackOperationUntilNoConflicts(t => ((IITFileOrCDTrack)t).SetMusicBeeAlbumRating(mbFile.RatingAlbum));
			}
		}

		static DateTime? MusicBeeLastPlayed(this IITTrack track)
		{
			if (track.PlayedDate <= MinPlayedDate) return null;
			return track.PlayedDate.AddSeconds(-track.PlayedDate.Second).ToUniversalTime();
		}

		private static readonly DateTime MinPlayedDate = new DateTime(1899, 12, 30);

		static DateTime MusicBeeToITunes(this DateTime? date)
		{
			if (date == null) return MinPlayedDate;
			return date.Value.ToLocalTime();
		}

		static int MusicBeeDuration(this IITTrack track)
		{
			return track.Duration * 1000;
		}

		static int MusicBeeRating(this IITTrack track)
		{
			return track.Rating / 20;
		}

		static void SetMusicBeeRating(this IITTrack track, int value)
		{
			track.Rating = value * 20;
		}

		static int MusicBeeAlbumRating(this IITFileOrCDTrack track)
		{
			return track.AlbumRating / 20;
		}

		static void SetMusicBeeAlbumRating(this IITFileOrCDTrack track, int value)
		{
			track.AlbumRating = value * 20;
		}
	}
}
