using iTunesLib;
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

		public static void DeleteAllTracks(this IITUserPlaylist playlist)
		{
			while (playlist.Tracks.Count > 0)
			{
				var track = playlist.Tracks[1];
				track.Delete();
				Marshal.ReleaseComObject(track);
			}
		}

		public static void AddTrackToPlaylistByPersistentId(this iTunesApp iTunes, IITUserPlaylist playlist, long id)
		{
			var track = iTunes.GetTrackByPersistentId(id);
			playlist.AddTrack(track);
			Marshal.ReleaseComObject(track);
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
