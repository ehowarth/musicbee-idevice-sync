using iTunesLib;
using System.Threading;

namespace MusicBeePlugin
{
	static class iTunesExtensions
	{
		public static IITPlaylist GetPlaylist(this iTunesApp iTunes, string name)
		{
			foreach (IITPlaylist playlist in iTunes.LibrarySource.Playlists)
			{
				if (
					 playlist.Name == name
					 && playlist.Kind == ITPlaylistKind.ITPlaylistKindUser
					 && ((IITUserPlaylist)playlist).SpecialKind == ITUserPlaylistSpecialKind.ITUserPlaylistSpecialKindNone
					 && !((IITUserPlaylist)playlist).Smart
					 )
				{
					return playlist;
				}
			}
			return null;
		}

		public static void Await(this IITOperationStatus operation)
		{
			while (operation.InProgress)
			{
				Thread.Sleep(100);
			}
		}
	}
}
