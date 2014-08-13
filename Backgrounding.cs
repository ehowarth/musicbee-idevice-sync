using System;
using System.Diagnostics;
using System.Threading;

namespace MusicBeeDeviceSyncPlugin
{
	static class Backgrounding
	{
		public static void RunInBackground(this Action action)
		{
			new Thread(() => action.RunAndLogAnyException()).Start();
		}

		private static void RunAndLogAnyException(this Action action)
		{
			try
			{
				Thread.Sleep(10);
				action();
			}
			catch (Exception x)
			{
				Trace.WriteLine(x);
				Trace.Flush();
			}
		}
	}
}
