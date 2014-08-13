using MusicBeePlugin;
using System.Drawing;
using System.Reflection;
using System.Resources;

namespace MusicBeeDeviceSyncPlugin
{
	static class Info
	{
		public static readonly Assembly Assembly = Assembly.GetExecutingAssembly();
		public static readonly AssemblyName AssemblyName = Assembly.GetName();
		//public static readonly ResourceManager Resources = new ResourceManager("MusicBeePlugin.Images", Assembly);

		public static readonly Plugin.PluginInfo Current = new Plugin.PluginInfo
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
	}
}
