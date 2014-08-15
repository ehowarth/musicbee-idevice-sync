using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace MusicBeeDeviceSyncPlugin
{
	static class Text
	{
		public static string L(string developerText, params object[] parameters)
		{
			var format = developerText;

#if DEBUG
			format = FakeText(format);
			Trace.WriteLineIf(
				!Translations.ContainsKey(developerText),
				"Missing translation entry for: \"" + developerText + "\""
				);
#else
			if (LanguageIndex >= 0)
			{
				string[] localizations;
				if (Translations.TryGetValue(developerText, out localizations))
					if (localizations[LanguageIndex] != null)
						format = localizations[LanguageIndex];
			}
#endif

			return string.Format(format, parameters);
		}

		private static string[] l10n(
			string ru = null
			// other languages go here as additional parameters
			)
		{
			return new string[]
			{
				ru,
				// Other languages go here ordered same as LanguageIndex
			};
		}

		private static readonly string Language = Thread.CurrentThread.CurrentCulture.TwoLetterISOLanguageName;

		private static readonly int LanguageIndex = Language == "ru"
			// Add new languages as new array indexes
			? 0
			: -1;

		private static readonly IDictionary<string, string[]> Translations = new Dictionary<string, string[]>
		{
			{ "iPod && iPhone Sync", l10n(
				ru: "Драйвер iPod & iPhone")},
			{ "Takes control of the iTunes application to synchronize an iPod, iPhone, or iPad", l10n(
				ru: "Плагин позволяет использовать iPod/iPhone посредством iTunes")},
			{ "Open iPod & iPhone Sync", l10n(
				ru: "Включить драйвер iPod && iPhone")},
			{ "Opens iTunes to begin sync", l10n(
				ru: "Драйвер iPod & iPhone: Включить/выключить")},
			{ "Starting iTunes...", l10n(
				)},
			{ "Failed to add track to iTunes library: \"{0}\"", l10n(
				ru:"Не удалось добавить композицию в медиатеку iTunes: \"{0}\"")},
			{ "No appropriate category found in iTunes library",l10n(
				ru:"Не найдено подходящей категории в медиатеке iTunes")},
			{ "Delete playlist \"{0}\" from device?",l10n(
				ru: "Удалить плейлист \"{0}\" из устройства?")},
			{ "Delete track \"{0}\" from device?",l10n(
				ru: "Удалить композицию \"{0}\" из устройства?")},
			{ "Some duplicated tracks were skipped.", l10n(
				ru: "Некоторые повторяющиеся композиции были пропущены.")},
			{"Track source file not found: \"{0}\"",l10n(
				ru: "Исходный файл композиции не найден: \"{0}\"")},
			{ "Removing track: \"{0}\"", l10n(
				ru: "Удаляется композиция: \"{0}\"")},
			{ "Adding track to playlist {0}/{1}", l10n(
				ru: "В плейлист добавляется композиция {0}/{1}")},
			{ "Syncing iTunes with iPod/iPhone...", l10n(
				ru: "Синхронизация iTunes с iPod/iPhone...")},
			{ "Syncing play counts and ratings from iTunes...", l10n(
				ru: "Синхронизация счетчиков воспроизведения и рейтингов из iPod/iPhone")},
			{ "Syncing play counts and ratings from iTunes: \"{0}\"", l10n(
				ru: "Синхронизация счетчиков воспроизведения и рейтингов из iPod/iPhone: \"{0}\"")},
			{ "Filling synced playlists...", l10n(
				ru: "Заполнение синхронизированных плейлистов...")},
		};

#if DEBUG

		/// <summary>
		/// Converts an English phrase to Unicode gobbledygook that looks like the original phrase and has the same length.
		/// </summary>
		/// <param name="original"></param>
		/// <returns></returns>
		static string MungeText(string original)
		{
			return original.Replace(
								 'A', 'Å').Replace(
								 'B', 'ß').Replace(
								 'C', 'C').Replace(
								 'D', 'Đ').Replace(
								 'E', 'Ē').Replace(
								 'F', 'F').Replace(
								 'G', 'Ğ').Replace(
								 'H', 'Ħ').Replace(
								 'I', 'Ĩ').Replace(
								 'J', 'Ĵ').Replace(
								 'K', 'Ķ').Replace(
								 'L', 'Ŀ').Replace(
								 'M', 'M').Replace(
								 'N', 'Ń').Replace(
								 'O', 'Ø').Replace(
								 'P', 'P').Replace(
								 'Q', 'Q').Replace(
								 'R', 'Ŗ').Replace(
								 'S', 'Ŝ').Replace(
								 'T', 'Ŧ').Replace(
								 'U', 'Ů').Replace(
								 'V', 'V').Replace(
								 'W', 'Ŵ').Replace(
								 'X', 'X').Replace(
								 'Y', 'Ÿ').Replace(
								 'Z', 'Ż').Replace(
								 'a', 'ä').Replace(
								 'b', 'þ').Replace(
								 'c', 'č').Replace(
								 'd', 'đ').Replace(
								 'e', 'ę').Replace(
								 'f', 'ƒ').Replace(
								 'g', 'ģ').Replace(
								 'h', 'ĥ').Replace(
								 'i', 'į').Replace(
								 'j', 'ĵ').Replace(
								 'k', 'ĸ').Replace(
								 'l', 'ľ').Replace(
								 'm', 'm').Replace(
								 'n', 'ŉ').Replace(
								 'o', 'ő').Replace(
								 'p', 'p').Replace(
								 'q', 'q').Replace(
								 'r', 'ř').Replace(
								 's', 'ş').Replace(
								 't', 'ŧ').Replace(
								 'u', 'ū').Replace(
								 'v', 'v').Replace(
								 'w', 'ŵ').Replace(
								 'x', 'χ').Replace(
								 'y', 'y').Replace(
								 'z', 'ž');
		}

		private static readonly Regex[] Preservations =
		{
			new Regex(@"(#\w+;)"),
			new Regex(@"(&\w+;)"),
			new Regex(@"(\{\d+(:[^\}]+)?\})"),
		};

		/// <summary>
		/// Calls MungeText -- but preserves escape sequences such as &amp; #LineBreak; and {0:f} -- and appends extra text.
		/// </summary>
		/// <param name="original"></param>
		/// <returns></returns>
		static string FakeText(string original)
		{
			var update = MungeText(original);

			foreach (var regex in Preservations)
			{
				foreach (Match match in regex.Matches(original))
				{
					update = update.Remove(match.Index, match.Length).Insert(match.Index, match.Value);
				}
			}

			return update + string.Concat(Enumerable.Repeat(" !!", update.Length / 9));
		}

#endif
	}
}
