using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;


namespace cen_3DEXPERIENCE_New_Document
{
    //==============================================================================
    //
    //        Filename: FileHaendler.cs
    //
    //        Created by: CENIT AG (Thomas Wassel)
    //              Version: CATIA V5-6 2016x 
    //              Date: 11-04-2017  (Format: mm-dd-yyyy)
    //              Time: 08:30 (Format: hh-mm)
    //
    //==============================================================================
    public class FileHandler
	{
		// Damit im Konstruktor der universelle Datentyp "ILogger"
		// entgegengenommen werden kann, muss auch die
		// Instanzvariable "logger" diesem Datentyp angepasst werden.
		private FileLogger logger;
		//private ILogger logger;
		private string indent = "   ";
		// Instanzvariablen sind immer initialisiert
		private int newDirectories, newFiles, filesRefreshed, filesNotRefreshed;

		// Property "Cancel" zur Prüfung, ob Backup abgebrochen werden soll
		private bool cancel = false;

		public bool Cancel
		{
			set { this.cancel = value; }
		}

		// Damit den universellen Datentyp "ILogger"
		// akzeptiert muss er im Konstruktor auch angegeben werden.
		// public FileHandler(FileLogger logger)
		public FileHandler(FileLogger logger)
		{
			//logger = new FileLogger(LogLevels.Debug, @"..\..\Test.log");
			this.logger = logger;
		}

		public string FileInfos(string filePath)
		{
			FileInfo fileInfo = new FileInfo(filePath);
			StringBuilder sb = new StringBuilder(); // aus: System.Text

			sb.Append("Dateipfad: ");
			// DirectoryName zeigt den Pfad der uebergebenen Datei
			sb.Append(fileInfo.DirectoryName);
			sb.Append(Environment.NewLine);
			sb.Append("Dateiname: ");
			sb.AppendLine(fileInfo.Name);
			sb.Append("Existiert? ");
			bool exists = fileInfo.Exists;
			sb.Append(exists);
			sb.AppendLine();
			if (exists)
			{
				sb.Append("erzeugt: ");
				sb.Append(fileInfo.CreationTime);
				sb.AppendLine();
				sb.Append("geändert: ");
				sb.Append(fileInfo.LastWriteTime);
				sb.AppendLine();
				sb.Append("Zugriff: ");
				sb.Append(fileInfo.LastAccessTime);
				sb.AppendLine();
			}
			return sb.ToString();
		}

		// Kopie einer Einzeldatei nach der Fallunterscheidung
		public void Copy(string sourcePath, string destinationPath)
		{
			this.Copy(sourcePath, destinationPath, "");
		}

		private void Copy(string sourcePath, string destinationPath, string indent)
		{
			// TODO was ist, wenn sourcePath nicht existiert?
			if (sourcePath == null || !File.Exists(sourcePath))
			{   // In der Klasse "File" sind alle Methoden static.
				// Folglich werden sie alle mit Klassennamen aufgerufen.
				// Console.WriteLine("Datei \"{0}\" existiert nicht!", sourcePath);
				logger.Log(LogLevels.Debug, "{0}Datei \"{1}\" existiert nicht!", indent, sourcePath);
				return;
			}
			if (!File.Exists(destinationPath))
			{   // "File.Copy" kopiert nur, wenn Datei noch nicht vorhanden ist.
				// Eine vorhandene Datei wird nicht ueberschrieben.
				File.Copy(sourcePath, destinationPath);
				newFiles++;
				logger.Log(LogLevels.Debug, "{0}Datei \"{1}\" kopiert",
					indent, new FileInfo(sourcePath).Name);
				// Console.WriteLine("Datei kopiert");
			}
			else if (File.GetLastWriteTime(sourcePath) >
			  File.GetLastWriteTime(destinationPath))
			{
				// Diese Ueberladung ueberschreibt vorhandenen Dateien, wenn
				// der dritte Parameter true ist.
				File.Copy(sourcePath, destinationPath, true);
				filesRefreshed++;
				logger.Log(LogLevels.Debug, "{0}Datei \"{1}\" aktualisiert",
					indent, new FileInfo(sourcePath).Name);
				// Console.WriteLine("Datei aktualisiert");
			}
			else
			{
				filesNotRefreshed++;
				logger.Log(LogLevels.Debug, "{0}Datei \"{1}\" unverändert (keine Aktion)",
					indent, new FileInfo(sourcePath).Name);
				// Console.WriteLine("Datei unverändert (keine Aktion)");
			}
		}
		public bool CopyFileSystem(string sourceDirPath, string destinationDirPath)
		{
			return this.CopyFileSystem(sourceDirPath, destinationDirPath, "");
		}

		private bool CopyFileSystem(string sourceDirPath, string destinationDirPath, string indent)
		{
			if (!Directory.Exists(destinationDirPath))
			{
				DirectoryInfo dirInfo = Directory.CreateDirectory(destinationDirPath);
				newDirectories++;
				logger.Log(LogLevels.Debug, "{0}Ordner \"{1}\" erstellt", indent, dirInfo.Name);
			}
			else
			{
				logger.Log(LogLevels.Debug, "{0}Ordner \"{1}\" existiert", indent,
					new DirectoryInfo(destinationDirPath).Name);
			}
			string[] subdirs = Directory.GetDirectories(sourceDirPath);
			string[] files = Directory.GetFiles(sourceDirPath);
			// Einrücktiefe erhöhen
			indent += this.indent;
			// "file" enthält den kompletten Pfad mit Dateinamen
			foreach (string file in files)
			{
				if (cancel)
					return false;
				this.Copy(file, file.Replace(sourceDirPath, destinationDirPath), indent);
			}
			// Einstieg in die Rekursion - Aufruf d i e s e r Methode für jede subdir
			foreach (string dir in subdirs)
			{
				if (cancel)
					return false;
				this.CopyFileSystem(dir, dir.Replace(sourceDirPath, destinationDirPath), indent);
			}
			return true;
		}

		public void Statistics()
		{
			// Zähler auswerten
			logger.Log(LogLevels.Run,
				"{0} Ordner angelegt, {1} Dateien kopiert, {2} Dateien akualisiert, {3} Dateien unverändert",
				newDirectories, newFiles, filesRefreshed, filesNotRefreshed);
			// Zähler zurücksetzen
			newDirectories = 0;
			newFiles = 0;
			filesRefreshed = 0;
			filesNotRefreshed = 0;
		}
	}
}
