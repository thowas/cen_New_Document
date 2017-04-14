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
    //        Filename: FileLogger.cs
    //
    //        Created by: CENIT AG (Thomas Wassel)
    //              Version: CATIA V5-6 2016x 
    //              Date: 11-04-2017  (Format: mm-dd-yyyy)
    //              Time: 08:30 (Format: hh-mm)
    //
    //==============================================================================

    // "enum" ist eine eigenstaendige Typeklasse und ist
    // auf der gleichen Programm-Logischen Ebene wie "class".
    // Folglich ist es nicht innerhalb des Programmblocks
    // von "class", sondern ausserhalb.
    // Damit ist es auch von anderen Klassen benutzbar.
    /*
	 * Ist jetzt im Interface
	public enum LogLevels { Debug, Run, Error }
	*/

    // Wenn eine Klasse eine Interface implementiert, sollte die
    // Klasse nur noch die Methoden des Interfaces enthalten
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1001:TypesThatOwnDisposableFieldsShouldBeDisposable")]
	public  class FileLogger : ILogger
	{
		private  string logFile;
		private  StreamWriter writer;
		private  string user;
		private string maschine;
        private string sprache;

		private LogLevels logLevel = LogLevels.Debug;
		// Achtung: Dies ist ein Property:
		// 1. Zugriffsmodifier
		//     2. Datentyp
		//               3. Name ==> der gleich Name, wie die Variabel,
		//                  jedoch mit Grossbuchstaben.
		// "Property" - Zugriffsstruktur auf eine private
		// (gekapselte) Variabel
		public LogLevels LogLevel
		{   // "set"- bzw. "get"-Anweisungen in geschweifte Klammern.
			// Anweisungen in den Klammern mit Strich-Punkt abschliessen.
			set { logLevel = value; }
			get { return logLevel; }
		}

		public FileLogger(LogLevels logLevel, string logFile)
		{
			this.logFile = logFile;
			// Null-Prüfung
			// verlagert in die Log-Methode, abhängig von writer == null
			//writer = new StreamWriter(logFile, true); // File.Open(logFile, FileMode.Append)
			this.logLevel = logLevel;
		}

		public  void Log(LogLevels logLevel, string message)
		{
            if (Properties.Settings.Default.LogFile == true)
            {
                if (writer == null)
                {
                    writer = new StreamWriter(logFile, true);
                }
                if (logLevel >= this.logLevel)
                {
                    //writer.WriteLine(message);
                    user = Environment.UserName;
                    maschine = Environment.MachineName;
                    sprache = my_Static.lanuange;
                    writer.Write("\r\nLog Entry : ");
                    writer.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString() + "  " + user + "  " + maschine);
                    writer.WriteLine("  :{0}", "cen_New_Document_3D_Exp_2015x Version: " + my_Static.Assembly_Version + " gestartet in  V" + my_Static.catia_Version + " R" + my_Static.catia_Release + " SP" + my_Static.catia_SP);
                    writer.WriteLine("  :{0}", "CATIA Sprache: " + sprache);
                    writer.WriteLine("  :{0}", message);
                    writer.WriteLine("-------------------------------");
                }
            }
		}
		public void Log(LogLevels logLevel, string message, params object[] args)
		{   // Hier wird die Klassen-Eigenen ueberladenen Methode aufgerufen.
			// Die Klasse "String" kennt die Methode "Format"
			// Diese Methode ersetzt die Platzhalter im String "message"
			// mit den uebergebene Argumenten in "args"
			this.Log(logLevel, String.Format(message, args));
		}

		public void EndLogging()
		{
            if (Properties.Settings.Default.LogFile == true)
            {

                // "StreamWriter"-Objekte können nach dem Schließen nicht mehr
                // geöffnet werden. Sie müssen immer wieder neu erzeugt werden.
                // Folglich wird nach dem "Close" die Variable genullt.
                // Somit kann das Objekt entfernt und bei Bedarf neu erzeugt
                // werden.
                writer.Close();

                FileInfo txtfile = new FileInfo(logFile);
                if (txtfile.Length > (1 * 1024 * 1024))       // ## NOTE: 1MB max file size
                {
                    var lines = File.ReadAllLines(logFile).Skip(20).ToArray();  // ## Set to 20 lines
                    File.WriteAllLines(logFile, lines);
                }

                writer = null;
            }
		}

		public string Logging
		{

			get
			{

				// Zuerst muss der aktive Schreibvorgang geschlossen
				// werden
				writer.Close();
				// Zur spaeteren Abbruefung muss Variabel genullt werden
				writer = null;
				// Reader-Objekt erzeugen, damit das Log-File ausgelesen
				// werden kann.
				StreamReader reader = new StreamReader(logFile);
				// Mit der Methode "ReadToEnd" den gesamten Inhalt der
				// Log-Datei auslesen
				string logging = reader.ReadToEnd();
				// Und das Reader-Objekt muss auch wieder geschlossen
				// werden.
				reader.Close();
				return logging;
			}
		}
	}
}
