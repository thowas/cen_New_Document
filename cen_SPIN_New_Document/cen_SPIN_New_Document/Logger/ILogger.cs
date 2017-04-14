using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace cen_3DEXPERIENCE_New_Document
{
    //==============================================================================
    //
    //        Filename: ILogger.cs
    //
    //        Created by: CENIT AG (Thomas Wassel)
    //              Version: CATIA V5-6 2016x 
    //              Date: 11-04-2017  (Format: mm-dd-yyyy)
    //              Time: 08:30 (Format: hh-mm)
    //
    //==============================================================================

    //[ComImport]
    //[Guid("E8B01F74-96A2-4017-9840-1DEAD6EB5ED0")]
    //public interface ISampleInterface
    //{
    //	void GetUserInput();
    //	string UserInput { get; }
    //}


    // Ist ausserhalb der Klassendefinition
    public enum LogLevels { Debug, Run, Error }
	//[ComImport]
	//[Guid("E8B01F74-96A2-4017-9840-1DEAD6EB5ED0")]
	public interface ILogger
	{

		// Variante für einfache String
		void Log(LogLevels logLevel, string message);

		// Variante für formatierten String
		void Log(LogLevels logLevel, string message, params object[] args);

		void EndLogging();

		// Property
		// Zur Ausgabe eines Logging-Strings
		string Logging { get; }

		// Property für die LogLevel-Variable
		// ;=> LogLevels ist Datentyp
		// |      ;=> Variablenname, aber mit grossen
		//            Anfangsbuchstaben
		LogLevels LogLevel { get; set; }

	}

}
