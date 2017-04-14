using System;
using System.IO;
using System.Reflection;
using System.Diagnostics;
//######################################
//Notwendige Verweise für LateBinding ----
//COM-Bibliotheken von V5 -------------------
using System.Runtime.InteropServices;
using INFITF;
using DRAFTINGITF;
using MECMOD;
using ProductStructureTypeLib;
using KnowledgewareTypeLib;



namespace cen_3DEXPERIENCE_New_Document
{

    //==============================================================================
    //
    //        Filename: my_Staics.cs
    //
    //        Created by: CENIT AG (Thomas Wassel)
    //              Version: CATIA V5-6 2016x 
    //              Date: 11-04-2017  (Format: mm-dd-yyyy)
    //              Time: 08:30 (Format: hh-mm)
    //
    //==============================================================================
    public class my_Static
    {
        public static ILogger logger;
        public static string[] oParamArry;
        public static PartDocument oPartDoc;
        public static ProductDocument oProdDoc;
        public static DrawingDocument oDrwDoc;
        public static Product prod;
        public static CATBaseDispatch product;
        public static string DocType = "Nothing";
        public static bool bTreeParam = false;
        public static bool bPropParam = false;
        public static Parameters oParam;
        public static string lanuange = "EN";
        public static string Bulid_Daten = "13.04.2017";
        public static string Titel = Assembly.GetExecutingAssembly().GetName().Name;
        public static string Assembly_Version = Assembly.GetExecutingAssembly().GetName().Version.ToString();
        public static string KundenName = Properties.Settings.Default.Kunden_Name;
        public static string User_Tmp_Path = Path.GetTempPath();
        public static int catia_Release;
        public static int catia_Version;
        public static int catia_SP;

        public static bool CheckIfAProcessIsRunning(string processname)
        {
            return Process.GetProcessesByName(processname).Length > 1;
        }

        public static bool AlreadyRunning()
        {
            try
            {
                Process current = Process.GetCurrentProcess();
                Process[] processes = Process.GetProcessesByName(current.ProcessName);
                foreach (Process process in processes)
                {
                    if (process.Id != current.Id)
                    {
                        if (Assembly.GetExecutingAssembly().Location
                            .Replace("/", "\\") == current.MainModule.FileName)
                        {
                            return true;
                        }
                    }
                }
                return false;
            }
            catch
            {

                return false;
            }
        }

        public static void checkCatiaType ()
        {

            //Allgemein Variablen ----------------------
            Application catiaapp;
            PartDocument activedocpart;
            DrawingDocument activedocdrawing;
            ProductDocument activedocproduct;
            Product prod;

            try
            {
                Object CATIA = Marshal.GetActiveObject("CATIA.Application");
                catiaapp = (Application)CATIA;

                    activedocproduct = catiaapp.ActiveDocument as ProductDocument;
                    
                    if (activedocproduct == null)
                    {

                        activedocpart = catiaapp.ActiveDocument as PartDocument;

                        if (activedocpart == null)
                        {

                            activedocdrawing = catiaapp.ActiveDocument as DrawingDocument;

                            if (activedocdrawing == null)
                            {
                                my_Static.DocType = "Nothing";
                                my_Static.logger = new FileLogger(LogLevels.Error, my_Static.User_Tmp_Path + my_Static.KundenName + "\\.log\\Error.log");
                                my_Static.logger.Log(LogLevels.Error, Environment.NewLine + "  kein CATPart ; CATProduct ; CATDrawing geladen !");
                                my_Static.logger.EndLogging();
                            }
                            else
                            {
                                my_Static.DocType = "Drawing";

                            }
                        }
                        else
                        {
                            my_Static.DocType = "Part";
                            prod = activedocpart.Product;
                        }
                    }
                    else
                    {
                        my_Static.DocType = "Product";
                        prod = activedocproduct.Product;
                    }
                }
            catch
            {
                my_Static.DocType = "Nothing";
                //my_Static.logger = new FileLogger(LogLevels.Error, my_Static.User_Tmp_Path + my_Static.KundenName + "\\.log\\Error.log");
                //my_Static.logger.Log(LogLevels.Error, Environment.NewLine + "  kein Dokument geladen !");
                //my_Static.logger.EndLogging();

            }
       
        }


        public static string Languange()
        {
            Application catiaapp;
            Object CATIA = Marshal.GetActiveObject("CATIA.Application");
            catiaapp = (Application)CATIA;
            Documents Doc;
            Document Partdoc;

            try
            {
                string strStatusBar = catiaapp.get_StatusBar();
                if (strStatusBar.Contains("Select") == true)
                {
                    return lanuange = "EN";
                }
                if (strStatusBar.Contains("Ein") == true)
                {
                    return lanuange = "D";
                }
                else
                {
                    Doc = catiaapp.Documents;
                    Partdoc = Doc.Add("Part");
                    strStatusBar = catiaapp.get_StatusBar();

                    if (strStatusBar.Contains("Select") == true)
                    {
                        Partdoc.Close();
                        return lanuange = "EN";
                    }
                    else
                    {
                        Partdoc.Close();
                        return lanuange = "D";

                    }
                }
            }
            catch 
            {
                return lanuange = "EN";
            }

        }


    }
}
