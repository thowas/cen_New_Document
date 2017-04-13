using System;
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


namespace cen_SPIN_New_Document
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
            INFITF.Application catiaapp;
            PartDocument activedocpart;
            Document activedocdrawing;
            ProductDocument activedocproduct;
            Product prod;

            try
            {
                Object CATIA = Marshal.GetActiveObject("CATIA.Application");
                catiaapp = (INFITF.Application)CATIA;



                activedocproduct = catiaapp.ActiveDocument as ProductDocument;

                if (activedocproduct == null)
                {

                    activedocpart = catiaapp.ActiveDocument as PartDocument;

                    if (activedocpart == null)
                    {

                        activedocdrawing = catiaapp.ActiveDocument as Document;

                        if (activedocdrawing == null)
                        {
                            my_Static.DocType = "Nothing";
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

               
            }
       
        }


        public static string Languange()
        {
            INFITF.Application catiaapp;
            Object CATIA = System.Runtime.InteropServices.Marshal.GetActiveObject("CATIA.Application");
            catiaapp = (INFITF.Application)CATIA;
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
            catch (Exception ex)
            {
                //logger = new FileLogger(LogLevels.Error, myStatic.Catia_User_Tmp_Path + myStatic.KundenName + "\\.log\\Error.log");
                //logger.Log(LogLevels.Error, ex.ToString() + Environment.NewLine + Environment.NewLine + "  Languange()");
                //logger.EndLogging();
                return lanuange = "EN";
            }

        }


    }
}
