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
     

    }
}
