using System;
using System.Windows.Forms;

//######################################
//Notwendige Verweise für LateBinding ----
using System.Runtime.InteropServices;
//COM-Bibliotheken von V5 -------------------
using INFITF;
using ProductStructureTypeLib;

namespace cen_SPIN_New_Document
{

    //==============================================================================
    //
    //        Filename: Program.cs
    //
    //        Created by: CENIT AG (Thomas Wassel)
    //              Version: CATIA V5-6 2016x 
    //              Date: 11-04-2017  (Format: mm-dd-yyyy)
    //              Time: 08:30 (Format: hh-mm)
    //
    //==============================================================================
    static class Program
    {
        /// <summary>
        /// Der Haupteinstiegspunkt für die Anwendung.
        /// </summary>
        [STAThread]
        static void Main()
        {
            DialogResult result;
            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);

            my_Static.Languange();

            if (my_Static.AlreadyRunning())
            {
                System.Windows.Forms.Application.Exit();
                return;
            }


            bool catia_running = my_Static.CheckIfAProcessIsRunning("CNEXT");
            if (catia_running == true)
            {
                string message = "CATIA Several Instances Run !...........CENIT AG";
                result = MessageBox.Show("Several CATIA Instances Run." + Environment.NewLine + Environment.NewLine + "Please Only One CATIA Start !", message, MessageBoxButtons.OK, MessageBoxIcon.Information);


                if (result == DialogResult.OK)
                {
                    Environment.Exit(0);
                }
            }

            //Allgemein Variablen ----------------------
            INFITF.Application catiaapp;
            MECMOD.PartDocument activedocpart;
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

                    activedocpart = catiaapp.ActiveDocument as MECMOD.PartDocument;

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
                        
            System.Windows.Forms.Application.Run(new New());
        }
    }
}
