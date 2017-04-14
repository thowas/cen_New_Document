using System;
using System.Windows.Forms;
using System.IO;
//######################################
//Notwendige Verweise für LateBinding ----
using System.Runtime.InteropServices;
//COM-Bibliotheken von V5 -------------------
using INFITF;
using ProductStructureTypeLib;
using DRAFTINGITF;

namespace cen_3DEXPERIENCE_New_Document
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

                my_Static.logger = new FileLogger(LogLevels.Error, my_Static.User_Tmp_Path + my_Static.KundenName + "\\.log\\Error.log");
                my_Static.logger.Log(LogLevels.Error, Environment.NewLine + "Several CATIA Instances Run.");
                my_Static.logger.EndLogging();



                if (result == DialogResult.OK)
                {
                    Environment.Exit(0);
                }
            }
           
            if (!(Directory.Exists(my_Static.User_Tmp_Path + my_Static.KundenName + "\\.log") == true))
            {
                Directory.CreateDirectory(my_Static.User_Tmp_Path + my_Static.KundenName + "\\.log");
                DirectoryInfo attributes = new DirectoryInfo(my_Static.User_Tmp_Path + my_Static.KundenName + "\\.log");
                attributes.Attributes = FileAttributes.Hidden;
            }


            //Allgemein Variablen ----------------------
            INFITF.Application catiaapp;
            MECMOD.PartDocument activedocpart;
            DrawingDocument activedocdrawing;
            ProductDocument activedocproduct;
            Product prod;

            try
            {
                Object CATIA = Marshal.GetActiveObject("CATIA.Application");
                catiaapp = (INFITF.Application)CATIA;

                my_Static.catia_Release = catiaapp.SystemConfiguration.Release;
                my_Static.catia_Version = catiaapp.SystemConfiguration.Version;
                my_Static.catia_SP = catiaapp.SystemConfiguration.ServicePack;

                ILogger logger_1 = new FileLogger(LogLevels.Run, my_Static.User_Tmp_Path + my_Static.KundenName + "\\.log\\Run.log");
                logger_1.Log(LogLevels.Run, "cen_New_Document_3D_Exp_2015x " + my_Static.Assembly_Version + " gestartet in  V" + my_Static.catia_Version + " R" + my_Static.catia_Release + " SP" + my_Static.catia_SP + Environment.NewLine);
                logger_1.EndLogging();

                
                activedocproduct = catiaapp.ActiveDocument as ProductDocument;

                if (activedocproduct == null)
                {

                    activedocpart = catiaapp.ActiveDocument as MECMOD.PartDocument;

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
                //my_Static.logger = new FileLogger(LogLevels.Error, my_Static.User_Tmp_Path + my_Static.KundenName + "\\.log\\Error.log");
                //my_Static.logger.Log(LogLevels.Error, Environment.NewLine + "  kein Dokument");
                //my_Static.logger.EndLogging();

            }

            System.Windows.Forms.Application.Run(new New());
        }
    }
}
