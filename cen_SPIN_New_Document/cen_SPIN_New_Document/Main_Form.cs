using System;
using System.Data;
using System.Windows.Forms;
//Notwendige Verweise für LateBinding ----
using System.Runtime.InteropServices;
//--------------------------------------------------------
//COM-Bibliotheken von V5 -------------------
using INFITF;
using MECMOD;
using ProductStructureTypeLib;
using KnowledgewareTypeLib;
using DRAFTINGITF;
using CD5Integ;
//------------------------------------------------------

namespace cen_SPIN_New_Document
{

    //==============================================================================
    //
    //        Filename: Main_Form.cs
    //
    //        Created by: CENIT AG (Thomas Wassel)
    //              Version: CATIA V5-6 2016x 
    //              Date: 11-04-2017  (Format: mm-dd-yyyy)
    //              Time: 08:30 (Format: hh-mm)
    //
    //==============================================================================

    public partial class New : Form
    {


        //Allgemein Variablen ----------------------
        public TreeNode tn0;
        public INFITF.Application catiaapp;
        public PartDocument activedocpart;
        public DrawingDocument activedocdrawing;
        public ProductDocument activedocproduct;
        public Product prod;
        public CD5EngineV6R2015 oCD5Engine;
        public string strcurrentSelection;
        public string strcurrentType;




        public New()
        {
            InitializeComponent();

        }


        private void button_Cancel_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void New_Load(object sender, EventArgs e)
        {
            Properties.Settings.Default.Save();
            string[] setting = new string[Properties.Settings.Default.Betriebsmittelart.Count];
            Properties.Settings.Default.Betriebsmittelart.CopyTo(setting, 0);

            DataTable dt = new DataTable();
            DataColumn dc1 = new DataColumn("item");

            dt.Columns.Add(dc1);

            foreach (string str in setting)
            {
                dt.Rows.Add(str);
            }

            comboBoxBetriebsmittelart.DataSource = dt;
            comboBoxBetriebsmittelart.DisplayMember = "item";


            Object CATIA = Marshal.GetActiveObject("CATIA.Application");
            catiaapp = (INFITF.Application)CATIA;

            my_Static.checkCatiaType();

            if (my_Static.DocType == "Product")
            {
                button_InsertCATPart.Visible = true;
            }

            CD5_connect();

        }



        public bool NodeExists(TreeNode levelNodes, string key)
        {
            var exists = false;
            foreach (TreeNode node in levelNodes.Nodes)
            {
                if (node.Text == key)
                {
                    exists = true;
                    break;
                }

            }

            return exists;
        }




        //-------------------------------------------------------
        private void getbodies(TreeNode tn0, MECMOD.Part Part1)
        {
            MECMOD.Bodies bodies1 = Part1.Bodies;
            for (int i = 1; i <= bodies1.Count; i++)
            {
                object index = i;
                MECMOD.Body body1 = bodies1.Item(ref index);
                if (body1.InBooleanOperation == false)
                {
                    TreeNode tn_body = new TreeNode(body1.get_Name());
                    tn0.Nodes.Add(tn_body);

                    MECMOD.Shapes shapes1 = body1.Shapes;
                    getshape(tn_body, shapes1);
                }
            }
        }
        //-------------------------------------------------------
        private void getshape(TreeNode tn_body, MECMOD.Shapes shapes1)
        {
            for (int m = 1; m <= shapes1.Count; m++)
            {
                object index = m;
                try
                {
                    PARTITF.BooleanShape shape1 = (PARTITF.BooleanShape)shapes1.Item(ref index);
                    TreeNode tn_body_bool = new TreeNode(shape1.Body.get_Name());
                    tn_body.Nodes.Add(tn_body_bool);


                    MECMOD.Shapes shapes_bool = shape1.Body.Shapes;
                    getshape(tn_body_bool, shapes_bool);
                }
                catch
                {
                    MECMOD.Shape shape1 = shapes1.Item(ref index);
                    tn_body.Nodes.Add(shape1.get_Name());
                }
            }
        }
        //-------------------------------------------------------
        private void gethybridbodies(TreeNode tn0, MECMOD.HybridBodies hbodies1)
        {
            for (int i = 1; i <= hbodies1.Count; i++)
            {
                object index = i;
                MECMOD.HybridBody hbody1 = hbodies1.Item(ref index);

                TreeNode tn_body = new TreeNode(hbody1.get_Name());
                tn0.Nodes.Add(tn_body);

                if (hbody1.HybridBodies.Count > 0)
                    gethybridbodies(tn_body, hbody1.HybridBodies);

                MECMOD.HybridShapes hshapes1 = hbody1.HybridShapes;
                getshybridhape(tn_body, hshapes1);
            }
        }
        //-------------------------------------------------------
        private void getshybridhape(TreeNode tn_body, MECMOD.HybridShapes hshapes1)
        {
            for (int m = 1; m <= hshapes1.Count; m++)
            {
                object index = m;

                MECMOD.HybridShape hshape1 = hshapes1.Item(ref index);
                tn_body.Nodes.Add(hshape1.get_Name());
            }
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }

        //---------------------------------------------------
        // Returns a value indicating whether the specified 
        // TreeNode has checked child nodes.
        public bool HasSelectChildNodes(object sender, TreeViewEventArgs node)
        {
            return true;
        }

        public void set_Catia_Properties()
        {
            INFITF.Application catiaapp;
            PartDocument activedocpart;
            ProductDocument activedocproduct;
            Product prod;
            Parameters pParams;
            Parameter pParam;
            bool bstr = false;
            string strProp = "";

            Object CATIA = Marshal.GetActiveObject("CATIA.Application");
            catiaapp = (INFITF.Application)CATIA;


            if (my_Static.lanuange == "EN")
            {

                strProp = "Properties\\";
            }
            else if (my_Static.lanuange == "D")
            {

                strProp = "Eigenschaften\\";
            }

            if (my_Static.DocType == "Product")
            {
                activedocproduct = catiaapp.ActiveDocument as ProductDocument;
                prod = activedocproduct.Product;
                pParams = prod.Parameters;

                for (int i = 1; i <= pParams.Count; i++)
                {

                    if ((bstr = pParams.Item(i).get_Name().Contains("Equipment-Nummer") == true))
                    {
                        try
                        {
                            pParam = pParams.Item(strProp + "Equipment-Nummer");
                            pParam.ValuateFromString(textBoxEquipmentNumber.Text);
                        }
                        catch
                        {
                            pParam = pParams.Item("Equipment-Nummer");
                            pParam.ValuateFromString(textBoxEquipmentNumber.Text);
                        }
                    }
                    else if ((bstr = pParams.Item(i).get_Name().Contains("Betriebsmittelart") == true))
                    {
                        try
                        {
                            pParam = pParams.Item(strProp + "Betriebsmittelart");
                            pParam.ValuateFromString(comboBoxBetriebsmittelart.Text);
                        }
                        catch
                        {
                            pParam = pParams.Item("Betriebsmittelart");
                            pParam.ValuateFromString(comboBoxBetriebsmittelart.Text);
                        }
                    }
                    else if ((bstr = pParams.Item(i).get_Name().Contains("Bezeichnung") == true))
                    {
                        try
                        {
                            pParam = pParams.Item(strProp + "Bezeichnung");
                            pParam.ValuateFromString(textBox_Bezeichnung2.Text);
                        }
                        catch
                        {
                            pParam = pParams.Item("Bezeichnung");
                            pParam.ValuateFromString(textBox_Bezeichnung2.Text);
                        }
                    }

                }

            }
            if (my_Static.DocType == "Part")
            {
                activedocpart = catiaapp.ActiveDocument as PartDocument;
                prod = activedocpart.Product;
                pParams = prod.Parameters;

                for (int i = 1; i <= pParams.Count; i++)
                {

                    if ((bstr = pParams.Item(i).get_Name().Contains("Equipment-Nummer") == true))
                    {
                        try
                        {
                            pParam = pParams.Item(strProp + "Equipment-Nummer");
                            pParam.ValuateFromString(textBoxEquipmentNumber.Text);
                        }
                        catch
                        {
                            pParam = pParams.Item("Equipment-Nummer");
                            pParam.ValuateFromString(textBoxEquipmentNumber.Text);
                        }
                    }
                    else if ((bstr = pParams.Item(i).get_Name().Contains("Betriebsmittelart") == true))
                    {
                        try
                        {
                            pParam = pParams.Item(strProp + "Betriebsmittelart");
                            pParam.ValuateFromString(comboBoxBetriebsmittelart.Text);
                        }
                        catch
                        {
                            pParam = pParams.Item("Betriebsmittelart");
                            pParam.ValuateFromString(comboBoxBetriebsmittelart.Text);
                        }
                    }
                    else if ((bstr = pParams.Item(i).get_Name().Contains("Bezeichnung") == true))
                    {
                        try
                        {
                            pParam = pParams.Item(strProp + "Bezeichnung");
                            pParam.ValuateFromString(textBox_Bezeichnung2.Text);
                        }
                        catch
                        {
                            pParam = pParams.Item("Bezeichnung");
                            pParam.ValuateFromString(textBox_Bezeichnung2.Text);
                        }
                    }
                }

            }


        }

        public void CD5_insert_CATPart()
        {
            INFITF.Application catiaapp;
            Document DocTempl;
            CD5Template oTemplate;
            CD5Templates oTemplates;
            CD5TemplateTypes oTemplateTypes;
            CD5TemplateType oTemplateType;
            ProductDocument activedocproduct;
            Document activedocpart;
            Product prod;
            Window RootWin;
            Window PartWin;

            Object CATIA = Marshal.GetActiveObject("CATIA.Application");
            catiaapp = (INFITF.Application)CATIA;

            activedocproduct = catiaapp.ActiveDocument as ProductDocument;
            prod = activedocproduct.Product;

            oCD5Engine = catiaapp.GetItem("CD5EngineV6R2015") as CD5EngineV6R2015;

            if (oCD5Engine.IsConnected())
            {
                oTemplate = null;
                oTemplateTypes = oCD5Engine.TemplateTypes;

                for (int iTypes = 1; iTypes <= oTemplateTypes.Count; iTypes++)
                {
                    oTemplateType = null;
                    oTemplateType = oTemplateTypes.Item(iTypes);

                    if (oTemplateType.TemplateTypeName == strcurrentType)
                    {
                        for (int iItem = 1; iItem <= oTemplateType.Templates.Count; iItem++)
                        {
                            oTemplates = oTemplateType.Templates;
                            oTemplate = oTemplates.Item(iItem);
                            if (oTemplate.TemplateName == strcurrentSelection)
                            {
                                int prodCount = prod.Products.Count + 1;
                                RootWin = catiaapp.ActiveWindow;
                                DocTempl = oCD5Engine.NewFrom(oTemplate, prod.get_PartNumber() + "_" + prodCount, "CATPart");
                                PartWin = catiaapp.ActiveWindow;
                                activedocpart = catiaapp.ActiveDocument;
                                object[] strArray = new object[1];
                                strArray[0] = DocTempl.FullName;
                                prod.Products.AddComponentsFromFiles(strArray, "All");
                                activedocpart.Close();
                                RootWin.Activate();

                                break;
                            }
                        }
                    }

                }

            }
            else
            {
                MessageBox.Show("Sie sind nicht an 3DExperience angemeldet.", "Abbruch");
                System.Windows.Forms.Application.Exit();
            }
        }


        public void CD5_openFile()
        {
            INFITF.Application catiaapp;
            Document DocTempl;
            CD5Template oTemplate;
            CD5Templates oTemplates;
            CD5TemplateTypes oTemplateTypes;
            CD5TemplateType oTemplateType;
            string strType = "";
            DialogResult result;

            Object CATIA = Marshal.GetActiveObject("CATIA.Application");
            catiaapp = (INFITF.Application)CATIA;

            oCD5Engine = catiaapp.GetItem("CD5EngineV6R2015") as CD5EngineV6R2015;

            if (oCD5Engine.IsConnected())
            {
                oTemplate = null;
                oTemplateTypes = oCD5Engine.TemplateTypes;

                for (int iTypes = 1; iTypes <= oTemplateTypes.Count; iTypes++)
                {
                    oTemplateType = null;
                    oTemplateType = oTemplateTypes.Item(iTypes);

                    if (oTemplateType.TemplateTypeName == strcurrentType)
                    {

                        switch (strcurrentType)
                        {
                            case "CATProduct Template":
                                strType = "CATProduct";
                                break;

                            case "CATPart Template":
                                strType = "CATPart";
                                break;

                            case "CATDrawing Template":
                                strType = "CATDrawing";
                                break;
                        }

                        for (int iItem = 1; iItem <= oTemplateType.Templates.Count; iItem++)
                        {
                            oTemplates = oTemplateType.Templates;
                            oTemplate = oTemplates.Item(iItem);
                            if (oTemplate.TemplateName == strcurrentSelection)
                            {
                                try
                                {
                                    DocTempl = oCD5Engine.NewFrom(oTemplate, textBox_Name.Text, strType);
                                    break;
                                }
                                catch
                                {

                                    string message = " Confirmation ";
                                    result = MessageBox.Show(" Following file exits in the checkout directory. " + Environment.NewLine + textBox_Name.Text + "." + strType + Environment.NewLine + "New File wird beendet !", message, MessageBoxButtons.OK, MessageBoxIcon.Question);


                                    if (result == DialogResult.OK)
                                    {
                                        Environment.Exit(0);
                                    }

                                }
                            }
                        }
                        break;
                    }
                }
            }
            else
            {
                MessageBox.Show("Sie sind nicht an 3DExperience angemeldet.", "Abbruch");
                System.Windows.Forms.Application.Exit();
            }
        }

        public void CD5_Item_count()
        {
            object[] AutonameSeries = new object[1];
            string strSizeType = "";

            strSizeType = comboBoxAutoN.Text;

            oCD5Engine = catiaapp.GetItem("CD5EngineV6R2015") as CD5EngineV6R2015;

            if (oCD5Engine.IsConnected())
            {
                //AutonameSeries[0] = oCD5Engine.GetAutonameSeries("CATIA Embedded Component");
                var Autoname = oCD5Engine.GenerateAutoname(strSizeType, 1);
                textBox_Name.Text = Autoname.GetValue(0).ToString();
            }

        }



        public void CD5_connect()
        {
            INFITF.Application catiaapp;
            CD5Template oTemplate;
            CD5Templates oTemplates;
            CD5TemplateTypes oTemplateTypes;
            CD5TemplateType oTemplateType;
            string strType = "";
            int indexImage = 0;

            Object CATIA = Marshal.GetActiveObject("CATIA.Application");
            catiaapp = (INFITF.Application)CATIA;

            oCD5Engine = catiaapp.GetItem("CD5EngineV6R2015") as CD5EngineV6R2015;

            if (oCD5Engine.IsConnected())
            {
                oTemplate = null;
                oTemplateTypes = oCD5Engine.TemplateTypes;

                for (int iTypes = 1; iTypes <= oTemplateTypes.Count; iTypes++)
                {

                    oTemplateType = null;
                    oTemplateType = oTemplateTypes.Item(iTypes);
                    strcurrentType = oTemplateType.TemplateTypeName;

                    switch (strcurrentType)
                    {
                        case "CATProduct Template":
                            strType = "CATProduct Template";
                            indexImage = 2;
                            break;

                        case "CATPart Template":
                            strType = "CATPart Template";
                            indexImage = 1;
                            break;

                        case "CATDrawing Template":
                            strType = "CATDrawing Template";
                            indexImage = 0;
                            break;
                    }

                    TreeNode tn0 = new TreeNode(strType);
                    tn0.ImageIndex = indexImage;
                    tn0.SelectedImageIndex = indexImage;
                    for (int iItem = 1; iItem <= oTemplateType.Templates.Count; iItem++)
                    {
                        oTemplates = oTemplateType.Templates;
                        oTemplate = oTemplates.Item(iItem);

                        TreeNode tn_product = new TreeNode(oTemplate.TemplateName);

                        tn0.Nodes.Add(tn_product);
                        tn_product.ImageIndex = indexImage;
                        tn_product.SelectedImageIndex = indexImage;
                    }
                    treeView1.Nodes.Add(tn0);
                    treeView1.ExpandAll();
                }

            }
            else
            {
                MessageBox.Show("Sie sind nicht an 3DExperience angemeldet.", "Abbruch");
                System.Windows.Forms.Application.Exit();
            }


        }

        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            strcurrentSelection = e.Node.Text;
            if (e.Node.SelectedImageIndex == 0)
            {
                if (e.Node.Text != "CATDrawing Template")
                {
                    strcurrentType = "CATDrawing Template";
                    comboBox_Type.Enabled = true;
                    comboBox_Type.Text = "CATDrawing";

                    if (textBox_Name.TextLength > 0 || comboBoxAutoN.Text != "")
                    {
                        if (textBoxEquipmentNumber.TextLength > 0 && comboBoxBetriebsmittelart.Text != "" && textBox_Bezeichnung2.TextLength > 0)
                        {
                            button_OK.Enabled = true;
                        }
                    }
                }
            }
            else if (e.Node.SelectedImageIndex == 1)
            {
                button_InsertCATPart.Enabled = true;
                if (e.Node.Text != "CATPart Template")
                {
                    strcurrentType = "CATPart Template";
                    comboBox_Type.Enabled = true;
                    comboBox_Type.Text = "CATPart";

                    if (textBox_Name.TextLength > 0 || comboBoxAutoN.Text != "")
                    {
                        if (textBoxEquipmentNumber.TextLength > 0 && comboBoxBetriebsmittelart.Text != "" && textBox_Bezeichnung2.TextLength > 0)
                        {
                            button_OK.Enabled = true;
                        }
                    }
                }
            }
            else if (e.Node.SelectedImageIndex == 2)
            {
                if (e.Node.Text != "CATProduct Template")
                {
                    strcurrentType = "CATProduct Template";
                    comboBox_Type.Enabled = true;
                    comboBox_Type.Text = "CATProduct";

                    if (textBox_Name.TextLength > 0 || comboBoxAutoN.Text != "")
                    {
                        if (textBoxEquipmentNumber.TextLength > 0 && comboBoxBetriebsmittelart.Text != "" && textBox_Bezeichnung2.TextLength > 0)
                        {
                            button_OK.Enabled = true;
                        }
                    }
                }

            }
            else
            {
                button_OK.Enabled = false;
                comboBox_Type.Enabled = false;
                comboBox_Type.Text = "";

            }

        }

        private void checkBox_AutoN_CheckedChanged(object sender, EventArgs e)
        {
            bool bcheck = checkBox_AutoN.Checked;
            if (bcheck == true)
            {
                comboBoxAutoN.Enabled = true;
                textBox_Name.Enabled = false;
                comboBoxAutoN.Text = "A Size";
                CD5_Item_count();
            }
            else
            {
                comboBoxAutoN.Enabled = false;
                textBox_Name.Enabled = true;
            }
        }

        private void toolStripContainer1_ContentPanel_Load(object sender, EventArgs e)
        {

        }

        private void button_OK_Click(object sender, EventArgs e)
        {
            CD5_openFile();
            my_Static.checkCatiaType();

            if (my_Static.DocType == "Product")
            {

                button_InsertCATPart.Visible = true;
                button_InsertCATPart.Enabled = true;

            }

            set_Catia_Properties();
        }

        private void radioButton_Insert_Part_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void comboBoxAutoN_SelectedIndexChanged(object sender, EventArgs e)
        {
            CD5_Item_count();
        }

        private void button_InsertCATPart_Click(object sender, EventArgs e)
        {
            my_Static.checkCatiaType();
            if (strcurrentType != "CATDrawing Template")
            {
                CD5_insert_CATPart();
            }
        }

        private void treeView1_BeforeSelect(object sender, TreeViewCancelEventArgs e)
        {

        }
    }
}
