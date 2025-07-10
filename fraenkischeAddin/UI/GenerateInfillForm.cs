using Fraenkische.SWAddin.Commands;
using Fraenkische.SWAddin.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Fraenkische.SWAddin.UI
{
    /// <summary>
    /// Formulář pro generování výplní pomocí jedné PART-šablony s konfiguracemi.
    /// </summary>
    public partial class GenerateInfillForm : Form
    {
        private readonly SldWorks _swApp;
        private readonly List<InfillType> _infillTypes;

        private int _pWidth;
        private int _pHeight;
        private string _descVal;
        private Button btn_ref;
        private Button btn_gen;
        private TextBox txt_w;
        private TextBox txt_h;
        private ComboBox cbox_types;
        private Label lbl_desc;
        private Label lbl_length;
        private CheckBox check_insert;
        private CheckBox check_openFolder;
        private Label label1;
        private Label label2;
        private Label label3;
        private Label label4;
        private string _lenVal;

        private string  assPath = string.Empty;
        private readonly string[] _insertedFaceNames = { "Leva", "Prava", "Horni", "Dolni" };

        public GenerateInfillForm(SldWorks swApp)
        {
            InitializeComponent();

            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            _swApp = swApp;
            _infillTypes = new List<InfillType>();

            InitializeInfillTypes();
            PopulateComboBox();

            // Event handlery
            cbox_types.SelectedIndexChanged += OnInputChanged;
            txt_w.TextChanged += OnInputChanged;
            txt_h.TextChanged += OnInputChanged;
            btn_ref.Click += OnRefreshClicked;
            btn_gen.Click += OnGenerateClicked;

            // Počáteční stav
            btn_gen.Enabled = false;
            lbl_desc.Text = lbl_length.Text = "-- STISKNĚTE OBNOVIT --";

            ModelDoc2 actDoc = _swApp.ActiveDoc as ModelDoc2;

            if (actDoc != null)
            {
                if (actDoc.GetType() == 2)
                {
                    assPath = actDoc.GetPathName();
                    check_insert.Enabled = true;
                }
            }
        }

        /// <summary>
        /// Naplnění seznamu InfillType – jeden template, více konfigurací
        /// </summary>
        private void InitializeInfillTypes()
        {

            string template = Path.Combine(
               Path.GetDirectoryName(typeof(CMD_8_CreateGaugeDrawing).Assembly.Location),
               "Resources",
               "SWTemplates",
               "InfillTemplate.SLDPRT"
           );

            _infillTypes.Add(new InfillType(
                title: "Plexi matné (6mm)",
                templatePath: template,
                configName: "Plexi_6mm_matne",
                descPrefix: "Plexi_matne_6mm_",
                lenPrefix: "Tesneni_gumove_(823163): L = ",
                offset: 22));

            _infillTypes.Add(new InfillType(
                title: "Plexi odjímatelné (6mm)",
                templatePath: template,
                configName: "Plexi_6mm_odjimatelne",
                descPrefix: "Plexi_odjimatelne_6mm_",
                lenPrefix: "4x(XXX)+(XXX) L = ",
                offset: -14));

            _infillTypes.Add(new InfillType(
                title: "Síto drátěné",
                templatePath: template,
                configName: "Sito_dratene",
                descPrefix: "Sito_dratene_",
                lenPrefix: "Tesneni_gumove_(823122): L = ",
                offset: 20));

            _infillTypes.Add(new InfillType(
                title: "Plexi čiré (4mm)",
                templatePath: template,
                configName: "Plexi_4mm_cire",
                descPrefix: "Plexi_cire_4mm_",
                lenPrefix: "Tesneni_gumove_(823163): L = ",
                offset: 22));

            _infillTypes.Add(new InfillType(
                title: "Plexi matné (4mm)",
                templatePath: template,
                configName: "Plexi_4mm_matne",
                descPrefix: "Plexi_matne_4mm_",
                lenPrefix: "Tesneni_gumove_(823163): L = ",
                offset: 22));
        }

        /// <summary>
        /// Naplní ComboBox názvy typů
        /// </summary>
        private void PopulateComboBox()
        {
            cbox_types.Items.Clear();
            cbox_types.Items.AddRange(_infillTypes
                .Select(t => t.Title)
                .ToArray());
            cbox_types.SelectedIndex = 0;
        }

        /// <summary>
        /// Při změně vstupu zruší povolení tlačítka Generovat
        /// </summary>
        private void OnInputChanged(object sender, EventArgs e)
        {
            btn_gen.Enabled = false;
            lbl_desc.Text = lbl_length.Text = "-- STISKNĚTE OBNOVIT --";
        }

        /// <summary>
        /// Validuje rozměry a připraví texty Description a LENGTH
        /// </summary>
        private void OnRefreshClicked(object sender, EventArgs e)
        {
            if (!int.TryParse(txt_w.Text, out _pWidth) ||
                !int.TryParse(txt_h.Text, out _pHeight) ||
                _pWidth <= 0 || _pHeight <= 0)
            {
                MessageBox.Show(
                    "Rozměr musí být celé číslo větší než 0 mm",
                    "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Vždy větší × menší
            if (_pWidth < _pHeight)
            {
                (_pWidth, _pHeight) = (_pHeight, _pWidth);
                txt_w.Text = _pWidth.ToString();
                txt_h.Text = _pHeight.ToString();
            }

            var type = _infillTypes[cbox_types.SelectedIndex];
            _descVal = $"{type.DescPrefix}{_pWidth}x{_pHeight}";
            var perimeter = CalculatePerimeter(_pWidth, _pHeight, type.Offset);
            _lenVal = $"{type.LenPrefix}{perimeter}mm";

            lbl_desc.Text = _descVal;
            lbl_length.Text = _lenVal;
            btn_gen.Enabled = true;
        }

        /// <summary>
        /// Výpočet obvodu + offset
        /// </summary>
        private int CalculatePerimeter(int x, int y, int offset)
        {
            return 2 * (x + y) + offset;
        }

        /// <summary>
        /// Otevře šablonu, přepne konfiguraci, nastaví parametry, uloží kopii a případně vloží/vyhledá složku
        /// </summary>
        private void OnGenerateClicked(object sender, EventArgs e)
        {
            var type = _infillTypes[cbox_types.SelectedIndex];

            // 1) Rozhodnutí kam ukládat
            string outputDir;
            var active = (ModelDoc2)_swApp.ActiveDoc;

            if (active == null)
            {
                using (var dlg = new FolderBrowserDialog
                {
                    Description = "Vyberte složku pro uložení vygenerované výplně"
                })
                {
                    if (dlg.ShowDialog() != DialogResult.OK)
                    {
                        MessageBox.Show("Ukládání bylo zrušeno.", "Zrušeno",
                                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        _swApp.CommandInProgress = false;
                        return;
                    }
                    outputDir = dlg.SelectedPath;

                }
                
            }
            else
            {
                if (active.GetType() is 2)
                {
                    // běžíme v sestavě se skutečnou cestou
                    outputDir = Path.GetDirectoryName(active.GetPathName());

                }
                else
                {
                    // Part nebo prázdné — necháme uživatele vybrat složku
                    using (var dlg = new FolderBrowserDialog
                    {
                        Description = "Vyberte složku pro uložení vygenerované výplně"
                    })
                    {
                        if (dlg.ShowDialog() != DialogResult.OK)
                        {
                            MessageBox.Show("Ukládání bylo zrušeno.", "Zrušeno",
                                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            _swApp.CommandInProgress = false;
                            return;
                        }
                        outputDir = dlg.SelectedPath;

                    }
                    
                }
            }

            _swApp.DocumentVisible(false,1);

            // 2) Otevřít template (Read-Only)
            int errs = 0, warns = 0;
            var model = (ModelDoc2)_swApp.OpenDoc6(
                type.TemplatePath,
                (int)swDocumentTypes_e.swDocPART,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent |
                (int)swOpenDocOptions_e.swOpenDocOptions_ReadOnly,
                "", ref errs, ref warns);

            // 3) Přepnout konfiguraci
            model.ShowConfiguration2(type.ConfigName);
            model.EditRebuild3();

            // 4) Nastavit parametry (v metrech)
            model.Parameter("D1@DimensionsSketch").SetSystemValue3(_pWidth / 1000.0, 2, null);
            model.Parameter("D2@DimensionsSketch").SetSystemValue3(_pHeight / 1000.0, 2, null);
            model.EditRebuild3();

            // 5) Přidat vlastní vlastnosti
            var cust = model.Extension.CustomPropertyManager[""];
            cust.Add3("Description",
                (int)swCustomInfoType_e.swCustomInfoText,
                _descVal,
                (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
            cust.Add3("LENGTH",
                (int)swCustomInfoType_e.swCustomInfoText,
                _lenVal,
                (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);



            // 6) Uložit jako kopii
            var savePath = Path.Combine(outputDir, _descVal + ".SLDPRT");
            model.Extension.SaveAs3(
                savePath,
                0,
                (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen,
                null,
                null,
                ref errs, ref warns);

            // 7) Volitelné vložení a otevření složky
            if (check_insert.Checked == true)
            {
                //MessageBox.Show(assPath);
                //MessageBox.Show(savePath);



                var swModel = _swApp.ActivateDoc3(assPath,true,0,0) as ModelDoc2;
                var swAssy = (AssemblyDoc)swModel;
                Component2 newComp;

                newComp = swAssy.AddComponent5(savePath, 0, "", false, "", 0, 0, 0);

                var compDoc = (PartDoc)newComp.GetModelDoc2();
                //compDoc.GetEntityByName("Leva", (int)swSelectType_e.swSelFACES);

                swModel.ClearSelection2(true);

                // Najdeme v ní plochy podle jmen
                var insertedFaces = new IFace2[4];
                for (int i = 0; i < 4; i++)
                {
                    var entity = compDoc.GetEntityByName(
                        _insertedFaceNames[i],
                        (int)swSelectType_e.swSelFACES);
                    insertedFaces[i] = entity as IFace2;
                    MessageBox.Show($"Plocha '{_insertedFaceNames[i]}' nalezena.");
                    if (insertedFaces[i] == null)
                    {
                        MessageBox.Show($"Plocha '{_insertedFaceNames[i]}' ve vložené součásti nenalezena.",
                                        "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        SetBarText.Clear();
                        return;
                    }
                }

                _swApp.CloseDoc(savePath);
            }

            if (check_openFolder.Checked)
                Process.Start("explorer.exe", outputDir);

            // 8) Zavřít šablonu beze změn
            //_swApp.CloseDoc(model.GetTitle());
            SetBarText.Write("Výplň vygenerována");
            _swApp.DocumentVisible(true, 1);
            _swApp.CommandInProgress = false;


            MessageBox.Show("Výplň úspěšně vytvořena!", "Hotovo",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void InitializeComponent()
        {
            this.btn_ref = new System.Windows.Forms.Button();
            this.btn_gen = new System.Windows.Forms.Button();
            this.txt_w = new System.Windows.Forms.TextBox();
            this.txt_h = new System.Windows.Forms.TextBox();
            this.cbox_types = new System.Windows.Forms.ComboBox();
            this.lbl_desc = new System.Windows.Forms.Label();
            this.lbl_length = new System.Windows.Forms.Label();
            this.check_insert = new System.Windows.Forms.CheckBox();
            this.check_openFolder = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btn_ref
            // 
            this.btn_ref.Location = new System.Drawing.Point(212, 74);
            this.btn_ref.Name = "btn_ref";
            this.btn_ref.Size = new System.Drawing.Size(76, 26);
            this.btn_ref.TabIndex = 0;
            this.btn_ref.Text = "REFRESH";
            this.btn_ref.UseVisualStyleBackColor = true;
            // 
            // btn_gen
            // 
            this.btn_gen.Location = new System.Drawing.Point(12, 106);
            this.btn_gen.Name = "btn_gen";
            this.btn_gen.Size = new System.Drawing.Size(276, 40);
            this.btn_gen.TabIndex = 1;
            this.btn_gen.Text = "GENERATE";
            this.btn_gen.UseVisualStyleBackColor = true;
            // 
            // txt_w
            // 
            this.txt_w.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_w.Location = new System.Drawing.Point(12, 74);
            this.txt_w.Name = "txt_w";
            this.txt_w.Size = new System.Drawing.Size(100, 26);
            this.txt_w.TabIndex = 2;
            // 
            // txt_h
            // 
            this.txt_h.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_h.Location = new System.Drawing.Point(118, 74);
            this.txt_h.Name = "txt_h";
            this.txt_h.Size = new System.Drawing.Size(89, 26);
            this.txt_h.TabIndex = 3;
            // 
            // cbox_types
            // 
            this.cbox_types.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbox_types.FormattingEnabled = true;
            this.cbox_types.Location = new System.Drawing.Point(12, 12);
            this.cbox_types.Name = "cbox_types";
            this.cbox_types.Size = new System.Drawing.Size(276, 28);
            this.cbox_types.TabIndex = 4;
            // 
            // lbl_desc
            // 
            this.lbl_desc.AutoSize = true;
            this.lbl_desc.Location = new System.Drawing.Point(329, 44);
            this.lbl_desc.Name = "lbl_desc";
            this.lbl_desc.Size = new System.Drawing.Size(22, 13);
            this.lbl_desc.TabIndex = 5;
            this.lbl_desc.Text = "- - -";
            // 
            // lbl_length
            // 
            this.lbl_length.AutoSize = true;
            this.lbl_length.Location = new System.Drawing.Point(329, 97);
            this.lbl_length.Name = "lbl_length";
            this.lbl_length.Size = new System.Drawing.Size(22, 13);
            this.lbl_length.TabIndex = 6;
            this.lbl_length.Text = "- - -";
            // 
            // check_insert
            // 
            this.check_insert.AutoSize = true;
            this.check_insert.Enabled = false;
            this.check_insert.Location = new System.Drawing.Point(323, 129);
            this.check_insert.Name = "check_insert";
            this.check_insert.Size = new System.Drawing.Size(110, 17);
            this.check_insert.TabIndex = 7;
            this.check_insert.Text = "Insert to assembly";
            this.check_insert.UseVisualStyleBackColor = true;
            // 
            // check_openFolder
            // 
            this.check_openFolder.AutoSize = true;
            this.check_openFolder.Location = new System.Drawing.Point(453, 129);
            this.check_openFolder.Name = "check_openFolder";
            this.check_openFolder.Size = new System.Drawing.Size(81, 17);
            this.check_openFolder.TabIndex = 8;
            this.check_openFolder.Text = "Open folder";
            this.check_openFolder.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 56);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(47, 13);
            this.label1.TabIndex = 9;
            this.label1.Text = "WIDTH:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(115, 56);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(51, 13);
            this.label2.TabIndex = 10;
            this.label2.Text = "HEIGHT:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(320, 20);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(137, 13);
            this.label3.TabIndex = 11;
            this.label3.Text = "INFILL DESCRIPTION:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(320, 74);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(71, 13);
            this.label4.TabIndex = 12;
            this.label4.Text = "BOM INFO:";
            // 
            // GenerateInfillForm
            // 
            this.ClientSize = new System.Drawing.Size(555, 161);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.check_openFolder);
            this.Controls.Add(this.check_insert);
            this.Controls.Add(this.lbl_length);
            this.Controls.Add(this.lbl_desc);
            this.Controls.Add(this.cbox_types);
            this.Controls.Add(this.txt_h);
            this.Controls.Add(this.txt_w);
            this.Controls.Add(this.btn_gen);
            this.Controls.Add(this.btn_ref);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(630, 200);
            this.MinimizeBox = false;
            this.Name = "GenerateInfillForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "INFILL GENERATOR FORM";
            this.ResumeLayout(false);
            this.PerformLayout();

        }
    }

    /// <summary>
    /// Popis jednoho typu výplně s konfigurací
    /// </summary>
    internal class InfillType
    {
        public string Title { get; }
        public string TemplatePath { get; }
        public string ConfigName { get; }
        public string DescPrefix { get; }
        public string LenPrefix { get; }
        public int Offset { get; }

        public InfillType(string title, string templatePath, string configName,
                          string descPrefix, string lenPrefix, int offset)
        {
            Title = title;
            TemplatePath = templatePath;
            ConfigName = configName;
            DescPrefix = descPrefix;
            LenPrefix = lenPrefix;
            Offset = offset;
        }
    }
}
