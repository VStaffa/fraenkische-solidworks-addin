using Fraenkische.SWAddin.Commands;
using Fraenkische.SWAddin.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing.Text;
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

        private string assPath = string.Empty;
        private CheckBox check_offset;

        private readonly bool _autoMode;
        private readonly Face2[] _asmPair1;
        private readonly Face2[] _asmPair2;

        public GenerateInfillForm(SldWorks swApp)
        {
            InitializeComponent();

            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            _swApp = swApp;
            _infillTypes = new List<InfillType>();

            this.Text = "GENERATE INFILL FORM - CUSTOM SIZE MODE";

            InitializeInfillTypes();
            PopulateComboBox();

            // Event handlery
            cbox_types.SelectedIndexChanged += OnInputChanged;
            txt_w.TextChanged += OnInputChanged;
            txt_h.TextChanged += OnInputChanged;
            btn_ref.Click += OnRefreshClicked;
            btn_gen.Click += OnGenerateClicked;
            check_offset.CheckedChanged += OnInputChanged;

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
        /// Konstruktor pro automatický režim: pevné rozměry, uživatel jen vybere typ infillu a potvrdí.
        /// </summary>
        public GenerateInfillForm(SldWorks swApp, int widthMm, int heightMm, Face2[] asmPair1, Face2[] asmPair2)
            : this(swApp)   // zavolá původní ctor, který udělá InitializeComponent a eventy
        {
            // předvyplníme
            txt_w.Text = widthMm.ToString();
            txt_h.Text = heightMm.ToString();

            // zamkneme editaci
            txt_w.ReadOnly = true;
            txt_h.ReadOnly = true;
            label1.Enabled = false;  // popisky WIDTH/HEIGHT, pokud chcete
            label1.Text = "ROZMĚRY VYBRANÉHO OTVORU:";
            label2.Enabled = false;
            label2.Visible = false;

            check_offset.Checked = true; // přidáme offset, pokud je potřeba
            check_offset.Enabled = false; // vypneme možnost změny offsetu
            

            // tlačítko REFRESH necháme aktivní, nechá uživatele přepočítat popisky při změně typu
            btn_ref.Enabled = true;

            _autoMode = true;
            _asmPair1 = asmPair1;
            _asmPair2 = asmPair2;
            check_insert.Checked = true; // automaticky přidáme do sestavy
            check_insert.Visible = false;

            this.Text = "GENERATE INFILL FORM - INSERT + MATE MODE";
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
                lenPrefix: "4x (00823987-T) + (00823990-T + 00823991-T) L = ",
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
            var selIndex = _infillTypes.IndexOf(type);

            var perimeter = CalculatePerimeter(_pWidth, _pHeight, selIndex);

            if (check_offset.Checked)
            {
                // Přidáme offset, pokud je potřeba
                _pWidth += type.Offset;
                _pHeight += type.Offset;
                if (_pWidth < 0 || _pHeight < 0)
                {
                    MessageBox.Show(
                        "S offsetem musí být rozměry větší než 0 mm",
                        "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            _descVal = $"{type.DescPrefix}{_pWidth}x{_pHeight}";
            _lenVal = $"{type.LenPrefix}{perimeter}mm";

            lbl_desc.Text = _descVal;
            lbl_length.Text = _lenVal;
            btn_gen.Enabled = true;
        }

        /// <summary>
        /// Výpočet obvodu + offset
        /// </summary>
        private int CalculatePerimeter(int x, int y, int t)
        {
            switch (t)
            {
                case 1:
                    if (2 * (x + y) - 424 <= 0)
                    {
                        return 0;
                    }
                    else return 2 * (x + y) - 424;

                default:
                    return 2 * (x + y);

            }

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

            _swApp.DocumentVisible(false, 1);

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
                512,
                null,
                null,
                ref errs, ref warns);


            // 6.1) Přidat do sestavy, pokud je to možné
            var activeDoc = (ModelDoc2)_swApp.ActivateDoc(assPath);
            var asm = activeDoc as IAssemblyDoc;
            Component2 comp = null;

            if (asm != null)
            {
                comp = asm.AddComponent4(savePath, "", 0, 0, 0);
                activeDoc.EditRebuild3();
            }

            // ZAVAZBIT NOVĚ VLOŽENÝ DÍL
            if (_autoMode && comp != null)
            {
                var swExt = activeDoc.Extension;
                SelectionMgr selMgr = activeDoc.SelectionManager;

                MateFeatureData mateData;
                SymmetricMateFeatureData symmetricMateFeatureData;

                mateData = (MateFeatureData)asm.CreateMateData(8);
                symmetricMateFeatureData = (SymmetricMateFeatureData)mateData;


                AddSymetricalMate(_asmPair2, "Horizontal");

                // Pomocná lokální funkce pro vložení jednoho Symetrical Mate
                void AddSymetricalMate(Face2[] asmFaces, string direction)
                {
                    SymmetricMateFeatureData symmetricMateFeatureData;
                    CoincidentMateFeatureData coincidentMateFeatureData;

                    object faceVar;

                    object plane = null;

                    //0 - Coicident, 8 - Symmetrical
                    coincidentMateFeatureData = (CoincidentMateFeatureData)asm.CreateMateData(0);
                    symmetricMateFeatureData = (SymmetricMateFeatureData)asm.CreateMateData(8);
                    
                    symmetricMateFeatureData.SymmetryPlane = null;
                    symmetricMateFeatureData.EntitiesToMate = null;
                    symmetricMateFeatureData.MateAlignment = (int)swMateReferenceAlignment_e.swMateReferenceAlignment_Aligned;



                    // a) vymažeme výběr
                    activeDoc.ClearSelection2(true);

                    string compName = comp.Name2;
                    string assemblyName = Path.GetFileNameWithoutExtension(activeDoc.GetPathName());

                    MessageBox.Show($"{direction}@{compName}@{assemblyName}");

                    plane = swExt.SelectByID2(
                        $"{direction}@{compName}@{assemblyName}",
                        "PLANE",
                        0, 0, 0, false, 4, null, 0);

                    faceVar = asmFaces;

                    foreach (Face2 f in asmFaces) {
                        ISurface surf = f.GetSurface();
                    }

                    symmetricMateFeatureData.SymmetryPlane = plane;
                    symmetricMateFeatureData.EntitiesToMate = faceVar;

                    asm.CreateMate(symmetricMateFeatureData);
                }
                
            }


            if (check_openFolder.Checked)
                Process.Start("explorer.exe", outputDir);

            // 7) Zavřít šablonu beze změn

            SetBarText.Write("Výplň vygenerována");
            _swApp.DocumentVisible(true, 1);
            _swApp.CommandInProgress = false;


            MessageBox.Show("Výplň úspěšně vytvořena!", "Hotovo",
                MessageBoxButtons.OK, MessageBoxIcon.Information);

            return;
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
            this.check_offset = new System.Windows.Forms.CheckBox();
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
            this.label1.Location = new System.Drawing.Point(12, 53);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(47, 13);
            this.label1.TabIndex = 9;
            this.label1.Text = "WIDTH:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(115, 52);
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
            // check_offset
            // 
            this.check_offset.AutoSize = true;
            this.check_offset.Location = new System.Drawing.Point(195, 52);
            this.check_offset.Name = "check_offset";
            this.check_offset.Size = new System.Drawing.Size(93, 17);
            this.check_offset.TabIndex = 13;
            this.check_offset.Text = "ADD OFFSET";
            this.check_offset.UseVisualStyleBackColor = true;
            // 
            // GenerateInfillForm
            // 
            this.ClientSize = new System.Drawing.Size(554, 161);
            this.Controls.Add(this.check_offset);
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
