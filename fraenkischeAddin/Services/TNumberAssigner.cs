﻿using System.IO;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using SolidWorks.Interop.sldworks;

namespace Fraenkische.SWAddin.Services
{
    public class TNumberAssigner
    {
        private readonly ISldWorks _swApp;
        private readonly TNumberExcelReader _excelReader;
        private readonly CustomPropertyEditor _propertyEditor;


        public TNumberAssigner(
            ISldWorks swApp,
            TNumberExcelReader excelReader,
            CustomPropertyEditor propertyEditor
            )

        {
            _swApp = swApp;
            _excelReader = excelReader;
            _propertyEditor = propertyEditor;
        }

        public void UpdateTNumber(ModelDoc2 swModel)
        {
            // 1. Zkontrolovat Custom Property
            string tNumber = _propertyEditor.GetTNumber(swModel);
            if (!string.IsNullOrWhiteSpace(tNumber))
            {
                MessageBox.Show($"TENTO DÍL JIŽ MÁ T-ČÍSLO!\nT-ČÍSLO: {tNumber}", "UPOZORNĚNÍ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 2. Získat název komponenty a hledat v Excelu
            string fullName = swModel.GetTitle();
            string componentName = Path.GetFileNameWithoutExtension(fullName);

            string userInputName = Interaction.InputBox("Enter a name:", "Input Required", componentName);

            if (string.IsNullOrWhiteSpace(userInputName))
            {
                MessageBox.Show("Nemuze byt prazdne!", "CHYBA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string foundTNumber = _excelReader.GetTNumberForComponent(userInputName);

            if (!string.IsNullOrWhiteSpace(foundTNumber))
            {
                // 3. Zapsat do modelu
                _propertyEditor.SetTNumber(swModel, foundTNumber);
            }

        }
    }
}
