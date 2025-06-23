using System.Windows;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Fraenkische.SWAddin.Services
{
    public class CustomPropertyEditor
    {
        /// <summary>
        /// Získá hodnotu vlastnosti "T-Number" z modelu.
        /// </summary>
        public string GetTNumber(ModelDoc2 model)

        {
            var propMgr = model.Extension.CustomPropertyManager[""]; // prázdný string = aktuální konfigurace
            propMgr.Get5("T-Number", false, out string value, out _, out bool wsRes);
            return value;
        }

        /// <summary>
        /// Nastaví vlastnost "T-Number" v modelu.
        /// </summary>
        public void SetTNumber(ModelDoc2 model, string tNumber)
        {
            var propMgr = model.Extension.CustomPropertyManager[""];

            // přidá nebo nahradí hodnotu vlastnosti
            propMgr.Add3(
                "T-Number",
                (int)swCustomInfoType_e.swCustomInfoText,
                tNumber,
                (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);

            // provede rebuild a uloží model
            model.ForceRebuild3(false);
            model.Save3(
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                0,
                0);
            MessageBox.Show($"T-Cislo: {tNumber}, Pridano dilu:{model.GetTitle()} ");
        }
    }
}
