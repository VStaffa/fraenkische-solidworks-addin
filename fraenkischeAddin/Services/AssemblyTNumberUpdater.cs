using System;
using System.Collections.Generic;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swcommands;
using SolidWorks.Interop.swpublished;

namespace Fraenkische.SWAddin.Services
{
    public class AssemblyTNumberUpdater
    {
        private readonly ISldWorks _swApp;
        private readonly TNumberExcelReader _excelReader;
        private readonly CustomPropertyEditor _propertyEditor;


        public AssemblyTNumberUpdater(
            ISldWorks swApp,
            TNumberExcelReader excelReader,
            CustomPropertyEditor propertyEditor
            ) 

        {
            _swApp = swApp;
            _excelReader = excelReader;
            _propertyEditor = propertyEditor;
        }

        public void UpdateAllComponentsTNumbers(IAssemblyDoc assemblyDoc)
        {
            // 1. Získat komponenty
            IComponent2[] components = (IComponent2[])assemblyDoc.GetComponents(true);

            foreach (IComponent2 component in components)
            {
                var model = (ModelDoc2)component.GetModelDoc2();
                if (model == null)
                    continue;

                // 2. Zkontrolovat Custom Property
                string tNumber = _propertyEditor.GetTNumber(model);
                if (!string.IsNullOrWhiteSpace(tNumber))
                    continue; // Už má T-číslo

                // 3. Získat název komponenty a hledat v Excelu
                string componentName = "test";
                string foundTNumber = _excelReader.GetTNumberForComponent(componentName);

                if (!string.IsNullOrWhiteSpace(foundTNumber))
                {
                    // 4. Zapsat do modelu
                    _propertyEditor.SetTNumber(model, foundTNumber);
                }
            }
        }
    }
}
