using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Fraenkische.SWAddin.Commands
{
    internal class CMD_9_InfillGenerator : ICommand
    {
        private readonly List<InfillType> _infillTypes;
        private readonly SldWorks _swApp;
        public CMD_9_InfillGenerator(SldWorks swApp)
        {
            _swApp = swApp;
            _infillTypes = new List<InfillType>();
            InitializeInfillTypes();
        }

        public void Execute()
        {
            MessageBox.Show("Infill Generator Command Executed", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void Register(CommandManagerService cmdMgr)
        {
            cmdMgr.AddCommand(
            commandTitle: "Generovat výplně",
            tooltip: "Otevře formulář pro vygenerování výplně ze šablony",
            iconI: 8,
            callback: Execute
        );
        }
        private void InitializeInfillTypes()
        {
            var _infillTypes = new List<InfillType>
            {
                new InfillType("Plexi_4mm_cire",        "Plexi_čiré_4mm_",              "Tesneni_gumove_(823163): L = ",    22),
                new InfillType("Plexi_4mm_matne",       "Plexi_matne_4mm_",             "Tesneni_gumove_(823163): L = ",    22),
                new InfillType("Plexi_6mm_matne",       "Plexi_matne_6mm_",             "Tesneni_gumove_(823163): L = ",    22),
                new InfillType("Plexi_4mm_odjimatelne", "Plexi_odjimatelne_4mm_",       "4x(XXX)+(XXX) L = ",               -14),
                new InfillType("Sito_dratene",          "Sito_dratene_",                "Tesneni_gumove_(823122): L = ",    22),
             };
        }
    }
}
