using SolidWorks.Interop.swpublished;
using static System.Net.Mime.MediaTypeNames;

namespace Fraenkische.SWAddin.Core
{
    public static class SetBarText
    {
        /// <summary>
        /// Zapíše text do hlavního stavového řádku SolidWorks vlevo.<br/>
        /// Používá jednou uloženou instanci Frame.
        /// </summary>
        public static void Write(string text)
        {
            SWAddinClass.myFrame.SetStatusBarText(text);
        }

        /// <summary>
        /// Smaže zprávu ze stavového řádku.
        /// </summary>
        public static void Clear()
        {
            SWAddinClass.myFrame.SetStatusBarText(string.Empty);
        }
    }
}
