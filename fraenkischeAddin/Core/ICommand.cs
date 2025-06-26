using Fraenkische.SWAddin.Commands;

namespace Fraenkische.SWAddin
{
    internal interface ICommand
    {
        void Register(CommandManagerService cmdMgrService);
        void Execute();

    }
}
