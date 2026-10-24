namespace Gongwen.AddinRegistrar;

internal enum RegistrationOperation { Invalid, Register, Unregister }

internal static class RegistrarCommandLine
{
    internal static bool TryParse(string[] arguments, out RegistrationOperation operation)
    {
        operation = RegistrationOperation.Invalid;
        if (arguments is null || arguments.Length != 1) return false;
        if (arguments[0] == "--register") operation = RegistrationOperation.Register;
        else if (arguments[0] == "--unregister") operation = RegistrationOperation.Unregister;
        return operation != RegistrationOperation.Invalid;
    }
}
