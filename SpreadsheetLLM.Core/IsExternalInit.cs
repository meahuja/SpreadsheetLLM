// Polyfill required for C# 9 `init` property accessors when targeting netstandard2.0.
// The runtime type System.Runtime.CompilerServices.IsExternalInit is only present in
// .NET 5+; for older targets the compiler still emits a modreq referencing it, so we
// supply a dummy internal class with the exact same name and namespace.
namespace System.Runtime.CompilerServices
{
    internal static class IsExternalInit { }
}
