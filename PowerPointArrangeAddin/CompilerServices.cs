// ReSharper disable CheckNamespace

// Refer to https://qiita.com/kenichiuda/items/fada6068ea265fd6a389.

namespace System.Runtime.CompilerServices {

    // ReSharper disable UnusedMember.Global

    internal static class IsExternalInit { }

    [AttributeUsage(AttributeTargets.Class |
                           AttributeTargets.Constructor |
                           AttributeTargets.Event |
                           AttributeTargets.Interface |
                           AttributeTargets.Method |
                           AttributeTargets.Module |
                           AttributeTargets.Property |
                           AttributeTargets.Struct, Inherited = false)]
    internal sealed class SkipLocalsInitAttribute : Attribute { }


    [AttributeUsage(AttributeTargets.Method, Inherited = false)]
    internal sealed class ModuleInitializerAttribute : Attribute { }

}
