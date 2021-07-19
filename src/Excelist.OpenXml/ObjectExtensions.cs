namespace System.Collections.Generic
{
    internal static class ObjectExtensions
    {
        internal static object GetPropValue(this object src, string propName)
            => src.GetType().GetProperty(propName).GetValue(src, null);
    }
}