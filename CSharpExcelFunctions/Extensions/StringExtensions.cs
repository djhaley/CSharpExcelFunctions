namespace CSharpExcelFunctions.Extensions
{
    public static class StringExtensions
    {
        public static string InitialLetterLowercase(this string @this)
        {
            if (string.IsNullOrEmpty(@this)) return "";
            if (@this.Length == 1) return @this.ToLower();

            return @this.Substring(0, 1).ToLower() + @this.Substring(1);
        }
    }
}
