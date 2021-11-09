using CSharpExcelFunctions.Extensions;
using CSharpExcelFunctions.Services;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.JSInterop;

namespace MyData.Web.Pages.Excel
{
    public class FunctionsModel : PageModel
    {
        public string JavaScript { get; set; } = "";

        public void OnGet()
        {
            var excelType = typeof(ExcelFunctionsService);

            var methods = excelType.GetMethods(System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public);
            var invokableMethods = methods
                .Where(m => m.CustomAttributes.Any(a => a.AttributeType == typeof(JSInvokableAttribute)))
                .ToList();

            var output = "";

            foreach (var method in invokableMethods)
            {
                var parameters = method.GetParameters();
                var paramList = string.Join(",", parameters.Select(p => p.Name!.InitialLetterLowercase()).ToList());

                output += $"async function {method.Name.InitialLetterLowercase()}({paramList}){Environment.NewLine}{{{Environment.NewLine}";

                output += $"\treturn await functionsRef.invokeMethodAsync('{method.Name}'";
                
                if (paramList != "") output += $", {paramList}";
                
                output += $"){Environment.NewLine}}}";
            }

            JavaScript = output;
        }
    }
}
