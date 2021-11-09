using CSharpExcelFunctions.Extensions;
using CSharpExcelFunctions.Services;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.JSInterop;
using Newtonsoft.Json;

namespace MyData.Web.Pages.Excel
{
    public class FunctionsJsonModel : PageModel
    {
        public string Json { get; set; } = "";

        public void OnGet()
        {
            var excelType = typeof(ExcelFunctionsService);

            var methods = excelType.GetMethods(System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public);
            var invokableMethods = methods
                .Where(m => m.CustomAttributes.Any(a => a.AttributeType == typeof(JSInvokableAttribute)))
                .ToList();

            var functions = new List<ExcelFunctionForJson>();

            foreach (var method in invokableMethods)
            {
                var currentFunction = new ExcelFunctionForJson();
                currentFunction.description = method.Name;
                currentFunction.id = method.Name.InitialLetterLowercase();
                currentFunction.name = method.Name.ToUpper();
                currentFunction.parameters = new List<ExcelFunctionParameterForJson>();
                currentFunction.result = new ExcelFunctionResultForJson { type = ConvertTypeToExcelType(method.ReturnType) };

                var parameters = method.GetParameters();

                foreach (var param in parameters)
                {
                    var curentParam = new ExcelFunctionParameterForJson();
                    curentParam.description = param.Name!;
                    curentParam.name = param.Name!;
                    curentParam.type = ConvertTypeToExcelType(param.ParameterType);

                    currentFunction.parameters.Add(curentParam);
                }

                functions.Add(currentFunction);
            }

            Json = JsonConvert.SerializeObject(new { functions = functions });
        }


        private string ConvertTypeToExcelType(Type type)
        {
            var numberTypes = new List<string> { "int", "decimal", "double" };
            var typeName = type.Name.ToLower();

            if (numberTypes.Any(nt => typeName.Contains(nt))) return "number";
            return "string";
        }
    }

    public class ExcelFunctionForJson
    {
        public string description { get; set; } = "";
        public string id { get; set; } = "";
        public string name { get; set; } = "";
        public List<ExcelFunctionParameterForJson> parameters { get; set; } = new List<ExcelFunctionParameterForJson>();
        public ExcelFunctionResultForJson result { get; set; } = new ExcelFunctionResultForJson();
    }

    public class ExcelFunctionParameterForJson
    {
        public string description { get; set; } = "";
        public string name { get; set; } = "";
        public string type { get; set; } = "";
    }

    public class ExcelFunctionResultForJson
    {
        public string type { get; set; } = "";
    }
}
