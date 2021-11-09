using Microsoft.JSInterop;

namespace CSharpExcelFunctions.Services
{
    public class ExcelFunctionsService
    {
        [JSInvokable]
        public int Add(int firstValue, int secondValue) => firstValue + secondValue;

        [JSInvokable]
        public int Subtract(int firstValue, int secondValue) => firstValue - secondValue;

        [JSInvokable]
        public int Multiply(int firstValue, int secondValue) => firstValue * secondValue;

        [JSInvokable]
        public string Concat(string firstValue, string secondValue) => firstValue + secondValue;
    }
}
