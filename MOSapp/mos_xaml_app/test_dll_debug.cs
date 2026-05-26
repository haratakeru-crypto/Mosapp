using System;
using System.Reflection;

class Program
{
    static void Main()
    {
        try
        {
            string dllPath = @"bin\Debug\Libraries\Group1\ExcelChecker1_2.dll";
            Console.WriteLine($"Loading DLL: {dllPath}");
            Assembly assembly = Assembly.LoadFrom(dllPath);
            Type checkerType = assembly.GetType("Libraries.Group1.ExcelChecker1_2");
            Console.WriteLine($"Type found: {checkerType?.Name ?? "null"}");
            
            if (checkerType != null)
            {
                object instance = Activator.CreateInstance(checkerType);
                
                // メソッドを直接呼び出してテスト
                var method = checkerType.GetMethod("CheckTask_1_2_01");
                if (method != null)
                {
                    Console.WriteLine("Calling CheckTask_1_2_01...");
                    bool result = (bool)method.Invoke(instance, null);
                    Console.WriteLine($"CheckTask_1_2_01 result: {result}");
                }
                else
                {
                    Console.WriteLine("Method CheckTask_1_2_01 not found");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
        }
    }
}
