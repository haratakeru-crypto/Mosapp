using System;
using System.Reflection;
using System.Linq;

class Program
{
    static void Main()
    {
        try
        {
            Console.WriteLine("Testing ExcelChecker1_2.dll...");
            
            // Load the DLL
            Assembly assembly = Assembly.LoadFrom(@"Libraries\Group1\ExcelChecker1_2.dll");
            Console.WriteLine("DLL loaded successfully");
            
            // Get the type
            Type checkerType = assembly.GetTypes().FirstOrDefault(t => t.Name == "ExcelChecker1_2");
            if (checkerType != null)
            {
                Console.WriteLine($"Type found: {checkerType.Name}");
                Console.WriteLine($"Full name: {checkerType.FullName}");
                
                // List all CheckTask methods
                Console.WriteLine("CheckTask methods:");
                foreach (MethodInfo method in checkerType.GetMethods())
                {
                    if (method.Name.StartsWith("CheckTask"))
                    {
                        Console.WriteLine($"  - {method.Name}");
                    }
                }
                
                // Test specific methods
                string[] testMethods = {
                    "CheckTask_1_2_01",
                    "CheckTask_1_2_02", 
                    "CheckTask_1_2_03",
                    "CheckTask_1_2_04",
                    "CheckTask_1_2_05"
                };
                
                object checkerInstance = Activator.CreateInstance(checkerType);
                
                foreach (string methodName in testMethods)
                {
                    MethodInfo method = checkerType.GetMethod(methodName);
                    if (method != null)
                    {
                        Console.WriteLine($"Testing {methodName}...");
                        try
                        {
                            bool result = (bool)method.Invoke(checkerInstance, null);
                            Console.WriteLine($"  Result: {result}");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"  Exception: {ex.Message}");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Method {methodName} NOT found!");
                    }
                }
            }
            else
            {
                Console.WriteLine("Type not found!");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
        
        Console.WriteLine("Press any key to continue...");
        Console.ReadKey();
    }
}
