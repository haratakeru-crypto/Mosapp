using System;
using System.Reflection;
using System.Linq;

class Program
{
    static void Main()
    {
        try
        {
            Console.WriteLine("Testing ExcelChecker1_1.dll...");
            
            // Load the DLL
            Assembly assembly = Assembly.LoadFrom(@"Libraries\Group1\ExcelChecker1_1.dll");
            Console.WriteLine("DLL loaded successfully");
            
            // Get the type
            Type checkerType = assembly.GetTypes().FirstOrDefault(t => t.Name == "ExcelChecker1_1");
            if (checkerType != null)
            {
                Console.WriteLine($"Type found: {checkerType.Name}");
                Console.WriteLine($"Full name: {checkerType.FullName}");
                
                // List all methods
                Console.WriteLine("All methods:");
                foreach (MethodInfo method in checkerType.GetMethods())
                {
                    if (method.Name.StartsWith("CheckTask"))
                    {
                        Console.WriteLine($"  - {method.Name}");
                    }
                }
                
                // Test specific method
                MethodInfo testMethod = checkerType.GetMethod("CheckTask_1_1_01");
                if (testMethod != null)
                {
                    Console.WriteLine("CheckTask_1_1_01 method found!");
                }
                else
                {
                    Console.WriteLine("CheckTask_1_1_01 method NOT found!");
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
