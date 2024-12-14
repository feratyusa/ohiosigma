public class Helpers
{
    public static void ShowErrorMessage(Exception e)
    {
        string keyMessage = "Press any key to continue";
        
        Console.WriteLine();
        Console.WriteLine("-----ERROR------");
        Console.WriteLine(e.Message);
        Console.WriteLine();
        Console.WriteLine(keyMessage);
        Console.ReadKey(); 
    }
}