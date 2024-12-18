namespace OpenXMLCore
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("please input the file name:");

            string fileName = Console.ReadLine();

            if (!fileName.EndsWith(".docx"))
            {
                fileName += ".docx";
            }

            string filePath = @"C:\Users\usr\Code\MyProject\OpenXMLDemo\" + fileName;

            OpenXMLExecute.RunProcess(filePath);

            //Console.WriteLine($"表达式计算：{OpenXMLExecute.ValidExpression<long>(9223372036854775807, 50, "/")}");
        }
    }
}
