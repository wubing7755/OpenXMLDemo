namespace OpenXMLCore
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("please input the file name:");

            string fileName = Console.ReadLine();

            if(!fileName.EndsWith(".docx"))
            {
                fileName += ".docx";
            }

            //OpenXMLExecute.CreateDocx(@"C:\Users\usr\Code\MyProject\OpenXMLDemo\" + fileName);
            //OpenXMLExecute.AddNewPart(@"C:\Users\usr\Code\MyProject\OpenXMLDemo\" + fileName);
        }
    }
}
