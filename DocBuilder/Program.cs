namespace DocBuilder
{
    class Program
    {
        private const string SamplePath = @"C:\Users\brian.mcnulty\Documents\SampleDocs";

        static void Main(string[] args)
        {
            using (var obj = new DocumentProcessor())
            {
                obj.MergeDocuments();
            }
        }
    }
}
