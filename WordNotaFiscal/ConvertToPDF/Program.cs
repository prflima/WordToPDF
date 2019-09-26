using System;
using Spire.Doc;
using Spire.Doc.Documents;

namespace Convertendo_pra_PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            Document documentoPdf = new Document();
            documentoPdf.LoadFromFile(@"..\ArquivoWord\Arquivo\notafiscal.docx");

            documentoPdf.SaveToFile("notafiscal.pdf", FileFormat.PDF);
        }
    }
}
