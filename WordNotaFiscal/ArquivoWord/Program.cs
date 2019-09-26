using System;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using Spire.Doc.Fields;
using Spire.Doc.Formatting;

namespace Criando_o_arquivo_word
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Programa de Nota Fiscal no Word");

            string clienteNome = "";
            string clienteEndereco = "";
            double valorCompra = 0;
            string dataCompra; 

            Console.WriteLine("Por favor digite seu nome:");
            clienteNome = Console.ReadLine();

            Console.WriteLine("Por favor digite seu endereço:");
            clienteEndereco = Console.ReadLine();

            Console.WriteLine("Qual é o valor da compra?");
            valorCompra = double.Parse(Console.ReadLine());

            Console.WriteLine("Qual a data da compra?");
            dataCompra = DateTime.Parse(Console.ReadLine()).ToString("dd-MM-yyyy");

            //Criando um novo documento com o nome documento.
            Document documentoNotaFiscal =  new Document();

            // Criando uma sessão dentro do documento.
            // A cada sessão criada, uma nova página é adicionada.
            Section secaoNotaFiscal = documentoNotaFiscal.AddSection();

            // Insere um título na primeira página.   
            Paragraph paragrafoNotaFiscal = secaoNotaFiscal.AddParagraph();


            TextBox textBoxNota = paragrafoNotaFiscal.AppendTextBox(300, 100);
            textBoxNota.Format.VerticalOrigin = VerticalOrigin.Margin;
            textBoxNota.Format.VerticalPosition = 100;
            textBoxNota.Format.HorizontalOrigin = HorizontalOrigin.Margin;
            textBoxNota.Format.HorizontalPosition = 100;
            textBoxNota.Format.NoLine = true;



            TextBox textBox = paragrafoNotaFiscal.AppendTextBox(300, 600);
            textBox.Format.VerticalOrigin = VerticalOrigin.Margin;
            textBox.Format.VerticalPosition = 140;
            textBox.Format.HorizontalOrigin = HorizontalOrigin.Margin;
            textBox.Format.HorizontalPosition = 50;
            textBox.Format.NoLine = true;

            CharacterFormat formatoTitulos = new CharacterFormat(documentoNotaFiscal);
            formatoTitulos.FontName = "Comic Sans";
            formatoTitulos.FontSize = 18;
            formatoTitulos.Bold = true;


            CharacterFormat formatoVariaveis = new CharacterFormat(documentoNotaFiscal);
            formatoVariaveis.FontName = "Comic Sans";
            formatoVariaveis.FontSize = 18;
            formatoVariaveis.Italic = true;

            Paragraph tituloNotaFiscal = textBoxNota.Body.AddParagraph();
            tituloNotaFiscal.AppendText("Nota Fiscal").ApplyCharacterFormat(formatoTitulos);

            Paragraph tituloNome = textBox.Body.AddParagraph();
            tituloNome.AppendText("Nome: ").ApplyCharacterFormat(formatoTitulos);
            tituloNome.AppendText(clienteNome).ApplyCharacterFormat(formatoVariaveis);

            Paragraph tituloEndereco = textBox.Body.AddParagraph();
            tituloEndereco.AppendText("Endereço: ").ApplyCharacterFormat(formatoTitulos);
            tituloEndereco.AppendText(clienteEndereco).ApplyCharacterFormat(formatoVariaveis);

            Paragraph tituloValor = textBox.Body.AddParagraph();
            tituloValor.AppendText("Valor: ").ApplyCharacterFormat(formatoTitulos);
            tituloValor.AppendText($"R$ {valorCompra}").ApplyCharacterFormat(formatoVariaveis);

            Paragraph tituloData = textBox.Body.AddParagraph();
            tituloData.AppendText("Data Compra: ").ApplyCharacterFormat(formatoTitulos);
            tituloData.AppendText(dataCompra).ApplyCharacterFormat(formatoVariaveis);


            documentoNotaFiscal.SaveToFile(@"Arquivo\notafiscal.docx", FileFormat.Docx);

        }
    }
}
