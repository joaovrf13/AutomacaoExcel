using ClosedXML.Excel;
using System;



//biblioteca ClosedXML

namespace AutomacaoExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            int opcao;

           do
           {
                Console.WriteLine("=======================================");
                Console.WriteLine("         Bem-vindo ao Programa!        ");
                Console.WriteLine("=======================================");
                Console.WriteLine();
                //Menu


                Console.WriteLine("Selecione sua opção para prosseguir: ");
                Console.WriteLine();
                Console.WriteLine("(1) Relatório Formatado");
                Console.WriteLine("(2) Relatório para análise");
                Console.WriteLine("(3) Comparação de Relatórios (em desenvolvimento)"); //em desenvolvimento
                Console.WriteLine("(4) Sair do Programa ");
                Console.WriteLine();

               opcao = int.Parse(Console.ReadLine());
                switch (opcao)
                {
                    case 1:
                        //digite o local do seu arquivo Excel aqui
                        string caminhoArq = @"C:Caminho\Onde\Seu\Arquivo\Esta\arquivoformatado.xlsx";
                        //string caminhoArq = @"C:Caminho\Onde\Seu\Arquivo\Esta\arquivoformatado.xlsx";

                        var workbook = new XLWorkbook(caminhoArq);

                        var worksheet = workbook.Worksheet("Página 1");

                        //limpando coluna F até o final 
                        int ultimalinha = worksheet.LastRowUsed().RowNumber();
                        worksheet.Range(5, 6, ultimalinha, 6).Clear(XLClearOptions.All);



                        //Adicionando coluna 
                        // Entre F e G Data do Pagamento
                        worksheet.Column(7).InsertColumnsBefore(1);
                        // Entre G e H (que agora virou a 8) Agente
                        worksheet.Column(9).InsertColumnsBefore(1);

                        //formatando colunas 
                        worksheet.Cell(5, 6).Value = "Valor Pago";
                        worksheet.Cell(5, 7).Value = "Data do Pagamento";
                        worksheet.Cell(5, 9).Value = "Agente";


                        //digite o local do seu arquivo Excel aqui
                        workbook.SaveAs(@"C:Caminho\Onde\Seu\Arquivo\Esta\arquivoformatado.xlsx");
                        //workbook.SaveAs(@"C:Caminho\Onde\Seu\Arquivo\Esta\arquivoformatado.xlsx");

                        Console.WriteLine("Sucesso na Formatação!");


                        break;

                    case 2:

                        string caminhoArq2 = @"C:Caminho\Onde\Seu\Arquivo\Esta\arquivoformatado.xlsx";
                        

                        var workbook2 = new XLWorkbook(caminhoArq2);
                        var worksheet2 = workbook2.Worksheet("Relatório");

                        List<string> colunasParaManter = new List<string>
                     {
                         "Nome do Consumidor", "Documento", "Número de Instalação", "Referência","Valor Assinatura",
                         "Data de Vencimento Assinatura", "Data de Emissão Assinatura", "Status de Pagamento"
                    };

                        var colunas = worksheet2.ColumnsUsed();

                        for (int i = colunas.Count(); i > 0; i--)
                        {
                            var coluna = colunas.ElementAt(i - 1);
                            string Cabecalho = coluna.Cell(1).GetString();

                            if (!colunasParaManter.Contains(Cabecalho))
                            {
                                coluna.Delete();
                            }
                        }

                       
                       workbook2.SaveAs(@"C:Caminho\Onde\Seu\Arquivo\Esta\arquivoformatado.xlsx");
                       Console.WriteLine("Colunas deletadas com sucesso!");

                        break;

                    case 3:
                        Console.WriteLine("3");
                        break;

                    case 4:
                        Console.WriteLine("Saindo");
                        //fazer execução para fechar o programa
                        Environment.Exit(0);
                        break; //não vai ser usado mas evita do programa dar erro

                    default:
                        Console.WriteLine("Opção Inválida");
                        break;
                }
           } while (opcao != 4);
        }
    }
}


