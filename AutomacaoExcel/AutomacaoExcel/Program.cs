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
                Console.WriteLine("(1) Relatório para análise");
                Console.WriteLine("(2) Sair do Programa ");
                Console.WriteLine();
                string pastaAutomacao = @"C:Local/Da/Pasta/Automação";
                string caminhoArq;

                opcao = int.Parse(Console.ReadLine());
                switch (opcao)
                {
                    case 1:
                        {
                            // Pega o arquivo Excel mais recente da pasta
                            var arquivos = Directory.GetFiles(pastaAutomacao, "*.xlsx");
                            if (arquivos.Length == 0)
                            {
                                Console.WriteLine("Nenhum arquivo .xlsx encontrado na pasta Automação!");
                                break;
                            }

                             caminhoArq = arquivos
                                .OrderByDescending(f => File.GetLastWriteTime(f))
                                .First();

                            Console.WriteLine($"Arquivo selecionado: {Path.GetFileName(caminhoArq)}");

                            var workbook2 = new XLWorkbook(caminhoArq);

                            // Pega a primeira aba do arquivo
                            var worksheet2 = workbook2.Worksheet(1);

                            List<string> colunasParaManter = new List<string>
                            {
                                "Nome do Consumidor", "Documento", "Número de Instalação", "Referência", "Valor Assinatura",
                                   "Data de Vencimento Assinatura", "Data de Emissão Assinatura", "Status de Pagamento"
                            };

                            var colunas = worksheet2.ColumnsUsed();

                            for (int i = colunas.Count(); i > 0; i--)
                            {
                                var coluna = colunas.ElementAt(i - 1);
                                string cabecalho = coluna.Cell(1).GetString();

                                if (!colunasParaManter.Contains(cabecalho))
                                {
                                    coluna.Delete();
                                }
                            }

                            // Cria pasta com data atual
                            string pastaData = Path.Combine(pastaAutomacao, DateTime.Now.ToString("yyyy-MM-dd"));
                            Directory.CreateDirectory(pastaData);

                            string novoArquivo = Path.Combine(pastaData, "RelatorioDGFormatado.xlsx");
                            workbook2.SaveAs(novoArquivo);

                            Console.WriteLine("Colunas deletadas com sucesso!");
                            Console.WriteLine($"Arquivo salvo em: {novoArquivo}");
                            break;
                        }
                    case 2:
                        Console.WriteLine("Saindo");
                        //fazer execução para fechar o programa
                        Environment.Exit(0);
                        break; //não vai ser usado mas evita do programa dar erro

                    default:
                        Console.WriteLine("Opção Inválida");
                        break;
                }
            } while (opcao != 2);
        }
    }
}


