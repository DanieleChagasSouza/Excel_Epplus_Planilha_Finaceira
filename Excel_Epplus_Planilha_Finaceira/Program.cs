using OfficeOpenXml;

class Program
{
    static void Main(string[] args)
    {
        string caminhoDoArquivo = "C:\\Users\\DanieleChagasSouza\\source\\repos\\Excel_Epplus_Planilha_Finaceira\\Planilha_Financeira.xlsx";

        Console.WriteLine("Pressione uma tecla para iniciar!");
        Console.ReadKey();
        CriarPlanilhaExcel(caminhoDoArquivo);

        Console.WriteLine("Pressione uma tecla para Abrir a planilha/n");
        Console.ReadKey();
        AbrePlanilhaExcel(caminhoDoArquivo);

        Console.ReadKey();
    }

    private static void CriarPlanilhaExcel(string caminhoDoArquivo)
    {
        // Define dados fictícios
        var receitasEmpresas = new[]
        {
            new {Data = "11/01/2024", Descricao = "venda de Produto", Categoria = "Vendas", Valor = 1500.00 },
            new {Data = "10/01/2024", Descricao = "Investimento", Categoria = "Investimentos", Valor = 5000.00 },
            new {Data = "04/01/2024", Descricao = "Administrativo", Categoria = "Administrativos", Valor = 2500.00 },
            new {Data = "01/01/2024", Descricao = "Recursos Humanos", Categoria = "Gestão e Gente", Valor = 3500.00 },
            new {Data = "19/06/2024", Descricao = "Setor operacional", Categoria = "Operações", Valor = 1200.00 }

       };
        // Define dados fictícios
        var despesas = new[]
        {
            new { Data = "02/01/2024", Descricao = "Pagamento de Salários", Categoria = "Salários", Valor = 3000.00 },
            new { Data = "10/01/2024", Descricao = "Campanha de Marketing", Categoria = "Marketing", Valor = 1200.00 },
            new { Data = "15/01/2024", Descricao = "Pagamento de Faturas", Categoria = "Pagamentos", Valor = 2000.00 },
            new { Data = "30/01/2024", Descricao = "Administrativo", Categoria = "Administrativos", Valor = 1600.00 }
        };

        // Define o contexto de licença para uso não comercial do EPPlus
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (ExcelPackage excel =  new ExcelPackage())
        {
            // Adiciona uma nova planilha ao pacote Excel
            var sheet = excel.Workbook.Worksheets.Add("Financeiro");

            // Cabeçalho da planilha
            sheet.Cells[1, 1].Value = "Nome da Empresa: New Show Tech Brazil";
            sheet.Cells[1, 2].Value = "Período Financeiro: Janeiro 2024";

            // Adiciona cor de fundo ao cabeçalho de resumo financeiro
            sheet.Cells["A1:D1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            sheet.Cells["A1:D1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);

            // Títulos da seção de receitas
            sheet.Cells[4, 1].Value = "Receitas";
            sheet.Cells[5, 1].Value = "Data";
            sheet.Cells[5, 2].Value = "Descrição";
            sheet.Cells[5, 3].Value = "Categoria";
            sheet.Cells[5, 4].Value = "Valor";

            // Adiciona cor de fundo ao cabeçalho
            sheet.Cells["A5:D5"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            sheet.Cells["A5:D5"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);

            // Itera sobre os dados de receitas e preenche as células
            int linhaReceita = 6;

            foreach (var receita in receitasEmpresas)
            {
                sheet.Cells[linhaReceita, 1].Value = receita.Data;
                sheet.Cells[linhaReceita, 2].Value = receita.Descricao;
                sheet.Cells[linhaReceita, 3].Value = receita.Categoria;
                sheet.Cells[linhaReceita, 4].Value = receita.Valor;
                linhaReceita++;
            }

            // Títulos da seção de despesas
            int linhaDespesaTitulo = linhaReceita + 1;

            sheet.Cells[linhaDespesaTitulo, 1].Value = "Despesas";
            sheet.Cells[linhaDespesaTitulo + 1, 1].Value = "Data";
            sheet.Cells[linhaDespesaTitulo + 1, 2].Value = "Descrição";
            sheet.Cells[linhaDespesaTitulo + 1, 3].Value = "Categoria";
            sheet.Cells[linhaDespesaTitulo + 1, 4].Value = "Valor";

            // Adiciona cor de fundo ao cabeçalho
            sheet.Cells[$"A{linhaDespesaTitulo + 1}:D{linhaDespesaTitulo + 1}"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            sheet.Cells[$"A{linhaDespesaTitulo + 1}:D{linhaDespesaTitulo + 1}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightCoral);

            // Itera sobre os dados de despesas e preenche as células
            int linhaDespesa = linhaDespesaTitulo + 2;
            foreach (var despesa in despesas)
            {
                sheet.Cells[linhaDespesa, 1].Value = despesa.Data;
                sheet.Cells[linhaDespesa, 2].Value = despesa.Descricao;
                sheet.Cells[linhaDespesa, 3].Value = despesa.Categoria;
                sheet.Cells[linhaDespesa, 4].Value = despesa.Valor;
                linhaDespesa++;
            }

            // Define o formato das células de valores
            sheet.Column(4).Style.Numberformat.Format = "0.00";

            // Resumo financeiro
            int linhaResumo = linhaDespesa + 1;
            sheet.Cells[linhaResumo, 1].Value = "Resumo Financeiro";
            sheet.Cells[linhaResumo + 1, 1].Value = "Receita Total";
            sheet.Cells[linhaResumo + 1, 4].Formula = $"SUM(D6:D{linhaReceita - 1})";

            // Adiciona cor de fundo ao cabeçalho de resumo financeiro
            sheet.Cells["A19:D19"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            sheet.Cells["A19:D19"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkOrange);

            sheet.Cells[linhaResumo + 2, 1].Value = "Despesa Total";
            sheet.Cells[linhaResumo + 2, 4].Formula = $"SUM(D{linhaDespesaTitulo + 2}:D{linhaDespesa - 1})";

            sheet.Cells[linhaResumo + 3, 1].Value = "Saldo";
            sheet.Cells[linhaResumo + 3, 4].Formula = $"B{linhaResumo + 1} - B{linhaResumo + 2}";

            // Ajusta automaticamente a largura das colunas
            sheet.Column(1).AutoFit();
            sheet.Column(2).AutoFit();
            sheet.Column(3).AutoFit();
            sheet.Column(4).AutoFit();

            // Salva o arquivo Excel no disco físico
            if (File.Exists(caminhoDoArquivo)) File.Delete(caminhoDoArquivo);
            File.WriteAllBytes(caminhoDoArquivo, excel.GetAsByteArray());


        }
        Console.WriteLine($"Planilha criada com sucesso: {caminhoDoArquivo}\n");

    }
    private static void AbrePlanilhaExcel(string caminhoDoArquivo)
    {
        // Abre o arquivo Excel existente
        using (var arquivoExcel = new ExcelPackage(new FileInfo(caminhoDoArquivo)))
        {
            // Obtém a primeira planilha do pacote Excel
            ExcelWorksheet planilha = arquivoExcel.Workbook.Worksheets.FirstOrDefault();

            // Verifica se a planilha foi encontrada
            if (planilha == null)
            {
                Console.WriteLine("Nem uma planilha encontrada!");
                return;
            }

            // Obter o número de linhas e colunas
            int roms = planilha.Dimension.Rows;
            int cols = planilha.Dimension.Columns;

            // Percorre as linhas e colunas da planilha
            for (int i = 1; i <= roms; i++)
            {
                for (int j = 1; j <= cols; j++)
                {
                    string conteudo = planilha.Cells[i, j].Value?.ToString() ?? string.Empty;
                    Console.WriteLine(conteudo); // Obtém o valor da célula e exibe no console
                }
            }
        }
    }
}