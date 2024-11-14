using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.Globalization;
using FinanceiroScript.Dominio;

public class ExcelHelper
{
    public static bool ValidarNFSe(NFSe dadosNfse, string caminhoArquivoExcel)
    {
        if (dadosNfse == null) return false;
        if (string.IsNullOrEmpty(caminhoArquivoExcel)) return false;

        try
        {
            if (string.IsNullOrEmpty(dadosNfse.Prestador.Cnpj) || string.IsNullOrEmpty(dadosNfse.DataCompetencia))
            {
                throw new ArgumentException("CNPJ e Competência são necessários para a busca.");
            }

            IWorkbook planilha = CarregarArquivoExcel(caminhoArquivoExcel);

            var aba = planilha.GetSheetAt(0) ?? throw new Exception("Não foi possível acessar a planilha no arquivo Excel.");

            int indiceColunaCnpj = ObterIndiceColunaPorTitulo(aba, "CNPJ");
            int indiceColunaCompetencia = ObterIndiceColunaPorTitulo(aba, "Competência");
            int indiceColunaSalario = ObterIndiceColunaPorTitulo(aba, "Salário");
            int indiceColunaRazaoSocial = ObterIndiceColunaPorTitulo(aba, "Razão Social");

            for (int indiceLinha = 1; indiceLinha <= aba.LastRowNum; indiceLinha++)
            {
                var linha = aba.GetRow(indiceLinha);
                if (linha == null) continue;

                string? cnpj = linha.GetCell(indiceColunaCnpj)?.ToString()?.Trim();
                string? competencia = linha.GetCell(indiceColunaCompetencia)?.ToString()?.Trim();
                string? salario = FormatarSalario(linha.GetCell(indiceColunaSalario)?.ToString());
                string? razaoSocial = linha.GetCell(indiceColunaRazaoSocial)?.ToString()?.Trim();

                if (!string.IsNullOrEmpty(cnpj) && !string.IsNullOrEmpty(competencia) &&
                    !string.IsNullOrEmpty(salario) && !string.IsNullOrEmpty(razaoSocial) &&
                    VerificarCorrespondencia(cnpj, competencia, salario, razaoSocial, dadosNfse))
                {
                    Console.WriteLine("Tudo certo. Todos os dados necessários foram validados.");
                    return true;
                }
            }

            Console.WriteLine("Dados não encontrados no Excel para o CNPJ e Competência especificados.");
            return false;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro ao validar NFSe: {ex.Message}");
            return false;
        }
    }

    private static IWorkbook CarregarArquivoExcel(string caminhoArquivoExcel)
    {
        using (var fluxoArquivo = new FileStream(caminhoArquivoExcel, FileMode.Open, FileAccess.Read))
        {
            return caminhoArquivoExcel.EndsWith(".xls") ? (IWorkbook)new HSSFWorkbook(fluxoArquivo) : new XSSFWorkbook(fluxoArquivo);
        }
    }

    private static bool VerificarCorrespondencia(string cnpj, string competencia, string salario, string razaoSocial, NFSe dadosNfse)
    {
        string? valorServicoFormatado = FormatarSalario(dadosNfse.ValorServico);
        string? dataCompetenciaFormatada = FormatarDataCompetencia(dadosNfse.DataCompetencia);

        cnpj = cnpj?.Trim();
        competencia = competencia?.Trim();
        salario = salario?.Trim();
        razaoSocial = razaoSocial?.Trim();
        valorServicoFormatado = valorServicoFormatado?.Trim();
        dataCompetenciaFormatada = dataCompetenciaFormatada?.Trim();

        Console.WriteLine($"CNPJ no Excel: {cnpj}, CNPJ Procurado: {dadosNfse.Prestador.Cnpj.Trim()}");
        Console.WriteLine($"Competência no Excel: {competencia}, Competência Procurada: {dataCompetenciaFormatada}");
        Console.WriteLine($"Salário no Excel: {salario}, Salário Procurado: {valorServicoFormatado}");
        Console.WriteLine($"Razão Social no Excel: {razaoSocial}, Razão Social Procurada: {dadosNfse.Prestador.RazaoSocial.Trim()}");

        bool correspondeCnpj = cnpj == dadosNfse.Prestador.Cnpj.Trim();
        bool correspondeCompetencia = competencia?.Equals(dataCompetenciaFormatada, StringComparison.OrdinalIgnoreCase) ?? false;
        bool correspondeSalario = salario == valorServicoFormatado;
        bool correspondeRazaoSocial = razaoSocial == dadosNfse.Prestador.RazaoSocial.Trim();

        return correspondeCnpj && correspondeCompetencia && correspondeSalario && correspondeRazaoSocial;
    }

    private static int ObterIndiceColunaPorTitulo(ISheet aba, string titulo)
    {
        IRow linhaCabecalho = aba.GetRow(0);
        if (linhaCabecalho == null)
        {
            throw new Exception("Cabeçalho não encontrado na planilha.");
        }

        string tituloNormalizado = NormalizarString(titulo);
        for (int i = 0; i < linhaCabecalho.LastCellNum; i++)
        {
            var celula = linhaCabecalho.GetCell(i);
            if (celula != null)
            {
                string valorCelula = NormalizarString(celula.StringCellValue);
                if (valorCelula.Equals(tituloNormalizado, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }
        }

        throw new Exception($"Título '{titulo}' não encontrado na planilha.");
    }

    private static string NormalizarString(string entrada)
    {
        if (entrada == null) return string.Empty;

        return new string(entrada
            .Where(c => !char.IsWhiteSpace(c) && !char.IsControl(c))
            .ToArray());
    }

    private static string FormatarDataCompetencia(string dataCompetencia)
    {
        try
        {
            return DateTime.ParseExact(dataCompetencia, "dd/MM/yyyy", CultureInfo.InvariantCulture)
                            .ToString("dd-MMM-yyyy", new CultureInfo("en-US"));
        }
        catch (FormatException ex)
        {
            Console.WriteLine($"Erro ao formatar a data de competência: {ex.Message}");
            return string.Empty;
        }
    }

    private static string FormatarSalario(string salario)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(salario) || salario.StartsWith("base.folha"))
            {
                return string.Empty;
            }
            if (decimal.TryParse(salario, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal salarioParseado))
            {
                return salarioParseado.ToString("N2", new CultureInfo("pt-BR"));
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro ao converter o valor do salário para decimal: {ex.Message}");
        }
        return salario;
    }
}
