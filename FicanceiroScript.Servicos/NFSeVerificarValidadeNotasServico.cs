using FinanceiroScript.Dominio;
using FinanceiroScript.Dominio.Interfaces.Helpers;
using FinanceiroScript.Dominio.Interfaces.Servicos;

public class NFSeVerificarValidadeNotasServico : INFSeVerificarValidadeNotasServico
{
    private readonly INFSeServico _nfseServico;
    private readonly IDiretorioHelper _directoryHelper;
    private readonly LogHelper logHelper;

    public NFSeVerificarValidadeNotasServico(INFSeServico nfseServico, IDiretorioHelper directoryHelper, LogHelper logHelper)
    {
        _nfseServico = nfseServico;
        _directoryHelper = directoryHelper;
        this.logHelper = logHelper;
    }

    public void VerificarValidadeNotasFiscais()
    {
        string caminhoRaizApp = _directoryHelper.GetAppRootPath();
        string caminhoNotasDir = Path.Combine(caminhoRaizApp, "Notas");
        string caminhoNotasValidasDir = _directoryHelper.GetValidosDirectory();
        string caminhoNotasErrosDir = _directoryHelper.GetErrosDirectory();
        string[] arquivosPdf = _nfseServico.ObterTodasNFSes(caminhoNotasDir);
        string caminhoArquivoExcel = Path.Combine(caminhoRaizApp, "Excel", "PlanilhaFinanceiro.xlsx");

        if (arquivosPdf == null || arquivosPdf.Length < 1)
        {
            Console.WriteLine("Não foram encontrados arquivos PDF.");
            return;
        }

        foreach (string arquivoPdf in arquivosPdf)
        {
            try
            {
                using var fluxoPdf = new FileStream(arquivoPdf, FileMode.Open, FileAccess.Read);
                NFSe dadosNFSe = _nfseServico.ExtrairDadosNFSeDoPdf(fluxoPdf);

                bool isValid = ExcelHelper.ValidarNFSe(dadosNFSe, caminhoArquivoExcel);
                string diretorioDestino = isValid ? caminhoNotasValidasDir : caminhoNotasErrosDir;

                string novoCaminhoArquivo = _nfseServico.RenomearEMoverNFSePdf(arquivoPdf, dadosNFSe, diretorioDestino);

                string status = isValid ? "válida" : "inválida";

                string motivoInvalido = isValid ? string.Empty : ObterMotivoInvalido(dadosNFSe, caminhoArquivoExcel);

                string mensagemLog = $"NFSe {status}: {Path.GetFileName(novoCaminhoArquivo)}";
                if (!string.IsNullOrEmpty(motivoInvalido))
                {
                    mensagemLog += $" - Motivo: {motivoInvalido}";
                }

                logHelper.LogMessage(mensagemLog);

                Console.WriteLine(mensagemLog);
            }
            catch (IOException ioEx)
            {
                string mensagemErro = $"Erro ao acessar o arquivo '{Path.GetFileName(arquivoPdf)}': {ioEx.Message}";
                logHelper.LogError(ioEx, mensagemErro);
                Console.Error.WriteLine(mensagemErro);
            }
            catch (Exception ex)
            {
                string mensagemErro = $"Erro ao processar o arquivo '{Path.GetFileName(arquivoPdf)}': {ex.Message}";
                logHelper.LogError(ex, mensagemErro);
                Console.Error.WriteLine(mensagemErro);
            }
        }
    }

    private string ObterMotivoInvalido(NFSe dadosNFSe, string caminhoArquivoExcel)
    {
        if (string.IsNullOrEmpty(dadosNFSe.Prestador.Cnpj))
        {
            return "CNPJ vazio.";
        }

        if (!ExcelHelper.ValidarNFSe(dadosNFSe, caminhoArquivoExcel))
        {
            return "Os dados da NFSe não correspondem às informações da planilha de validação.";
        }

        return "Motivo desconhecido.";
    }
}
