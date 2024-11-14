using FinanceiroScript.Dominio.Interfaces.Helpers;
using System;
using System.IO;

public class LogHelper
{
    private readonly string diretorioLogs;
    private readonly IDiretorioHelper _directoryHelper;

    public LogHelper(IDiretorioHelper directoryHelper)
    {
        _directoryHelper = directoryHelper;
        diretorioLogs = _directoryHelper.GetResultDirectory();
    }

    public void LogarMensagem(string mensagem, string nivelLog = "INFO", string? infoAdicional = null)
    {
        string caminhoArquivoLog = Path.Combine(diretorioLogs, "Log.txt");

        using (StreamWriter writer = new StreamWriter(caminhoArquivoLog, append: true))
        {
            writer.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{nivelLog}] {mensagem} {infoAdicional ?? string.Empty}");
        }
    }

    public void LogarErro(Exception ex, string contexto)
    {
        string caminhoArquivoLog = Path.Combine(diretorioLogs, "Log.txt");
        using (StreamWriter writer = new StreamWriter(caminhoArquivoLog, append: true))
        {
            writer.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [ERROR] Erro no contexto: {contexto}");
            writer.WriteLine($"Mensagem: {ex.Message}");
            writer.WriteLine($"StackTrace: {ex.StackTrace}");
        }
    }
}
