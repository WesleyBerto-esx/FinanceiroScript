using System.Text;
using System.Text.RegularExpressions;
using FinanceiroScript.Dominio;
using FinanceiroScript.Dominio.Interfaces.Servicos;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using Microsoft.Extensions.Logging;

namespace FinanceiroScript.Servicos
{
    public class NFSeServico : INFSeServico
    {
        private readonly ILogger<NFSeServico> _logger;

        public NFSeServico(ILogger<NFSeServico> logger)
        {
            _logger = logger;
        }

        public string[]? ObterTodasNFSes(string caminhoPDFsNFSe)
        {
            _logger.LogInformation("Função listar todas as NFSes em PDF");

            if (!Directory.Exists(caminhoPDFsNFSe))
            {
                _logger.LogError($"O diretório '{caminhoPDFsNFSe}' não foi encontrado.");
                return null;
            }

            string[] arquivosPdf = Directory.GetFiles(caminhoPDFsNFSe, "*.pdf");

            if (arquivosPdf.Length == 0)
            {
                _logger.LogWarning($"Nenhum arquivo PDF encontrado no diretório '{caminhoPDFsNFSe}'.");
                return null;
            }

            return arquivosPdf;
        }

        public NFSe ExtrairDadosNFSeDoPdf(Stream pdfStream)
        {
            var dadosNfse = new NFSe();
            string textoPdf = ObterTextoDoPdfStream(pdfStream);

            ExtrairCamposNFSe(textoPdf, dadosNfse);
            ExtrairDadosPessoaJuridica(textoPdf, dadosNfse.Prestador, "Prestador");
            ExtrairDadosPessoaJuridica(textoPdf, dadosNfse.Tomador, "Tomador");

            return dadosNfse;
        }

        private void ExtrairCamposNFSe(string textoPdf, NFSe dadosNfse)
        {
            var chavesCampos = new List<string>
            {
                "ChaveAcesso", "Numero", "DataCompetencia", "DataEmissao",
                "CodigoServico", "DescricaoServico", "StatusImpostoMunicipal",
                "IncidenciaMunicipal", "ValorServico", "ValorLiquidoNotaFiscal"
            };

            foreach (var chaveCampo in chavesCampos)
            {
                if (_mapeamentoCamposNfse.TryGetValue(chaveCampo, out var configuracaoCampo))
                {
                    string padrao = configuracaoCampo["padrao"];
                    string rotulo = configuracaoCampo["rotulo"];

                    var valorCampo = ExtrairCampoDoTexto(textoPdf, "", new List<string> { rotulo }, padrao);
                    dadosNfse.GetType().GetProperty(chaveCampo)?.SetValue(dadosNfse, valorCampo);
                }
            }
        }

        private void ExtrairDadosPessoaJuridica(string textoPdf, PessoaJuridica pessoa, string prefixoEntidade)
        {
            var chavesCampos = new List<string> { "Cnpj", "RazaoSocial", "Email", "Endereco", "Municipio", "Cep" };

            foreach (var chaveCampo in chavesCampos)
            {
                if (_mapeamentoCamposNfse.TryGetValue(chaveCampo, out var configuracaoCampo))
                {
                    string padrao = configuracaoCampo["padrao"];
                    string rotulo = configuracaoCampo["rotulo"];
                    var valorCampo = ExtrairCampoDoTexto(textoPdf, prefixoEntidade, new List<string> { rotulo }, padrao);
                    pessoa.GetType().GetProperty(chaveCampo)?.SetValue(pessoa, valorCampo);
                }
            }
        }

        private string? ExtrairCampoDoTexto(string texto, string prefixoEntidade, List<string> rotulos, string padrao)
        {
            if (!string.IsNullOrEmpty(prefixoEntidade))
            {
                prefixoEntidade = $@"(?:{prefixoEntidade}[\s\S]*?)";
            }

            foreach (var rotulo in rotulos)
            {
                var padraoCompleto = $@"(?i){prefixoEntidade}(?:{rotulo}\s*[:\-]?\s*)[\s\S]*?{padrao}";

                Regex regex = new Regex(padraoCompleto, RegexOptions.IgnoreCase | RegexOptions.Singleline);

                var correspondencia = regex.Match(texto);

                if (correspondencia.Success)
                {
                    return correspondencia.Groups[1].Value.Trim();
                }
            }

            return null;
        }

        private string ObterTextoDoPdfStream(Stream pdfStream)
        {
            var resultado = new StringBuilder();
            using var leitorPdf = new PdfReader(pdfStream);
            using var docPdf = new PdfDocument(leitorPdf);
            for (int pagina = 1; pagina <= docPdf.GetNumberOfPages(); pagina++)
            {
                var estrategia = new SimpleTextExtractionStrategy();
                string conteudo = PdfTextExtractor.GetTextFromPage(docPdf.GetPage(pagina), estrategia);
                resultado.Append(conteudo);
            }
            return resultado.ToString();
        }

        public string RenomearEMoverNFSePdf(string caminhoArquivo, NFSe nfse, string diretórioDestino)
        {
            string novoNomeArquivo = $"erro_" + Path.GetFileName(caminhoArquivo);

            if (!string.IsNullOrEmpty(nfse?.Numero) && !string.IsNullOrEmpty(nfse?.Prestador?.RazaoSocial))
            {
                string numeroFormatado = nfse.Numero.PadLeft(4, '0');

                string razaoSocialFormatada = Regex.Replace(nfse.Prestador.RazaoSocial, @"[^a-zA-Z\s]", "");
                razaoSocialFormatada = Regex.Replace(razaoSocialFormatada.Trim().ToUpper(), @"\s+", "_");

                novoNomeArquivo = $"{numeroFormatado}_{razaoSocialFormatada}.pdf";
            }
            else
            {
                Console.WriteLine("Aviso: 'Numero' ou 'RazaoSocial' é nulo ou vazio. Usando o nome original do arquivo.");
            }

            string caminhoArquivoCopia = Path.Combine(Path.GetDirectoryName(caminhoArquivo), novoNomeArquivo);
            File.Copy(caminhoArquivo, caminhoArquivoCopia, overwrite: true);

            string novoCaminhoArquivo = Path.Combine(diretórioDestino, Path.GetFileName(caminhoArquivoCopia));
            File.Move(caminhoArquivoCopia, novoCaminhoArquivo);
            return novoCaminhoArquivo;
        }

        private readonly Dictionary<string, Dictionary<string, string>> _mapeamentoCamposNfse = new()
        {
            { "ChaveAcesso", new Dictionary<string, string>
                {
                    { "rotulo", "Chave de Acesso da NFS-e" },
                    { "padrao", @"([\d]+)" }
                }
            },
            { "Razão Social", new Dictionary<string, string>
                {
                    { "rotulo", "Razão Social da NFS-e" },
                    { "padrao", @"([\d]+)" }
                }
            },
            { "Numero", new Dictionary<string, string>
                {
                    { "rotulo", "Número da NFS-e" },
                    { "padrao", @"(\d+)" }
                }
            },
            { "DataCompetencia", new Dictionary<string, string>
                {
                    { "rotulo", "Competência da NFS-e" },
                    { "padrao", @"([\d/]+)" }
                }
            },
            { "DataEmissao", new Dictionary<string, string>
                {
                    { "rotulo", "Data e Hora da emissão" },
                    { "padrao", @"([\d/]+ \d{2}:\d{2}:\d{2})" }
                }
            },
            { "CodigoServico", new Dictionary<string, string>
                {
                    { "rotulo", "Código de Tributação Nacional" },
                    { "padrao", @"([\d.]+)" }
                }
            },
            { "DescricaoServico", new Dictionary<string, string>
                {
                    { "rotulo", "Descrição do Serviço" },
                    { "padrao", @"([^\n]+)" }
                }
            },
            { "StatusImpostoMunicipal", new Dictionary<string, string>
                {
                    { "rotulo", "Tributação do ISSQN" },
                    { "padrao", @"([^\n]+)" }
                }
            },
            { "IncidenciaMunicipal", new Dictionary<string, string>
                {
                    { "rotulo", "Município de Incidência do ISSQN" },
                    { "padrao", @"([^\n]+)" }
                }
            },
            { "ValorServico", new Dictionary<string, string>
                {
                    { "rotulo", "Valor do Serviço" },
                    { "padrao", @"R\$\s*([\d,\.]+)" }
                }
            },
            { "ValorLiquidoNotaFiscal", new Dictionary<string, string>
                {
                    { "rotulo", "Valor Líquido da NFS-e" },
                    { "padrao", @"R\$\s*([\d.,]+)" }
                }
            },
            { "Cnpj", new Dictionary<string, string>
                {
                    { "rotulo", @"CNPJ" },
                    { "padrao", @"(\d{2}\.?(\d{3}\.?){2}([/])?\d{4}-?\d{2})" }
                }
            },
            { "RazaoSocial", new Dictionary<string, string>
                {
                    { "rotulo", @"(?:Nome\s*?:)" },
                    { "padrao", @"(.+)" }
                }
            },
            { "Email", new Dictionary<string, string>
                {
                    { "rotulo", @"E-mail" },
                    { "padrao", @"(?:[\w\.-]+@[\w\.-]+)" }
                }
            },
            { "Endereco", new Dictionary<string, string>
                {
                    { "rotulo", "Endereço" },
                    { "padrao", @"([^\n]+)" }
                }
            },
            { "Municipio", new Dictionary<string, string>
                {
                    { "rotulo", "Município" },
                    { "padrao", @"([^\n]+)" }
                }
            },
            { "Cep", new Dictionary<string, string>
                {
                    { "rotulo", "CEP" },
                    { "padrao", @"(\d{5}-?\d{3})" }
                }
            }
        };
    }
}
