using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using MimeKit;
using Serilog;
using System.Xml;
using System.Xml.Linq;

public class EmailDownloaderService : IHostedService, IDisposable
{
    private readonly IConfiguration _configuration;
    private readonly ILogger _logger;
    private Timer? _timer;

    public EmailDownloaderService(IConfiguration configuration, ILogger logger)
    {
        _configuration = configuration;
        _logger = logger;
    }

    public Task StartAsync(CancellationToken cancellationToken)
    {
        int intervaloEmHoras = _configuration.GetValue<int>("ConfiguracoesWorker:IntervaloEmHoras");
        _timer = new Timer(DoWork, null, TimeSpan.Zero, TimeSpan.FromHours(intervaloEmHoras));
        return Task.CompletedTask;
    }

    private async void DoWork(object? state)
    {
        var configuracoesEmail = _configuration.GetSection("ConfiguracoesEmail");
        var pastaRaiz = _configuration.GetValue<string>("ConfiguracoesDownload:PastaRaiz") ?? string.Empty;
        var pastaTemporaria = Path.Combine(pastaRaiz, _configuration.GetValue<string>("ConfiguracoesDownload:PastaArquivosTemporarios") ?? string.Empty);
        var pastaInvalida = Path.Combine(pastaRaiz, _configuration.GetValue<string>("ConfiguracoesDownload:PastaInvalida") ?? string.Empty);
        var filtrosPesquisa = _configuration.GetSection("FiltrosPesquisa:AssuntoContem").Get<string[]>();

        try
        {
            // Cria a pasta para arquivos inválidos caso não exista
            if (!Directory.Exists(pastaInvalida))
            {
                Directory.CreateDirectory(pastaInvalida);
                _logger.Information("Pasta para arquivos inválidos criada em {PastaInvalida}.", pastaInvalida);
            }

            // Cria a pasta para arquivos temporários caso não exista
            if (!Directory.Exists(pastaTemporaria))
            {
                Directory.CreateDirectory(pastaTemporaria);
                _logger.Information("Pasta para arquivos temporários criada em {PastaTemporaria}.", pastaTemporaria);
            }

            using (var client = new ImapClient())
            {
                // Conecta ao servidor de email
                await client.ConnectAsync(configuracoesEmail["Host"], configuracoesEmail.GetValue<int>("Port"), true);
                await client.AuthenticateAsync(configuracoesEmail["Username"], configuracoesEmail["Password"]);

                // Abre a caixa de entrada
                var inbox = client.Inbox;
                await inbox.OpenAsync(FolderAccess.ReadWrite);

                // Constrói a consulta de pesquisa dinamicamente
                var pesquisa = SearchQuery.NotSeen;
                if (filtrosPesquisa != null)
                {
                    foreach (var filtro in filtrosPesquisa)
                    {
                        pesquisa = pesquisa.Or(SearchQuery.SubjectContains(filtro));
                    }
                }

                var resultados = await inbox.SearchAsync(pesquisa);

                if (resultados.Count == 0)
                {
                    _logger.Information("Nenhum email novo encontrado.");
                }

                // Processa os emails encontrados
                foreach (var idUnico in resultados)
                {
                    var email = await inbox.GetMessageAsync(idUnico);
                    bool temAnexoValido = false;

                    // Processa os anexos do email
                    foreach (var anexo in email.Attachments)
                    {
                        //Verificando se o anexo é XML
                        if (anexo is MimePart part && part.FileName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                        {
                            // Salvar o anexo em uma pasta temporária
                            var caminhoPastaTemporaria = Path.Combine(pastaTemporaria, part.FileName);
                            using (var stream = File.Create(caminhoPastaTemporaria))
                            {
                                await part.Content.DecodeToAsync(stream);
                                _logger.Information("Arquivo XML {FileName} baixado com sucesso para {TempFilePath}.", part.FileName, caminhoPastaTemporaria);
                            }

                            // Ler o XML para obter o CNPJ
                            var cnpj = GetCNPJFromXml(caminhoPastaTemporaria);
                            var pastaDestino = string.IsNullOrEmpty(cnpj) ? pastaInvalida : Path.Combine(pastaRaiz, cnpj);

                            // Cria a pasta de destino caso não exista
                            if (!Directory.Exists(pastaDestino))
                            {
                                Directory.CreateDirectory(pastaDestino);
                                _logger.Information("Pasta para o CNPJ {CNPJ} criada em {pastaDestino}.", cnpj, pastaDestino);
                            }

                            // Mover o arquivo XML para a pasta de destino
                            var caminhoPastaDestino = Path.Combine(pastaDestino, part.FileName);
                            File.Move(caminhoPastaTemporaria, caminhoPastaDestino);
                            _logger.Information("Arquivo XML {FileName} movido para {DestinationFilePath}.", part.FileName, caminhoPastaDestino);

                            temAnexoValido = true;
                        }
                    }

                    if (!temAnexoValido)
                    {
                        _logger.Information("Nenhum anexo válido encontrado no email com ID {UniqueId}.", idUnico);
                    }

                    // Marca o email como lido
                    await inbox.AddFlagsAsync(idUnico, MessageFlags.Seen, true);
                    _logger.Information("Email com ID {UniqueId} marcado como lido.", idUnico);
                }

                await client.DisconnectAsync(true);
            }
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Erro ao processar emails.");
        }
    }

    private string GetCNPJFromXml(string filePath)
    {
        try
        {
            // Lê o XML e obtém o CNPJ do emitente
            var xml = XDocument.Load(filePath);
            var cnpjElement = xml.Descendants("emit").Elements("CNPJ").FirstOrDefault();
            if (cnpjElement != null)
            {
                return cnpjElement.Value;
            }
            else
            {
                return string.Empty; // Retorna uma string vazia em vez de null
            }
        }
        catch (XmlException ex)
        {
            _logger.Error(ex, "Erro ao ler o XML para obter o CNPJ: XML malformado.");
            return string.Empty;
        }
        catch (FileNotFoundException ex)
        {
            _logger.Error(ex, "Erro ao ler o XML para obter o CNPJ: Arquivo não encontrado.");
            return string.Empty;
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Erro ao ler o XML para obter o CNPJ.");
            return string.Empty;
        }
    }

    public Task StopAsync(CancellationToken cancellationToken)
    {
        _timer?.Change(Timeout.Infinite, 0);
        return Task.CompletedTask;
    }

    public void Dispose()
    {
        _timer?.Dispose();
    }
}