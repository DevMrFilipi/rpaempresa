using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Globalization;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using iText.Kernel.Pdf;
using iText.Forms;
using iText.Forms.Fields;
using iText.Kernel.Font;
using iText.IO.Font.Constants;
using iText.IO.Font;
using iText.Kernel.Colors;
using iText.Kernel.Pdf.Annot;

class Program
{
    public static List<string> Vetor { get; set; } = new();

    public static class Logger
    {
        private static StreamWriter? _writer;
        private static readonly object _lock = new();

        public static void Init(string logFilePath)
        {
            _writer = new StreamWriter(logFilePath, append: true);
            _writer.AutoFlush = true;
        }

        public static void Log(string message)
        {
            string text = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";
            lock (_lock)
            {
                Console.WriteLine(text);
                if (_writer != null)
                    _writer.WriteLine(text);
            }
        }

        public static void Close()
        {
            lock (_lock)
            {
                _writer?.Flush();
                _writer?.Close();
                _writer = null;
            }
        }
    }

    static void Main(string[] args)
    {
        string logFile = @"C:\\Users\\PdfWriterApp\\app_log_" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt";
        Logger.Init(logFile);
        Logger.Log("Iniciando execução do PDF Writer App");

        try
        {
            Vetor = args.ToList();
            Logger.Log($"Argumentos recebidos: {string.Join(" | ", Vetor)}");

            if (Vetor.Count < 11)
            {
                Logger.Log("Erro: número insuficiente de argumentos. São necessários pelo menos 11.");
                return;
            }

            // Variáveis que serão montadas a partir dos argumentos
            string dataMonitorUm = "";
            string dataMonitorDois = "";
            string dataEquipamento = "";
            string dataDockstation = "";
            List<string> equipamentos = new List<string>{};
            List<string> perifericos = GetListArg(12);

            if (!string.IsNullOrWhiteSpace(GetListArg(7).ElementAtOrDefault(0)) && !string.IsNullOrWhiteSpace(GetListArg(8).ElementAtOrDefault(0)))
            {
                dataMonitorUm = GetListArg(7).ElementAtOrDefault(0) + " / " + GetListArg(8).ElementAtOrDefault(0);
            }

            if (!string.IsNullOrWhiteSpace(GetListArg(7).ElementAtOrDefault(1)) && !string.IsNullOrWhiteSpace(GetListArg(8).ElementAtOrDefault(1)))
            {
                dataMonitorDois = GetListArg(7).ElementAtOrDefault(1) + " / " + GetListArg(8).ElementAtOrDefault(1);
            }

            if (!string.IsNullOrWhiteSpace(GetArg(1)) && !string.IsNullOrWhiteSpace(GetArg(2)))
            {
                dataEquipamento = GetArg(1) + " / " + GetArg(2);
            }

            if (!string.IsNullOrWhiteSpace(GetArg(0)) && GetArg(0) == "NOTEBOOK") 
            {
                equipamentos.Add("NOTEBOOK");
            }

            if (!string.IsNullOrWhiteSpace(GetArg(0)) && GetArg(0) == "DESKTOP") 
            {
                equipamentos.Add("DESKTOP");
            }

            if (!string.IsNullOrWhiteSpace(GetArg(11)))
            {
                string[] tempDock = GetArg(11).Split(new string[] { " / " }, StringSplitOptions.None);

                if (tempDock.Length >= 2)
                {
                    dataDockstation = tempDock[0] + "-" + tempDock[1];
                }
                else
                {
                    dataDockstation = tempDock[0]; // ou outro tratamento padrão
                }

                equipamentos.Add("DOCKSTATION");
            }

            Logger.Log($"Valor do Equipamento: {dataEquipamento}");
            Logger.Log($"Valor do Monitor 1: {dataMonitorUm} | Valor do Array: {GetListArg(7).ElementAtOrDefault(0)} e {GetListArg(7).ElementAtOrDefault(1)}");
            Logger.Log($"Valor do Monitor 2: {dataMonitorDois} | Valor do Array: {GetListArg(8).ElementAtOrDefault(0)} e {GetListArg(8).ElementAtOrDefault(1)}");
            Logger.Log($"Valor do Dockstation: {dataDockstation} | Valor do Array: {GetArg(11)}");
            // Criação do dicionário de variáveis
           var variaveis = new Dictionary<string, object>
            {
                { "movimentacao", GetArg(4) },
                { "localidade", GetArg(5) },
                { "tipo_equipamento", GetArg(0) },
                { "data_equipamento", dataEquipamento },
                { "data_monitor_um", dataMonitorUm },
                { "data_monitor_dois", dataMonitorDois },
                { "patrimonio_monitor_um", GetListArg(9).ElementAtOrDefault(0) ?? "" },
                { "patrimonio_monitor_dois", GetListArg(9).ElementAtOrDefault(1) ?? "" },
                { "patrimonio_computador", GetArg(10) },
                { "data_dockstation", GetArg(11) },
                { "perifericos", GetListArg(12) },
                { "usuario", GetArg(3) },
                { "userAdmin", GetArg(6) },
                { "day", DateTime.Now.Day.ToString("D2") },
                { "month", DateTime.Now.ToString("MMMM", new CultureInfo("pt-BR")) },
                { "year", DateTime.Now.Year.ToString() }
            };


            Logger.Log("Variáveis definidas:");
            foreach (var v in variaveis)
            {
                Logger.Log($"{v.Key} = '{v.Value}'");
            }

            string exeDir = AppDomain.CurrentDomain.BaseDirectory;
            string jsonPath = Path.Combine(exeDir, "map.json");

            Logger.Log($"Lendo JSON de: {jsonPath}");

            if (!File.Exists(jsonPath))
            {
                Logger.Log($"Erro: Arquivo não encontrado: {jsonPath}");
                return;
            }

            string jsonMapping = File.ReadAllText(jsonPath);
            Logger.Log("Arquivo JSON lido com sucesso.");

            if (string.IsNullOrWhiteSpace(jsonMapping))
            {
                Logger.Log("Erro: O arquivo map.json está vazio!");
                return;
            }

            // Substitui todas as variáveis do JSON bruto
            string jsonComVariaveis = ReplaceVariables(jsonMapping, variaveis);

            Logger.Log("JSON final após substituição:");
            Logger.Log(jsonComVariaveis);

            // Validação do JSON final substituído
            Dictionary<string, object>? mapping = null;
            try
            {
                mapping = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonComVariaveis);
                Logger.Log("JSON com variáveis substituídas:");
                Logger.Log(JObject.Parse(jsonComVariaveis).ToString());
            }
            catch (Exception ex)
            {
                Logger.Log("Erro ao desserializar JSON após substituição:");
                Logger.Log(ex.Message);
                return;
            }

            if (mapping == null)
            {
                Logger.Log("Erro: Falha ao interpretar o map.json após substituição.");
                return;
            }

            string inputPdfPath = @"C:\Users\PdfWriterApp\TERMO.USUARIO_MODEL_TAG_MODEL_TAG_CIDADE.pdf";
            if (!File.Exists(inputPdfPath))
            {
                Logger.Log($"Erro: PDF de entrada não encontrado em: {inputPdfPath}");
                return;
            }

            string outputDir = @"C:\Users\filipi.serpa\Downloads";
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            string outputFile = "";

            if (string.IsNullOrWhiteSpace(dataMonitorUm))
            {
                outputFile = Path.Combine(outputDir,
                $"{variaveis["usuario"]} {GetArg(1)}-{GetArg(2)} {variaveis["localidade"]}.pdf");
            } else if (string.IsNullOrWhiteSpace(dataMonitorDois) && string.IsNullOrWhiteSpace(dataDockstation))
            {
                outputFile = Path.Combine(outputDir,
                $"{variaveis["usuario"]} {GetArg(1)}-{GetArg(2)} {GetListArg(7).ElementAtOrDefault(0)}-{GetListArg(8).ElementAtOrDefault(0)} {GetArg(5)}.pdf");
            } else if (string.IsNullOrWhiteSpace(dataDockstation))
            {
                outputFile = Path.Combine(outputDir,
                $"{variaveis["usuario"]} {GetArg(1)}-{GetArg(2)} {GetListArg(7).ElementAtOrDefault(0)}-{GetListArg(8).ElementAtOrDefault(0)} {GetListArg(7).ElementAtOrDefault(1)}-{GetListArg(8).ElementAtOrDefault(1)} {GetArg(5)}.pdf");
            } else if (string.IsNullOrWhiteSpace(dataMonitorDois) && !string.IsNullOrWhiteSpace(dataDockstation))
            {
                outputFile = Path.Combine(outputDir,
                $"{variaveis["usuario"]} {GetArg(1)}-{GetArg(2)} {GetListArg(7).ElementAtOrDefault(0)}-{GetListArg(8).ElementAtOrDefault(0)} {dataDockstation} {GetArg(5)}.pdf");
            }
            else
            {
                outputFile = Path.Combine(outputDir,
                $"{variaveis["usuario"]} {GetArg(1)}-{GetArg(2)} {GetListArg(7).ElementAtOrDefault(0)}-{GetListArg(8).ElementAtOrDefault(0)} {GetListArg(7).ElementAtOrDefault(1)}-{GetListArg(8).ElementAtOrDefault(1)} {dataDockstation} {GetArg(5)}.pdf");
            }

            Logger.Log($"Gerando PDF em: {outputFile}");

            // Caminho temporário local (pasta do sistema)
            string tempPath = Path.Combine(Path.GetTempPath(), Path.GetFileName(outputFile));

            using (var reader = new PdfReader(inputPdfPath))
            using (var writer = new PdfWriter(tempPath)) // sem criptografia
            using (var pdfDoc = new PdfDocument(reader, writer, new StampingProperties().UseAppendMode()))
            {
                var form = PdfAcroForm.GetAcroForm(pdfDoc, true);

                // --- Log dos campos sem usar GetFormFields (compatível com sua versão do iText) ---
                try
                {
                    var acroDict = (PdfDictionary)form.GetPdfObject();
                    var fieldsArray = acroDict?.GetAsArray(PdfName.Fields);
                    if (fieldsArray != null)
                    {
                        Logger.Log("Campos encontrados no PDF:");
                        for (int i = 0; i < fieldsArray.Size(); i++)
                        {
                            var fieldDict = fieldsArray.GetAsDictionary(i);
                            var name = fieldDict?.GetAsString(PdfName.T)?.ToString();
                            if (!string.IsNullOrEmpty(name))
                                Logger.Log($"Campo: {name}");
                        }
                    }
                    else
                    {
                        Logger.Log("Nenhum campo de formulário encontrado (PdfName.Fields ausente).");
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log("Falha ao enumerar campos de formulário via dicionário: " + ex.Message);
                }
                // ------------------------------------------------------------------------------

                // PreencherComNomesDosCampos(form);

                string arialPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arial.ttf");
                PdfFont fonteArial = PdfFontFactory.CreateFont(arialPath, PdfEncodings.WINANSI, pdfDoc);
                float tamanhoFonte = 9f;

                string valorMovimentacao = variaveis["movimentacao"]?.ToString() ?? "";
                string valorLocalidade = variaveis["localidade"]?.ToString() ?? "";
                
                PreencherCampo(mapping, "movimentacao", variaveis, form, valorMovimentacao, fonteArial, tamanhoFonte);
                PreencherCampoContains(mapping, "localidade", variaveis, form, valorLocalidade, fonteArial, tamanhoFonte);

                foreach (var tipo in equipamentos)
                {
                    PreencherCampo(mapping, "tipo_equipamento", variaveis, form, tipo, fonteArial, tamanhoFonte);
                }

                if (variaveis.TryGetValue("perifericos", out var obj) && obj is List<string> perifericosLista)
                {
                    Logger.Log("Periféricos selecionados:");

                    foreach (string prf in perifericosLista)
                    {
                        Logger.Log($"Item: {prf}");
                        PreencherCampo(mapping, "perifericos", variaveis, form, prf, fonteArial, tamanhoFonte);
                    }
                }
                else
                {
                    Logger.Log("Nenhum periférico encontrado nas variáveis.");
                }


                if (!string.IsNullOrWhiteSpace(variaveis["data_monitor_um"]?.ToString()) || !string.IsNullOrWhiteSpace(variaveis["data_monitor_dois"]?.ToString()))
                {
                    PreencherMonitores(mapping, variaveis, form, fonteArial, tamanhoFonte);
                }

                if (mapping.ContainsKey("data_pdf"))
                    SetField(form, mapping["data_pdf"], variaveis, fonteArial, tamanhoFonte);

                if (mapping.ContainsKey("assinatura"))
                    SetField(form, mapping["assinatura"], variaveis, fonteArial, tamanhoFonte);
            }

            // Garante que nenhum handle do arquivo esteja aberto antes de copiar/mover
            GC.Collect();
            GC.WaitForPendingFinalizers();

            string outputFilePath = outputFile;

            // Novo: garante que a pasta existe
            string? outputFolder = Path.GetDirectoryName(outputFilePath);
            if (!string.IsNullOrEmpty(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }
            
            File.Copy(tempPath, outputFilePath, overwrite: true);
            try { File.Delete(tempPath); } catch { /* ignore */ }

            Logger.Log("PDF gerado com sucesso!");
        }
        catch (Exception ex)
        {
            Logger.Log("❌ Ocorreu um erro ao gerar o PDF:");
            Logger.Log(" " + ex + " ");
        }
        finally
        {
            Logger.Log("Finalizando execução.");
            Logger.Close();
        }
    }

    // Retorna um argumento por índice, ou string vazia
    static string GetArg(int index)
    {
        return (index >= 0 && index < Vetor.Count) ? Vetor[index] : "";
    }

    // Retorna um argumento como lista, separando por vírgula
    static List<string> GetListArg(int index)
    {
        string arg = GetArg(index);
        return string.IsNullOrWhiteSpace(arg)
            ? new List<string>()
            : arg.Split(',').Select(s => s.Trim()).ToList();
    }

        static void PreencherCampo(Dictionary<string, object> mapping, string key, Dictionary<string, object> variaveis, PdfAcroForm form, string valor, PdfFont fonte, float fonteSize)
    {
        Logger.Log($"▶ Iniciando preenchimento de campos para '{key}' com valor '{valor}'...");

        if (!mapping.TryGetValue(key, out var campoObj))
        {
            Logger.Log($"❌ '{key}' não encontrado no mapping.");
            return;
        }

        if (campoObj is not JObject campoMap)
        {
            Logger.Log($"❌ '{key}' não é um JObject.");
            return;
        }

        if (!campoMap.TryGetValue(valor, out var camposToken))
        {
            Logger.Log($"❌ Valor '{valor}' não encontrado no mapping de '{key}'.");
            return;
        }

        if (camposToken is not JObject campos)
        {
            Logger.Log($"❌ Campos de '{valor}' não são um JObject.");
            return;
        }

        foreach (var campo in campos)
        {
            string nomeCampoPdf;
            string template = campo.Value?.ToString() ?? string.Empty;

            if (campo.Key.Equals("botao", StringComparison.OrdinalIgnoreCase))
            {
                nomeCampoPdf = template;
                template = "";
            }
            else
            {
                nomeCampoPdf = campo.Key;
            }

            Logger.Log($"📝 Campo: {nomeCampoPdf} | Template original: {template}");

            try
            {
                foreach (var par in variaveis)
                {
                    template = template.Replace($"${{{par.Key}}}", par.Value?.ToString() ?? "");

                }

                var field = form.GetField(nomeCampoPdf);
                if (field == null)
                {
                    Logger.Log($"⚠️ Campo '{nomeCampoPdf}' não encontrado no PDF.");
                    continue;
                }

                Logger.Log($"🔎 Opções para campo {nomeCampoPdf}: {string.Join(", ", field.GetAppearanceStates())}");

                
                if (nomeCampoPdf.StartsWith("Button"))
                {
                    var options = field.GetAppearanceStates();
                    string valorBotao = options.Contains("Yes") ? "Yes" : options.FirstOrDefault() ?? "On";
                    field.SetValue(valorBotao, fonte, fonteSize);
                    Logger.Log($"✅ Botão '{nomeCampoPdf}' marcado com valor '{valorBotao}'");
                }
                else
                {
                    field.SetValue(template ?? string.Empty, fonte, fonteSize);
                    Logger.Log($"✅ Campo preenchido: {nomeCampoPdf} = '{template}'");
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"❌ Erro ao preencher campo '{nomeCampoPdf}': {ex.Message}");
            }
        }

        Logger.Log($"✔️ Finalizado preenchimento de '{key}' para '{valor}'.");
    }

        static void PreencherCampoContains(Dictionary<string, object> mapping, string key, Dictionary<string, object> variaveis, PdfAcroForm form, string valor, PdfFont fonte, float fonteSize)
    {
        Logger.Log($"▶ Iniciando preenchimento com Contains para '{key}' com valor '{valor}'...");

        if (!mapping.TryGetValue(key, out var obj)) return;
        if (obj is not JObject mapa) return;

        foreach (var prop in mapa.Properties())
        {
            if (!valor.Contains(prop.Name, StringComparison.OrdinalIgnoreCase)) continue;

            var campos = prop.Value as JObject;
            if (campos == null) continue;

            foreach (var campo in campos)
            {
                string nomeCampoPdf;
                string template = campo.Value?.ToString() ?? string.Empty;

                if (campo.Key.Equals("botao", StringComparison.OrdinalIgnoreCase))
                {
                    nomeCampoPdf = template;
                    template = "";
                }
                else
                {
                    nomeCampoPdf = campo.Key;
                }

                Logger.Log($"📝 Campo: {nomeCampoPdf} | Template original: {template}");

                try
                {
                    // Substituição com base nos valores do dicionário de variáveis
                    foreach (var par in variaveis)
                    {
                        string valorSubstituto = par.Value switch
                        {
                            null => "",
                            string s => s,
                            IEnumerable<string> list => string.Join(", ", list),
                            _ => par.Value?.ToString() ?? ""
                        };


                        template = template.Replace($"${{{par.Key}}}", valorSubstituto);
                    }

                    var field = form.GetField(nomeCampoPdf);
                    if (field == null)
                    {
                        Logger.Log($"⚠️ Campo '{nomeCampoPdf}' não encontrado no PDF.");
                        continue;
                    }

                    Logger.Log($"🔎 Opções para campo {nomeCampoPdf}: {string.Join(", ", field.GetAppearanceStates())}");

                    if (nomeCampoPdf.StartsWith("Button"))
                    {
                        var options = field.GetAppearanceStates();
                        string valorBotao = options.Contains("Yes") ? "Yes" : options.FirstOrDefault() ?? "On";
                        field.SetValue(valorBotao, fonte, fonteSize);
                        Logger.Log($"✅ Botão '{nomeCampoPdf}' marcado com valor '{valorBotao}'");
                    }
                    else
                    {
                        field.SetValue(template ?? string.Empty, fonte, fonteSize);
                        Logger.Log($"✅ Campo preenchido: {nomeCampoPdf} = '{template}'");
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log($"❌ Erro ao preencher campo '{nomeCampoPdf}': {ex.Message}");
                }
            }
        }

        Logger.Log($"✔️ Finalizado preenchimento para '{key}'.");
    }


        static void PreencherMonitores(Dictionary<string, object> mapping, Dictionary<string, object> variaveis, PdfAcroForm form, PdfFont fonte, float fonteSize)
    {
        Logger.Log("▶ Iniciando preenchimento de campos de MONITORES...");

        if (!mapping.TryGetValue("tipo_equipamento", out var tipoObj) || tipoObj is not JObject tipoEquipamentoMap)
        {
            Logger.Log("❌ 'tipo_equipamento' inválido ou ausente.");
            return;
        }

        if (!tipoEquipamentoMap.TryGetValue("MONITORES", out var monitoresToken) || monitoresToken is not JObject monitores)
        {
            Logger.Log("❌ Mapeamento de 'MONITORES' inválido ou ausente.");
            return;
        }

        foreach (var monitorTipo in new[] { "MONITOR_UM", "MONITOR_DOIS" })
        {
            if (monitorTipo == "MONITOR_DOIS")
            {
                bool dadosVazios =
                    string.IsNullOrWhiteSpace(variaveis.GetValueOrDefault("data_monitor_dois")?.ToString()) &&
                    string.IsNullOrWhiteSpace(variaveis.GetValueOrDefault("patrimonio_monitor_dois")?.ToString());

                if (dadosVazios)
                {
                    Logger.Log("ℹ️ Ignorando preenchimento de MONITOR_DOIS por estar vazio.");
                    continue;
                }
            }

            Logger.Log($"🔍 Processando {monitorTipo}...");

            if (!monitores.TryGetValue(monitorTipo, out var monitorFieldsToken) || monitorFieldsToken is not JObject monitorFields)
            {
                Logger.Log($"⚠️ '{monitorTipo}' não encontrado ou inválido.");
                continue;
            }

            foreach (var campo in monitorFields)
            {
                string nomeCampoPdf;
                string template = campo.Value?.ToString() ?? string.Empty;

                if (campo.Key.Equals("botao", StringComparison.OrdinalIgnoreCase))
                {
                    nomeCampoPdf = template;
                    template = "";
                }
                else
                {
                    nomeCampoPdf = campo.Key;
                }

                Logger.Log($"📝 Campo: {nomeCampoPdf} | Template original: {template}");

                try
                {
                    foreach (var par in variaveis)
                    {
                        string valorSubstituto = par.Value switch
                        {
                            null => "",
                            string s => s,
                            IEnumerable<string> list => string.Join(", ", list),
                            _ => par.Value?.ToString() ?? ""
                        };

                        template = template.Replace($"${{{par.Key}}}", valorSubstituto);
                    }

                    var field = form.GetField(nomeCampoPdf);
                    if (field == null)
                    {
                        Logger.Log($"⚠️ Campo '{nomeCampoPdf}' não encontrado.");
                        continue;
                    }

                    Logger.Log($"🔎 Opções para campo {nomeCampoPdf}: {string.Join(", ", field.GetAppearanceStates())}");

                    if (nomeCampoPdf.StartsWith("Button"))
                    {
                        var options = field.GetAppearanceStates();
                        string valorBotao = options.Contains("Yes") ? "Yes" : options.FirstOrDefault() ?? "On";
                        field.SetValue(valorBotao, fonte, fonteSize);
                        Logger.Log($"✅ Botão '{nomeCampoPdf}' marcado com valor '{valorBotao}'");
                    }
                    else
                    {
                        field.SetValue(template ?? string.Empty, fonte, fonteSize);
                        Logger.Log($"✅ Preenchido: {nomeCampoPdf} = '{template}'");
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log($"⚠️ Erro ao preencher campo '{nomeCampoPdf}': {ex.Message}");
                }
            }
        }

        Logger.Log("✔️ Finalizado preenchimento de MONITORES.");
    }



    // Define campos específicos a partir do JSON usando substituição de variáveis
    static void SetField(PdfAcroForm form, object mapObject, Dictionary<string, object> variaveis, PdfFont fonte, float fonteSize)
    {
        if (mapObject is not JObject obj) return;

        foreach (var prop in obj)
        {
            string nomeCampoPdf = prop.Key;
            string template = prop.Value?.ToString() ?? string.Empty;


            // Substituir variáveis no template, suportando strings e listas
            foreach (var par in variaveis)
            {
                string valor = par.Value switch
                {
                    null => "",
                    string s => s,
                    IEnumerable<string> lista => string.Join(", ", lista),
                    _ => par.Value?.ToString() ?? ""
                };

                template = template.Replace($"${{{par.Key}}}", valor);
            }

            try
            {
                var field = form.GetField(nomeCampoPdf);

                if (field != null)
                {
                    field.SetValue(template ?? string.Empty, fonte, fonteSize);
                    Logger.Log($"SetField: Preenchendo {nomeCampoPdf} com '{template}'");
                }
                else
                {
                    Logger.Log($"SetField: Campo '{nomeCampoPdf}' não encontrado no PDF.");
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"❌ Erro ao preencher campo '{nomeCampoPdf}': {ex.Message}");
            }
        }
    }


    static string ReplaceVariables(string template, Dictionary<string, object> variaveis)
    {
        if (string.IsNullOrEmpty(template) || variaveis == null)
            return template;

        foreach (var par in variaveis)
        {
            string chave = par.Key;
            string valorSubstituto = par.Value switch
            {
                null => "",
                string s => s,
                IEnumerable<string> lista => string.Join(", ", lista),
                _ => par.Value?.ToString() ?? ""
            };

            template = template.Replace($"${{{chave}}}", valorSubstituto);
        }

        return template;
    }


    /*
    static void PreencherComNomesDosCampos(PdfAcroForm form, PdfFont fonte, float fonteSize)
    {
        Logger.Log("🧪 Preenchendo todos os campos com seus próprios nomes...");

        try
        {
            IDictionary<string, PdfFormField> fields = form.GetAllFormFields();

            foreach (var entry in fields)
            {
                string nomeCampo = entry.Key;
                PdfFormField campo = entry.Value;

                try
                {
                    campo.SetValue(nomeCampo, fonte, fonteSize);
                    Logger.Log($"🔤 Campo '{nomeCampo}' preenchido com '{nomeCampo}'");
                }
                catch (Exception exCampo)
                {
                    Logger.Log($"⚠️ Erro ao preencher campo '{nomeCampo}': {exCampo.Message}");
                }
            }

            Logger.Log("✅ Todos os campos foram preenchidos com seus nomes.");
        }
        catch (Exception ex)
        {
            Logger.Log($"❌ Erro ao acessar os campos do formulário: {ex.Message}");
        }
    }
    */

}
