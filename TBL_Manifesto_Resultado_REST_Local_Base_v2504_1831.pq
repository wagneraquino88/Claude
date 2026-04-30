// Versão: 25/04/2026 18:31 | Carga dinâmica de TBL_Manifesto_Resultado na pasta Base
// Suporte: caminhos locais (C:\ ou \\servidor) e URLs SharePoint / OneDrive Personal

let
    // ── CONFIGURAÇÕES ─────────────────────────────────────────────────────────
    NomeTabelaCaminho = "TabelaCaminho",
    NomeColunaCaminho = "CaminhoArquivo",
    NomePastaBase     = "Base",
    TextoArquivoAlvo  = "TBL_Manifesto_Resultado",
    NomeAbaAlvo       = "TBL_Manifesto_Resultado",
    ExtensoesValidas  = {".xlsx", ".xlsm", ".xls", ".xlsb"},

    // ── FUNÇÕES AUXILIARES ────────────────────────────────────────────────────

    // Decodifica segmentos de URL codificados (ex: %20 → espaço, %C3%A7 → ç)
    fnDecodificarUrl = (texto as text) as text =>
        try Uri.Parts("http://x/?q=" & Text.Replace(texto, "&", "%26"))[Query][q]
        otherwise texto,

    // Retorna true se o arquivo é o alvo correto (nome, extensão e não temporário)
    fnEhArquivoAlvo = (nome as text) as logical =>
        Text.Contains(nome, TextoArquivoAlvo, Comparer.OrdinalIgnoreCase)
        and not Text.StartsWith(nome, "~$")
        and List.AnyTrue(
            List.Transform(ExtensoesValidas, each Text.EndsWith(Text.Lower(nome), _))
        ),

    // Garante que o caminho local aponte para a subpasta Base
    fnCaminhoComBase = (caminho as text) as text =>
        let
            Limpo       = Text.TrimEnd(Text.Trim(caminho), "\"),
            UltimaParte = try Text.AfterDelimiter(Limpo, "\", {0, RelativePosition.FromEnd}) otherwise Limpo
        in
            if Text.Lower(UltimaParte) = Text.Lower(NomePastaBase)
            then Limpo
            else Limpo & "\" & NomePastaBase,

    // Extrai WebRoot e ServerRelativeUrl da pasta Base a partir de uma URL SharePoint/OneDrive
    fnInfoUrl = (urlOriginal as text) as record =>
        let
            Url    = if Text.StartsWith(Text.Lower(urlOriginal), "http") then urlOriginal else "https://" & urlOriginal,
            Partes = try Uri.Parts(Url) otherwise error "URL inválida: " & Url,
            Host   = Partes[Host],
            Scheme = Partes[Scheme],

            // Decodifica e filtra segmentos do caminho
            Segmentos =
                List.Select(
                    List.Transform(
                        Text.Split(Text.TrimStart(Partes[Path], "/"), "/"),
                        each Text.Trim(fnDecodificarUrl(Text.From(_)))
                    ),
                    each _ <> ""
                ),

            // Links compartilhados têm prefixo /:x:/r/ — remove os dois primeiros segmentos
            SegsFiltrados =
                if List.Count(Segmentos) >= 4
                    and Text.StartsWith(Segmentos{0}, ":")
                    and Text.Lower(Segmentos{1}) = "r"
                then List.Skip(Segmentos, 2)
                else Segmentos,

            SegmLower   = List.Transform(SegsFiltrados, Text.Lower),
            PosSites    = List.PositionOf(SegmLower, "sites"),
            PosPersonal = List.PositionOf(SegmLower, "personal"),

            PosRaiz =
                if PosSites    >= 0 then PosSites
                else if PosPersonal >= 0 then PosPersonal
                else error "URL sem '/sites/' nem '/personal/'. Verifique o caminho informado.",

            APartirDaRaiz = List.Skip(SegsFiltrados, PosRaiz),
            ValidaRaiz    =
                if List.Count(APartirDaRaiz) < 2
                then error "URL incompleta: não foi possível identificar a raiz do site."
                else APartirDaRaiz,

            WebRoot        = Scheme & "://" & Host & "/" & Text.Combine(List.FirstN(ValidaRaiz, 2), "/"),
            ServerRelativo = "/" & Text.Combine(ValidaRaiz, "/"),
            JaNaBase       = Text.Lower(Text.Trim(List.Last(ValidaRaiz))) = Text.Lower(NomePastaBase),
            ServerBase     = if JaNaBase then ServerRelativo else ServerRelativo & "/" & NomePastaBase
        in
            [WebRoot = WebRoot, ServerRelativeBase = ServerBase],

    // ── FONTES ───────────────────────────────────────────────────────────────

    // Lista arquivos da pasta Base em caminhos de rede locais
    fnFonteLocal = (caminho as text) as table =>
        let
            Pasta     = fnCaminhoComBase(caminho),
            Arquivos  =
                try Folder.Contents(Pasta)
                otherwise error "Pasta local inacessível: " & Pasta,
            Filtrados = Table.SelectRows(Arquivos, each fnEhArquivoAlvo(Text.From([Name]))),
            ComMeta1  = Table.AddColumn(Filtrados, "_OrigemFonte",      each "Local", type text),
            ComMeta2  = Table.AddColumn(ComMeta1,  "_WebRoot",          each null,    type nullable text),
            ComMeta3  = Table.AddColumn(ComMeta2,  "ServerRelativeUrl", each null,    type nullable text)
        in
            ComMeta3,

    // Lista arquivos da pasta Base via REST API do SharePoint / OneDrive
    fnFonteSharePoint = (urlOriginal as text) as table =>
        let
            Info       = fnInfoUrl(urlOriginal),
            WebRoot    = Info[WebRoot],
            PastaOData = Text.Replace(Info[ServerRelativeBase], "'", "''"),

            Resposta =
                try
                    Json.Document(
                        Web.Contents(WebRoot, [
                            RelativePath = "_api/web/GetFolderByServerRelativeUrl('" & PastaOData & "')/Files",
                            Query        = [#"$select" = "Name,ServerRelativeUrl,Length,TimeLastModified"],
                            Headers      = [Accept = "application/json;odata=nometadata"]
                        ])
                    )
                otherwise error "Falha ao listar arquivos em: " & Info[ServerRelativeBase],

            Arquivos  =
                if Record.HasFields(Resposta, "value") then Resposta[value]
                else error "Resposta inesperada da API para: " & Info[ServerRelativeBase],

            Tabela    =
                if List.Count(Arquivos) = 0
                then #table({"Name", "ServerRelativeUrl", "Length", "TimeLastModified"}, {})
                else Table.FromRecords(Arquivos),
            Colunas   = Table.SelectColumns(Tabela, {"Name", "ServerRelativeUrl", "Length", "TimeLastModified"}, MissingField.UseNull),
            Filtrados = Table.SelectRows(Colunas, each fnEhArquivoAlvo(Text.From([Name]))),
            ComMeta1  = Table.AddColumn(Filtrados, "_OrigemFonte", each "SharePointRest", type text),
            ComMeta2  = Table.AddColumn(ComMeta1,  "_WebRoot",     each WebRoot,          type text)
        in
            ComMeta2,

    // Baixa o binário de um arquivo via SharePoint REST
    fnBaixarBinario = (WebRoot as text, ServerRelativeUrl as text, NomeArquivo as text) as binary =>
        try
            Binary.Buffer(
                Web.Contents(WebRoot, [
                    RelativePath = "_api/web/GetFileByServerRelativeUrl('" & Text.Replace(ServerRelativeUrl, "'", "''") & "')/$value"
                ])
            )
        otherwise error "Falha ao baixar: " & NomeArquivo,

    // ── LEITURA DO CAMINHO CONFIGURADO ────────────────────────────────────────
    TabelaBruta =
        try Excel.CurrentWorkbook(){[Name = NomeTabelaCaminho]}[Content]
        otherwise error "Tabela '" & NomeTabelaCaminho & "' não encontrada no Excel.",

    TabelaCaminho =
        if Table.RowCount(TabelaBruta) = 0
        then error "Tabela '" & NomeTabelaCaminho & "' está vazia."
        else TabelaBruta,

    CaminhoConfig =
        if Table.HasColumns(TabelaCaminho, NomeColunaCaminho)
        then Text.Trim(Text.From(TabelaCaminho{0}[CaminhoArquivo]))
        else Text.Trim(Text.From(Record.FieldValues(TabelaCaminho{0}){0})),

    Caminho =
        if CaminhoConfig = ""
        then error "O caminho na tabela '" & NomeTabelaCaminho & "' está em branco."
        else CaminhoConfig,

    // ── DETECÇÃO E BUSCA DO ARQUIVO ───────────────────────────────────────────
    EhLocal =
        Text.Contains(Caminho, ":\") or Text.StartsWith(Caminho, "\\"),

    EhUrl =
        Text.StartsWith(Text.Lower(Caminho), "http")
        or Text.Contains(Text.Lower(Caminho), "sharepoint.com"),

    FonteArquivos =
        if EhLocal    then fnFonteLocal(Caminho)
        else if EhUrl then fnFonteSharePoint(Caminho)
        else error "Caminho não reconhecido como local nem como URL SharePoint/OneDrive.",

    Arquivo =
        if Table.RowCount(FonteArquivos) = 0 then
            error "Arquivo '" & TextoArquivoAlvo & "' não encontrado na pasta '" & NomePastaBase & "'."
        else if Table.RowCount(FonteArquivos) > 1 then
            error "Mais de um arquivo '" & TextoArquivoAlvo & "' encontrado. Deixe apenas um na pasta '" & NomePastaBase & "'."
        else
            FonteArquivos{0},

    Binario =
        if Arquivo[_OrigemFonte] = "Local"
        then Arquivo[Content]
        else fnBaixarBinario(
            Text.From(Arquivo[_WebRoot]),
            Text.From(Arquivo[ServerRelativeUrl]),
            Text.From(Arquivo[Name])
        ),

    // ── TRANSFORMAÇÕES ────────────────────────────────────────────────────────
    Fonte = Excel.Workbook(Binario, null, true),

    Abas =
        if Table.HasColumns(Fonte, "Kind")
        then Table.SelectRows(Fonte, each [Item] = NomeAbaAlvo and [Kind] = "Sheet")
        else Table.SelectRows(Fonte, each [Item] = NomeAbaAlvo),

    DadosBrutos =
        if Table.RowCount(Abas) = 1 then Abas{0}[Data]
        else if Table.RowCount(Abas) = 0 then
            error "Aba '" & NomeAbaAlvo & "' não encontrada em '" & TextoArquivoAlvo & "'."
        else
            error "Mais de uma aba '" & NomeAbaAlvo & "' encontrada em '" & TextoArquivoAlvo & "'.",

    #"Changed Type2" = Table.TransformColumnTypes(DadosBrutos,{{"Column1", type text}, {"Column2", type text}, {"Column3", type text}, {"Column4", type text}, {"Column5", type text}, {"Column6", type text}, {"Column7", type text}, {"Column8", type text}, {"Column9", type text}, {"Column10", type text}, {"Column11", type text}, {"Column12", type text}, {"Column13", type text}, {"Column14", type text}, {"Column15", type text}, {"Column16", type text}, {"Column17", type text}, {"Column18", type text}, {"Column19", type text}, {"Column20", type number}, {"Column21", type text}}),
    #"Promoted Headers" = Table.PromoteHeaders(#"Changed Type2", [PromoteAllScalars=true]),
    #"Changed Type1" = Table.TransformColumnTypes(#"Promoted Headers",{{"House", type text}, {"Conceito", type text}, {"US$", type number}, {"EUR", type number}, {"R$", type number}, {".", type text}, {"US$2", type number}, {"Total", type text}, {"..", type text}, {"Tt PP", type number}, {"Tt CC", type number}, {"Obs", type text}, {"Resultado [Não]", type text}, {"$", type text}, {"Conta Contabil", Int64.Type}, {"Cotação", type text}, {"Conceito3", type text}, {"US$4", Int64.Type}, {"EUR5", type number}, {"6", type number}, {"Cliente", type text}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type1",{{"Resultado [Não]", "-2"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns",{{"Ordem Arquivo", Int64.Type}, {"Ordem Linha", Int64.Type}})
in
    #"Changed Type"
