// Versão: 01/05/2026 | Carga dos itens da pasta Base_Manifesto_Resultado (dentro de Base)
// Suporte: caminhos locais (C:\ ou \\servidor) e URLs SharePoint / OneDrive Personal

let
    // ── CONFIGURAÇÕES ─────────────────────────────────────────────────────────
    NomeTabelaCaminho = "TabelaCaminho",
    NomeColunaCaminho = "CaminhoArquivo",
    NomePastaBase     = "Base",
    NomePastaAlvo     = "Base_Manifesto_Resultado",

    // ── FUNÇÕES AUXILIARES ────────────────────────────────────────────────────

    // Decodifica segmentos de URL codificados (ex: %20 → espaço, %C3%A7 → ç)
    fnDecodificarUrl = (texto as text) as text =>
        try Uri.Parts("http://x/?q=" & Text.Replace(texto, "&", "%26"))[Query][q]
        otherwise texto,

    // Garante que o caminho local aponte para Base\Base_Manifesto_Resultado
    fnCaminhoAlvo = (caminho as text) as text =>
        let
            Limpo       = Text.TrimEnd(Text.Trim(caminho), "\"),
            UltimaParte = try Text.AfterDelimiter(Limpo, "\", {0, RelativePosition.FromEnd}) otherwise Limpo,
            JaNaAlvo    = Text.Lower(UltimaParte) = Text.Lower(NomePastaAlvo),
            JaNaBase    = Text.Lower(UltimaParte) = Text.Lower(NomePastaBase)
        in
            if JaNaAlvo    then Limpo
            else if JaNaBase then Limpo & "\" & NomePastaAlvo
            else               Limpo & "\" & NomePastaBase & "\" & NomePastaAlvo,

    // Extrai WebRoot e ServerRelativeUrl da pasta Base_Manifesto_Resultado a partir de uma URL SharePoint/OneDrive
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
            UltimaSegmento = Text.Lower(Text.Trim(List.Last(ValidaRaiz))),
            JaNaAlvo       = UltimaSegmento = Text.Lower(NomePastaAlvo),
            JaNaBase       = UltimaSegmento = Text.Lower(NomePastaBase),

            ServerAlvo =
                if JaNaAlvo    then ServerRelativo
                else if JaNaBase then ServerRelativo & "/" & NomePastaAlvo
                else               ServerRelativo & "/" & NomePastaBase & "/" & NomePastaAlvo
        in
            [WebRoot = WebRoot, ServerRelativeBase = ServerAlvo],

    // ── FONTES ───────────────────────────────────────────────────────────────

    // Lista todos os itens da pasta Base_Manifesto_Resultado em caminhos locais
    fnFonteLocal = (caminho as text) as table =>
        let
            PastaAlvo = fnCaminhoAlvo(caminho),
            Conteudo  =
                try Folder.Contents(PastaAlvo)
                otherwise error "Pasta local inacessível: " & PastaAlvo,
            ComOrigem = Table.AddColumn(Conteudo, "_OrigemFonte", each "Local", type text)
        in
            ComOrigem,

    // Lista todos os arquivos da pasta Base_Manifesto_Resultado via REST API do SharePoint / OneDrive
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

            Arquivos =
                if Record.HasFields(Resposta, "value") then Resposta[value]
                else error "Resposta inesperada da API para: " & Info[ServerRelativeBase],

            Tabela    =
                if List.Count(Arquivos) = 0
                then #table({"Name", "ServerRelativeUrl", "Length", "TimeLastModified"}, {})
                else Table.FromRecords(Arquivos),
            Colunas   = Table.SelectColumns(Tabela, {"Name", "ServerRelativeUrl", "Length", "TimeLastModified"}, MissingField.UseNull),
            ComOrigem = Table.AddColumn(Colunas,   "_OrigemFonte", each "SharePointRest", type text),
            ComWebRoot= Table.AddColumn(ComOrigem, "_WebRoot",     each WebRoot,          type text)
        in
            ComWebRoot,

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

    // ── DETECÇÃO E CARREGAMENTO DA PASTA ALVO ────────────────────────────────
    EhLocal =
        Text.Contains(Caminho, ":\") or Text.StartsWith(Caminho, "\\"),

    EhUrl =
        Text.StartsWith(Text.Lower(Caminho), "http")
        or Text.Contains(Text.Lower(Caminho), "sharepoint.com"),

    Resultado =
        if EhLocal    then fnFonteLocal(Caminho)
        else if EhUrl then fnFonteSharePoint(Caminho)
        else error "Caminho não reconhecido como local nem como URL SharePoint/OneDrive."
in
    Resultado
