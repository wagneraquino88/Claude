let
    NomePastaArquivos = "Base\Base_Resultado_Manifesto",
    PrefixosArquivoAlvo = {"BR", "conceitos"},
    ExtensoesExcelAceitas = {".xlsx", ".xlsm", ".xls", ".xlsb"},
    LinhaCabecalhoEsperada = 5,
    NomesColunasResultado = {
        "House", "Conceito", "US$", "EUR", "R$", ".", "US$2", "Total", "..",
        "Tt PP", "Tt CC", "Obs", "Resultado [Não]", "$", "Conta Contabil",
        "Cotação", "Conceito3", "US$4", "EUR5", "R$6", "Cliente"
    },

    TiposDepoisCabecalho = {
        {"US$", type number}, {"EUR", type number}, {"R$", type number},
        {"US$2", type number}, {"Tt PP", type number}, {"Tt CC", type number},
        {"Conta Contabil", Int64.Type}, {"US$4", type number},
        {"EUR5", type number}, {"R$6", type number}
    },

    TiposOrdem = {
        {"Ordem Arquivo", Int64.Type},
        {"Ordem Linha", Int64.Type}
    },

    TabelaCaminho =
        try Excel.CurrentWorkbook(){[Name = "TabelaCaminho"]}[Content]
        otherwise error "Não encontrei a tabela 'TabelaCaminho' na aba Config.",

    ValorCaminho =
        if Table.RowCount(TabelaCaminho) = 0 then
            error "A TabelaCaminho está vazia. Confira a fórmula da aba Config."
        else if Table.HasColumns(TabelaCaminho, "CaminhoArquivo") then
            Record.Field(TabelaCaminho{0}, "CaminhoArquivo")
        else
            Record.FieldValues(TabelaCaminho{0}){0},

    CaminhoInformado = Text.Trim(Text.From(ValorCaminho) ?? ""),
    CaminhoPrincipal =
        if CaminhoInformado = "" then
            error "A TabelaCaminho retornou um caminho em branco. Confira a fórmula da aba Config."
        else if Text.EndsWith(CaminhoInformado, "\") or Text.EndsWith(CaminhoInformado, "/") then
            Text.Start(CaminhoInformado, Text.Length(CaminhoInformado) - 1)
        else
            CaminhoInformado,

    CaminhoArquivos = CaminhoPrincipal & "\" & NomePastaArquivos,

    ArquivosDaPasta =
        try Folder.Contents(CaminhoArquivos)
        otherwise error "Não consegui acessar a pasta: " & CaminhoArquivos & ". Confira se a pasta Base_Resultado_Manifesto existe dentro do caminho da aba Config.",

    ArquivosValidos =
        Table.SelectRows(ArquivosDaPasta, each
            let
                NomeLimpo = Text.Trim(Text.From([Name]) ?? ""),
                ExtLimpa = Text.Lower(Text.Trim(Text.From([Extension]) ?? ""))
            in
                List.AnyTrue(
                    List.Transform(
                        PrefixosArquivoAlvo,
                        (prefixo) => Text.StartsWith(NomeLimpo, prefixo, Comparer.OrdinalIgnoreCase)
                    )
                )
                and not Text.StartsWith(NomeLimpo, "~$")
                and List.Contains(ExtensoesExcelAceitas, ExtLimpa)
        ),

    ArquivosParaCarga =
        if Table.RowCount(ArquivosValidos) = 0 then
            error "Nenhum arquivo Excel iniciado por "
                & Text.Combine(List.Transform(PrefixosArquivoAlvo, each "'" & Text.From(_) & "'"), ", ")
                & " foi encontrado em: " & CaminhoArquivos
        else
            Table.AddIndexColumn(
                Table.Sort(ArquivosValidos, {{"Name", Order.Ascending}}),
                "_OrdemArquivo",
                1,
                1,
                Int64.Type
            ),

    TabelasProcessadas = List.Transform(Table.ToRecords(ArquivosParaCarga), (arquivo) =>
        let
            NomeArquivo    = Text.From(Record.Field(arquivo, "Name")),
            OrdemArquivo   = Int64.From(Record.Field(arquivo, "_OrdemArquivo")),
            BinarioArquivo = Binary.Buffer(Record.Field(arquivo, "Content")),

            PastaDeTrabalho =
                try Excel.Workbook(BinarioArquivo, null, true)
                otherwise error "Encontrei o arquivo '" & NomeArquivo & "', mas não consegui abri-lo como Excel. Verifique se ele não está corrompido, protegido por senha ou aberto de forma exclusiva.",

            SomenteAbas    = Table.SelectRows(PastaDeTrabalho, each [Kind] = "Sheet"),
            DadosOriginais =
                if Table.RowCount(SomenteAbas) = 1 then
                    SomenteAbas{0}[Data]
                else
                    error "O arquivo '" & NomeArquivo & "' precisa ter apenas uma aba.",

            ColunasOriginais = Table.ColumnNames(DadosOriginais),
            QuantEsperada    = List.Count(NomesColunasResultado),
            QuantEncontrada  = List.Count(ColunasOriginais),

            ValidaColunas =
                if QuantEncontrada < QuantEsperada then
                    error "A aba selecionada possui menos colunas do que o esperado. Esperado: " & Text.From(QuantEsperada) & "; encontrado: " & Text.From(QuantEncontrada) & "."
                else
                    DadosOriginais,

            LinhaCabecalho =
                try Table.First(Table.Skip(ValidaColunas, LinhaCabecalhoEsperada - 1))
                otherwise error "Não consegui validar a linha de cabeçalho " & Text.From(LinhaCabecalhoEsperada) & ". Confira se a aba selecionada possui dados.",

            CabecalhoValido =
                Text.Upper(Text.Trim(Text.From(Record.Field(LinhaCabecalho, ColunasOriginais{0})) ?? "")) = "HOUSE"
                and Text.Upper(Text.Trim(Text.From(Record.Field(LinhaCabecalho, ColunasOriginais{1})) ?? "")) = "CONCEITO",

            DadosSemTitulos =
                if CabecalhoValido then
                    Table.Skip(ValidaColunas, LinhaCabecalhoEsperada)
                else
                    error "A linha " & Text.From(LinhaCabecalhoEsperada) & " não parece ser o cabeçalho esperado. A primeira coluna precisa ser 'House' e a segunda 'Conceito'.",

            ColsParaRenomear = List.FirstN(ColunasOriginais, QuantEsperada),
            DadosPadrao      = Table.RenameColumns(DadosSemTitulos, List.Zip({ColsParaRenomear, NomesColunasResultado}), MissingField.Ignore),
            DadosColunas     = Table.SelectColumns(DadosPadrao, NomesColunasResultado, MissingField.UseNull),
            Registros        = List.Buffer(Table.ToRecords(DadosColunas)),
            QtdRegistros     = List.Count(Registros),

            ProcessarLinha = (indice as number, houseAtual as any, cotacaoAtual as any, clienteAtual as any) as record =>
                let
                    Linha        = Registros{indice},
                    HouseValor   = Record.Field(Linha, "House"),
                    HouseTexto   = Text.Trim(Text.From(HouseValor) ?? ""),
                    HouseChave   = Text.Upper(HouseTexto),

                    InicioDocumento =
                        Text.StartsWith(HouseChave, "COST IMPORTAÇÃO")
                        or Text.StartsWith(HouseChave, "COST IMPORTACAO"),

                    ValoresDaLinha   = List.Transform(NomesColunasResultado, (c) => Record.Field(Linha, c)),
                    LinhaTemConteudo = List.AnyTrue(
                        List.Transform(ValoresDaLinha, (v) => v <> null and Text.Trim(Text.From(v)) <> "")
                    ),
                    ConceitoMaiusc = Text.Upper(Text.Trim(Text.From(Record.Field(Linha, "Conceito")) ?? "")),
                    EhTituloOuCabecalho =
                        (HouseChave = "HOUSE" and ConceitoMaiusc = "CONCEITO")
                        or Text.StartsWith(HouseChave, "COST IMPORTAÇÃO")
                        or Text.StartsWith(HouseChave, "COST IMPORTACAO")
                        or Text.StartsWith(HouseChave, "ORIGEM:")
                        or Text.StartsWith(HouseChave, "AGENTE:"),
                    LinhaIgnorada = not LinhaTemConteudo or EhTituloOuCabecalho,

                    HouseAtualBase   = if InicioDocumento then null else houseAtual,
                    CotacaoAtualBase = if InicioDocumento then null else cotacaoAtual,
                    ClienteAtualBase = if InicioDocumento then null else clienteAtual,
                    HouseAtualChave  = Text.Upper(Text.Trim(Text.From(HouseAtualBase) ?? "")),
                    TemHouseNaLinha  = HouseTexto <> "",
                    MudouHouse       = TemHouseNaLinha and HouseChave <> HouseAtualChave,

                    EhTotal          = ConceitoMaiusc = "TOTAL",
                    CotacaoValor     = Record.Field(Linha, "Cotação"),
                    TemCotacaoNaLinha = Text.Trim(Text.From(CotacaoValor) ?? "") <> "",
                    ClienteValor     = Record.Field(Linha, "Cliente"),
                    TemClienteNaLinha = Text.Trim(Text.From(ClienteValor) ?? "") <> "",

                    NovaHouseAtual =
                        if LinhaIgnorada then HouseAtualBase
                        else if TemHouseNaLinha then HouseValor
                        else HouseAtualBase,

                    NovaCotacaoAtual =
                        if LinhaIgnorada then CotacaoAtualBase
                        else if MudouHouse then
                            if EhTotal then null
                            else if TemCotacaoNaLinha then CotacaoValor
                            else null
                        else if EhTotal then CotacaoAtualBase
                        else if TemCotacaoNaLinha then CotacaoValor
                        else CotacaoAtualBase,

                    NovaClienteAtual =
                        if LinhaIgnorada then ClienteAtualBase
                        else if MudouHouse then
                            if TemClienteNaLinha then ClienteValor else null
                        else if TemClienteNaLinha then ClienteValor
                        else ClienteAtualBase,

                    HouseSaida   = if LinhaIgnorada then null else if TemHouseNaLinha then HouseValor else NovaHouseAtual,
                    CotacaoSaida = if LinhaIgnorada then null else if TemCotacaoNaLinha then CotacaoValor else NovaCotacaoAtual,
                    ClienteSaida = if LinhaIgnorada then null else if TemClienteNaLinha then ClienteValor else NovaClienteAtual,

                    LinhaSaida =
                        if LinhaIgnorada then
                            null
                        else
                            Record.Combine({
                                Record.RemoveFields(Linha, {"House", "Cotação", "Cliente"}, MissingField.Ignore),
                                [House = HouseSaida, #"Cotação" = CotacaoSaida, Cliente = ClienteSaida]
                            })
                in
                    [
                        Indice       = indice,
                        HouseAtual   = NovaHouseAtual,
                        CotacaoAtual = NovaCotacaoAtual,
                        ClienteAtual = NovaClienteAtual,
                        Saida        = LinhaSaida
                    ],

            RegistrosProcessados =
                if QtdRegistros = 0 then
                    {}
                else
                    let
                        EstadoInicial = ProcessarLinha(0, null, null, null),
                        Estados = List.Generate(
                            () => EstadoInicial,
                            each [Indice] < QtdRegistros,
                            each
                                if [Indice] + 1 < QtdRegistros then
                                    ProcessarLinha([Indice] + 1, [HouseAtual], [CotacaoAtual], [ClienteAtual])
                                else
                                    [Indice = QtdRegistros, HouseAtual = [HouseAtual], CotacaoAtual = [CotacaoAtual], ClienteAtual = [ClienteAtual], Saida = null],
                            each [Saida]
                        )
                    in
                        List.RemoveNulls(Estados),

            DadosComCabecalho =
                if List.Count(RegistrosProcessados) = 0 then
                    #table(NomesColunasResultado, {})
                else
                    Table.SelectColumns(Table.FromRecords(RegistrosProcessados), NomesColunasResultado, MissingField.UseNull),

            TentativaTransformacao =
                try
                    let
                        ComTipos        = Table.TransformColumnTypes(DadosComCabecalho, TiposDepoisCabecalho, [Culture = "pt-BR", MissingField = MissingField.Ignore]),
                        Renomeado       = Table.RenameColumns(ComTipos, {{"Resultado [Não]", "-2"}}, MissingField.Ignore),
                        ComOrdemLinha   = Table.AddIndexColumn(Renomeado, "Ordem Linha", 1, 1, Int64.Type),
                        ComOrdemArquivo = Table.AddColumn(ComOrdemLinha, "Ordem Arquivo", each OrdemArquivo, Int64.Type),
                        Reordenado      = Table.ReorderColumns(ComOrdemArquivo, {"Ordem Arquivo", "Ordem Linha"} & Table.ColumnNames(Renomeado), MissingField.Ignore)
                    in
                        Table.TransformColumnTypes(Reordenado, TiposOrdem, [Culture = "pt-BR", MissingField = MissingField.Ignore])
        in
            if TentativaTransformacao[HasError] then
                error "O arquivo '" & NomeArquivo & "' foi encontrado, mas a aba selecionada não está no layout esperado. Detalhe técnico: " & TentativaTransformacao[Error][Message]
            else
                TentativaTransformacao[Value]
    ),

    ResultadoFinal = Table.Combine(TabelasProcessadas)
in
    ResultadoFinal
