# 📋 Documentação — Carregador de TBL_Manifesto_Resultado

Um guia simples e prático para usar o código Power Query que **busca automaticamente um arquivo Excel em diferentes locais** (seu computador, servidor da empresa ou nuvem SharePoint/OneDrive).

---

## 🎯 Para que serve?

Este código resolve um problema comum: **você precisa trazer dados de um arquivo Excel que muda de lugar ou de nome frequentemente, sem ter que procurar e carregar manualmente toda vez**.

### Exemplos de uso:

- 📁 O arquivo de "Manifesto de Resultado" fica em `C:\Documentos\Base\`
- 🌐 O arquivo está no SharePoint da empresa (nuvem)
- 👥 O arquivo está no OneDrive corporativo
- O arquivo pode estar em qualquer um desses lugares — o código encontra e carrega automaticamente

---

## 🚀 Como começar (Setup)

### Passo 1: Prepare a tabela de configuração no Excel

Você precisa criar uma **pequena tabela** no seu arquivo Excel que indique **onde** o arquivo está.

1. **Abra seu arquivo Excel**
2. **Crie uma nova planilha** (ou use uma existente)
3. **Crie uma tabela com dois dados:**

| TabelaCaminho | CaminhoArquivo |
|---|---|
| (deixe em branco) | **Cole aqui o caminho onde o arquivo está** |

**Exemplos de caminhos válidos:**

```
C:\Documentos\Base
\\servidor\compartilhado\Base
https://empresa.sharepoint.com/sites/financeiro/documentos/Base
https://empresa-my.sharepoint.com/personal/seu_usuario/DocumentosCompartilhados/Base
```

> **Dica:** O caminho **não precisa incluir o nome do arquivo** — o código procura automaticamente por qualquer arquivo com "TBL_Manifesto_Resultado" no nome dentro da pasta `Base`.

---

### Passo 2: Integre o código no Power Query

1. **No Excel, vá para:** `Dados` → `Obter Dados` → `De Outras Fontes` → `Editor em Branco`

2. **Copie todo o código melhorado** (o arquivo `.pq`)

3. **Cole no editor** que apareceu

4. **Pressione `Enter`** ou clique em ✓ para carregar

5. **Pronto!** O Excel vai buscar automaticamente o arquivo

---

## 📖 Entender o que acontece (o básico)

```
┌─────────────────────────────────────────┐
│ 1. Você informa onde o arquivo está     │
│    (tabela TabelaCaminho)               │
└────────────┬────────────────────────────┘
             │
             ▼
┌─────────────────────────────────────────┐
│ 2. O código identifica o tipo de lugar: │
│    • Pasta local? (C:\)                 │
│    • SharePoint? (nuvem)                │
│    • OneDrive? (nuvem pessoal)          │
└────────────┬────────────────────────────┘
             │
             ▼
┌─────────────────────────────────────────┐
│ 3. O código procura pela pasta "Base"   │
│    naquele local                        │
└────────────┬────────────────────────────┘
             │
             ▼
┌─────────────────────────────────────────┐
│ 4. Dentro da "Base", procura por:       │
│    "TBL_Manifesto_Resultado"            │
│    (arquivo .xlsx, .xlsm, .xls, .xlsb) │
└────────────┬────────────────────────────┘
             │
             ▼
┌─────────────────────────────────────────┐
│ 5. Encontrou? Abre a aba                │
│    "TBL_Manifesto_Resultado"            │
│    e carrega os dados                   │
└─────────────────────────────────────────┘
```

---

## ⚙️ O que é cada nome no código?

Termos que aparecem e o que significam em português descomplicado:

| Termo | O que é | Exemplo |
|---|---|---|
| **TabelaCaminho** | Tabela que você cria para indicar onde os dados estão | Uma tabela com 1 linha dizendo `C:\Documentos\Base` |
| **CaminhoArquivo** | Coluna da tabela com o endereço/localização | `C:\Documentos\Base` ou `https://empresa.sharepoint.com/...` |
| **Base** | Nome da pasta onde o arquivo deve estar | Sua pasta precisa se chamar `Base` |
| **TBL_Manifesto_Resultado** | Parte do nome do arquivo que o código procura | Válido: `TBL_Manifesto_Resultado_jan.xlsx` ou `TBL_Manifesto_Resultado.xlsx` |
| **Aba/Sheet** | As "abas" ou "planilhas" dentro do arquivo Excel | Como a "Plan1", "Plan2" etc. do Excel |
| **SharePoint** | Serviço de nuvem da Microsoft para empresas | Tipo de "OneDrive corporativo" onde os arquivos da empresa ficam |
| **OneDrive** | Serviço de nuvem da Microsoft para armazenar arquivos | Tipo de "Drive" (seu armazenamento na nuvem) |
| **REST API** | Forma de se comunicar com servidores do SharePoint/OneDrive | Não precisa entender — é automático |
| **Binário** | Dados do arquivo "empacotados" para processamento | Não precisa fazer nada — é automático |

---

## ✅ Checklist antes de começar

Antes de rodar o código, confirme que você tem:

- [ ] **Tabela `TabelaCaminho`** criada no Excel com pelo menos uma linha
- [ ] **Coluna `CaminhoArquivo`** preenchida com um caminho válido
- [ ] **Pasta chamada `Base`** existindo naquele local
  - ✓ `C:\Documentos\Base` ← válido
  - ✓ `\\servidor\pasta\Base` ← válido
  - ✓ `https://empresa.sharepoint.com/sites/financeiro/documentos/Base` ← válido
  - ✗ `C:\Documentos` ← **inválido** (falta `\Base` no final)
- [ ] **Arquivo com "TBL_Manifesto_Resultado" no nome** dentro da pasta Base
  - ✓ `TBL_Manifesto_Resultado.xlsx` ← válido
  - ✓ `TBL_Manifesto_Resultado_2024.xlsx` ← válido
  - ✗ `Manifesto.xlsx` ← **inválido** (não tem "TBL_Manifesto_Resultado")
- [ ] **Aba chamada "TBL_Manifesto_Resultado"** dentro do arquivo
- [ ] **Conectividade** (se usar nuvem):
  - Estar conectado à internet
  - Ter acesso ao SharePoint/OneDrive da empresa
  - Estar logado no Excel com a conta corporativa

---

## 🆘 Erros comuns e soluções

### ❌ "Tabela 'TabelaCaminho' não encontrada"

**O que significa:** O código procura por uma tabela chamada `TabelaCaminho` e não achou.

**Como resolver:**
1. Verifique se a tabela está realmente criada no Excel
2. **Atento:** O nome tem que ser exatamente `TabelaCaminho` (maiúscula/minúscula importa)
3. Se a tabela está em outra planilha, tudo bem — Power Query acessa qualquer planilha

---

### ❌ "O caminho informado está em branco"

**O que significa:** Você preencheu a tabela, mas a coluna `CaminhoArquivo` está vazia.

**Como resolver:**
1. Abra a tabela `TabelaCaminho`
2. Preencha a célula na coluna `CaminhoArquivo` com um caminho válido
3. Atualizar o Power Query novamente

---

### ❌ "Pasta local inacessível: C:\Documentos\Base"

**O que significa:** O código conseguiu ler que é um caminho local, mas não conseguiu acessar aquela pasta. Pode ser:
- A pasta não existe
- O caminho está errado/digitado errado
- Sem permissão para acessar
- A unidade `C:\` não existe neste computador

**Como resolver:**
1. **Abra o Explorador de Arquivos** (Windows) ou Finder (Mac)
2. **Navegue até aquela pasta manualmente** — se você conseguir abrir, o caminho é válido
3. **Copie o caminho exato** da barra de endereço e cole na tabela
4. **Verifique as permissões** — você tem direito de ler/abrir arquivos de lá?

---

### ❌ "Nenhum arquivo Excel contendo 'TBL_Manifesto_Resultado' foi encontrado"

**O que significa:** A pasta `Base` foi encontrada, mas não há arquivo com esse nome dentro dela.

**Como resolver:**
1. **Abra a pasta manualmente** (use Explorador de Arquivos)
2. **Procure por "TBL_Manifesto_Resultado"** — é assim que o arquivo começa?
3. **Se o arquivo tem outro nome**, renomeie para incluir "TBL_Manifesto_Resultado"
   - ✓ Exemplo de nome válido: `TBL_Manifesto_Resultado_janeiro_2024.xlsx`
4. **Se tem múltiplos arquivos com esse nome**, deixe apenas um

---

### ❌ "Mais de um arquivo foi encontrado"

**O que significa:** Há **dois ou mais arquivos** dentro da pasta `Base` que contêm "TBL_Manifesto_Resultado" no nome.

**Como resolver:**
1. **Abra a pasta `Base` manualmente**
2. **Procure por todos os arquivos que começam com "TBL_Manifesto_Resultado"**
3. **Apague ou renomeie** os arquivos que não são o atual
4. **Deixe apenas um** arquivo com esse nome na pasta

Exemplo:
```
✓ TBL_Manifesto_Resultado.xlsx          ← DEIXE este
✗ TBL_Manifesto_Resultado_backup.xlsx   ← DELETE este
✗ TBL_Manifesto_Resultado_2023.xlsx     ← DELETE este (ou renomeie)
```

---

### ❌ "URL inválida" ou "URL sem '/sites/' nem '/personal/'"

**O que significa:** Você passou um link de SharePoint/OneDrive, mas o código não conseguiu ler.

**Como resolver:**

**Se é SharePoint corporativo:**
```
✓ https://empresa.sharepoint.com/sites/financeiro/documentos/Base
✓ https://empresa.sharepoint.com/sites/financeiro
✓ https://meu-grupo.sharepoint.com/sites/projeto
```

**Se é OneDrive corporativo:**
```
✓ https://empresa-my.sharepoint.com/personal/joao_empresa_com/Documents/Base
✓ https://empresa-my.sharepoint.com/personal/maria_empresa_com
```

**Se é OneDrive pessoal (consumer):**
```
✗ https://onedrive.live.com/...        ← NÃO FUNCIONA (este não é suportado)
```

---

### ❌ "Falha ao baixar o arquivo"

**O que significa:** O código encontrou o arquivo no SharePoint/OneDrive, mas não conseguiu baixá-lo.

**Causas possíveis:**
- Sem internet
- Sem acesso/permissão para aquele arquivo
- Problema temporário com o servidor

**Como resolver:**
1. **Verifique sua internet** — está conectado?
2. **Abra o link manualmente no navegador** — consegue abrir?
3. **Verifique as permissões** — seu chefe/TI deu acesso àquele arquivo?
4. **Tente novamente** — pode ser erro temporário do servidor

---

### ❌ "A aba 'TBL_Manifesto_Resultado' não foi encontrada"

**O que significa:** O arquivo foi encontrado, mas não tem uma aba/planilha com esse nome dentro dele.

**Como resolver:**
1. **Abra o arquivo manualmente**
2. **Procure pela aba** (as abas ficam embaixo do Excel, à esquerda)
3. **Se existe mas com outro nome**, renomeie a aba para `TBL_Manifesto_Resultado`
4. **Se a aba não existe**, crie uma com esse nome e move os dados para lá

---

## 💡 Dicas práticas

### ✨ Dica 1: Testar o caminho antes de rodar o código

```
1. Copie o caminho que você vai usar
2. Abra um Explorador de Arquivos (Windows) ou Finder (Mac)
3. Cole o caminho na barra de endereço
4. Pressione Enter
5. Se conseguir abrir a pasta, o caminho está correto
```

### ✨ Dica 2: Usar compartilhamentos de rede

Se o arquivo fica em um compartilhamento de rede (um "drive" da empresa):

```
✓ Use o caminho UNC: \\servidor\compartilhado\Base
✓ Ou mapeie o compartilhamento a uma letra (Z:\, etc.) e use: Z:\Base
```

### ✨ Dica 3: Atualizar dados automaticamente

Você pode configurar o Excel para **atualizar automaticamente** quando abre o arquivo:

1. `Dados` → `Atualizar Tudo` (ou `Conexões de Dados`)
2. Clique em um resultado do Power Query
3. `Dados` → `Propriedades da Conexão`
4. Marque: `Atualizar este arquivo quando for aberto`

### ✨ Dica 4: Usar nomes curtos para caminhos

Se seu caminho tem **espaços ou caracteres especiais**, tudo bem — o código processa normalmente. Exemplo:

```
C:\Meus Documentos\Relatório Final\Base  ← Funciona!
C:\Docs\R$ e €\Base                      ← Funciona!
```

---

## 📞 Precisa de ajuda? Checklist de diagnóstico

Se algo deu errado, tente isso nessa ordem:

1. **Erro é sobre a tabela `TabelaCaminho`?**
   - [ ] Tabela existe e se chama exatamente `TabelaCaminho`?
   - [ ] Tem pelo menos 1 linha (além do cabeçalho)?
   - [ ] A coluna se chama `CaminhoArquivo`?

2. **Erro é sobre o caminho/pasta?**
   - [ ] O caminho está preenchido e não está em branco?
   - [ ] A pasta realmente existe naquele local?
   - [ ] Você tem permissão para acessar?
   - [ ] Se é na nuvem, tem internet?

3. **Erro é sobre o arquivo?**
   - [ ] O arquivo tem "TBL_Manifesto_Resultado" no nome?
   - [ ] O arquivo está dentro da pasta `Base`?
   - [ ] Tem apenas UM arquivo com esse nome (não vários)?
   - [ ] A aba dentro do arquivo se chama `TBL_Manifesto_Resultado`?

4. **Se chegou até aqui e ainda não funciona:**
   - Anote a mensagem de erro exata
   - Tire um screenshot
   - Contate suporte técnico com essas informações

---

## 📚 Glossário rápido

| Palavra | Significa |
|---|---|
| **Aba** | Abas/planilhas do Excel (aquelas abas embaixo) |
| **Célula** | Um "quadradinho" do Excel com um dado |
| **Coluna** | Vertical (A, B, C, etc.) |
| **Linha** | Horizontal (1, 2, 3, etc.) |
| **Tabela** | Dados organizados em linhas e colunas |
| **Power Query** | Ferramenta do Excel que busca, processa e carrega dados |
| **SharePoint** | Serviço da Microsoft para empresas armazenar/compartilhar arquivos |
| **OneDrive** | Seu armazenamento pessoal na nuvem (igual Google Drive) |
| **Caminho** | O endereço de um arquivo (tipo: `C:\Documentos\arquivo.xlsx`) |
| **Binário** | Dados "compactados" (não precisa saber detalhes) |

---

## ✅ Você está pronto!

Siga o **Checklist antes de começar** acima e tente. Se ficar preso, consulte a seção **🆘 Erros comuns**.

**Sucesso! 🎉**
