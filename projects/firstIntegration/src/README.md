# 📋 Gerador de Relatórios de Solicitações

Macro VBA para Excel que integra banco de dados SQL Server, geração de documentos Word e exportação em PDF — tudo em um único clique.

---

## 📌 O que o projeto faz

1. Lê o número do sócio informado na planilha `pesquisa`
2. Conecta ao banco de dados e busca a foto do sócio (armazenada como BLOB)
3. Converte o BLOB em arquivo `.jpg` e exibe na planilha
4. Abre um template Word (`.docx`) e preenche os bookmarks com os dados da planilha
5. Insere a foto do sócio no documento Word
6. Exporta o documento preenchido como `.pdf` na pasta de destino

---

## 🗂️ Estrutura do Projeto

```
VBAProject (base_pesquisa.xls)
│
├── Planilhas
│   ├── pesquisa          ← planilha principal de entrada de dados
│   └── ...               ← demais planilhas do sistema
│
└── Módulos
    ├── Main              ← ponto de entrada único (botão da planilha)
    ├── DatabaseService   ← conexão e consulta ao banco de dados
    ├── ImageService      ← conversão de BLOB, exibição na planilha
    └── WordService       ← preenchimento do template Word e exportação PDF
```

---

## ⚙️ Pré-requisitos

- Microsoft Excel (recomendado 2016 ou superior)
- Microsoft Word instalado na mesma máquina
- Acesso à rede onde o banco de dados SQL Server está hospedado
- Driver **SQLOLEDB** ou **MSOLEDBSQL** instalado
- Referências VBA habilitadas:
  - `Microsoft ActiveX Data Objects` (ADO) — para conexão com banco
  - `Microsoft Scripting Runtime` — para uso do `Dictionary`

### Como habilitar as referências

No Editor VBA (`Alt + F11`):

```
Ferramentas → Referências → marcar:
  ✅ Microsoft ActiveX Data Objects X.X Library
  ✅ Microsoft Scripting Runtime
```

---

## 🔧 Configuração antes de usar

### 1. String de conexão — `DatabaseService`

Localize a constante `CONN_STRING` e preencha com os dados do seu ambiente:

```vb
Private Const CONN_STRING As String = _
    "Provider=SEU_PROVIDER;" & _
    "Data Source=SEU_SERVIDOR;" & _
    "Initial Catalog=SEU_BANCO;" & _
    "User ID=SEU_USUARIO;" & _
    "Password=SUA_SENHA;"
```

| Campo | Descrição |
|---|---|
| `SEU_PROVIDER` | Ex: `SQLOLEDB` ou `MSOLEDBSQL` |
| `SEU_SERVIDOR` | IP ou nome do servidor SQL na rede |
| `SEU_BANCO` | Nome do banco de dados |
| `SEU_USUARIO` | Usuário com permissão de leitura |
| `SUA_SENHA` | Senha do usuário |

---

### 2. Consulta SQL — `DatabaseService.BuscarFotoSocio`

Ajuste a query com o nome real da tabela e das colunas:

```vb
sql = "SELECT SUA_COLUNA_FOTO FROM SUA_TABELA WHERE SUA_COLUNA_ID = " & numSocio
```

---

### 3. Caminhos de arquivo — `ImageService` e `WordService`

Atualize as constantes com os caminhos reais da sua rede:

**`ImageService`:**
```vb
Private Const CAMINHO_FOTO As String = "CAMINHO_DE_REDE\fotos-arquivadas\foto.jpg"
```

**`WordService`:**
```vb
Private Const CAMINHO_TEMPLATE As String = "CAMINHO_DE_REDE\template_solicitacoes.docx"
Private Const CAMINHO_PDF As String = "CAMINHO_DE_REDE\destino-pdfs\"
```

---

### 4. Template Word

O arquivo `.docx` deve conter **bookmarks** com os seguintes nomes exatos:

| Bookmark | Célula na planilha `pesquisa` |
|---|---|
| `data_solicitacao` | C5 |
| `num_solicitacao` | A5 |
| `nome_socio` | H5 |
| `num_socio` | G5 |
| `celular_socio` | J5 |
| `email_socio` | I5 |
| `assunto_solicitacao` | L5 |
| `tipo_solicitacao` | L5 |
| `status` | F5 |
| `texto_solicitacao` | K5 |
| `foto_socio` | *(imagem do banco)* |

#### Como criar um bookmark no Word
```
Selecione o texto/espaço → Inserir → Indicador → digite o nome → Adicionar
```

---

## ▶️ Como executar

1. Abra o arquivo `base_pesquisa.xls`
2. Navegue até a planilha **`pesquisa`**
3. Preencha o número do sócio na célula **`B5`**
4. Clique no botão **"Gerar Relatório"** — ele chama `Main.GerarRelatorio`

O macro irá automaticamente:
- Buscar e exibir a foto na célula `Q5`
- Preencher e exportar o PDF na pasta configurada

---

## 📁 Arquivos gerados

| Arquivo | Localização |
|---|---|
| `foto.jpg` | Pasta de fotos configurada em `ImageService` |
| `SOLICITAÇÃO_<nome>.pdf` | Pasta de PDFs configurada em `WordService` |

---

## 🐛 Solução de problemas comuns

| Erro | Causa provável | Solução |
|---|---|---|
| *Falha na conexão com o banco* | String de conexão incorreta ou sem acesso à rede | Verifique `CONN_STRING` e conectividade |
| *Foto do sócio não encontrada* | Sócio sem foto cadastrada ou ID incorreto | Confirme o valor em `B5` e o registro no banco |
| *Bookmark não encontrado* | Nome do bookmark no Word diferente do mapeado | Confira os nomes em `MapaBookmarks` e no template |
| *Erro ao abrir template* | Caminho do `.docx` incorreto ou arquivo em uso | Verifique `CAMINHO_TEMPLATE` e feche o Word |
| *Referência ADO ausente* | Biblioteca ADO não habilitada | Habilite em Ferramentas → Referências |

---
