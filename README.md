# Sistema TCE/PI — PGE-PI

Automação do fluxo de cobrança judicial de débitos apurados pelo **Tribunal de Contas do Estado do Piauí (TCE/PI)**, desenvolvido para a **Procuradoria Geral do Estado do Piauí — Procuradoria Tributária**.

O sistema lê certidões de condenação em PDF, extrai os dados dos responsáveis, gera a petição inicial e a protocola automaticamente no **PJe TJPI** via intercomunicação SOAP (MNI 2.2.2).

---

## Módulos

| Arquivo | Responsabilidade |
|---|---|
| `extrator_certidao.py` | Interface principal — extrai dados do PDF da certidão TCE e gera a petição inicial (.docx / .pdf) |
| `gestor_enderecos.py` | Extrai endereços do PDF de imputação de débito via OCR e os persiste no banco |
| `enviador_peticao.py` | Seleciona a certidão, monta o envelope SOAP e protocola no PJe TJPI |

Os três módulos compartilham um único banco SQLite (`certidoes_tce.db`), criado automaticamente na primeira execução.

---

## Requisitos

### Python 3.12+

```
pip install pdfplumber pdf2image pytesseract python-docx requests
```

### Dependências externas

| Ferramenta | Uso | Instalação |
|---|---|---|
| **Tesseract OCR** | Leitura de PDFs com encoding privado | [github.com/tesseract-ocr](https://github.com/tesseract-ocr/tesseract) |
| **Poppler** | Conversão PDF→imagem (pdf2image) | Windows: [github.com/oschwartz10612/poppler-windows](https://github.com/oschwartz10612/poppler-windows) |
| **LibreOffice** | Conversão .docx → .pdf headless | [libreoffice.org](https://www.libreoffice.org/) |

> **Tessdata em português:** coloque `por.traineddata` na pasta `tessdata/` junto aos executáveis do Poppler (o sistema detecta automaticamente) ou no diretório padrão do Tesseract.

---

## Execução

```bash
# Interface principal (extração + geração de petição)
python extrator_certidao.py

# Carregar certidão direto pela linha de comando
python extrator_certidao.py certidao.pdf
python extrator_certidao.py certidao.pdf planilha.pdf

# Módulo de envio ao PJe (standalone)
python enviador_peticao.py
python enviador_peticao.py certidoes_tce.db   # banco em caminho específico

# Demo do parser de endereços
python gestor_enderecos.py
```

---

## Fluxo de uso

```
1. Abrir extrator_certidao.py
2. Carregar PDF da certidão TCE/PI  →  dados extraídos automaticamente
3. Carregar PDF de imputação de débito  →  endereços extraídos via OCR
4. Sistema gera petição inicial (.docx + .pdf)
5. Assinar o PDF digitalmente (token / certificado A3 — fora do sistema)
6. Abrir enviador_peticao.py
7. Selecionar a certidão e a comarca
8. Selecionar o PDF assinado
9. Confirmar envio  →  protocolo registrado no banco com número do processo PJe
```

---

## Banco de dados

Arquivo: `certidoes_tce.db` (SQLite, criado automaticamente)

| Tabela | Conteúdo |
|---|---|
| `certidoes` | Uma linha por certidão processada (processo, valor, acórdão, caminhos dos PDFs) |
| `responsaveis` | Réus vinculados à certidão (nome, CPF/CNPJ, acórdão individual) |
| `enderecos_responsavel` | Endereços estruturados extraídos via OCR, indexados por CPF/CNPJ |
| `cod_comarcas` | Mapeamento comarca → código de localidade PJe + competência da vara |
| `peticoes_enviadas` | Log completo de todos os envios ao PJe (sucesso e erro) |

---

## Configuração de comarcas

O módulo de envio inclui uma interface **Gerenciar Comarcas** para cadastrar o par `(cod_comarca, competencia)` de cada localidade do TJPI. Os valores são consultados via:

```
POST https://pje.tjpi.jus.br/1g/ConsultaPJe
operação: consultarCompetencias  |  arg0.id = <cod_comarca>
```

Comarcas pré-cadastradas: Altos, Barras, Campo Maior, Corrente, Esperantina, Floriano, Oeiras, Parnaíba, Picos, Piripiri, São Raimundo Nonato, Teresina, União e Uruçuí.

---

## Integração PJe TJPI

- **Endpoint:** `https://pje.tjpi.jus.br/1g/intercomunicacao?wsdl`
- **Protocolo:** SOAP / MNI intercomunicação 2.2.2
- **Polo Ativo:** Estado do Piauí — CNPJ `06.553.481/0001-49`
- **Classe processual:** 1116 (Execução Fiscal)
- **Assunto CNJ:** 10872

> **Nota sobre `tipoDocumento`:** o PJe TJPI usa o código `CMF` tanto para CPF quanto para CNPJ. O campo `emissorDocumento` diferencia: `MF` para CPF e `SRFB` para CNPJ. `codigoDocumento` deve ter pontuação (`152.308.643-20`); `numeroDocumentoPrincipal` sem pontuação (`15230864320`).

---

## Estrutura de arquivos

```
.
├── extrator_certidao.py     # Módulo principal
├── gestor_enderecos.py      # Extração OCR de endereços
├── enviador_peticao.py      # Envio ao PJe TJPI
├── certidoes_tce.db         # Banco de dados (gerado automaticamente)
├── tessdata/                # Modelos do Tesseract (por.traineddata)
│   └── por.traineddata
└── README.md
```

---

## Observações técnicas

- PDFs do TCE/PI usam **fonte com encoding privado** (`(cid:XX)`). O sistema usa pdfplumber para detectar a estrutura de tabelas e Tesseract OCR para ler o conteúdo célula a célula (250 DPI, `por+eng`).
- A tabela `enderecos_responsavel` pode ter múltiplas entradas para o mesmo CPF/CNPJ (um por processo TCE). O polo passivo usa `LIMIT 1` por subquery para garantir exatamente um réu por entrada no XML.
- O banco é compatível com versões anteriores: colunas adicionadas em atualizações são criadas via `ALTER TABLE` na abertura da conexão.

---

*PGE-PI / Procuradoria Tributária — 2026*
