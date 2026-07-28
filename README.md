# minuta-decreto

Ferramenta de automação para geração de **minutas de decreto de crédito
adicional** (suplementação/anulação) a partir dos arquivos exportados do
SIAFI, com **conferência automática dos incisos** de saldo/excesso e
verificações de consistência antes da geração.

O script unifica em um único fluxo duas etapas que antes eram executadas
separadamente:

1. **Conferência** — lê os `.docx` da pasta da minuta, extrai os trechos de
   *"do saldo"* / *"do excesso" ... "no valor de R$..."*, consolida os valores
   por chave e grava o `conferencia.xlsx`.
2. **Geração** — a partir do `conferencia.xlsx` e do arquivo `Minuta*.TXT`
   (exportado do SIAFI), preenche o template e produz o `.docx` final da minuta.

Entre as duas etapas o operador confirma se deve prosseguir (`s`/`n`). Se
responder `n`, o processo encerra com o `conferencia.xlsx` já salvo.

---

## Fluxo

```
Informe a subpasta  ─►  Etapa 1: Conferência  ─►  "Prosseguir? [s/n]"
                                                        │
                                          n ◄───────────┴───────────► s
                                          │                           │
                                    encerra (xlsx salvo)      Etapa 2: Geração
                                                                      │
                                                        verificações  │  gera minuta N.docx
                                                        (só param se houver algo a revisar)
```

### Verificações da Etapa 2

Executadas automaticamente antes de gravar o `.docx`. Só interrompem pedindo
confirmação quando encontram algo — caso contrário, passam silenciosamente:

- **Fontes especiais** — sinaliza a presença de recursos que exigem atenção:
  - IPU `0` ou fonte `89` → **PROPAG**
  - fonte `95` → **Brumadinho**
  - fonte `80` → **Mariana**
- **Balanço** — confere se `Suplementação (Art. 1º) == Anulação (TXT) + Incisos
  (conferencia.xlsx)`, com tolerância de 1 centavo, avisando se o Art. 2º ficou
  a maior ou a menor.

---

## Requisitos

- Python 3.9+
- Dependências em `requirements.txt`:

```bash
pip install -r requirements.txt
```

## Configuração dos caminhos

O caminho-raiz dos arquivos de decreto **não** fica fixo no código. A resolução
é, nesta ordem:

1. Variável de ambiente `MINUTA_BASE`, se definida.
2. Padrão: `<pasta do usuário>\OneDrive - CAMG\General\@dcmefo\<ano>\Decretos`
   (o `<pasta do usuário>` é resolvido automaticamente por máquina).

Para apontar para outro local:

```powershell
# Windows (persistente)
setx MINUTA_BASE "C:\Users\SEU_ID\OneDrive - CAMG\General\@dcmefo\2026\Decretos"
```

```bash
# Linux / WSL
export MINUTA_BASE="/mnt/c/Users/SEU_ID/OneDrive - CAMG/General/@dcmefo/2026/Decretos"
```

O template usado é `<BASE>\Manual decreto\minutas padrão\minuta xxx.docx`.

### Estrutura esperada das pastas

```
<MINUTA_BASE>/
├── Manual decreto/
│   └── minutas padrão/
│       └── minuta xxx.docx          # template
└── Minutas Finalizadas/
    └── 05 maio/
        └── minuta 83/
            ├── Minuta0083.TXT       # export do SIAFI
            ├── <docx de trabalho>   # lidos na conferência
            ├── conferencia.xlsx     # gerado na Etapa 1
            └── minuta 83.docx       # gerado na Etapa 2
```

---

## Uso

Informando a subpasta interativamente:

```bash
python gerar_minuta.py
# Informe a subpasta da minuta (ex: 05 maio\minuta 83):
# > 05 maio\minuta 83
```

Ou passando a subpasta como argumento:

```bash
python gerar_minuta.py "05 maio\minuta 83"
```

---

## Observações

- Os artefatos gerados (`conferencia.xlsx`, `minuta *.docx`, temporários
  `~$...`) estão no `.gitignore` e não devem ser versionados.
- Se não houver incisos de saldo/excesso, o `conferencia.xlsx` é criado apenas
  com o cabeçalho e a minuta é montada no formato de Art. 2º sem incisos.
- O número da minuta é extraído do nome do `Minuta*.TXT` (ex.: `Minuta0083.TXT`
  → `minuta 83.docx`).
