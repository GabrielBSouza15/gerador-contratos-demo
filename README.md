# Gerador de contratos demonstrativos

Automação local em Python que seleciona um registro pelo ID de uma planilha e preenche um modelo Word. O projeto apresenta um exemplo de aplicação de Python e Pandas à preparação de documentos, mantendo o fluxo de uso pelo terminal.

A origem do projeto foi uma rotina de geração de contratos a partir de informações recebidas por Google Forms. A planilha deste repositório é uma representação didática desses dados, com uma única aba e quatro registros inteiramente fictícios. Não há conexão com formulários ou contas Google.

## O que o exemplo faz

1. Pergunta o ID do registro, como `CT-001`.
2. Solicita modalidade de pagamento, tipo de pessoa e serviço.
3. Busca exatamente uma linha da aba `Dados`.
4. Formata os campos, calcula valores por extenso e preenche um modelo DOCX.
5. Salva um novo documento na pasta `gerados`, com ID, data e horário no nome.

IDs inexistentes ou duplicados e campos obrigatórios ausentes impedem a geração. O script não gera todos os registros automaticamente e não sobrescreve arquivos existentes.

## Executar

Requer Python 3.10 ou superior. No terminal, entre na pasta deste repositório:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe gerar_contrato.py
```

No Linux/macOS, use `python3` para criar o ambiente e `.venv/bin/python` nos comandos equivalentes. Os arquivos são localizados em relação ao script, sem caminhos pessoais fixos.

Para começar, escolha `CT-001`, pagamento `2` (PIX), pessoa `1` (jurídica) e serviço `1` (Excel). Outra possibilidade é `CT-002`, pagamento `1` (cartão), pessoa `2` (física) e serviço `2` (Power BI).

## Planilha e campos

`dados/Informacoes.xlsx` contém somente a aba `Dados`. Cada linha é um registro independente, identificado por um ID único no formato `CT-001`. Os campos pessoais e empresariais são exemplos sintéticos. Os documentos CPF/CNPJ e os CEPs usam zeros deliberadamente, sem identificação válida de uma pessoa ou empresa. Por isso, os exemplos exibem avisos de validação matemática; o aviso não bloqueia a demonstração.

Para pessoas físicas, use CT-002 ou CT-004. Para pessoas jurídicas, CT-001 ou CT-003 possuem o bloco da empresa preenchido. Digite nacionalidade e estado civil exatamente como devem aparecer no documento. O programa não infere essas informações pelo nome.

Mantenha CPF, CNPJ e CEP como texto para preservar zeros iniciais. `Valor` é numérico, positivo e tem no máximo duas casas decimais. `Prazo` é um inteiro de dias. Na modalidade PIX, o programa calcula uma entrada aproximada de 50% e o saldo como diferença; se houver centavo indivisível, a entrada recebe o centavo adicional. Na modalidade cartão, o exemplo representa pagamento único.

## Uso de IA na preparação dos dados

No fluxo que motivou o projeto, funções `AI.ASK` de um complemento do Excel eram usadas para escrever valores e o mês por extenso. Aqui, a conversão é feita localmente em Python com `num2words` e uma lista de meses, sem exigir complemento, API de IA ou credenciais. O código não envia o conteúdo da planilha a serviços de IA.

Esta edição demonstrativa contou com apoio do OpenAI Codex na preparação, revisão e testes. Isso não atribui ao Codex a autoria da automação histórica nem indica uso de agente dentro do programa.

## Arquivos

- `gerar_contrato.py`: leitura por ID, menus, formatação e geração.
- `criar_modelos.py`: fonte textual dos dois modelos fictícios; permite reconstruí-los.
- `modelos/`: modelos DOCX de PIX e cartão, com campos de substituição.
- `dados/`: planilha sintética com uma aba.
- `exemplos/`: dois documentos preenchidos com registros fictícios.
- `tests/`: testes com documentos e dados temporários.

Os contratos são ilustrativos e não constituem modelos jurídicos para contratação. Partes, condições e textos foram criados para o cenário didático; não representam termos de uma empresa real. Arquivos gerados durante o uso ficam fora do Git.

## Testar

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

É possível reconstruir os modelos com `python criar_modelos.py`. Esse comando substitui somente os dois modelos demonstrativos em `modelos`; não é necessário executá-lo para gerar contratos.
