# Optional PandasAI Integration

Sidecar for Excel remains a native Office.js task-pane add-in. [PandasAI](https://github.com/Sinaptik-AI/pandas-ai) is an optional **Hermes-side analysis tool** for requests that are awkward to express with workbook formulas, tables, pivots, or the add-in's deterministic range statistics.

PandasAI is deliberately not bundled into the task pane. It generates and executes Python, has a large Python dependency tree, and runs outside the Office WebView. Hermes can send a trusted CSV, workbook, or Parquet file to the wrapper, interpret its JSON response, and then write requested results back through the existing constrained Excel tools.

## Appropriate uses

Use PandasAI for:

- exploratory analysis and anomaly investigation;
- joins or comparisons across multiple exported datasets;
- data cleansing and feature generation;
- Python/Pandas transformations that would be brittle as Excel formulas;
- analytical charts that do not need to remain native Excel chart objects.

Keep these operations in the add-in's native Office.js tools:

- reading and writing workbook ranges;
- live Excel formulas;
- workbook formatting, tables, filters, and sorting;
- native charts and pivot tables;
- sheet creation, protection, and comments;
- recalculation and post-write verification.

## Installation

PandasAI 3 requires Python 3.8–3.11. Use an isolated Python 3.11 environment rather than the add-in's Node dependency tree:

```bash
python3.11 -m venv ~/.hermes/tools/pandasai/.venv
~/.hermes/tools/pandasai/.venv/bin/python -m pip install \
  -r scripts/pandasai-requirements.txt
```

With `uv`:

```bash
uv venv --python python3.11 ~/.hermes/tools/pandasai/.venv
uv pip install \
  --python ~/.hermes/tools/pandasai/.venv/bin/python \
  -r scripts/pandasai-requirements.txt
```

Copying the wrapper is optional; it can run directly from this repository.

## Usage

PandasAI uses LiteLLM. For an OpenAI-compatible local endpoint such as LM Studio, prefix the served model ID with `openai/`:

```bash
SCARF_NO_ANALYTICS=true \
PANDASAI_MODEL='openai/<served-model-id>' \
PANDASAI_API_BASE='http://<local-host>:<port>/v1' \
PANDASAI_API_KEY='lm-studio' \
~/.hermes/tools/pandasai/.venv/bin/python \
  scripts/pandasai-analyze.py workbook.xlsx \
  'Identify unusual values and summarize the main trends' \
  --sheet Sheet1
```

Supported inputs are CSV, XLSX/XLSM, and Parquet. The wrapper writes one JSON response to standard output and structured errors to standard error. It defaults to `temperature=0` and disables PandasAI's Scarf analytics.

Example successful response:

```json
{"ok": true, "type": "str", "value": "The largest month-over-month increase was in March."}
```

## Security boundary

PandasAI executes LLM-generated Python. That is substantially more powerful—and riskier—than the add-in's allowlisted Office.js tools.

- Use the wrapper only with trusted prompts and trusted files.
- Prefer a local model endpoint for confidential workbooks.
- Never send payroll, claims, health, financial, customer, or other confidential data to a cloud model without explicit authorization.
- Do not pass secrets in command-line arguments. Use environment variables or the existing local model's dummy key.
- For public-facing, untrusted, sensitive production, or multi-tenant usage, use PandasAI's official `pandasai-docker` sandbox. The simple wrapper in this repository is not a security sandbox.
- Do not allow PandasAI to write directly into a live workbook. Hermes should review the returned result, write through Office.js, recalculate, and read back the affected range.

## Architecture

```text
Workbook or exported range
          |
          v
      Hermes Agent
          |
          +---- ordinary workbook work ----> Office.js tools
          |
          +---- deep analysis -------------> PandasAI/Python
                                                  |
                                                  v
                                        structured JSON result
                                                  |
                                                  v
                               Hermes writes and verifies in Excel
```

This separation keeps the add-in small, auditable, and native while making deeper analysis available when it is genuinely useful.

## Licensing

PandasAI's open-source code is MIT-licensed except for code under its enterprise directory, which has separate production terms. This repository installs the published open-source packages and does not copy or bundle PandasAI source code.
