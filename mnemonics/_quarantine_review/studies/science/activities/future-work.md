# Future work

## High context

🔲 Maybe remove the broswering from the main agent (high context).

🔲 The broswering evidence shoulde have more granuar evidences separeted by a dot. This shoulde be done in the validate file too.

🔲 Test the MCP browseing.

🔲 Look over the dataset file. There is a yellow line. I marked it in yellow.

🔲 Finish. Look looking over the main file.

🔲 All evidence should be in the validate file.

🔲 Web methods additional comments:

33 - Check in dead linke checkers
62 - Perforamance
![img](images/20251006-194723.png)

🔲⚡ Go to the data set md file and update the methods used. I marked in yellow where is necessary to take a look.

🔲 Each prompt should have its own ID.

🔲 Can we instruct the AI to browse to assessment tools, such as accessibility or performance tools, for evaluation?
For example, to check attributes like responsiveness, broken links, performance, the HTTPS indicator, etc.

🔲 The MCP browser should also perform these actions on the replayed web page.

🔲 We have 26 prompts. Some of these prompts are similar, such as those for navigation or saving data. We need a common source of truth so that when it is changed, all related prompts are updated accordingly.

## Hooks

🔲 Use hooks to create analytics of prompt duration.
You need a hook for the start of the prompt and a hook for the end.
[Link](https://cursor.com/docs/agent/hooks#examples)

Onde hooks ajudam imediatamente

Telemetria de prompts - duração, custo estimado, tokens, tool usage.

Guarda-chuvas de evidência - exigir 3+ screenshots e diretórios whitelisted.

Persistência padronizada - escrever/validar CSVs e atualizar SQL após cada execução.

Redução de tokens - sumarizar automaticamente contextos grandes antes de enviar ao modelo.

Reprodutibilidade - carimbar run_id, website_id, scenario_type, ai_model, mcp_server.

Qualidade - checagem de esquema dos outputs e “fail fast” com relatório.

Retry controlado - backoff para falhas de navegação/auditoria.

Auditoria - logar Console/Network e anexar ao relatório.

### Rules

🔲 Create a project scoped rule.
[Rules](https://cursor.com/docs/context/rules)

This rule should always be applied.

The rules should have the content that will go through all prompts context.

**Best practices**
Good rules are focused, actionable, and scoped.

- Keep rules under 500 lines
- Split large rules into multiple, composable rules
- Provide concrete examples or referenced files
- Avoid vague guidance. Write rules like clear internal docs
- Reuse rules when repeating prompts in chat

You can use this agents.md file as a plan b:
![img](images/20251001-222406.png)
![img](images/20251001-222826.png)

1. Project Rule – Agentic AI Evidence & Data Integrity

Name: Agentic AI Evidence & Data Integrity
Scope: Project
Applies to: todos os prompts/agentes deste repo (Agentic + Reactive)

Policy (must/should)

Data outputs

MUST escrever resultados somente em data/ e sessions/.

MUST manter CSVs append-only:

data/agentic_ai_experiment.csv

data/agentic_ai_questions.csv

MUST rodar o validador após cada append e abortar se falhar.

Metadata canônica (por execução)

MUST preencher: website_id, scenario_type, ai_model, mcp_server, start_time, end_time, duration.

scenario_type ∈ {reactive-ai,empyric,agentic}

Se scenario_type = agentic: mcp_server ∈ {@agentdeskai/playwright-mcp,@agentdeskai/browser-tools-mcp}

Evidências

MUST capturar ≥ 3 screenshots por execução.

MUST salvar logs de Console e Network quando disponíveis.

MUST usar apenas diretórios whitelisted:

sessions/**/screenshots/, sessions/**/logs/, sessions/**/analysis/, sessions/**/styles/, sessions/**/structure/, sessions/**/questions/

MUST bloquear escrita fora desses diretórios.

Ambiente & Cenários

MUST não acessar web “live” quando cenário for replay (WACZ).

SHOULD registrar website_version e website_source (replay|live) no nível do experimento.

Token efficiency

SHOULD resumir automaticamente documentos contextuais longos (>6k chars) antes do invoke.

SHOULD reutilizar sources of truth compartilhados (dataset/métodos/validator), sem duplicar texto nos prompts.

Qualidade & Confiabilidade

MUST “fail fast” se: screenshots < 3, CSV inválido, diretório fora de whitelist, ou campos canônicos ausentes.

SHOULD aplicar retry exponencial (max 2) para falhas de navegação/auditoria.

SHOULD registrar run_report.json com telemetria: run_id, duration_ms, tokens_delta, contagem de screenshots.

## Medium context
