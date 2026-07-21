Agent_Main: the current working Morning Clips pipeline.

Architecture: single agent, multiple inline prompts. One agent runs three
AI Builder prompt tools in sequence within a single Main topic -
Extractor (PDF -> raw text via Code Interpreter) -> Parser (raw text ->
===ARTICLE=== blocks) -> Reporter (blocks -> formatted report).

Why this design: the earlier multi-agent approach (separate connected
agents per step) hit a platform blocker - connecting agents strips the
data bindings between them ("variable data type not eligible to receive
or return values"), Global variables don't survive crossing agent
boundaries ("Identifier not recognized"), and child agents relay return
values to the parent, causing duplicate outputs. Consolidating the three
steps into one agent as inline prompt tools keeps the data bindings
intact and deterministic; prompt tools return silently into variables.
Classic orchestration (not generative) keeps runs repeatable.

Not the same as Agent_Main_TestV - that was an earlier abandoned attempt.
