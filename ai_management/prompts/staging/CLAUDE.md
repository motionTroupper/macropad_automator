## Alter Memory, knowledge or skills
- **A rule you learn** goes to `ai_management/prompts/proposed/` of the repository being worked on — its main one
  when several are used, `~/.claude/ai_management/prompts/proposed/` when there is none
- **One concept per file**: imperative, no justification, no dates. Add the minimum, "telegraphic" context, and only when it's valuable to follow the imperative part. If it supersedes arule, name it.
- **NEVER edit a `CLAUDE.md`, a `SKILL.md` or anything under a `.claude/skills` folder.** They are
  generated from atoms. Write the change to `ai_management/prompts/proposed/` instead.
- **NEVER edit a memory file or a `MEMORY.md` directly.** They are generated from
  `ai_management/memories/atoms`. A fact you learn goes to `ai_management/memories/proposed/` of the
  repository being worked on — its main one when several are open — one concept per file, in memory
  shape: frontmatter plus prose. A rule of behaviour goes to `ai_management/prompts/proposed/` instead.
- **Memory holds facts, never rules.** `memory/` and `MEMORY.md` take context only — state of a
  project, who is who, where something lives. A rule of behaviour is proposed instead, even
  when asked for it as a memory.
- **What YOU propose goes under `ai_proposed`.** A rule goes to
  `ai_management/prompts/proposed/ai_proposed/<what-it-holds>/` of the repository being worked on, a
  fact to `ai_management/memories/proposed/ai_proposed/<what-it-holds>/`, making the folders that do
  not exist yet. Name the inner folder for what it holds, NEVER for the session or the date. What a
  hand proposes stays out of `ai_proposed`.

## Knowledge cultivation
You are running on a knowledge-cultivation mode to efficiently enrich your skills and memories.

### The breakpoint

Investigating the repository is free. Delivering what was learned is not.
- **Investigating is free** — read, grep and search as the task needs. Asking permission to look
  is not required.
- **Stop before using what was learned** — when producing the action, the answer or the plan
  required finding out how the repository is structured, how it operates or what it does: write the
  atoms, link them, and **WAIT**. It holds when the delivery writes nothing.
- **One atom per citable statement** — one file per statement that stands on its own and serves a
  task other than this one. Statements that only ever apply together are ONE atom. NEVER summarise
  a file.
- **Link what was written** — ALWAYS write the atom to its file and link the file. NEVER propose one
  in prose.
- **Provenance stays in the message** — say what was read to assert each atom, beside its link.
  NEVER inside the atom.

### What goes up
- **Only what the answer rests on** — write the atoms the delivery depends on. What was read on the
  way and carries nothing stays out.
- **Check the tree first** — search the atom tree and what is proposed before writing. NEVER write again what
  is already written. When the finding contradicts an applied atom, name that atom and **WAIT**.
- **A fact is a memory, a rule is a prompt** — the state of the repository, who is who and where
  something lives go to `ai_management/memories/proposed/ai_proposed/`. How to act goes to
  `ai_management/prompts/proposed/ai_proposed/`.
- **Only permanent atoms go up** — write what still holds once the conversation ends. Transient
  state — what the next build, commit or fix erases — stays out.

### Resuming
- **Resume from what was promoted** — on continuing, read the promoted atom, NEVER the draft left
  proposed. When nothing was promoted, ask before acting on the draft.
