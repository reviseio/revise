/** The agent skill, printed by `revise skill` and installed as a folder by
 * `revise skill install`: SKILL.md plus reference files read on demand. The
 * text lives in skills/revise-cli/ as Markdown; this is the only place it is
 * published from (the reviseio/skills repo copies the release build's
 * output). Keep SKILL.md short: it loads into an agent's context every time
 * the skill triggers. */
import SKILL from "./skills/revise-cli/SKILL.md?raw";
import LIBRARY from "./skills/revise-cli/reference/library.md?raw";
import JSON_REFERENCE from "./skills/revise-cli/reference/json.md?raw";

export const SKILL_MARKDOWN: string = SKILL;

/** Reference files, by the name `revise skill --reference` takes. */
export const SKILL_REFERENCES: Readonly<Record<string, string>> = {
  library: LIBRARY,
  json: JSON_REFERENCE,
};
