/**
 * Color for human output. Plain text unless a person is likely watching:
 * a stream gets color only when it is a terminal, and never under --json.
 *
 *   First matching rule wins, for stdout and stderr separately:
 *   --json, --no-color, NO_COLOR (non-empty), FORCE_COLOR=0, TERM=dumb → off
 *   FORCE_COLOR (any other value)                       → on, even piped
 *   otherwise                                           → on for a terminal
 *
 * Color never carries meaning on its own: every colored thing also has a
 * word or a sign (+, -, error, warning) that says the same.
 */

export interface Style {
  readonly enabled: boolean;
  bold(text: string): string;
  dim(text: string): string;
  red(text: string): string;
  green(text: string): string;
  yellow(text: string): string;
  cyan(text: string): string;
}

function wrap(open: number, close: number) {
  return (text: string) => (text ? `\x1b[${open}m${text}\x1b[${close}m` : text);
}

const ON: Style = {
  enabled: true,
  bold: wrap(1, 22),
  dim: wrap(2, 22),
  red: wrap(31, 39),
  green: wrap(32, 39),
  yellow: wrap(33, 39),
  cyan: wrap(36, 39),
};

const identity = (text: string) => text;

export const PLAIN: Style = {
  enabled: false,
  bold: identity,
  dim: identity,
  red: identity,
  green: identity,
  yellow: identity,
  cyan: identity,
};

export interface ColorInputs {
  /** Whether the stream is a terminal. */
  isTTY: boolean;
  env: Record<string, string | undefined>;
  /** --no-color was given. */
  noColor: boolean;
  json: boolean;
}

export function colorEnabled({
  isTTY,
  env,
  noColor,
  json,
}: ColorInputs): boolean {
  if (json || noColor || env.NO_COLOR) return false;
  const force = env.FORCE_COLOR;
  if (force === "0" || force === "false") return false;
  if (env.TERM === "dumb") return false;
  if (force !== undefined) return true;
  return isTTY;
}

export function styleFor(inputs: ColorInputs): Style {
  return colorEnabled(inputs) ? ON : PLAIN;
}

/** Remove color codes (for width math and tests). */
export function stripStyle(text: string): string {
  return text.replace(/\x1b\[\d+m/g, "");
}
