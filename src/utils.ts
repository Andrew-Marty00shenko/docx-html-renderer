export function escapeClassName(className: string) {
  return className?.replace(/[ .]+/g, "-").replace(/[&]+/g, "and").toLowerCase();
}

export function encloseFontFamily(fontFamily: string): string {
  return /^[^"'].*\s.*[^"']$/.test(fontFamily) ? `'${fontFamily}'` : fontFamily;
}

export function splitPath(path: string): [string, string] {
  const si = path.lastIndexOf("/") + 1;
  const folder = si == 0 ? "" : path.substring(0, si);
  const fileName = si == 0 ? path : path.substring(si);

  return [folder, fileName];
}

export function resolvePath(path: string, base: string): string {
  try {
    const prefix = "http://docx/";
    const url = new URL(path, prefix + base).toString();
    return url.substring(prefix.length);
  } catch {
    return `${base}${path}`;
  }
}

export function keyBy<T>(
  array: T[],
  by: (x: T) => string | number | null | undefined,
): Record<string, T> {
  return array.reduce(
    (a, x) => {
      const key = by(x);
      if (key != null) {
        a[key.toString()] = x;
      }
      return a;
    },
    {} as Record<string, T>,
  );
}

export function blobToBase64(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result as string);
    reader.onerror = () => reject();
    reader.readAsDataURL(blob);
  });
}

export function isObject(item: unknown): item is Record<string, unknown> {
  return item && typeof item === "object" && !Array.isArray(item);
}

export function isString(item: unknown): item is string {
  return typeof item === "string" || item instanceof String;
}

export function mergeDeep(
  target: Record<string, unknown>,
  ...sources: Record<string, unknown>[]
): Record<string, unknown> {
  if (!sources.length) return target;

  const source = sources.shift();

  if (isObject(target) && isObject(source)) {
    for (const key in source) {
      if (isObject(source[key])) {
        const val = target[key] ?? (target[key] = {});
        mergeDeep(val as Record<string, unknown>, source[key] as Record<string, unknown>);
      } else {
        target[key] = source[key];
      }
    }
  }

  return mergeDeep(target, ...sources);
}

export function parseCssRules(text: string): Record<string, string> {
  const result: Record<string, string> = {};

  for (const rule of text.split(";")) {
    const [key, val] = rule.split(":");
    result[key] = val;
  }

  return result;
}

export function formatCssRules(style: Record<string, string>): string {
  return Object.entries(style)
    .map((k, v) => `${k}: ${v}`)
    .join(";");
}

export function asArray<T>(val: T | T[]): T[] {
  return Array.isArray(val) ? val : [val];
}

export function clamp(val: number, min: number, max: number): number {
  return min > val ? min : max < val ? max : val;
}
