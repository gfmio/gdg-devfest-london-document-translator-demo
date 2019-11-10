/** Determines a new file name from the original file and the language pair in the translation  */
export default function newFileName(file: File | null, to: string, from: string | null) {
  if (!file) {
    return "";
  }

  const parts = file.name.split(".");
  const suffix = from ? ` (${from} to ${to})` : ` (${to})`;

  switch (parts.length) {
    case 0:
      return `Document${suffix}`;
    case 1:
      return `${parts[0]}${suffix}`;
    default:
      return [
        ...parts.slice(0, parts.length - 2),
        `${parts[parts.length - 2]}${suffix}`,
        parts[parts.length - 1]
      ].join(".");
  }
};
