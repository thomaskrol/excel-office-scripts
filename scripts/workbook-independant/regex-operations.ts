/**
 * Performs regex operations on a string.
 * 
 * @param operation Operation type: "all matches" returns matching strings, "test match" returns boolean, "replace" substitutes matches with replaceString, "groups" returns captured groups
 * @param searchString String to test against regex pattern
 * @param regexPattern Regex pattern for matching
 * @param regexFlags Regex flags (g, i, m, s, u, y, d)
 * @param replaceString Replacement string (required for "replace" operation)
 */
function main(
  workbook: ExcelScript.Workbook,
  operation: "all matches" | "test match" | "replace" | "groups" = "all matches",
  searchString: string,
  regexPattern: string,
  regexFlags?: string,
  replaceString?: string
): string | Array<string> | boolean {
  // validate regexFlags
  if (regexFlags) {
    const validFlags = ['g', 'i', 'm', 's', 'u', 'y', 'd'];
    const invalidFlags = regexFlags.split("").filter(f => !validFlags.includes(f));
    if (invalidFlags.length > 0) {
      throw new Error(`Invalid regex flag(s): ${invalidFlags.join(', ')}. Valid flags are: ${validFlags.join(', ')}`);
    }
  }

  const regex = new RegExp(regexPattern, regexFlags || undefined);

  switch (operation.toLowerCase()) {
    case "all matches": {
      const matches = searchString.match(regex);
      return matches || [];
    }
    case "test match": {
      return regex.test(searchString);
    }
    case "replace": {
      return searchString.replace(regex, replaceString || "");
    }
    case "groups": {
      const match = regex.exec(searchString);

      // Return all captured groups except the full match at index 0
      return match?.slice(1) || [];
    }
    default: {
      throw new Error(`Unknown operation: ${operation}`);
    }
  }
}