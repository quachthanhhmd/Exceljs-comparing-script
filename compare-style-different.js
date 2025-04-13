const getStyleDifferences = (source, target) => {
  const sourceDiffs = {};
  const targetDiffs = {};

  Object.keys(source).forEach((key) => {
    if (JSON.stringify(source?.[key] ?? '') !== JSON.stringify(target?.[key] ?? '')) {
      if (
        typeof source[key] === "object" &&
        source[key] !== null &&
        target[key] !== null
      ) {
        // Recursively check nested objects
        const nestedDiffs = getStyleDifferences(source[key], target[key]);
        if (Object.keys(nestedDiffs.source).length > 0)
          sourceDiffs[key] = nestedDiffs.source;
        if (Object.keys(nestedDiffs.target).length > 0)
          targetDiffs[key] = nestedDiffs.target;
      } else {
        // Store different values
        sourceDiffs[key] = source?.[key] ?? '';
        targetDiffs[key] = target?.[key] ?? '';
      }
    }
  });

  return { source: sourceDiffs, target: targetDiffs };
};


module.exports = {
  getStyleDifferences
}