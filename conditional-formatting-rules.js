const { getStyleDifferences } = require("./compare-style-different");
const { getReadableCellStyle } = require("./readable-style");

const formulaDictionary = {
  aboveAverage: "Format only values that are above or below average",
  uniqueValues: "Format only unique or duplicate values",
  duplicateValues: "Format only unique or duplicate values",
  top10: "Format the top or bottom ranked values",
  expression: "Use a formula to determine which cells to format",
};

const compareStyle = (source, target) => {
  if (JSON.stringify(source?.style) !== JSON.stringify(target?.style)) {
    return getStyleDifferences(source?.style, target?.style);
  }
};

const compareHook = (source, target, index) => {
  const differencesRules = [];
  console.log(source);

  // if the type of 2 rules are different
  if (source.type !== target.type) {
    return [
      {
        cell: "Conditional Formatting Rules",
        type: "Different conditional formatting rules",
        source: source.type,
        target: target.type,
        description: "Different conditional formatting rules",
        index,
      },
    ];
  }
  switch (source.type) {
    case "top10": {
      const message = (isBottom, isPercent, rank) => {
        return `${isBottom ? "Bottom" : "Top"} ${rank}${isPercent ? "%" : ""}`;
      };
      if (
        source?.percent !== target?.percent ||
        source?.rank !== target?.rank ||
        source?.bottom !== target?.bottom
      ) {
        differencesRules.push({
          source: message(source?.bottom, source?.percent, source?.rank),
          target: message(target?.bottom, target?.percent, target?.rank),
          type: formulaDictionary.top10,
          description: "Different conditional formatting rules",
          index,
        });
      }
      break;
    }
    case "expression": {
      if (source?.formula !== target?.formula) {
        differencesRules.push({
          cell: "Conditional Formatting Rules",
          source: source.formula,
          target: target.formula,
          type: formulaDictionary.expression,
          index,
        });
      }
      break;
    }
    case "colorScale": {
      // start compare
      const compareScale = (
        scaleSource,
        scaleTarget,
        index,
        scaleSourceColor,
        scaleTargetColor
      ) => {
        if (
          scaleSource &&
          scaleTarget &&
          (scaleSource.type !== scaleTarget.type ||
            scaleSource.value !== scaleTarget.value)
        ) {
          return {
            source: `${index === 0 ? "Minimum" : "Maximum"} ${
              scaleSource.value
            } ${scaleSource.type}${
              scaleSourceColor ? ` color #${scaleSourceColor}` : ""
            }`,
            target: `${index === 0 ? "Minimum" : "Maximum"} ${
              scaleTarget.value
            } ${scaleTarget.type}${
              scaleTargetColor ? ` color #${scaleTargetColor}` : ""
            }`,
          };
        }
      };
      const minCompare = compareScale(source.cfvo?.[0], target.cfvo?.[0], 0);
      const maxCompare = compareScale(source.cfvo?.[1], target.cfvo?.[1], 1);
      if (minCompare || maxCompare) {
        differencesRules.push({
          source: `${minCompare?.source ?? ""} ${maxCompare?.source ?? ""}`,
          target: `${minCompare?.target ?? ""} ${maxCompare?.target ?? ""}`,
          type: formulaDictionary.colorScale,
          index,
        });
      }
      break;
    }
    case "aboveAverage": {
      if (
        source?.aboveAverage !== target?.aboveAverage ||
        source?.equalAverage !== target?.equalAverage ||
        source?.stdDev !== target?.stdDev
      ) {
        differencesRules.push({
          source: JSON.stringify({
            aboveAverage: source?.aboveAverage,
            equalAverage: source?.equalAverage,
            stdDev: source?.stdDev,
          }),
          target: JSON.stringify({
            aboveAverage: target?.aboveAverage,
            equalAverage: target?.equalAverage,
            stdDev: target?.stdDev,
          }),
          type: formulaDictionary.aboveAverage,
          description: "Different above/below average rule settings",
          index,
        });
      }
      break;
    }
  }

  const styleDiff = compareStyle(source, target);
  if (styleDiff) {
    const { source: sourceDiff, target: targetDiff } = styleDiff;

    differencesRules.push({
      cell: "Conditional Formatting Rules",
      source: getReadableCellStyle(sourceDiff),
      target: getReadableCellStyle(targetDiff),
      type: formulaDictionary[source.type],
      description: "Different rules style",
      index,
    });
  }

  return differencesRules;
};

const isSameReferencesOfFormula = (sourceRef, targetRef) =>
  sourceRef !== targetRef;

/**
 * Compare rules in a specific formatting rule
 * @param {*} ruleSourceList
 * @param {*} ruleTargetList
 * @param {*} ruleIndex
 * @returns
 */
const compareRules = (ruleSourceList, ruleTargetList, ruleIndex) => {
  const maxLength = Math.max(ruleSourceList.length, ruleTargetList.length);

  const differences = [];

  for (let index = 0; index < maxLength; index++) {
    const sourceRule = ruleSourceList[index] || {};
    const targetRule = ruleTargetList[index] || {};
    const diff = compareHook(sourceRule, targetRule, ruleIndex);
    differences.push(...diff);
  }

  return differences;
};

function getConditionalFormattingDifferences(cf1, cf2, index) {
  const differences = [];
  if (isSameReferencesOfFormula(cf1.ref, cf2.ref)) {
    differences.push({
      source: cf1.ref,
      target: cf2.ref,
      type: "Reference",
      description: "Different conditional formatting rules reference",
      index,
    });
  }

  differences.push(...compareRules(cf1.rules, cf2.rules, index));

  return differences;
}

const compareConditionalFormattingRules = (sourceCfList, targetCfList) => {
  const maxLength = Math.max(sourceCfList.length, targetCfList.length);

  const differences = [];
  for (let index = 0; index < maxLength; index++) {
    const sourceCf = sourceCfList[index] || {};
    const targetCf = targetCfList[index] || {};

    const differencesForIndex = getConditionalFormattingDifferences(
      sourceCf,
      targetCf,
      index + 1
    );
    differences.push(...differencesForIndex);
  }
  return differences;
};

module.exports = {
  getConditionalFormattingDifferences,
  compareConditionalFormattingRules,
};
