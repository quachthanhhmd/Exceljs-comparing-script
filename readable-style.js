const getReadableCellStyle = (style) => {
  if (!style) return "";

  let styleDesc = [];

  // Font
  if (style.font) {
    let fontDesc = [];
    if (style.font?.name) fontDesc.push(`Font: ${style.font.name}`);
    if (style.font?.size) fontDesc.push(`Size: ${style.font.size}`);
    if (style.font?.bold) fontDesc.push("Bold");
    if (style.font?.italic) fontDesc.push("Italic");
    if (style.font?.underline) fontDesc.push("Underline");
    if (style.font?.strike) fontDesc.push("Strikethrough");
    if (style.font?.color && style.font?.color?.argb)
      fontDesc.push(`Color: #${style.font.color.argb}`);
    if (style.font?.vertAlign)
      fontDesc.push(`Vertical Align: ${style.font.vertAlign}`);
    styleDesc.push(fontDesc.join(", "));
  }

  // Fill (Background color)
  if (style.fill) {
    let fillDesc = [];
    if (style.fill?.type) fillDesc.push(`Fill Type: ${style.fill.type}`);
    if (style.fill?.pattern) fillDesc.push(`Pattern: ${style.fill.pattern}`);
    if (style.fill?.fgColor && style.fill?.fgColor?.argb)
      fillDesc.push(`Foreground Color: #${style.fill.fgColor.argb}`);
    if (style.fill?.bgColor && style.fill?.bgColor?.argb)
      fillDesc.push(`Background Color: #${style.fill.bgColor.argb}`);
    styleDesc.push(fillDesc.join(", "));
  }

  // Alignment
  if (style.alignment) {
    let alignments = [];
    if (style.alignment?.horizontal)
      alignments.push(`Horizontal: ${style.alignment.horizontal}`);
    if (style.alignment?.vertical)
      alignments.push(`Vertical: ${style.alignment.vertical}`);
    if (style.alignment?.wrapText) alignments.push("Wrap Text");
    if (style.alignment?.shrinkToFit) alignments.push("Shrink to Fit");
    if (style.alignment?.indent)
      alignments.push(`Indent: ${style.alignment.indent}`);
    styleDesc.push(`Alignment: ${alignments.join(", ")}`);
  }

  // Border
  if (style.border) {
    let borderDesc = [];
    Object.keys(style.border).forEach((side) => {
      let border = style.border[side];
      if (border && border?.style)
        borderDesc.push(
          `${side}: ${border.style} ${
            border.color && border.color?.argb
              ? `(Color: #${border.color?.argb})`
              : ""
          }`
        );
    });
    styleDesc.push(`Border: ${borderDesc.join(", ")}`);
  }

  // Number Format
  if (style.numFmt) {
    styleDesc.push(`Number Format: ${style.numFmt}`);
  }

  // Protection
  if (style.protection) {
    let protectionDesc = [];
    if (style.protection.locked !== undefined)
      protectionDesc.push(`Locked: ${style.protection.locked}`);
    if (style.protection.hidden !== undefined)
      protectionDesc.push(`Hidden: ${style.protection.hidden}`);
    styleDesc.push(`Protection: ${protectionDesc.join(", ")}`);
  }

  return styleDesc.join(" | ");
};

module.exports = {
  getReadableCellStyle,
};
