/**
 * WPS 表格选中区域 -> LaTeX (三线表通用版)
 *
 * Usage: select range -> DevTools -> Run macro -> WPSToLatex
 * Requires in preamble: \usepackage{booktabs} \usepackage{multirow}
 */

function WPS2Latex() {
    var sheet = Application.ActiveSheet;
    var sel = Application.Selection;
    var range = sheet.Range(
        sel.Cells.Item(1, 1),
        sel.Cells.Item(sel.Rows.Count, sel.Columns.Count)
    );

    var rows = range.Rows.Count;
    var cols = range.Columns.Count;

    var xlEdgeLeft = 7, xlEdgeTop = 8, xlEdgeBottom = 9, xlEdgeRight = 10;
    var xlNone = -4142;
    var xlHAlignLeft = -4131, xlHAlignCenter = -4108, xlHAlignRight = -4152, xlHAlignGeneral = 1;

    function getCell(r, c) {
        return sheet.Range(sel.Cells.Item(r, c), sel.Cells.Item(r, c));
    }

    function getCellVal(r, c) {
        var cell = range.Cells.Item(r, c);
        if (cell.Value2 != null) return String(cell.Value2).trim();
        return "";
    }

    function escapeLatex(s) {
        if (!s) return s;
        s = s.replace(/\\/g, "\\textbackslash{}");
        s = s.replace(/&/g, "\\&");
        s = s.replace(/%/g, "\\%");
        s = s.replace(/#/g, "\\#");
        s = s.replace(/_/g,"\\_");
        s = s.replace(/\$/g, "\\$");
        s = s.replace(/\{/g, "\\{");
        s = s.replace(/\}/g, "\\}");
        s = s.replace(/~/g, "\\textasciitilde{}");
        s = s.replace(/\^/g, "\\textasciicircum{}");
        return s;
    }

    function cellHasBorder(r, c, idx) {
        try {
            var ls = getCell(r, c).Borders.Item(idx).LineStyle;
            return (ls != null && ls !== xlNone);
        } catch(e) { return false; }
    }

    var horzLines = [];
    for (var r = 1; r <= rows; r++) {
        var hasBorder = false;
        for (var c = 1; c <= cols; c++) {
            if (cellHasBorder(r, c, xlEdgeBottom)) hasBorder = true;
            if (r < rows && cellHasBorder(r + 1, c, xlEdgeTop)) hasBorder = true;
        }
        horzLines[r] = hasBorder;
    }
    var hasTopBorder = cellHasBorder(1, 1, xlEdgeTop);

    // 检测竖向框线（列之间的边框）
    var vertLines = [];
    for (var c = 1; c <= cols; c++) {
        var hasVBorder = false;
        for (var r = 1; r <= rows; r++) {
            if (cellHasBorder(r, c, xlEdgeRight)) hasVBorder = true;
            if (c < cols && cellHasBorder(r, c + 1, xlEdgeLeft)) hasVBorder = true;
        }
        vertLines[c] = hasVBorder;
    }

    // 初始化 cellInfo
    var cellInfo = [];
    for (var r = 1; r <= rows; r++) {
        cellInfo[r] = [];
        for (var c = 1; c <= cols; c++) {
            cellInfo[r][c] = {
                skip: false, isFirst: false,
                rowSpan: 1, colSpan: 1,
                bold: false, italic: false,
                hAlign: null
            };
        }
    }

    // 1st pass: 收集所有合并单元格信息，并记录格式
    var mergeMap = {};  // key: "r,c" -> {firstRow, firstCol, rowSpan, colSpan}

    for (var r = 1; r <= rows; r++) {
        for (var c = 1; c <= cols; c++) {
            var cell = getCell(r, c);
            
            // Bold / Italic
            try { cellInfo[r][c].bold = (cell.Font.Bold === true); } catch(e) {}
            try { cellInfo[r][c].italic = (cell.Font.Italic === true); } catch(e) {}

            // Alignment
            try {
                var ha = cell.HorizontalAlignment;
                if (ha === xlHAlignLeft) cellInfo[r][c].hAlign = "l";
                else if (ha === xlHAlignRight) cellInfo[r][c].hAlign = "r";
                else if (ha === xlHAlignCenter) cellInfo[r][c].hAlign = "c";
                else {
                    var val = getCellVal(r, c);
                    cellInfo[r][c].hAlign = (val !== "" && !isNaN(Number(val))) ? "r" : "l";
                }
            } catch(e) { cellInfo[r][c].hAlign = "l"; }

            // 合并单元格检测
            if (cell.MergeCells) {
                var ma = cell.MergeArea;
                if (!ma) continue;
                
                var firstRow = ma.Row - range.Row + 1;
                var firstCol = ma.Column - range.Column + 1;
                var key = firstRow + "," + firstCol;
                
                if (!(key in mergeMap)) {
                    mergeMap[key] = {
                        firstRow: firstRow,
                        firstCol: firstCol,
                        rowSpan: ma.Rows.Count,
                        colSpan: ma.Columns.Count
                    };
                }
            }
        }
    }

    // 2nd pass: 标记每个单元格的合并状态
    for (var r = 1; r <= rows; r++) {
        for (var c = 1; c <= cols; c++) {
            var info = cellInfo[r][c];
            
            // 检查当前单元格是否在某个合并区域内
            for (var key in mergeMap) {
                var mc = mergeMap[key];
                // 判断 (r,c) 是否在合并区域内
                if (r >= mc.firstRow && r < mc.firstRow + mc.rowSpan &&
                    c >= mc.firstCol && c < mc.firstCol + mc.colSpan) {
                    
                    if (r === mc.firstRow && c === mc.firstCol) {
                        // 这是合并区域的第一个单元格
                        info.isFirst = true;
                        info.rowSpan = mc.rowSpan;
                        info.colSpan = mc.colSpan;
                    } else {
                        // 这是被合并覆盖的单元格
                        info.skip = true;
                        info.rowSpan = mc.rowSpan;
                        info.colSpan = mc.colSpan;
                    }
                    break;  // 一个单元格只可能在一个合并区域内
                }
            }
        }
    }

    // Column alignment
    var colAlign = [];
    for (var c = 1; c <= cols; c++) {
        var counts = { l: 0, c: 0, r: 0 };
        for (var r = 1; r <= rows; r++) {
            if (!cellInfo[r][c].skip) {
                var a = cellInfo[r][c].hAlign || "l";
                counts[a]++;
            }
        }
        if (counts.r >= counts.l && counts.r >= counts.c) colAlign[c] = "r";
        else if (counts.c >= counts.l) colAlign[c] = "c";
        else colAlign[c] = "l";
    }

    var alignSpec = "{";
    for (var c = 1; c <= cols; c++) {
        alignSpec += colAlign[c];
        if (vertLines[c]) alignSpec += "|";
    }
    alignSpec += "}";

    // Cover mark for multirow cells
    var coverInfo = [];
    for (var r = 1; r <= rows; r++) {
        coverInfo[r] = [];
        for (var c = 1; c <= cols; c++) coverInfo[r][c] = false;
    }
    for (var r = 1; r <= rows; r++) {
        for (var c = 1; c <= cols; c++) {
            var info = cellInfo[r][c];
            if (info.isFirst && info.rowSpan > 1) {
                for (var k = 1; k < info.rowSpan; k++) {
                    coverInfo[r + k][c] = true;
                }
            }
        }
    }

    // Build LaTeX
    var lines = [];
    lines.push("\\begin{table}[htbp]");
    lines.push("  \\centering");
    lines.push("  \\caption{TODO: caption \\label{tab:TODO}}");
    lines.push("  \\begin{tabular}" + alignSpec);

    if (hasTopBorder) lines.push("    \\toprule");

    for (var r = 1; r <= rows; r++) {
        var rowParts = [];
        for (var c = 1; c <= cols; c++) {
            var info = cellInfo[r][c];
            
            if (coverInfo[r][c]) {  // 被 multirow 覆盖，输出空占位
                // 如果该列右侧有竖线，用 \multicolumn 覆盖以避免竖线贯穿
                if (vertLines[c]) {
                    rowParts.push("\\multicolumn{1}{c}{}");
                } else {
                    rowParts.push("");
                }
                continue;
            }
            if (info.skip) continue;

            var cellVal = getCellVal(r, c);
            var cellStr = escapeLatex(cellVal);

            // 空白单元格不加粗/斜体
            if (cellStr !== "") {
                if (info.bold) cellStr = "\\textbf{" + cellStr + "}";
                if (info.italic) cellStr = "\\textit{" + cellStr + "}";
            }

            if (info.colSpan > 1) {
                var mcSpec = colAlign[c] || "c";
                // 如果该列右侧有竖线，multicolumn 格式也需要包含 |
                if (vertLines[c + info.colSpan - 1]) mcSpec += "|";
                cellStr = "\\multicolumn{" + info.colSpan + "}{" + mcSpec + "}{" + cellStr + "}";
            }

            if (info.rowSpan > 1) {
                // [4] 是 fixup 参数，用于调整 multirow 的垂直居中位置
                cellStr = "\\multirow{" + info.rowSpan + "}[4]{*}{" + cellStr + "}";
                // multirow 首单元格如果列右侧有竖线，用 \multicolumn 覆盖
                if (vertLines[c] && info.colSpan <= 1) {
                    cellStr = "\\multicolumn{1}{c}{" + cellStr + "}";
                }
            }

            rowParts.push(cellStr);
        }

        if (rowParts.length > 0) {
            lines.push("    " + rowParts.join(" & ") + "\\\\");
        }

        if (horzLines[r]) {
            if (r === rows) {
                lines.push("    \\bottomrule");
            } else {
                // 检测横线是否需要跳过 multirow 区域
                var lineSegments = [];
                var inSegment = false;
                var segStart = 0;
                
                for (var c = 1; c <= cols; c++) {
                    var hasBorder = false;
                    // 检查当前行第c列的底部边框
                    if (cellHasBorder(r, c, xlEdgeBottom)) hasBorder = true;
                    // 或下一行第c列的顶部边框
                    if (r < rows && cellHasBorder(r + 1, c, xlEdgeTop)) hasBorder = true;
                    
                    // 如果该列被 multirow 覆盖或是 multirow 首列，跳过
                    var isMultirowCol = false;
                    if (cellInfo[r][c].isFirst && cellInfo[r][c].rowSpan > 1) isMultirowCol = true;
                    if (coverInfo[r + 1] && coverInfo[r + 1][c]) isMultirowCol = true;
                    
                    if (hasBorder && !isMultirowCol) {
                        if (!inSegment) {
                            segStart = c;
                            inSegment = true;
                        }
                    } else {
                        if (inSegment) {
                            lineSegments.push({start: segStart, end: c - 1});
                            inSegment = false;
                        }
                    }
                }
                if (inSegment) {
                    lineSegments.push({start: segStart, end: cols});
                }
                
                if (lineSegments.length === 1 && lineSegments[0].start === 1 && lineSegments[0].end === cols) {
                    lines.push("    \\midrule");
                } else if (lineSegments.length > 0) {
                    var cmidruleParts = [];
                    for (var s = 0; s < lineSegments.length; s++) {
                        cmidruleParts.push("\\cmidrule{" + lineSegments[s].start + "-" + lineSegments[s].end + "}");
                    }
                    lines.push("    " + cmidruleParts.join(""));
                }
            }
        }
    }

    if (!hasTopBorder && horzLines[rows]) lines.push("    \\bottomrule");

    lines.push("  \\end{tabular}");
    lines.push("\\end{table}");

    var latexText = lines.join("\n");

    if (!latexText || latexText.trim() === "") {
        MsgBox("Error: empty LaTeX output!");
        return;
    }

    try {
        var tb = sheet.Shapes.AddTextbox(1, 0, 0, 100, 100);
        tb.TextFrame2.TextRange.Text = latexText;
        tb.TextFrame2.TextRange.Copy();
        tb.Delete();
        MsgBox("LaTeX copied! Length: " + latexText.length);
    } catch(e) {
        var tmpCell = sheet.Range("ZZ9999");
        tmpCell.Value2 = latexText;
        tmpCell.Select();
        Application.Selection.Copy();
        MsgBox("LaTeX copied (fallback)! Length: " + latexText.length);
    }
}
