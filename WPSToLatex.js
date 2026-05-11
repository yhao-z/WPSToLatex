/**
 * WPS 表格选中区域 -> LaTeX (三线表通用版)
 *
 * Usage: select range -> DevTools -> Run macro -> WPSToLatex
 * Requires in preamble: \usepackage{booktabs} \usepackage{multirow}
 */

function WPSToLatex() {
    var firstCells = {};
    var firstCells2 = {};
    var mergedCells = {};

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

    // 1st pass: merge detection
    var cellInfo = [];
    var mergeAreaInfo = [];  // store [{ma, firstRow, firstCol, ...}, ...]

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

    for (var r = 1; r <= rows; r++) {
        // *** 修复：在这里重新获取 ma ***
        for (var c = 1; c <= cols; c++) {
            var cell = getCell(r, c);
            
            if (cell.MergeCells) {
                // *** 关键修复：每次循环都从 cell 获取 ma ***
                var ma = cell.MergeArea;
                if (!ma) continue;
                
                var firstRow = ma.Row - range.Row + 1;
                var firstCol = ma.Column - range.Column + 1;
                var key = firstRow + "," + firstCol;
                var key2 = (firstCol - 1) * 100 + (firstRow - 1);
                
                if (!(key in firstCells)) {
                    firstCells[key] = {
                        isFirst: true,
                        rowSpan: ma.Rows.Count,
                        colSpan: ma.Columns.Count,
                        bold: (cell.Font.Bold === true) ? cell : false,
                        italic: (cell.Font.Italic === true) ? cell : false,
                        ma: ma
                    };
                    firstCells2[key2] = firstCells[key];
                }
            }

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
        }
    }

    // 2nd pass: fill merge info
    for (var r = 1; r <= rows; r++) {
        for (var c = 1; c <= cols; c++) {
            var cell = getCell(r, c);
            var info = cellInfo[r][c];
            
            // *** 修复：在这里也重新获取 ma ***
            var ma = null;
            if (cell.MergeCells) ma = cell.MergeArea;
            
            var key = r + "," + c;
            var key2 = (c - 1) * 100 + (r - 1);
            
            if (key in firstCells) {
                var mc = firstCells[key];
                var msr = mc.ma.Rows.Count;
                var msc = mc.ma.Columns.Count;
                var mk = mc.ma.Row - range.Row + 1 + "," + (mc.ma.Column - range.Column + 1);
                mergedCells[mk] = { isFirst: true, rowSpan: msr, colSpan: msc };
            } else if (key2 in firstCells2) {
                var mk2 = firstCells2[key2].ma.Row - range.Row + 1 + "," + (firstCells2[key2].ma.Column - range.Column + 1);
                if (!mergedCells[mk2]) {
                    mergedCells[mk2] = { isFirst: false, rowSpan: firstCells2[key2].rowSpan, colSpan: firstCells2[key2].colSpan };
                }
            }

            if (key in mergedCells) {
                info.rowSpan = mergedCells[key].rowSpan;
                info.colSpan = mergedCells[key].colSpan;
                if (!mergedCells[key].isFirst) info.skip = true;
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
    for (var c = 1; c <= cols; c++) alignSpec += colAlign[c];
    alignSpec += "}";

    // Cover mark for multirow cells
    var coverInfo = [];
    for (var r = 1; r <= rows; r++) {
        coverInfo[r] = [];
        for (var c = 1; c <= cols; c++) coverInfo[r][c] = false;

        var mc = mergedCells[r + "," + c];
        if (mc && mc.isFirst && mc.rowSpan > 1) {
            for (var k = 1; k < mc.rowSpan; k++) {
                coverInfo[r + k][c] = true;
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
            
            if (info.skip) continue;
            if (coverInfo[r][c]) {  // 被 multirow 覆盖
                rowParts.push("");  // 占位
                continue;
            }

            var cellVal = getCellVal(r, c);
            var cellStr = escapeLatex(cellVal);

            if (info.colSpan > 1) {
                var mcSpec = colAlign[c] || "c";
                cellStr = "\\multicolumn{" + info.colSpan + "}{" + mcSpec + "}{" + cellStr + "}";
            }

            // *** 之前的 bug: 删除了 && info.colSpan <= 1 ***
            if (info.rowSpan > 1) {
                cellStr = "\\multirow{" + info.rowSpan + "}{*}{" + cellStr + "}";
            }

            if (info.bold) cellStr = "\\textbf{" + cellStr + "}";
            if (info.italic) cellStr = "\\textit{" + cellStr + "}";

            rowParts.push(cellStr);
        }

        if (rowParts.length > 0) {
            lines.push("    " + rowParts.join(" & ") + "\\\\");
        }

        if (horzLines[r]) {
            if (r === rows) lines.push("    \\bottomrule");
            else lines.push("    \\midrule");
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