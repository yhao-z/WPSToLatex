/**
 * WPS 表格选中区域 → LaTeX（完全通用版）
 * 
 * 使用方法：
 *   1. 选中要转换的区域（含表头）
 *   2. 顶部菜单「开发工具」→「运行宏」
 *   3. 选择 rangeToLatex 运行
 * 
 * 自动检测：
 *   - 边框 → | 和 \hline
 *   - 合并单元格 → \multirow 和 \multicolumn
 *   - 加粗/斜体 → \textbf{} / \textit{}
 *   - 水平对齐 → 列对齐方式 l/c/r
 *   - 标题行（第一行）→ 作为 caption
 * 
 * 需要 LaTeX 导言区添加：\usepackage{multirow}
 */

function WPSToLatex() {
    var sheet = Application.ActiveSheet;
    var sel = Application.Selection;
    var range = sheet.Range(
        sel.Cells.Item(1, 1),
        sel.Cells.Item(sel.Rows.Count, sel.Columns.Count)
    );
    
    var rows = range.Rows.Count;
    var cols = range.Columns.Count;
    
    // ========== 枚举常量（来自 WPS 文档）==========
    var xlEdgeLeft = 7, xlEdgeTop = 8, xlEdgeBottom = 9, xlEdgeRight = 10;
    var xlNone = -4142;
    var xlHAlignLeft = -4131, xlHAlignCenter = -4108, xlHAlignRight = -4152, xlHAlignGeneral = 1;
    var xlVAlignTop = -4160, xlVAlignCenter = -4108, xlVAlignBottom = -4107;
    
    // ========== 工具函数 ==========
    function getCell(r, c) {
        return sheet.Range(sel.Cells.Item(r, c), sel.Cells.Item(r, c));
    }
    
    function getCellVal(r, c) {
        var cell = range.Cells.Item(r, c);
        if (cell.Value2 != null) return String(cell.Value2).trim();
        return "";
    }
    
    function escapeLatex(s) {
        s = s.replace(/\\/g, "\\textbackslash{}");
        s = s.replace(/&/g, "\\&");
        s = s.replace(/%/g, "\\%");
        s = s.replace(/#/g, "\\#");
        s = s.replace(/_/g, "\\_");
        s = s.replace(/\$/g, "\\$");
        s = s.replace(/\{/g, "\\{");
        s = s.replace(/\}/g, "\\}");
        s = s.replace(/~/g, "\\textasciitilde{}");
        s = s.replace(/\^/g, "\\textasciicircum{}");
        return s;
    }
    
    // ========== 边框检测 ==========
    function cellHasBorder(r, c, idx) {
        try {
            var ls = getCell(r, c).Borders.Item(idx).LineStyle;
            return (ls != null && ls !== xlNone);
        } catch(e) { return false; }
    }
    
    // 竖线：逐列检测右边框
    var vertLines = [];
    for (var c = 1; c <= cols; c++) {
        vertLines[c] = cellHasBorder(1, c, xlEdgeRight);
    }
    var hasLeftBorder = cellHasBorder(1, 1, xlEdgeLeft);
    
    // 横线：逐行检测底边框
    var horzLines = [];
    for (var r = 1; r <= rows; r++) {
        horzLines[r] = cellHasBorder(r, 1, xlEdgeBottom);
    }
    var hasTopBorder = cellHasBorder(1, 1, xlEdgeTop);
    
    // ========== 合并单元格检测 ==========
    // cellInfo[r][c] 记录每个单元格的状态
    var cellInfo = [];
    for (var r = 1; r <= rows; r++) {
        cellInfo[r] = [];
        for (var c = 1; c <= cols; c++) {
            cellInfo[r][c] = {
                skip: false,        // 跳过（被合并区域的非首格）
                isFirst: true,      // 是否合并区域的首格
                rowSpan: 1,         // 跨行数
                colSpan: 1,         // 跨列数
                bold: false,
                italic: false,
                hAlign: null,       // 水平对齐
                vAlign: null        // 垂直对齐
            };
        }
    }
    
    for (var r = 1; r <= rows; r++) {
        for (var c = 1; c <= cols; c++) {
            var cell = getCell(r, c);
            var info = cellInfo[r][c];
            
            // 合并检测
            if (cell.MergeCells) {
                var ma = cell.MergeArea;
                var fr = ma.Row - range.Row + 1;  // 合并区域首行（相对坐标）
                var fc = ma.Column - range.Column + 1;  // 合并区域首列
                var rs = ma.Rows.Count;
                var cs = ma.Columns.Count;
                
                if (r === fr && c === fc) {
                    info.isFirst = true;
                    info.rowSpan = rs;
                    info.colSpan = cs;
                } else {
                    info.skip = true;
                    info.isFirst = false;
                }
            }
            
            // 加粗检测
            try { info.bold = (cell.Font.Bold === true); } catch(e) {}
            
            // 斜体检测
            try { info.italic = (cell.Font.Italic === true); } catch(e) {}
            
            // 水平对齐检测
            try {
                var ha = cell.HorizontalAlignment;
                if (ha === xlHAlignLeft) info.hAlign = "l";
                else if (ha === xlHAlignRight) info.hAlign = "r";
                else if (ha === xlHAlignCenter) info.hAlign = "c";
                // xlHAlignGeneral: 按数据类型推断
                else if (ha === xlHAlignGeneral || ha == null) {
                    var val = getCellVal(r, c);
                    if (val !== "" && !isNaN(Number(val))) info.hAlign = "r";  // 数字右对齐
                    else info.hAlign = "l";  // 文本左对齐
                }
            } catch(e) { info.hAlign = "l"; }
        }
    }
    
    // ========== 列对齐推断（取每列多数单元格的对齐方式）==========
    var colAlign = [];
    for (var c = 1; c <= cols; c++) {
        var counts = { l: 0, c: 0, r: 0 };
        for (var r = 1; r <= rows; r++) {
            if (!cellInfo[r][c].skip) {
                var a = cellInfo[r][c].hAlign || "l";
                counts[a] = (counts[a] || 0) + 1;
            }
        }
        // 取多数
        if (counts.r >= counts.l && counts.r >= counts.c) colAlign[c] = "r";
        else if (counts.c >= counts.l) colAlign[c] = "c";
        else colAlign[c] = "l";
    }
    
    // ========== 列对齐 spec ==========
    var alignSpec = "{";
    for (var c = 1; c <= cols; c++) {
        if (c === 1 && hasLeftBorder) alignSpec += "|";
        if (c > 1 && vertLines[c - 1]) alignSpec += "|";
        alignSpec += colAlign[c];
    }
    if (vertLines[cols]) alignSpec += "|";
    alignSpec += "}";
    
    // ========== 构建 LaTeX ==========
    var lines = [];
    
    // caption 和 label 在 tabular 上方
    lines.push("\\begin{table}[htbp]");
    lines.push("  \\centering");
    lines.push("  \\caption{TODO: caption}");
    lines.push("  \\label{tab:TODO}");
    lines.push("  \\begin{tabular}" + alignSpec);
    
    if (hasTopBorder) lines.push("    \\hline");
    
    for (var r = 1; r <= rows; r++) {
        var rowParts = [];
        var colIdx = 0;  // 实际列索引（跳过被跳过的格子）
        
        for (var c = 1; c <= cols; c++) {
            var info = cellInfo[r][c];
            
            // 跳过被合并区域的非首格
            if (info.skip) continue;
            
            colIdx++;
            var cellVal = getCellVal(r, c);
            var cellStr = escapeLatex(cellVal);
            
            // 多列合并：\multicolumn{n}{align}{text}
            // 第二个参数是合并后整个区域的对齐方式，如 c/l/r，或 |c| 带竖线
            if (info.colSpan > 1) {
                var mcSpec = "";
                // 左竖线
                if (c === 1 && hasLeftBorder) mcSpec += "|";
                else if (c > 1 && vertLines[c - 1]) mcSpec += "|";
                // 对齐方式（取第一列的对齐）
                mcSpec += colAlign[c] || "c";
                // 右竖线
                if (vertLines[c + info.colSpan - 1]) mcSpec += "|";
                cellStr = "\\multicolumn{" + info.colSpan + "}{" + mcSpec + "}{" + cellStr + "}";
            }
            
            // 多行合并：\multirow{n}{*}{text}
            if (info.rowSpan > 1 && info.colSpan <= 1) {
                cellStr = "\\multirow{" + info.rowSpan + "}{*}{" + cellStr + "}";
            }
            
            // 加粗
            if (info.bold) cellStr = "\\textbf{" + cellStr + "}";
            
            // 斜体
            if (info.italic) cellStr = "\\textit{" + cellStr + "}";
            
            rowParts.push(cellStr);
        }
        
        lines.push("    " + rowParts.join(" & ") + " \\\\");
        
        // 组间横线
        if (horzLines[r] && r < rows) {
            lines.push("    \\hline");
        }
    }
    
    if (horzLines[rows]) lines.push("    \\hline");
    
    lines.push("  \\end{tabular}");
    lines.push("\\end{table}");
    
    // ========== 输出到剪贴板 ==========
    var latexText = lines.join("\n");
    
    // 调试：检查 latexText 是否为空
    if (!latexText || latexText.trim() === "") {
        MsgBox("错误：生成的 LaTeX 文本为空！");
        return;
    }
    
    try {
        // 用临时文本框来复制纯文本到剪贴板
        // 创建文本框
        var tb = sheet.Shapes.AddTextbox(1, 0, 0, 100, 100);
        // 设置文本内容
        tb.TextFrame2.TextRange.Text = latexText;
        // 复制到剪贴板
        tb.TextFrame2.TextRange.Copy();
        // 删除文本框
        tb.Delete();
        
        MsgBox("LaTeX 已复制到剪贴板！文本长度：" + latexText.length);
    } catch(e) {
        // 备选：用单元格中转
        var tmpCell = sheet.Range("ZZ9999");
        tmpCell.Value2 = latexText;
        tmpCell.Select();
        Application.Selection.Copy();
        MsgBox("LaTeX 已复制到剪贴板！文本长度：" + latexText.length + "\n" +
               "（使用单元格中转，内容在 ZZ9999）");
    }
}