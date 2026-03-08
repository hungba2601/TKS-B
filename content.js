chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === "extract_table") {
        sendResponse(extractData());
    } else if (request.action === "update_table") {
        updateTable(request.data);
        sendResponse({ success: true });
    } else if (request.action === "get_full_table") {
        sendResponse(getFullTableHtml());
    }
    return true;
});

let unconfirmedColIndex = -1;
let undeclaredColIndex = -1;
let classNameColIndex = 1; // Default to 1

function findTableAndColumns() {
    let allDocs = [document];
    document.querySelectorAll('iframe').forEach(ifr => {
        try {
            if (ifr.contentDocument) {
                allDocs.push(ifr.contentDocument);
            }
        } catch (e) { }
    });

    let targetTable = null;

    for (let doc of allDocs) {
        const tables = doc.querySelectorAll('table');
        for (let table of tables) {
            const innerHtml = table.innerHTML || "";
            const innerText = table.innerText || "";
            if (innerText.includes('Tên lớp') || innerHtml.includes('chưa xác nhận') || innerHtml.includes('chưa khai báo')) {
                if (table.querySelectorAll('tr').length > 2) {
                    targetTable = table;
                    break;
                }
            }
        }
        if (targetTable) break;
    }

    if (!targetTable) return null;

    // Try to find the correct indices by guessing standard table structure
    const firstDataRow = Array.from(targetTable.querySelectorAll('tr')).find(tr => tr.querySelectorAll('td').length >= 10);
    if (firstDataRow) {
        const tds = firstDataRow.querySelectorAll('td');
        // Usually, the last column is "Ngày/ Tiết chưa khai báo", and 2nd to last is "Tổng số tiết", 3rd to last is "Ngày/ Tiết chưa xác nhận".
        // Let's rely on standard lengths.
        if (tds.length === 13) {
            unconfirmedColIndex = 10;
            undeclaredColIndex = 12;
        } else {
            // fallback generic
            unconfirmedColIndex = tds.length - 3;
            undeclaredColIndex = tds.length - 1;
        }
    }

    return targetTable;
}

function extractData() {
    const table = findTableAndColumns();
    if (!table) return { error: "Không tìm thấy bảng Sổ đầu bài trên trang web." };

    const trs = table.querySelectorAll('tr');
    let result = [];

    trs.forEach((tr, index) => {
        const tds = tr.querySelectorAll('td');
        if (tds.length >= 10) {
            let className = tds[classNameColIndex] ? tds[classNameColIndex].innerText.trim() : "";
            if (className && !className.toLowerCase().includes('tổng')) {
                let unconfNode = tds[unconfirmedColIndex];
                let undeclNode = tds[undeclaredColIndex];

                let unconfirmedText = unconfNode ? unconfNode.innerText.trim() : "";
                let undeclaredText = undeclNode ? undeclNode.innerText.trim() : "";

                if (unconfirmedText || undeclaredText) {
                    result.push({
                        rowIndex: index,
                        className: className,
                        unconfirmed: unconfirmedText,
                        undeclared: undeclaredText
                    });
                }
            }
        }
    });

    return { data: result };
}

function updateTable(mappedData) {
    const table = findTableAndColumns();
    if (!table) return;

    const trs = table.querySelectorAll('tr');

    mappedData.forEach(item => {
        let tr = trs[item.rowIndex];
        if (tr) {
            const tds = tr.querySelectorAll('td');
            if (tds[unconfirmedColIndex] && item.unconfirmed) {
                tds[unconfirmedColIndex].innerHTML = item.unconfirmed.replace(/\n/g, '<br>');
            }
            if (tds[undeclaredColIndex] && item.undeclared) {
                tds[undeclaredColIndex].innerHTML = item.undeclared.replace(/\n/g, '<br>');
            }
        }
    });
}

function getFullTableHtml() {
    const table = findTableAndColumns();
    return table ? table.outerHTML : null;
}
