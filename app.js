document.addEventListener('DOMContentLoaded', () => {
    // --- THEME TOGGLE ---
    const themeToggleBtn = document.getElementById('themeToggleBtn');
    if (localStorage.getItem('darkMode') === 'true') {
        document.body.classList.add('dark-mode');
        if (themeToggleBtn) themeToggleBtn.innerHTML = '☀️ Chế độ Sáng';
    }
    if (themeToggleBtn) {
        themeToggleBtn.addEventListener('click', () => {
            document.body.classList.toggle('dark-mode');
            if (document.body.classList.contains('dark-mode')) {
                localStorage.setItem('darkMode', 'true');
                themeToggleBtn.innerHTML = '☀️ Chế độ Sáng';
            } else {
                localStorage.setItem('darkMode', 'false');
                themeToggleBtn.innerHTML = '🌙 Chế độ Tối';
            }
        });
    }

    // --- UI LOGIN / REGISTER ---
    const loginSection = document.getElementById('loginSection');
    const registerSection = document.getElementById('registerSection');
    const mainSection = document.getElementById('mainSection');

    const loginEmail = document.getElementById('loginEmail');
    const loginPassword = document.getElementById('loginPassword');
    const loginBtn = document.getElementById('loginBtn');
    const showRegisterBtn = document.getElementById('showRegisterBtn');
    const loginStatusArea = document.getElementById('loginStatusArea');

    const regEmail = document.getElementById('regEmail');
    const regPassword = document.getElementById('regPassword');
    const regRepassword = document.getElementById('regRepassword');
    const registerBtn = document.getElementById('registerBtn');
    const showLoginBtn = document.getElementById('showLoginBtn');
    const regStatusArea = document.getElementById('regStatusArea');
    const logoutBtn = document.getElementById('logoutBtn');

    // MẶC ĐỊNH BẠN PHẢI THAY LINK NÀY SAU KHI DEPLOY APP SCRIPT
    const APP_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxAkN02hwZ5F-IxR_et8U-MX2a1GgkgtI1AaWu7JpRuR170i8Q3OEMG7uvfj6OddimZ/exec"; // KẾT NỐI APP SCRIPT CỦA BẠN

    // Khởi đầu kiểm tra đăng nhập bằng Web LocalStorage
    const isLoggedIn = localStorage.getItem('isLoggedIn') === 'true';
    const userEmail = localStorage.getItem('userEmail');
    if (isLoggedIn) {
        loginSection.style.display = 'none';
        mainSection.style.display = 'block';
        if (userEmail) {
            document.getElementById('loggedInUserDisplay').innerHTML = `👤 <span style="color: #0056b3;">${userEmail}</span>`;
        }
    }

    showRegisterBtn.addEventListener('click', () => {
        loginSection.style.display = 'none';
        registerSection.style.display = 'block';
    });

    showLoginBtn.addEventListener('click', () => {
        registerSection.style.display = 'none';
        loginSection.style.display = 'block';
    });

    logoutBtn.addEventListener('click', () => {
        localStorage.removeItem('userEmail');
        localStorage.setItem('isLoggedIn', 'false');
        mainSection.style.display = 'none';
        loginSection.style.display = 'block';
        loginStatusArea.value = 'Đã đăng xuất.';
    });

    loginBtn.addEventListener('click', async () => {
        const email = loginEmail.value.trim();
        const pass = loginPassword.value.trim();
        if (!email || !pass) {
            loginStatusArea.value = 'Vui lòng nhập Email vả Mật khẩu!';
            return;
        }

        loginStatusArea.value = 'Đang kết nối để kiểm tra thông tin...';
        try {
            const formData = new URLSearchParams();
            formData.append('action', 'login');
            formData.append('email', email);
            formData.append('password', pass);

            const response = await fetch(APP_SCRIPT_URL, { method: 'POST', body: formData });
            const data = await response.json();

            if (data.success) {
                localStorage.setItem('isLoggedIn', 'true');
                localStorage.setItem('userEmail', email);
                loginSection.style.display = 'none';
                mainSection.style.display = 'block';
                document.getElementById('loggedInUserDisplay').innerHTML = `👤 <span style="color: #0056b3;">${email}</span>`;
                loginStatusArea.value = 'Vui lòng đăng nhập...';
                loginEmail.value = '';
                loginPassword.value = '';
            } else {
                loginStatusArea.value = data.message || 'Sai Email hoặc Mật khẩu!';
            }
        } catch (error) {
            loginStatusArea.value = 'Lỗi kết nối App Script: \n' + error.message;
        }
    });

    registerBtn.addEventListener('click', async () => {
        const email = regEmail.value.trim();
        const pass = regPassword.value.trim();
        const repass = regRepassword.value.trim();

        if (!email || !pass || !repass) {
            regStatusArea.value = 'Vui lòng nhập đủ thông tin!';
            return;
        }
        if (pass !== repass) {
            regStatusArea.value = 'Mật khẩu nhắc lại không khớp!';
            return;
        }

        regStatusArea.value = 'Đang đăng ký lên Google Sheet...';
        try {
            const formData = new URLSearchParams();
            formData.append('action', 'register');
            formData.append('email', email);
            formData.append('password', pass);

            const response = await fetch(APP_SCRIPT_URL, { method: 'POST', body: formData });
            const data = await response.json();

            if (data.success) {
                regStatusArea.value = 'Đăng ký thành công! Hãy Đăng nhập.';
                regEmail.value = '';
                regPassword.value = '';
                regRepassword.value = '';
                setTimeout(() => {
                    registerSection.style.display = 'none';
                    loginSection.style.display = 'block';
                    loginStatusArea.value = 'Đăng ký thành công! Vui lòng Đăng nhập.';
                }, 1500);
            } else {
                regStatusArea.value = data.message || 'Đăng ký không thành công!';
            }
        } catch (error) {
            regStatusArea.value = 'Lỗi kết nối API: \n' + error.message;
        }
    });

    // --- MAIN FUNCTIONALITY ---
    const apiKeyInput = document.getElementById('apiKey');
    const saveKeyBtn = document.getElementById('saveKeyBtn');
    const processBtn = document.getElementById('processBtn');
    const statusArea = document.getElementById('statusArea');
    const pdfFile = document.getElementById('pdfFile');
    const excelFile = document.getElementById('excelFile');

    // Load saved API key
    const savedKey = localStorage.getItem('geminiApiKey');
    if (savedKey) {
        apiKeyInput.value = savedKey;
    }

    saveKeyBtn.addEventListener('click', () => {
        const key = apiKeyInput.value.trim();
        if (key) {
            localStorage.setItem('geminiApiKey', key);
            alert('Đã lưu API Key.');
        }
    });

    processBtn.addEventListener('click', async () => {
        const key = apiKeyInput.value.trim();
        if (!key) {
            statusArea.value = 'Vui lòng nhập API Key!';
            return;
        }

        if (!excelFile.files.length) {
            statusArea.value = 'Vui lòng chọn file Excel tải từ trang web!';
            return;
        }

        if (!pdfFile.files.length) {
            statusArea.value = 'Vui lòng chọn file TKB (PDF hoặc Excel)!';
            return;
        }

        try {
            statusArea.value = '1. Đang phân tích File Excel...';

            const excelInputFile = excelFile.files[0];
            const arrayBuffer = await excelInputFile.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            // Convert to 2D array
            const aoa = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

            let classNameCol = 1;

            // Tìm cột 'Tên lớp'
            for (let r = 0; r < Math.min(10, aoa.length); r++) {
                if (!aoa[r]) continue;
                for (let c = 0; c < aoa[r].length; c++) {
                    let val = String(aoa[r][c] || "").toLowerCase();
                    if (val.includes("tên lớp")) classNameCol = c;
                }
            }

            let tasksToAI = [];
            for (let r = 0; r < aoa.length; r++) {
                let row = aoa[r];
                if (!row) continue;

                // Tìm kiếm Tên lớp
                let cName = row[classNameCol] ? String(row[classNameCol]).trim() : "";
                if (!cName) {
                    for (let c = 0; c < Math.min(4, row.length); c++) {
                        let val = String(row[c] || "").trim();
                        // Tránh STT, số trơn, và các ô rỗng
                        if (val && val.length > 1 && val.length < 15 && isNaN(Number(val)) && !val.toLowerCase().includes("stt")) {
                            cName = val;
                            break;
                        }
                    }
                }

                // Bỏ qua dòng tiêu đề và dòng 'Tổng'
                if (cName && !cName.toLowerCase().includes('tổng') && !cName.toLowerCase().includes('tên lớp')) {
                    // Quét toàn bộ cột ở phần sau của dòng đó
                    for (let c = 1; c < row.length; c++) {
                        if (c === classNameCol) continue;

                        let cellValue = row[c] ? String(row[c]).trim() : "";
                        let lowerCell = cellValue.toLowerCase();

                        // Hầu hết các ô tiết thiếu đều chứa chữ "thứ" hoặc "tiết"
                        if (lowerCell.includes("thứ") || lowerCell.includes("tiết")) {
                            // Bỏ qua các chuỗi của header
                            if (!lowerCell.includes("ngày") && !lowerCell.includes("tổng số") && !lowerCell.includes("sđb")) {
                                // Tách các dòng trong cùng 1 ô nếu có nhiều nội dung (Alt+Enter)
                                let lines = cellValue.split(/\r?\n/);
                                for (let line of lines) {
                                    let cleanLine = line.trim();
                                    if (cleanLine) {
                                        tasksToAI.push({ id: tasksToAI.length, row: r, col: c, className: cName, text: cleanLine });
                                    }
                                }
                            }
                        }
                    }
                }
            }

            if (tasksToAI.length === 0) {
                let debugInfo = `Rows=${aoa.length}, ClsCol=${classNameCol}`;
                let sampleRow = aoa[Math.min(5, Math.max(0, aoa.length - 1))] || [];
                debugInfo += `, Sample=${JSON.stringify(sampleRow.slice(0, 4))}`;
                statusArea.value = `Lỗi: Không tìm thấy dữ liệu tiết thiếu nào trong file Excel hoặc file không đúng định dạng. [${debugInfo}]`;
                return;
            }

            let pdfText = "";
            const tkbInputFile = pdfFile.files[0];
            const tkbName = tkbInputFile.name.toLowerCase();

            try {
                if (tkbName.endsWith('.pdf')) {
                    statusArea.value = `2. Đã tìm thấy ${tasksToAI.length} ô cần điền tên GV. Trích xuất text từ file TKB (PDF)...`;
                    // Tải worker cục bộ đã download
                    pdfjsLib.GlobalWorkerOptions.workerSrc = 'pdf.worker.min.js';
                    const pdfData = new Uint8Array(await tkbInputFile.arrayBuffer());
                    // Tạo danh sách lớp để lọc trang (nếu có thể) thay vì nối mọi trang vào chung 1 chuỗi dài lê thê
                    const missingClasses = [...new Set(tasksToAI.map(t => String(t.className).toLowerCase().trim()))];

                    const pdf = await pdfjsLib.getDocument({ data: pdfData }).promise;
                    for (let i = 1; i <= pdf.numPages; i++) {
                        const page = await pdf.getPage(i);
                        const content = await page.getTextContent();

                        // Lọc và sắp xếp các chữ (item) theo tọa độ để giữ vững cấu trúc bảng/từng dòng nằm ngang
                        let items = content.items;
                        items.sort((a, b) => {
                            // Cùng 1 dòng ngang (sai số 5 đơn vị y) thì ưu tiên từ trái sang phải
                            // Nếu y chênh quá 5, thì ưu tiên dòng trên (y lớn hơn)
                            let yDiff = b.transform[5] - a.transform[5];
                            if (Math.abs(yDiff) < 5) {
                                return a.transform[4] - b.transform[4]; // Sắp xếp x
                            }
                            return yDiff; // Sắp xếp y (từ trên xuống, gốc tọa độ PDF ở dưới cùng bên trái)
                        });

                        // Ghép các dòng theo logic (có cùng mức y) thay vì ráp bừa tứa lưa làm máy loạn não
                        let structuredPageText = "";
                        let currentY = -1;
                        let rowText = [];

                        for (let j = 0; j < items.length; j++) {
                            let textObj = items[j];
                            let startY = textObj.transform[5];
                            let txt = textObj.str.trim();
                            if (!txt) continue;

                            if (currentY === -1 || Math.abs(currentY - startY) > 5) {
                                // Xuống dòng mới
                                if (rowText.length > 0) {
                                    structuredPageText += rowText.join(" | ") + "\n";
                                }
                                rowText = [txt];
                                currentY = startY;
                            } else {
                                // Cùng dòng
                                rowText.push(txt);
                            }
                        }
                        // Thêm dòng cuối
                        if (rowText.length > 0) {
                            structuredPageText += rowText.join(" | ") + "\n";
                        }

                        // Ưu tiên: Chỉ lấy Text của các "Trang" mà có chứa (ít nhất một vài) tên Lớp thiếu tiết để cắt giảm size.
                        let isPageRelevant = false;
                        let lowerPageText = structuredPageText.toLowerCase();
                        let thresholdMatches = 0;

                        // Kiểm tra xem trang có Lớp cần tìm không
                        for (let mc of missingClasses) {
                            let cleanMc = mc.replace(/^l?ớp\s*/, '').trim(); // ví dụ từ 'lớp 6/1' thành '6/1'
                            if (lowerPageText.includes(` ${cleanMc} `) || lowerPageText.includes(`| ${cleanMc} |`) || lowerPageText.includes(mc)) {
                                thresholdMatches++;
                            }
                        }

                        // Nếu trang có dính líu đến LỚP thiếu tiết HOẶC chưa tìm ra đủ mọi lớp thiếu thì lấy trang đó đưa cho AI 
                        if (thresholdMatches > 0 || missingClasses.length === 0) {
                            pdfText += `\n--- Trang PDF liên quan (${i}) ---\n` + structuredPageText + "\n";
                        }
                    }
                } else if (tkbName.endsWith('.xls') || tkbName.endsWith('.xlsx')) {
                    statusArea.value = `2. Đã tìm thấy ${tasksToAI.length} ô cần điền tên GV. Trích xuất text từ file TKB (Excel)...`;
                    const tkbBuffer = await tkbInputFile.arrayBuffer();
                    const tkbWorkbook = XLSX.read(tkbBuffer, { type: 'array' });

                    const missingClasses = [...new Set(tasksToAI.map(t => String(t.className).toLowerCase().trim()))];

                    for (const sheetName of tkbWorkbook.SheetNames) {
                        const tkbSheet = tkbWorkbook.Sheets[sheetName];
                        const tkbAoa = XLSX.utils.sheet_to_json(tkbSheet, { header: 1, defval: "" });

                        if (tkbAoa.length === 0) continue;

                        // Tìm dòng chứa tên lớp (thường từ dòng 1 đến 10, index 0 đến 9)
                        let classRowIndex = -1;
                        let maxMatches = 0;

                        for (let r = 0; r < Math.min(15, tkbAoa.length); r++) {
                            let matches = 0;
                            for (let c = 0; c < tkbAoa[r].length; c++) {
                                let cellVal = String(tkbAoa[r][c]).toLowerCase().trim();
                                if (!cellVal || cellVal.length < 2) continue;

                                let cleanCell = cellVal.replace(/^l?ớp\s*/, '').trim(); // Lớp, lớp, ớp
                                let isMatch = missingClasses.some(mc => {
                                    let cleanMc = mc.replace(/^l?ớp\s*/, '').trim();
                                    return cleanCell === cleanMc ||
                                        (cleanCell.includes(cleanMc) && cleanCell.length <= cleanMc.length + 3) ||
                                        (cleanMc.includes(cleanCell) && cleanMc.length <= cleanCell.length + 3);
                                });

                                if (isMatch) matches++;
                            }
                            if (matches > maxMatches) {
                                maxMatches = matches;
                                classRowIndex = r;
                            }
                        }

                        // Nếu tìm thấy dòng có chứa lớp, ta chỉ trích các cột chứa lớp đó (và 1 cột kế tiếp)
                        if (classRowIndex !== -1 && maxMatches > 0) {
                            let colsToKeep = new Set();
                            // Luôn giữ ít nhất 3 cột đầu tiên (đề phòng cột STT, Thứ, Tiết)
                            colsToKeep.add(0);
                            colsToKeep.add(1);
                            colsToKeep.add(2);

                            // Thêm các cột chứa lớp bị thiếu
                            for (let c = 0; c < tkbAoa[classRowIndex].length; c++) {
                                let cellVal = String(tkbAoa[classRowIndex][c]).toLowerCase().trim();
                                if (!cellVal) continue;
                                let cleanCell = cellVal.replace(/^l?ớp\s*/, '').trim();
                                let isMatch = missingClasses.some(mc => {
                                    let cleanMc = mc.replace(/^l?ớp\s*/, '').trim();
                                    return cleanCell === cleanMc ||
                                        (cleanCell.includes(cleanMc) && cleanCell.length <= cleanMc.length + 3) ||
                                        (cleanMc.includes(cleanCell) && cleanMc.length <= cleanCell.length + 3);
                                });
                                if (isMatch) {
                                    colsToKeep.add(c);
                                    // Các cột Sáng/Chiều của cùng 1 lớp thường được gộp ô (Merge Cells) ở dòng Tên Lớp. 
                                    // Do đó các cột phía sau trong dòng đó sẽ rỗng. Bắt đúng các cột rỗng này thay vì +1, +2 cứng.
                                    let nextC = c + 1;
                                    while (nextC < tkbAoa[classRowIndex].length) {
                                        let nextCellVal = String(tkbAoa[classRowIndex][nextC]).trim();
                                        if (nextCellVal !== "") break; // Chạm phải cột chứa tên của lớp khác
                                        colsToKeep.add(nextC);
                                        nextC++;
                                    }
                                }
                            }

                            // Cắt mảng thành CSV nhưng chỉ chứa những cột cho phép
                            let filteredCsv = "";
                            for (let r = 0; r < tkbAoa.length; r++) {
                                let rowVals = [];
                                let hasData = false;
                                for (let c = 0; c < tkbAoa[r].length; c++) {
                                    if (colsToKeep.has(c)) {
                                        let val = String(tkbAoa[r][c]).trim();
                                        rowVals.push(val);
                                        // Có dữ liệu môn học
                                        if (val && c > 2) hasData = true;
                                    }
                                }
                                // Giữ lại dòng lớp, dòng sát lớp, và các dòng có nội dung môn
                                if (hasData || r <= classRowIndex + 1) {
                                    filteredCsv += rowVals.join(",") + "\n";
                                }
                            }
                            pdfText += `\n--- Sheet (Lấy Mẫu Cụ Thể Lớp): ${sheetName} ---\n` + filteredCsv + "\n";
                        } else {
                            // Nếu không rớt vào TH có vẻ như là cột lớp thì xài csv mặc định
                            const csvData = XLSX.utils.sheet_to_csv(tkbSheet);
                            pdfText += `\n--- Sheet Đầy đủ (Do Không Khớp Lớp): ${sheetName} ---\n` + csvData + "\n";
                        }
                    }
                } else {
                    statusArea.value = 'Lỗi: Thể loại file TKB không được hỗ trợ. Vui lòng chọn PDF hoặc Excel.';
                    return;
                }
            } catch (err) {
                statusArea.value = 'Lỗi đọc file TKB: ' + err.message;
                return;
            }

            const aiInput = tasksToAI.map(t => ({ id: t.id, class: t.className, time: t.text }));

            statusArea.value = `3. Bắt đầu dò TKB. Đang nạp toàn bộ Dữ liệu và chia gói hỏi AI...`;
            const fullPdfText = pdfText.replace(/\s+/g, ' ');
            const safePdfText = fullPdfText.substring(0, 90000); // Giới hạn tầm 30 trang TKB an toàn 1 lần gọi

            let geminiResults = [];

            // Tăng số lượng câu hỏi trong mỗi gói lên 10 để chạy nhanh gấp đôi
            let batchSize = 10;
            let totalBatches = Math.ceil(aiInput.length / batchSize);

            for (let b = 0; b < totalBatches; b++) {
                let batch = aiInput.slice(b * batchSize, (b + 1) * batchSize);

                let queryBlock = "";
                for (let t of batch) {
                    queryBlock += `[ID: ${t.id}] Lớp: ${t.class} | Dò tìm trong TKB ai dạy Giáo Viên và Môn học gì vào lúc: ${t.time}\n`;
                }

                const prompt = `Bạn là hệ thống AI siêu việt. Dưới đây là Toàn bộ Nội Dung Thời Khóa Biểu:
--- BẮT ĐẦU DỮ LIỆU TKB ---
${safePdfText}
--- KẾT THÚC DỮ LIỆU TKB ---

NHIỆM VỤ: Hãy nhìn vào bộ Dữ liệu TKB trên, và điền Môn, Tên Giáo viên cho ${batch.length} câu hỏi sau:
${queryBlock}
BẤT BUỘC: Trả về ĐÚNG CÁC DÒNG CÓ ĐỊNH DẠNG SAU:
ID|Kết quả (Ghép Tên Môn và Tên GV vào ngay sau thời gian)
Ví dụ nếu Lớp 6A1 môn Toán cô Hương dạy:
0|Thứ 3 - Sáng (Tiết 4) - Toán (Hương)
CHỈ TRẢ VỀ CÁC DÒNG CHỨA DẤU |, TUYỆT ĐỐI KHÔNG DÙNG FORMAT MARKDOWN HOẶC GIẢI THÍCH GÌ THÊM!`;

                let retries = 3;
                while (retries > 0) {
                    try {
                        statusArea.value = `Đang nhờ AI giải mã Gói ${b + 1}/${totalBatches} (Vui lòng đợi vài giây)...`;
                        let resp = await callGeminiTextAPI(key, prompt);

                        let lines = resp.trim().replace(/^```.*$/gm, "").split('\n');
                        for (let line of lines) {
                            if (!line.includes('|')) continue;
                            let parts = line.split('|');
                            let pId = parseInt(parts[0].trim());
                            let pText = parts[1].trim();
                            if (!isNaN(pId)) {
                                geminiResults.push({ id: pId, updatedText: pText });
                            }
                        }
                        break; // Thành công thì thoát vòng lặp retry
                    } catch (e) {
                        // Xử lý Lỗi Quá Tải của API Google Miễn phí
                        if (e.message.includes('429')) {
                            statusArea.value = `[Google báo 429 Quá Tải] Tự động nghỉ 10 giây để né spam API (Thử lại lần ${4 - retries})...`;
                            await new Promise(r => setTimeout(r, 10000)); // Nhấp trà 10 giây
                            retries--;
                            if (retries === 0) {
                                for (let t of batch) geminiResults.push({ id: t.id, updatedText: t.time + ` (Lỗi AI quá tải)` });
                            }
                        } else {
                            console.error("Lỗi AI gói", b, e);
                            for (let t of batch) geminiResults.push({ id: t.id, updatedText: t.time + ` (Lỗi AI: ${e.message})` });
                            break;
                        }
                    }
                }

                // Nghỉ ngơi 2 giây giữa các gói để tránh nhồi nhét Google
                if (b < totalBatches - 1) {
                    await new Promise(r => setTimeout(r, 2000));
                }
            }

            // Xử lý điền chữ cho các ô bị trượt sót
            for (let t of aiInput) {
                if (!geminiResults.find(r => r.id === t.id)) {
                    geminiResults.push({ id: t.id, updatedText: t.time + " (API AI lỗi nhả thiếu)" });
                }
            }

            statusArea.value = '4. Đang xuất kết quả và tạo file Excel ráp nối...';

            // Map AI result back to original tasks
            geminiResults.forEach(task => {
                let originalTask = tasksToAI.find(t => t.id === task.id);
                if (originalTask) {
                    originalTask.updatedText = task.updatedText;
                }
            });

            // Gộp lại các dòng trong cùng một ô
            let groupedCells = {};
            tasksToAI.forEach(t => {
                let key = `${t.row}_${t.col}`;
                if (!groupedCells[key]) {
                    groupedCells[key] = [];
                }
                groupedCells[key].push(t.updatedText || t.text);
            });

            // Gắn vào mảng dữ liệu aoa
            for (let key in groupedCells) {
                let [r, c] = key.split('_');
                aoa[r][c] = groupedCells[key].join('\n');
            }

            // Write to Excel and trigger download
            const newWs = XLSX.utils.aoa_to_sheet(aoa);

            // Giữ lại định dạng độ rộng cột (Column Widths) của file gốc nếu có
            if (worksheet['!cols']) {
                newWs['!cols'] = worksheet['!cols'];
            } else {
                // Nếu file gốc không có width chuẩn, tự tính toán độ rộng dựa trên độ dài nội dung (Auto-fit)
                let colWidths = [];
                for (let c = 0; c < aoa[0].length; c++) {
                    let maxLen = 10; // Chiều rộng tối thiểu
                    for (let r = 0; r < aoa.length; r++) {
                        if (aoa[r][c]) {
                            let cellLen = String(aoa[r][c]).length;
                            if (cellLen > maxLen) {
                                maxLen = cellLen;
                            }
                        }
                    }
                    // Thêm 2 ký tự pixel padding cho đẹp và giới hạn tối đa 60 để không bị quá to
                    colWidths.push({ wch: Math.min(60, maxLen + 2) });
                }
                newWs['!cols'] = colWidths;
            }

            workbook.Sheets[firstSheetName] = newWs;
            XLSX.writeFile(workbook, "SDB_TichHop_GV.xlsx");

            statusArea.value = 'Hoàn tất! File Excel mới "SDB_TichHop_GV.xlsx" đã được tải xuống máy của bạn.';

        } catch (error) {
            console.error(error);
            statusArea.value = 'Lỗi: ' + error.message;
        }
    });

    async function callGeminiTextAPI(apiKey, prompt) {
        // Use gemini-2.5-flash as requested by the user
        const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;

        const payload = {
            contents: [{
                parts: [
                    { text: prompt }
                ]
            }],
            generationConfig: {
                temperature: 0.1, // Use low temperature to ensure correctness
                responseMimeType: "text/plain"
            }
        };

        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(payload)
        });

        if (!response.ok) {
            const textContent = await response.text();
            throw new Error(`API Error: ${response.status} - ${textContent}`);
        }

        const data = await response.json();
        if (data.candidates && data.candidates[0] && data.candidates[0].content) {
            return data.candidates[0].content.parts.map(p => p.text).join("");
        } else {
            throw new Error("Không nhận được phản hồi hợp lệ từ Gemini.");
        }
    }
});
