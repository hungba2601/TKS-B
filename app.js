document.addEventListener('DOMContentLoaded', () => {
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
            statusArea.value = 'Vui lòng chọn file TKB (PDF)!';
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
                                tasksToAI.push({ id: tasksToAI.length, row: r, col: c, className: cName, text: cellValue });
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

            statusArea.value = `2. Đã tìm thấy ${tasksToAI.length} ô cần điền tên GV. Trích xuất text từ file PDF (Siêu tốc)...`;

            let pdfText = "";
            try {
                // Tải worker cục bộ đã download
                pdfjsLib.GlobalWorkerOptions.workerSrc = 'pdf.worker.min.js';
                const pdfData = new Uint8Array(await pdfFile.files[0].arrayBuffer());
                const pdf = await pdfjsLib.getDocument({ data: pdfData }).promise;
                for (let i = 1; i <= pdf.numPages; i++) {
                    const page = await pdf.getPage(i);
                    const content = await page.getTextContent();
                    const strings = content.items.map(item => item.str);
                    pdfText += strings.join(" ") + "\n";
                }
            } catch (err) {
                statusArea.value = 'Lỗi đọc file PDF: ' + err.message;
                return;
            }

            const aiInput = tasksToAI.map(t => ({ id: t.id, class: t.className, time: t.text }));

            statusArea.value = `3. Bắt đầu dò TKB. Đang nạp toàn bộ Dữ liệu và chia gói hỏi AI...`;
            const fullPdfText = pdfText.replace(/\s+/g, ' ');
            const safePdfText = fullPdfText.substring(0, 90000); // Giới hạn tầm 30 trang TKB an toàn 1 lần gọi

            let geminiResults = [];

            // Chia nhỏ thành các nhóm, mỗi nhóm 5 dòng để hỏi tránh quá tải
            let batchSize = 5;
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

            // Map AI result back to Array of Arrays (aoa)
            geminiResults.forEach(task => {
                let originalTask = tasksToAI.find(t => t.id === task.id);
                if (originalTask) {
                    aoa[originalTask.row][originalTask.col] = task.updatedText;
                }
            });

            // Write to Excel and trigger download
            const newWs = XLSX.utils.aoa_to_sheet(aoa);
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
