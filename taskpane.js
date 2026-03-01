const APP_VERSION = "v1.0.0";
const API_BASE = "https://fx-proxy.jang375a-03c.workers.dev";

let dateRangeAddress = "";
let targetRangeAddress = "";

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("btn-set-date").onclick = setDateRange;
        document.getElementById("btn-set-target").onclick = setTargetRange;
        document.getElementById("btn-run").onclick = runExchangeRateFetch;
        document.getElementById("app-version").textContent = APP_VERSION;
        loadCurrencies();
    }
});

async function loadCurrencies() {
    try {
        const response = await fetch(`${API_BASE}/get_currencies`);
        const list = await response.json();
        if (typeof window.setCurrencyList === "function") {
            window.setCurrencyList(list);
        }
    } catch (e) {
        // 통화 목록 로드 실패 시 무시 (수동 입력 가능)
    }
}

async function setDateRange() {
    await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("address");
        await context.sync();
        dateRangeAddress = range.address;
        document.getElementById("date-range-address").innerText = dateRangeAddress;
    });
}

async function setTargetRange() {
    await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("address");
        await context.sync();
        targetRangeAddress = range.address;
        document.getElementById("target-range-address").innerText = targetRangeAddress;
    });
}

function formatToAPIString(rawDateStr) {
    if (!rawDateStr) return "";
    const str = String(rawDateStr).trim();

    if (/^\d{8}$/.test(str)) return str;

    if (/^\d{5}$/.test(str)) {
        const d = new Date(Date.UTC(1899, 11, 30) + parseInt(str) * 86400000);
        return d.getUTCFullYear().toString() +
               String(d.getUTCMonth() + 1).padStart(2, '0') +
               String(d.getUTCDate()).padStart(2, '0');
    }

    let m = str.match(/(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})/);
    if (m) return m[1] + m[2].padStart(2, '0') + m[3].padStart(2, '0');

    m = str.match(/(\d{4})[.\-\/](\d{1,2})[.\-\/](\d{1,2})/);
    if (m) return m[1] + m[2].padStart(2, '0') + m[3].padStart(2, '0');

    return "";
}

function setRunStatus(msg, color) {
    const box = document.getElementById("run-status");
    box.style.display = "block";
    box.style.color = color || "#333";
    box.textContent = msg;
}

async function runExchangeRateFetch() {
    const currency = document.getElementById("currency-input").value.trim().toUpperCase();

    if (!currency) {
        setRunStatus("통화 코드를 입력해주세요. (예: USD)", "#c00");
        return;
    }
    if (!dateRangeAddress) {
        setRunStatus("1단계: 날짜 범위를 먼저 지정해주세요.", "#c00");
        return;
    }
    if (!targetRangeAddress) {
        setRunStatus("2단계: 결과 출력 범위를 먼저 지정해주세요.", "#c00");
        return;
    }

    setRunStatus("환율 조회 중...", "#555");
    document.getElementById("btn-run").disabled = true;

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const dateRange = sheet.getRange(dateRangeAddress);
            const targetRange = sheet.getRange(targetRangeAddress);

            dateRange.load("text");
            await context.sync();

            let dateValues = dateRange.text;
            let resultValues = [];
            let successCount = 0;

            for (let i = 0; i < dateValues.length; i++) {
                let rawDate = dateValues[i][0];
                let formattedDate = formatToAPIString(rawDate);

                if (formattedDate.length === 8) {
                    try {
                        const response = await fetch(`${API_BASE}/get_rate?date=${formattedDate}&currency=${currency}`);
                        const result = await response.json();

                        if (result.rate !== undefined) {
                            resultValues.push([result.rate]);
                            successCount++;
                        } else {
                            resultValues.push(["데이터없음"]);
                        }
                    } catch (error) {
                        resultValues.push(["연결실패"]);
                    }
                } else {
                    resultValues.push([""]);
                }
            }

            targetRange.values = resultValues;
            await context.sync();

            setRunStatus(`완료: ${successCount}/${dateValues.length}건 조회됨`, "#217346");
        });
    } catch (e) {
        setRunStatus("오류: " + e.message, "#c00");
    } finally {
        document.getElementById("btn-run").disabled = false;
    }
}
