const APP_VERSION = "v2.0.0";
const API_BASE = "https://fx-proxy.jang375a-03c.workers.dev";

// ── 상태 변수 ──
let dateRangeAddress = "";
let targetRangeAddress = "";
let currentMode = "daily"; // "daily" | "monthly" | "yearly"

// ── ECOS 통화 매핑 ──
// stat_code 731Y003: 주요국 통화의 대원화 환율 (USD, JPY, CNY 등)
// stat_code 731Y004: 기타 통화의 대원화 환율 (EUR, GBP 등)
// ※ item_code는 ECOS 실제 조회로 확인 후 보완 필요
const ECOS_CURRENCY_MAP = {
    "USD":  { stat_code: "731Y001", item_code: "0000001" },
    "JPY(100)": { stat_code: "731Y001", item_code: "0000002" },
    "EUR":  { stat_code: "731Y001", item_code: "0000003" },
    "GBP":  { stat_code: "731Y001", item_code: "0000012" },
    "CAD":  { stat_code: "731Y001", item_code: "0000013" },
    "CHF":  { stat_code: "731Y001", item_code: "0000014" },
    "HKD":  { stat_code: "731Y001", item_code: "0000015" },
    "SEK":  { stat_code: "731Y001", item_code: "0000016" },
    "AUD":  { stat_code: "731Y001", item_code: "0000017" },
    "DKK":  { stat_code: "731Y001", item_code: "0000018" },
    "NOK":  { stat_code: "731Y001", item_code: "0000019" },
    "SAR":  { stat_code: "731Y001", item_code: "0000020" },
    "KWD":  { stat_code: "731Y001", item_code: "0000021" },
    "BHD":  { stat_code: "731Y001", item_code: "0000022" },
    "AED":  { stat_code: "731Y001", item_code: "0000023" },
    "SGD":  { stat_code: "731Y001", item_code: "0000024" },
    "MYR":  { stat_code: "731Y001", item_code: "0000025" },
    "NZD":  { stat_code: "731Y001", item_code: "0000026" },
    "THB":  { stat_code: "731Y001", item_code: "0000028" },
    "IDR(100)": { stat_code: "731Y001", item_code: "0000029" },
    "TWD":  { stat_code: "731Y001", item_code: "0000031" },
    "CNH":  { stat_code: "731Y001", item_code: "0000053" },
};

// ── Office Ready ──
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("btn-set-date").onclick = setDateRange;
        document.getElementById("btn-set-target").onclick = setTargetRange;
        document.getElementById("btn-run").onclick = runExchangeRateFetch;
        document.getElementById("app-version").textContent = APP_VERSION;

        // 모드 탭 이벤트
        document.querySelectorAll(".mode-tab").forEach(tab => {
            tab.addEventListener("click", () => switchMode(tab.dataset.mode));
        });

        // ECOS 키 저장 버튼
        document.getElementById("btn-save-ecos-key").onclick = saveEcosKey;
        loadEcosKey();

        loadCurrencies();
    }
});

// ── 모드 전환 ──
function switchMode(mode) {
    currentMode = mode;

    // 탭 활성화
    document.querySelectorAll(".mode-tab").forEach(tab => {
        tab.classList.toggle("active", tab.dataset.mode === mode);
    });

    // ECOS 섹션 표시/숨김
    const ecosSection = document.getElementById("ecos-section");
    const dateHint = document.getElementById("date-hint");
    const btnRun = document.getElementById("btn-run");

    if (mode === "daily") {
        ecosSection.style.display = "none";
        dateHint.textContent = "날짜 형식: 2024-12-31, 20241231 등";
        btnRun.textContent = "환율 정보 가져오기 (수출입은행)";
    } else if (mode === "monthly") {
        ecosSection.style.display = "block";
        dateHint.textContent = "날짜 형식: 202401, 2024-01, 2024년 1월 등";
        btnRun.textContent = "월평균 환율 가져오기 (ECOS)";
    } else if (mode === "yearly") {
        ecosSection.style.display = "block";
        dateHint.textContent = "날짜 형식: 2024, 2023 등 (연도만)";
        btnRun.textContent = "연평균 환율 가져오기 (ECOS)";
    }

    // 상태 초기화
    document.getElementById("run-status").style.display = "none";
}

// ── ECOS 키 관리 (localStorage) ──
function saveEcosKey() {
    const key = document.getElementById("ecos-key-input").value.trim();
    if (!key) {
        showEcosKeyStatus("키를 입력해주세요.", "#c00");
        return;
    }
    try { localStorage.setItem("fxm_ecos_key", key); } catch(e) {}
    showEcosKeyStatus("저장 완료 ✓", "#217346");
}

function loadEcosKey() {
    try {
        const key = localStorage.getItem("fxm_ecos_key");
        if (key) document.getElementById("ecos-key-input").value = key;
    } catch(e) {}
}

function getEcosKey() {
    return document.getElementById("ecos-key-input").value.trim();
}

function showEcosKeyStatus(msg, color) {
    const el = document.getElementById("ecos-key-status");
    el.textContent = msg;
    el.style.color = color || "#333";
    el.style.display = "block";
    setTimeout(() => { el.style.display = "none"; }, 3000);
}

// ── 통화 목록 로드 (기존) ──
async function loadCurrencies() {
    try {
        const response = await fetch(`${API_BASE}/get_currencies`);
        const list = await response.json();
        if (typeof window.setCurrencyList === "function") {
            window.setCurrencyList(list);
        }
    } catch (e) {
        // 무시
    }
}

// ── 범위 지정 (기존) ──
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

// ═══════════════════════════════════════
//  날짜 포맷 변환
// ═══════════════════════════════════════

// 일별: → YYYYMMDD
function formatToDateString(rawDateStr) {
    if (!rawDateStr) return "";
    const str = String(rawDateStr).trim();

    if (/^\d{8}$/.test(str)) return str;

    // Excel 시리얼 넘버
    if (/^\d{5}$/.test(str)) {
        const d = new Date(Date.UTC(1899, 11, 30) + parseInt(str) * 86400000);
        return d.getUTCFullYear().toString() +
               String(d.getUTCMonth() + 1).padStart(2, '0') +
               String(d.getUTCDate()).padStart(2, '0');
    }

    // 2024년 3월 15일
    let m = str.match(/(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})/);
    if (m) return m[1] + m[2].padStart(2, '0') + m[3].padStart(2, '0');

    // 2024-03-15, 2024.03.15, 2024/03/15
    m = str.match(/(\d{4})[.\-\/](\d{1,2})[.\-\/](\d{1,2})/);
    if (m) return m[1] + m[2].padStart(2, '0') + m[3].padStart(2, '0');

    return "";
}

// 월별: → YYYYMM
function formatToMonthString(rawDateStr) {
    if (!rawDateStr) return "";
    const str = String(rawDateStr).trim();

    if (/^\d{6}$/.test(str)) return str;
    if (/^\d{8}$/.test(str)) return str.substring(0, 6);

    // 2024-01, 2024.01, 2024/01
    let m = str.match(/^(\d{4})[.\-\/](\d{1,2})$/);
    if (m) return m[1] + m[2].padStart(2, '0');

    // 2024년 1월
    m = str.match(/(\d{4})\s*년\s*(\d{1,2})\s*월/);
    if (m) return m[1] + m[2].padStart(2, '0');

    // 2024-01-15 → 202401
    m = str.match(/(\d{4})[.\-\/](\d{1,2})[.\-\/]\d{1,2}/);
    if (m) return m[1] + m[2].padStart(2, '0');

    // Excel 시리얼
    if (/^\d{5}$/.test(str)) {
        const d = new Date(Date.UTC(1899, 11, 30) + parseInt(str) * 86400000);
        return d.getUTCFullYear().toString() + String(d.getUTCMonth() + 1).padStart(2, '0');
    }

    return "";
}

// 연별: → YYYY
function formatToYearString(rawDateStr) {
    if (!rawDateStr) return "";
    const str = String(rawDateStr).trim();

    if (/^\d{4}$/.test(str)) return str;
    if (/^\d{6}$/.test(str)) return str.substring(0, 4);
    if (/^\d{8}$/.test(str)) return str.substring(0, 4);

    let m = str.match(/(\d{4})\s*년/);
    if (m) return m[1];

    m = str.match(/^(\d{4})[.\-\/]/);
    if (m) return m[1];

    if (/^\d{5}$/.test(str)) {
        const d = new Date(Date.UTC(1899, 11, 30) + parseInt(str) * 86400000);
        return d.getUTCFullYear().toString();
    }

    return "";
}

// ── ECOS 통화 매핑 조회 (JPY(100) → JPY(100), JPY → JPY(100) 폴백) ──
function findEcosMapping(code) {
    // 정확히 일치
    if (ECOS_CURRENCY_MAP[code]) return ECOS_CURRENCY_MAP[code];
    // 괄호 포함 키에서 앞부분만 매칭 (예: "JPY" → "JPY(100)")
    for (const key of Object.keys(ECOS_CURRENCY_MAP)) {
        if (key.startsWith(code + "(") || key === code) {
            return ECOS_CURRENCY_MAP[key];
        }
    }
    return null;
}

// ── 상태 표시 ──
function setRunStatus(msg, color) {
    const box = document.getElementById("run-status");
    box.style.display = "block";
    box.style.color = color || "#333";
    box.textContent = msg;
}

// ═══════════════════════════════════════
//  메인 실행 (모드별 분기)
// ═══════════════════════════════════════
async function runExchangeRateFetch() {
    if (currentMode === "daily") {
        await runDailyFetch();
    } else {
        await runEcosFetch();
    }
}

// ═══════════════════════════════════════
//  일별 환율 (기존 수출입은행 API)
// ═══════════════════════════════════════
async function runDailyFetch() {
    const currency = document.getElementById("currency-input").value.trim().toUpperCase();

    if (!currency) return setRunStatus("통화 코드를 입력해주세요. (예: USD)", "#c00");
    if (!dateRangeAddress) return setRunStatus("1단계: 날짜 범위를 먼저 지정해주세요.", "#c00");
    if (!targetRangeAddress) return setRunStatus("2단계: 결과 출력 범위를 먼저 지정해주세요.", "#c00");

    setRunStatus("환율 조회 중...", "#555");
    document.getElementById("btn-run").disabled = true;

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const dateRange = sheet.getRange(dateRangeAddress);
            const targetRange = sheet.getRange(targetRangeAddress);

            dateRange.load(["text", "rowCount", "columnCount"]);
            await context.sync();

            let dateValues = dateRange.text;
            const rows = dateRange.rowCount;
            const cols = dateRange.columnCount;
            const isHorizontal = (rows === 1 && cols > 1);

            // 날짜 추출 (가로/세로)
            let rawDates = [];
            if (isHorizontal) {
                for (let c = 0; c < cols; c++) rawDates.push(dateValues[0][c]);
            } else {
                for (let r = 0; r < rows; r++) rawDates.push(dateValues[r][0]);
            }

            let results = [];
            let successCount = 0;

            for (let i = 0; i < rawDates.length; i++) {
                let formattedDate = formatToDateString(rawDates[i]);

                if (formattedDate.length === 8) {
                    try {
                        const response = await fetch(`${API_BASE}/get_rate?date=${formattedDate}&currency=${currency}`);
                        const result = await response.json();

                        if (result.rate !== undefined) {
                            results.push(result.rate);
                            successCount++;
                        } else {
                            results.push("데이터없음");
                        }
                    } catch (error) {
                        results.push("연결실패");
                    }
                } else {
                    results.push("");
                }
            }

            // 결과를 가로/세로에 맞게 변환
            let resultValues;
            if (isHorizontal) {
                resultValues = [results];           // [[v1, v2, v3, ...]]
            } else {
                resultValues = results.map(v => [v]); // [[v1],[v2],[v3],...]
            }

            targetRange.values = resultValues;
            await context.sync();

            setRunStatus(`완료: ${successCount}/${rawDates.length}건 조회됨`, "#217346");
        });
    } catch (e) {
        setRunStatus("오류: " + e.message, "#c00");
    } finally {
        document.getElementById("btn-run").disabled = false;
    }
}

// ═══════════════════════════════════════
//  ECOS 월평균 / 연평균 환율
// ═══════════════════════════════════════
async function runEcosFetch() {
    const currency = document.getElementById("currency-input").value.trim().toUpperCase();
    const ecosKey = getEcosKey();
    const isMonthly = (currentMode === "monthly");
    const cycle = isMonthly ? "M" : "A";
    const modeLabel = isMonthly ? "월평균" : "연평균";

    if (!currency) return setRunStatus("통화 코드를 입력해주세요. (예: USD)", "#c00");
    if (!ecosKey) return setRunStatus("ECOS 인증키를 입력해주세요.", "#c00");
    if (!dateRangeAddress) return setRunStatus("1단계: 날짜 범위를 먼저 지정해주세요.", "#c00");
    if (!targetRangeAddress) return setRunStatus("2단계: 결과 출력 범위를 먼저 지정해주세요.", "#c00");

    // ECOS 통화 매핑 확인 (JPY(100) 등 괄호 포함 코드도 처리)
    const mapping = findEcosMapping(currency);
    if (!mapping) {
        const supported = Object.keys(ECOS_CURRENCY_MAP).join(", ");
        return setRunStatus(`${currency}는 ECOS 미지원 통화입니다.\n지원: ${supported}`, "#c00");
    }

    setRunStatus(`${modeLabel} 환율 조회 중...`, "#555");
    document.getElementById("btn-run").disabled = true;

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const dateRange = sheet.getRange(dateRangeAddress);
            const targetRange = sheet.getRange(targetRangeAddress);

            dateRange.load(["text", "rowCount", "columnCount"]);
            await context.sync();

            const dateValues = dateRange.text;
            const rows = dateRange.rowCount;
            const cols = dateRange.columnCount;
            const isHorizontal = (rows === 1 && cols > 1);
            const formatFn = isMonthly ? formatToMonthString : formatToYearString;

            // 1) 날짜 추출 (가로/세로 모두 지원)
            let rawDates = [];
            if (isHorizontal) {
                for (let c = 0; c < cols; c++) rawDates.push(dateValues[0][c]);
            } else {
                for (let r = 0; r < rows; r++) rawDates.push(dateValues[r][0]);
            }

            const periods = rawDates.map(d => formatFn(d));
            const validPeriods = periods.filter(p => p.length > 0);

            if (validPeriods.length === 0) {
                setRunStatus("유효한 날짜가 없습니다.", "#c00");
                document.getElementById("btn-run").disabled = false;
                return;
            }

            const minPeriod = validPeriods.reduce((a, b) => a < b ? a : b);
            const maxPeriod = validPeriods.reduce((a, b) => a > b ? a : b);

            // 2) ECOS API 호출 (프록시 경유)
            const params = new URLSearchParams({
                apikey: ecosKey,
                stat_code: mapping.stat_code,
                item_code: mapping.item_code,
                cycle: cycle,
                start: minPeriod,
                end: maxPeriod
            });

            const response = await fetch(`${API_BASE}/ecos_search?${params}`);
            const data = await response.json();

            if (data.error) {
                setRunStatus("ECOS 오류: " + data.error, "#c00");
                document.getElementById("btn-run").disabled = false;
                return;
            }

            // 3) TIME → 환율값 매핑 생성
            const rateMap = {};
            if (data.rows && data.rows.length > 0) {
                data.rows.forEach(row => {
                    rateMap[row.time] = parseFloat(row.value);
                });
            }

            // 4) 결과 배열 생성 (가로/세로 맞춤)
            let resultValues;
            let successCount = 0;
            const totalCount = periods.length;

            if (isHorizontal) {
                let rowArr = [];
                for (let i = 0; i < totalCount; i++) {
                    const period = periods[i];
                    if (period && rateMap[period] !== undefined) {
                        rowArr.push(rateMap[period]);
                        successCount++;
                    } else if (period) {
                        rowArr.push("데이터없음");
                    } else {
                        rowArr.push("");
                    }
                }
                resultValues = [rowArr];  // [[v1, v2, v3, ...]]
            } else {
                resultValues = [];
                for (let i = 0; i < totalCount; i++) {
                    const period = periods[i];
                    if (period && rateMap[period] !== undefined) {
                        resultValues.push([rateMap[period]]);
                        successCount++;
                    } else if (period) {
                        resultValues.push(["데이터없음"]);
                    } else {
                        resultValues.push([""]);
                    }
                }
            }

            targetRange.values = resultValues;
            await context.sync();

            setRunStatus(`${modeLabel} 완료: ${successCount}/${totalCount}건 조회됨`, "#217346");
        });
    } catch (e) {
        setRunStatus("오류: " + e.message, "#c00");
    } finally {
        document.getElementById("btn-run").disabled = false;
    }
}
