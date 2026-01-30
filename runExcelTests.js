const { chromium } = require("playwright");
const XLSX = require("xlsx");

// =====================
// SETTINGS (edit if needed)
// =====================
const EXCEL_PATH = "./Assignment_1_IT23543164.xlsx";
const SHEET_NAME = "Test cases"; // based on your screenshot tab name
const INPUT_COL = "D";          // your sheet: Input is column D
const ACTUAL_COL = "F";         // Actual output is column F
const STATUS_COL = "G";         // Status is column G
const START_ROW = 2;
const END_ROW = 200;
// =====================

async function findBestInputAndOutput(page) {
    // Wait a bit more for SPA pages
    await page.waitForLoadState("domcontentloaded");
    await page.waitForTimeout(1500);

    // Debug counts
    const taCount = await page.locator("textarea").count();
    const inputCount = await page.locator("input").count();
    const ceCount = await page.locator("[contenteditable='true']").count();
    console.log(`Debug: textarea=${taCount}, input=${inputCount}, contenteditable=${ceCount}`);

    // 1) Try common pattern: two textareas (input + output)
    if (taCount >= 2) {
        return {
            input: page.locator("textarea").first(),
            output: page.locator("textarea").nth(1),
            mode: "textarea-pair"
        };
    }

    // 2) Try labeled/placeholder based input
    const candidateInputs = [
        page.getByPlaceholder(/singlish|type|input|enter/i),
        page.getByLabel(/singlish|input/i),
        page.locator("textarea, input[type='text'], input:not([type]), [contenteditable='true']").first()
    ];

    let inputEl = null;
    for (const c of candidateInputs) {
        if (await c.count() > 0) { inputEl = c.first(); break; }
    }

    // 3) Try to find output by nearby text / label
    const candidateOutputs = [
        page.getByText(/sinhala/i).locator("xpath=following::textarea[1]"),
        page.getByText(/output|result|translation/i).locator("xpath=following::textarea[1]"),
        page.locator("textarea, [contenteditable='true']").last()
    ];

    let outputEl = null;
    for (const c of candidateOutputs) {
        if (await c.count() > 0) { outputEl = c.first(); break; }
    }

    if (!inputEl || !outputEl) {
        // Save screenshot for debugging
        await page.screenshot({ path: "swifttranslator_debug.png", fullPage: true });
        throw new Error(
            "Could not reliably detect input/output fields. Saved screenshot: swifttranslator_debug.png"
        );
    }

    return { input: inputEl, output: outputEl, mode: "heuristic" };
}

async function readOutput(outputEl) {
    // output could be textarea/input OR contenteditable div
    const tag = await outputEl.evaluate((el) => el.tagName.toLowerCase());
    if (tag === "textarea" || tag === "input") {
        return await outputEl.inputValue();
    }
    // contenteditable or other element
    return (await outputEl.innerText())?.trim();
}

async function writeInput(inputEl, value) {
    const tag = await inputEl.evaluate((el) => el.tagName.toLowerCase());
    if (tag === "textarea" || tag === "input") {
        await inputEl.fill("");
        await inputEl.type(String(value), { delay: 25 });
        return;
    }
    // contenteditable
    await inputEl.click();
    await inputEl.press("Control+A");
    await inputEl.type(String(value), { delay: 25 });
}

async function main() {
    const wb = XLSX.readFile(EXCEL_PATH);
    const ws = wb.Sheets[SHEET_NAME];
    if (!ws) throw new Error(`Sheet "${SHEET_NAME}" not found. (Check the tab name in Excel)`);

    const browser = await chromium.launch({ headless: false });
    const page = await browser.newPage();

    await page.goto("https://swifttranslator.com/", { waitUntil: "domcontentloaded" });

    const { input, output, mode } = await findBestInputAndOutput(page);
    console.log("Using mode:", mode);

    for (let row = START_ROW; row <= END_ROW; row++) {
        const inputCell = `${INPUT_COL}${row}`;
        const actualCell = `${ACTUAL_COL}${row}`;
        const statusCell = `${STATUS_COL}${row}`;

        const inputValue = ws[inputCell]?.v;
        if (!inputValue || String(inputValue).trim() === "") continue;

        console.log(`Row ${row} -> ${inputValue}`);

        await writeInput(input, inputValue);

        // Wait for auto update
        await page.waitForTimeout(2000);

        const out = await readOutput(output);

        ws[actualCell] = { t: "s", v: out || "" };
        ws[statusCell] = { t: "s", v: out ? "DONE" : "NO OUTPUT" };
    }

    const OUT_FILE = "Assignment1_AutoFilled.xlsx";
    XLSX.writeFile(wb, OUT_FILE);

    await browser.close();
    console.log(`✅ Finished. Saved: ${OUT_FILE}`);
}

main().catch((e) => {
    console.error("❌ Error:", e.message);
    process.exit(1);
});
