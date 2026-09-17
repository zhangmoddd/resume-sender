/* ============================================================
 * 「在线填表」撤销 / 重做 的回归测试用例
 *
 * 用法：python tests/test_sheet_undo.py
 * 说明：本文件只放"测试用例"，被测的函数是从 index.html 里现抽出来的，
 *      所以改了 index.html 再跑这个，测的就是最新代码，不会测到陈年副本。
 * ============================================================ */

let fails = 0;
function check(label, got, want) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (!ok) fails++;
  console.log((ok ? "[OK]   " : "[FAIL] ") + label + "  得到=" + JSON.stringify(got) + " 期望=" + JSON.stringify(want));
}
/* 有些场景以前会直接抛异常（比如行数被撤少了、选区还在旧行上），
   所以这里专门包一层：抛异常也算失败，而不是让整个脚本挂掉。 */
function safe(label, fn, want) {
  try { check(label, fn(), want); }
  catch (e) { fails++; console.log("[FAIL] " + label + "  ⚠ 抛异常了：" + e.message); }
}
function reset(rows) {
  SHEET_ROWS = rows; SEL = null; lastSelKey = "";
  SHEET_UNDO = []; SHEET_REDO = []; SHEET_EDIT_KEY = "";
  SHEET = Array.from({ length: rows }, (_, i) => [String.fromCharCode(97 + i), "", "", ""]);
}
const setA1 = v => { SHEET[0][0] = v; };
function edit(v) { pushSheetUndo("改 A1"); setA1(v); }
function snap() { return SHEET.map(r => r[0]).join(""); }

/* ---------- 1. 基本来回 ---------- */
reset(3);
check("初始内容", snap(), "abc");
edit("X");
check("改完一格", snap(), "Xbc");
sheetUndo();
check("撤销回原样", snap(), "abc");
sheetRedo();
check("重做回改后", snap(), "Xbc");
safe("重做到底后再点重做：不崩、只提示", () => { sheetRedo(); return lastToast.indexOf("没有可以重做") === 0; }, true);
safe("撤销到底后再点撤销：不崩、只提示", () => { sheetUndo(); sheetUndo(); return snap(); }, "abc");

/* ---------- 2. 撤销之后又动手改 → 重做作废（跟常见编辑器一致） ---------- */
reset(3);
edit("A"); sheetUndo();
check("撤完还有得重做", SHEET_REDO.length, 1);
edit("B");
check("重新动手后重做被清空", SHEET_REDO.length, 0);
safe("此时点重做会被拦住", () => { sheetRedo(); return lastToast.indexOf("没有可以重做") === 0; }, true);
check("新内容没被破坏", snap(), "Bbc");

/* ---------- 3. 行数被撤少的越界问题（专门修过的 bug，这里守住） ---------- */
reset(3);
SEL = { r1: 1, c1: 0, r2: 2, c2: 0 };
pushSheetUndo("粘贴撑大");
SHEET_ROWS = 8;
while (SHEET.length < 8) SHEET.push(["", "", "", ""]);
SEL = { r1: 6, c1: 0, r2: 7, c2: 3 };      // 选区落在"撑大后才有"的行上
safe("撤销把行数收回 3 行：不报错", () => { sheetUndo(); return [SHEET_ROWS, SHEET.length]; }, [3, 3]);
check("选区被收进有效范围", [SEL.r1, SEL.r2], [2, 2]);
safe("收完再刷新底部说明：不报错", () => { updateSheetFoot(); return typeof stubEl.innerHTML; }, "string");
safe("重做把行数还原回来", () => { sheetRedo(); return [SHEET_ROWS, SHEET.length]; }, [8, 8]);
safe("重做后选区也在范围内", () => [SEL.r1 <= SHEET_ROWS - 1, SEL.r2 <= SHEET_ROWS - 1], [true, true]);

/* ---------- 4. 选区指向不存在的行，也不该崩 ---------- */
reset(2);
SEL = { r1: 0, c1: 0, r2: 5, c2: 3 };
safe("框到不存在的行也扛得住", () => { updateSheetFoot(); return typeof stubEl.innerHTML; }, "string");
SEL = { r1: 0, c1: 0, r2: 9, c2: 3 };
safe("极端越界也不行崩", () => { updateSheetFoot(); return typeof stubEl.innerHTML; }, "string");

/* ---------- 5. 步数上限 ---------- */
reset(2);
for (let i = 0; i < 150; i++) edit("n" + i);
check("撤销栈封顶 100 步", SHEET_UNDO.length, 100);
for (let i = 0; i < 100; i++) sheetUndo();
check("撤满 100 步", [SHEET_UNDO.length, SHEET_REDO.length], [0, 100]);
safe("撤到底再撤：不崩", () => { sheetUndo(); return SHEET_REDO.length; }, 100);
for (let i = 0; i < 100; i++) sheetRedo();
check("重做满 100 步回到最新", snap(), "n149b");

/* ---------- 6. 来回横跳 250 轮，检查会不会错位 ---------- */
reset(3);
edit("P"); edit("Q"); edit("R");
let okAll = true, bad = "";
for (let i = 0; i < 250; i++) {
  sheetUndo();
  if (snap() !== "Qbc") { okAll = false; bad = "第 " + i + " 轮撤销后是 " + snap(); break; }
  sheetRedo();
  if (snap() !== "Rbc") { okAll = false; bad = "第 " + i + " 轮重做后是 " + snap(); break; }
}
check("来回横跳 250 轮都对" + (okAll ? "" : "（" + bad + "）"), okAll, true);
check("横跳后栈深度不变", [SHEET_UNDO.length, SHEET_REDO.length], [3, 0]);

/* ---------- 7. 一步一步退、一步一步进，每一步都要对得上 ---------- */
reset(3);
const hist = [];
edit("s1"); hist.push("s1bc");
edit("s2"); hist.push("s2bc");
edit("s3"); hist.push("s3bc");
let stepOk = true;
for (let i = hist.length - 1; i >= 0; i--) {
  sheetUndo();
  const want = i === 0 ? "abc" : hist[i - 1];
  if (snap() !== want) { stepOk = false; console.log("       回退时对不上：得到 " + snap() + " 期望 " + want); break; }
}
check("一步步回退到最初", stepOk, true);
stepOk = true;
for (const want of hist) {
  sheetRedo();
  if (snap() !== want) { stepOk = false; console.log("       重做时对不上：得到 " + snap() + " 期望 " + want); break; }
}
check("一步步重做到最新", stepOk, true);

console.log("\n结果：" + (fails ? "有 " + fails + " 项没过 ❌" : "全部通过 ✅"));
if (fails) process.exit(1);
