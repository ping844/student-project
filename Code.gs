/**
 * @OnlyCurrentDoc
 * สคริปต์สำหรับค้นหา PIN ของนักเรียนและนำมาเติมในชีต Scores
 * เวอร์ชันนี้จะตรวจสอบคอลัมน์ PIN ที่มีอยู่แล้ว และอัปเดตข้อมูลแทนการสร้างคอลัมน์ใหม่
 */
function processStudentPins() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentsSheet = ss.getSheetByName("Students");
  const scoresSheet = ss.getSheetByName("Scores");

  if (!studentsSheet || !scoresSheet) {
    SpreadsheetApp.getUi().alert("ไม่พบชีต 'Students' หรือ 'Scores'");
    return;
  }

  // 1. สร้าง Map ข้อมูลนักเรียนเพื่อการค้นหาที่รวดเร็ว
  const studentData = studentsSheet.getRange(2, 1, studentsSheet.getLastRow() - 1, 2).getValues();
  const studentPinMap = new Map();
  for (const row of studentData) {
    const name = row[0];
    const pin = row[1];
    if (name) {
      studentPinMap.set(name, pin);
    }
  }

  // 2. เตรียมข้อมูล PIN จากชีต Scores
  const scoresRange = scoresSheet.getRange(1, 1, scoresSheet.getLastRow(), scoresSheet.getLastColumn());
  const scoresData = scoresRange.getValues();
  const headerRow = scoresData[0];
  const nameColumnIndex = headerRow.indexOf("Name");

  if (nameColumnIndex === -1) {
    SpreadsheetApp.getUi().alert("ไม่พบคอลัมน์ 'Name' ในชีต 'Scores'");
    return;
  }
  
  const pinsToWrite = [];
  if (scoresData.length > 1) {
    for (let i = 1; i < scoresData.length; i++) {
      const studentName = scoresData[i][nameColumnIndex];
      const foundPin = studentPinMap.get(studentName) || "Not Found";
      pinsToWrite.push([foundPin]);
    }
  }
  
  if (pinsToWrite.length === 0) {
    SpreadsheetApp.getUi().alert("ไม่พบข้อมูลนักเรียนในชีต 'Scores' ที่จะประมวลผล");
    return;
  }

  // --- ส่วนที่ปรับปรุงใหม่ ---
  // 3. ตรวจสอบว่ามีคอลัมน์ 'PIN' อยู่แล้วหรือไม่
  let targetColumn;
  const existingPinColumnIndex = headerRow.indexOf("PIN");

  if (existingPinColumnIndex !== -1) {
    // ถ้ามีคอลัมน์ "PIN" อยู่แล้ว
    targetColumn = existingPinColumnIndex + 1; // +1 เพื่อแปลงเป็นเลขคอลัมน์ที่ถูกต้อง
    // ล้างข้อมูลเก่าในคอลัมน์ PIN (ตั้งแต่แถวที่ 2 ลงไป)
    if (scoresSheet.getLastRow() > 1) {
      scoresSheet.getRange(2, targetColumn, scoresSheet.getLastRow() - 1, 1).clearContent();
    }
  } else {
    // ถ้ายังไม่มี ให้สร้างคอลัมน์ใหม่
    targetColumn = scoresSheet.getLastColumn() + 1;
    scoresSheet.getRange(1, targetColumn).setValue("PIN"); // สร้างหัวข้อให้คอลัมน์ใหม่
  }
  
  // 4. เขียนข้อมูล PIN ทั้งหมดลงในคอลัมน์เป้าหมาย
  scoresSheet.getRange(2, targetColumn, pinsToWrite.length, 1).setValues(pinsToWrite);
  // --- สิ้นสุดส่วนที่ปรับปรุง ---

  SpreadsheetApp.getUi().alert("ประมวลผลสำเร็จ! จำนวน " + pinsToWrite.length + " รายการ");
}