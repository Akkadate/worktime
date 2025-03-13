// Google Apps Script สำหรับเชื่อมต่อกับ Google Sheets
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('แบบฟอร์มสำรวจเวลาทำงานของพนักงาน')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ฟังก์ชันดึงข้อมูลพนักงานจากรหัสพนักงาน
function getEmployeeData(employeeId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const employeeSheet = ss.getSheetByName('employees'); // ชื่อ sheet ที่เก็บข้อมูลพนักงาน
    
    if (!employeeSheet) {
      throw new Error('ไม่พบ Sheet "employees"');
    }
    
    // ดึงข้อมูลทั้งหมดจาก sheet
    const data = employeeSheet.getDataRange().getValues();
    
    // วนลูปเพื่อหาพนักงานที่มีรหัสตรงกับที่ระบุ
    for (let i = 1; i < data.length; i++) { // เริ่มจาก i = 1 เพื่อข้ามส่วนหัวตาราง
      if (data[i][0] === employeeId) { // สมมติว่าคอลัมน์แรก (index 0) คือรหัสพนักงาน
        return {
          name: data[i][1], // สมมติว่าคอลัมน์ที่สอง (index 1) คือชื่อ-นามสกุล
          department: data[i][2] // สมมติว่าคอลัมน์ที่สาม (index 2) คือหน่วยงาน
        };
      }
    }
    
    // กรณีไม่พบข้อมูลพนักงาน
    return null;
  } catch (error) {
    Logger.log('Error in getEmployeeData: ' + error.toString());
    throw new Error('เกิดข้อผิดพลาดในการดึงข้อมูลพนักงาน: ' + error.toString());
  }
}

// ฟังก์ชันดึงข้อมูลช่วงเวลาทำงาน
function getWorkTimeOptions() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const worktimeSheet = ss.getSheetByName('worktimes'); // ชื่อ sheet ที่ 2 ที่เก็บข้อมูลช่วงเวลาทำงาน
    
    if (!worktimeSheet) {
      throw new Error('ไม่พบ Sheet "worktimes"');
    }
    
    // ดึงข้อมูลทั้งหมดจาก sheet
    const data = worktimeSheet.getDataRange().getValues();
    const options = [];
    
    // วนลูปเพื่อดึงข้อมูลช่วงเวลาทำงาน
    for (let i = 1; i < data.length; i++) { // เริ่มจาก i = 1 เพื่อข้ามส่วนหัวตาราง
      options.push({
        id: data[i][0].toString(), // สมมติว่าคอลัมน์แรก (index 0) คือ id - แปลงเป็น string
        name: data[i][1].toString() // สมมติว่าคอลัมน์ที่สอง (index 1) คือชื่อช่วงเวลาทำงาน
      });
    }
    
    return options;
  } catch (error) {
    Logger.log('Error in getWorkTimeOptions: ' + error.toString());
    throw new Error('เกิดข้อผิดพลาดในการดึงข้อมูลช่วงเวลาทำงาน: ' + error.toString());
  }
}

// ฟังก์ชันบันทึกข้อมูลลงใน sheet worktime
function saveWorkTimeData(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let worktimeSheet = ss.getSheetByName('worktime'); // ชื่อ sheet ที่จะบันทึกข้อมูล
    
    // สร้าง sheet ใหม่หากยังไม่มี
    if (!worktimeSheet) {
      worktimeSheet = ss.insertSheet('worktime');
      
      // สร้างหัวตารางพื้นฐาน
      const headers = [
        'วันเวลาที่บันทึก', 
        'รหัสพนักงาน', 
        'ชื่อพนักงาน', 
        'หน่วยงาน', 
        'รหัสช่วงเวลาทำงาน', 
        'ชื่อช่วงเวลาทำงาน', 
        'วันที่เริ่มต้น',
        'วันที่สิ้นสุด'
      ];
      
      // เพิ่มหัวตารางของวันที่ 1-14
      for (let i = 1; i <= 14; i++) {
        headers.push(`วันที่ ${i}`);
      }
      
      // บันทึกหัวตาราง
      worktimeSheet.appendRow(headers);
    }
    
    // เตรียมข้อมูลพื้นฐาน
    const basicRowData = [
      new Date(), // วันเวลาที่บันทึก
      formData.employeeId, // รหัสพนักงาน
      formData.employeeName, // ชื่อพนักงาน
      formData.department, // หน่วยงาน
      formData.worktimeId, // รหัสช่วงเวลาทำงาน
      formData.worktime, // ชื่อช่วงเวลาทำงาน
      formData.startDate, // วันที่เริ่มต้น
      formData.endDate // วันที่สิ้นสุด
    ];
    
    // เพิ่มข้อมูลวันทำงานทีละวัน (คอลัมน์ที่ 9-22 สำหรับ 14 วัน)
    formData.dates.forEach((dateData, index) => {
      // ข้อมูลแต่ละวันจะอยู่ในรูปแบบ "01/01/2025 (จันทร์) - วันทำงาน"
      basicRowData.push(`${dateData.date} (${dateData.dayOfWeek}) - ${dateData.status}`);
    });
    
    // เติมช่องว่างให้ครบ 14 วัน ในกรณีที่มีวันน้อยกว่า 14 วัน
    while (basicRowData.length < 22) { // 8 คอลัมน์พื้นฐาน + 14 คอลัมน์วัน
      basicRowData.push("");
    }
    
    // เพิ่มข้อมูลใหม่ลงในตาราง
    worktimeSheet.appendRow(basicRowData);
    
    return { 
      success: true, 
      message: `บันทึกข้อมูลสำเร็จ (${formData.dates.length} วัน)`
    };
  } catch (error) {
    Logger.log('Error in saveWorkTimeData: ' + error.toString());
    return { 
      success: false, 
      message: "เกิดข้อผิดพลาด: " + error.toString() 
    };
  }
}
