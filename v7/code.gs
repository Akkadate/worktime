// Google Apps Script สำหรับเชื่อมต่อกับ Google Sheets
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('แบบฟอร์มสำรวจเวลาทำงานของพนักงาน')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ฟังก์ชันดึงข้อมูลพนักงานจากรหัสพนักงาน
function getEmployeeData(employeeId) {
  try {
    // สร้างข้อมูลทดสอบสำหรับตรวจสอบการทำงาน
    const testData = {
      'E001': { name: 'สมชาย ใจดี', department: 'ฝ่ายบุคคล' },
      'E002': { name: 'สมหญิง รักงาน', department: 'ฝ่ายการเงิน' },
      'E003': { name: 'สมศักดิ์ มุ่งมั่น', department: 'ฝ่ายขาย' }
    };
    
    // ตรวจสอบว่าผู้ใช้ต้องการใช้ข้อมูลทดสอบหรือไม่
    if (employeeId in testData) {
      return testData[employeeId];
    }
    
    // ถ้าไม่ใช่ข้อมูลทดสอบ ให้ดึงข้อมูลจาก Google Sheets
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const employeeSheet = ss.getSheetByName('employees');
    
    if (!employeeSheet) {
      // ถ้าไม่พบ sheet "employees" ให้สร้างและเพิ่มข้อมูลตัวอย่าง
      createEmployeeSheet(ss);
      return null; // ส่งค่าว่างในครั้งแรก
    }
    
    // ดึงข้อมูลทั้งหมดจาก sheet
    const data = employeeSheet.getDataRange().getValues();
    
    // ตรวจสอบว่ามีข้อมูลในแผ่นงานหรือไม่
    if (data.length <= 1) {
      // ถ้ามีเพียงแถวหัวตาราง แสดงว่ายังไม่มีข้อมูลพนักงาน
      return null;
    }
    
    // วนลูปเพื่อหาพนักงานที่มีรหัสตรงกับที่ระบุ
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === employeeId) {
        return {
          name: data[i][1],
          department: data[i][2]
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
    const worktimeSheet = ss.getSheetByName('worktimes');
    
    if (!worktimeSheet) {
      // ถ้าไม่พบ sheet "worktimes" ให้สร้างและเพิ่มข้อมูลตัวอย่าง
      createWorktimeSheet(ss);
      
      // ส่งค่าตัวอย่างกลับไปใช้งานก่อน
      return [
        { id: '1', name: 'กะเช้า (08:00 - 16:00)' },
        { id: '2', name: 'กะบ่าย (16:00 - 00:00)' },
        { id: '3', name: 'กะดึก (00:00 - 08:00)' },
        { id: '4', name: 'งานปกติ (09:00 - 18:00)' }
      ];
    }
    
    // ดึงข้อมูลทั้งหมดจาก sheet
    const data = worktimeSheet.getDataRange().getValues();
    const options = [];
    
    // วนลูปเพื่อดึงข้อมูลช่วงเวลาทำงาน
    for (let i = 1; i < data.length; i++) { // เริ่มจาก i = 1 เพื่อข้ามส่วนหัวตาราง
      options.push({
        id: data[i][0].toString(), // สมมติว่าคอลัมน์แรก (index 0) คือ id
        name: data[i][1].toString() // สมมติว่าคอลัมน์ที่สอง (index 1) คือชื่อช่วงเวลาทำงาน
      });
    }
    
    return options.length > 0 ? options : [
      { id: '1', name: 'กะเช้า (08:00 - 16:00)' },
      { id: '2', name: 'กะบ่าย (16:00 - 00:00)' },
      { id: '3', name: 'กะดึก (00:00 - 08:00)' },
      { id: '4', name: 'งานปกติ (09:00 - 18:00)' }
    ];
  } catch (error) {
    Logger.log('Error in getWorkTimeOptions: ' + error.toString());
    throw new Error('เกิดข้อผิดพลาดในการดึงข้อมูลช่วงเวลาทำงาน: ' + error.toString());
  }
}

// ฟังก์ชันสร้าง sheet "employees" และเพิ่มข้อมูลตัวอย่าง
function createEmployeeSheet(ss) {
  try {
    const sheet = ss.insertSheet('employees');
    
    // เพิ่มหัวตาราง
    sheet.appendRow(['รหัสพนักงาน', 'ชื่อ-นามสกุล', 'หน่วยงาน']);
    
    // เพิ่มข้อมูลตัวอย่าง
    sheet.appendRow(['E001', 'สมชาย ใจดี', 'ฝ่ายบุคคล']);
    sheet.appendRow(['E002', 'สมหญิง รักงาน', 'ฝ่ายการเงิน']);
    sheet.appendRow(['E003', 'สมศักดิ์ มุ่งมั่น', 'ฝ่ายขาย']);
    
    // จัดรูปแบบ
    sheet.getRange('A1:C1').setBackground('#D9EAD3').setFontWeight('bold');
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 150);
    
    return sheet;
  } catch (error) {
    Logger.log('Error in createEmployeeSheet: ' + error.toString());
    throw new Error('เกิดข้อผิดพลาดในการสร้าง sheet employees: ' + error.toString());
  }
}

// ฟังก์ชันสร้าง sheet "worktimes" และเพิ่มข้อมูลตัวอย่าง
function createWorktimeSheet(ss) {
  try {
    const sheet = ss.insertSheet('worktimes');
    
    // เพิ่มหัวตาราง
    sheet.appendRow(['ID', 'ชื่อช่วงเวลาทำงาน']);
    
    // เพิ่มข้อมูลตัวอย่าง
    sheet.appendRow(['1', 'กะเช้า (08:00 - 16:00)']);
    sheet.appendRow(['2', 'กะบ่าย (16:00 - 00:00)']);
    sheet.appendRow(['3', 'กะดึก (00:00 - 08:00)']);
    sheet.appendRow(['4', 'งานปกติ (09:00 - 18:00)']);
    
    // จัดรูปแบบ
    sheet.getRange('A1:B1').setBackground('#D9EAD3').setFontWeight('bold');
    sheet.setColumnWidth(1, 100);
    sheet.setColumnWidth(2, 250);
    
    return sheet;
  } catch (error) {
    Logger.log('Error in createWorktimeSheet: ' + error.toString());
    throw new Error('เกิดข้อผิดพลาดในการสร้าง sheet worktimes: ' + error.toString());
  }
}

// ฟังก์ชันบันทึกข้อมูลลงใน sheet worktime
function saveWorkTimeData(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let worktimeSheet = ss.getSheetByName('worktime');
    
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
        'รูปแบบวันทำงาน',
        'วันที่เริ่มต้น',
        'วันที่สิ้นสุด'
      ];
      
      // เพิ่มหัวตารางของวันที่ 1-14
      for (let i = 1; i <= 14; i++) {
        headers.push(`วันที่ ${i}`);
      }
      
      // บันทึกหัวตาราง
      worktimeSheet.appendRow(headers);
      
      // จัดรูปแบบ
      worktimeSheet.getRange('A1:W1').setBackground('#D9EAD3').setFontWeight('bold');
    }
    
    // เตรียมข้อมูลพื้นฐาน
    const basicRowData = [
      new Date(), // วันเวลาที่บันทึก
      formData.employeeId, // รหัสพนักงาน
      formData.employeeName, // ชื่อพนักงาน
      formData.department, // หน่วยงาน
      formData.worktimeId, // รหัสช่วงเวลาทำงาน
      formData.worktime, // ชื่อช่วงเวลาทำงาน
      formData.workPattern === 'same' ? 'วันทำงานเหมือนกันทุกสัปดาห์' : 'มีการสลับวันทำงานทุกสัปดาห์ (5.5 วัน)', // รูปแบบวันทำงาน
      formData.startDate, // วันที่เริ่มต้น
      formData.endDate // วันที่สิ้นสุด
    ];
    
    // เพิ่มข้อมูลวันทำงานทีละวัน (คอลัมน์ที่ 10-23 สำหรับ 14 วัน)
    if (formData.dates && formData.dates.length > 0) {
      formData.dates.forEach((dateData) => {
        // บันทึกเฉพาะค่าสถานะ 1 หรือ 0
        basicRowData.push(dateData.status);
      });
      
      // เติมช่องว่างให้ครบ 14 วัน ในกรณีที่มีวันน้อยกว่า 14 วัน
      while (basicRowData.length < 23) { // 9 คอลัมน์พื้นฐาน + 14 คอลัมน์วัน
        basicRowData.push("");
      }
    } else {
      // กรณีไม่มีข้อมูลวันที่
      for (let i = 0; i < 14; i++) {
        basicRowData.push("");
      }
    }
    
    // เพิ่มข้อมูลใหม่ลงในตาราง
    worktimeSheet.appendRow(basicRowData);
    
    return { 
      success: true, 
      message: `บันทึกข้อมูลสำเร็จ ${formData.dates ? '(' + formData.dates.length + ' วัน)' : ''}`
    };
  } catch (error) {
    Logger.log('Error in saveWorkTimeData: ' + error.toString());
    return { 
      success: false, 
      message: "เกิดข้อผิดพลาด: " + error.toString() 
    };
  }
}
