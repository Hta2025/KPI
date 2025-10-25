// Code.gs

// هذه الدالة تخدم صفحة الـ HTML عند فتح التطبيق
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('نظام إدخال مؤشرات الأداء')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// دالة لجلب أسماء الدوائر الفريدة من ملف الماستر
function getDepartments() {
  // تم تحديث المعرف
  const sheet = SpreadsheetApp.openById('1UgIiVqzaSCLi1o7iCtiie3L800BkP0rmJTq9_ftilLk').getSheetByName('KPIs');
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues(); // العمود الأول starting from row 2
  const uniqueDepartments = [...new Set(data.flat())];
  return uniqueDepartments.filter(dept => dept !== ''); // إزالة القيم الفارغة
}

// دالة لجلب المؤشرات الخاصة بدائرة معينة
function getKPIs(department) {
  // تم تحديث المعرف
  const sheet = SpreadsheetApp.openById('1UgIiVqzaSCLi1o7iCtiie3L800BkP0rmJTq9_ftilLk').getSheetByName('KPIs');
  const data = sheet.getDataRange().getValues();
  const headers = data.shift(); // إزالة رأس الجدول
  const kpis = data.filter(row => row[0] === department); // فلترة حسب اسم الدائرة
  
  return kpis.map(row => ({
    name: row[1],
    definition: row[2],
    type: row[3],
    email: row[4]
  }));
}

// الدالة الرئيسية لحفظ البيانات وإرسال الإيميل
function submitData(payload) {
  try {
    // تم تحديث المعرفات
    const masterSheetId = '1UgIiVqzaSCLi1o7iCtiie3L800BkP0rmJTq9_ftilLk';
    const outputSheetId = '1CNKRA3pSziBUSzovSaI4fhBDa64XH1O-awiLKfarv4Y';
    
    const { department, year, month, kpiData } = payload;
    
    // 1. حفظ البيانات في شيت الـ Output
    const outputSheet = SpreadsheetApp.openById(outputSheetId).getSheetByName('Sheet1');
    const timestamp = new Date();
    
    kpiData.forEach(kpi => {
      outputSheet.appendRow([
        timestamp,
        department,
        year,
        month,
        kpi.name,
        kpi.numerator || '',
        kpi.denominator || '',
        kpi.value || '',
        kpi.result
      ]);
    });

    // 2. إرسال إيميل التأكيد
    sendConfirmationEmail(department, year, masterSheetId, outputSheetId);

    return { success: true, message: "تم حفظ البيانات وإرسال الإيميل بنجاح!" };

  } catch (e) {
    Logger.log(e);
    return { success: false, message: "حدث خطأ: " + e.toString() };
  }
}

// دالة مساعدة لإرسال الإيميل
function sendConfirmationEmail(department, year, masterSheetId, outputSheetId) {
  // جلب إيميل المسؤول من شيت الماستر
  const masterSheet = SpreadsheetApp.openById(masterSheetId).getSheetByName('KPIs');
  const masterData = masterSheet.getDataRange().getValues();
  const responsibleEmail = masterData.find(row => row[0] === department)?.[4]; // البحث عن الإيميل
  
  if (!responsibleEmail) {
    Logger.log("لم يتم العثور على إيميل مسؤول للدائرة: " + department);
    return;
  }

  // جلب البيانات من شيت الـ Output لإنشاء الجدول
  const outputSheet = SpreadsheetApp.openById(outputSheetId).getSheetByName('Sheet1');
  const outputData = outputSheet.getDataRange().getValues();
  const headers = outputData.shift();
  
  // فلترة البيانات حسب الدائرة والسنة
  const filteredData = outputData.filter(row => row[1] === department && row[2] == year);
  
  // تحويل البيانات الطويلة إلى جدول عريض (Pivoting)
  const pivotData = {};
  const monthNames = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"];
  
  filteredData.forEach(row => {
    const kpiName = row[4];
    const monthName = monthNames[row[3] - 1]; // تحويل رقم الشهر إلى اسم
    const result = row[8];
    
    if (!pivotData[kpiName]) {
      pivotData[kpiName] = {};
    }
    pivotData[kpiName][monthName] = result;
  });

  // بناء جدول HTML
  let htmlTable = '<table border="1" style="border-collapse: collapse; width: 100%;"><tr><th>اسم المؤشر</th>';
  monthNames.forEach(month => htmlTable += `<th>${month}</th>`);
  htmlTable += '</tr>';

  for (const kpiName in pivotData) {
    htmlTable += `<tr><td>${kpiName}</td>`;
    monthNames.forEach(month => {
      const value = pivotData[kpiName][month] || '(فارغ)';
      htmlTable += `<td style="text-align: center;">${value}</td>`;
    });
    htmlTable += '</tr>';
  }
  htmlTable += '</table>';

  const subject = `تقرير مؤشرات الأداء الشهري - ${department} - ${year}`;
  const body = `
    <p>عزيزي/عزيزتي مسؤول الدائرة،</p>
    <p>تم تحديث بيانات مؤشرات الأداء لدائرة <strong>${department}</strong> للسنة <strong>${year}</strong>.</p>
    <p>تفضل بملخص البيانات المدخلة:</p>
    ${htmlTable}
    <p>مع أطيب التحيات،<br>نظام إدخال مؤشرات الأداء</p>
  `;

  MailApp.sendEmail({
    to: responsibleEmail,
    subject: subject,
    htmlBody: body
  });
}