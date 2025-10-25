// script.js

document.addEventListener('DOMContentLoaded', function() {
  const departmentSelect = document.getElementById('departmentSelect');
  const yearSelect = document.getElementById('yearSelect');
  const monthSelect = document.getElementById('monthSelect');
  const loadKpisBtn = document.getElementById('loadKpisBtn');
  const kpiContainer = document.getElementById('kpiContainer');
  const submitBtn = document.getElementById('submitBtn');
  const statusDiv = document.getElementById('status');

  // ملء قائمة السنوات (السنة الحالية + 3 سنوات سابقة)
  const currentYear = new Date().getFullYear();
  for (let i = 0; i < 4; i++) {
    const option = document.createElement('option');
    option.value = currentYear - i;
    option.textContent = currentYear - i;
    yearSelect.appendChild(option);
  }
  yearSelect.value = currentYear; // تحديد السنة الحالية افتراضياً

  // جلب أسماء الدوائر وملء القائمة المنسدلة
  google.script.run.withSuccessHandler(departments => {
    departments.forEach(dept => {
      const option = document.createElement('option');
      option.value = dept;
      option.textContent = dept;
      departmentSelect.appendChild(option);
    });
  }).getDepartments();

  // عند الضغط على زر "تحميل المؤشرات"
  loadKpisBtn.addEventListener('click', () => {
    const selectedDepartment = departmentSelect.value;
    if (!selectedDepartment) {
      alert('الرجاء اختيار الدائرة أولاً');
      return;
    }
    
    statusDiv.textContent = 'جاري التحميل...';
    google.script.run.withSuccessHandler(renderKPIs).getKPIs(selectedDepartment);
  });

  // دالة لعرض حقول إدخال المؤشرات
  function renderKPIs(kpis) {
    kpiContainer.innerHTML = ''; // مسح المحتوى القديم
    if (kpis.length === 0) {
      statusDiv.textContent = 'لا توجد مؤشرات لهذه الدائرة.';
      submitBtn.style.display = 'none';
      return;
    }
    
    kpis.forEach((kpi, index) => {
      const kpiDiv = document.createElement('div');
      kpiDiv.className = 'kpi-item';
      
      let inputFields = '';
      if (kpi.type === 'معادلة') {
        inputFields = `
          <div class="input-group">
            <label>البسط:</label>
            <input type="number" id="numerator_${index}" oninput="calculateResult(${index})">
            <label>المقام:</label>
            <input type="number" id="denominator_${index}" oninput="calculateResult(${index})">
            <span>النتيجة: </span><span id="result_${index}">0</span>%
          </div>
        `;
      } else { // رقمي
        inputFields = `
          <div class="input-group">
            <label>القيمة:</label>
            <input type="number" id="value_${index}" oninput="calculateResult(${index})">
            <span>النتيجة: </span><span id="result_${index}">0</span>
          </div>
        `;
      }
      
      kpiDiv.innerHTML = `
        <h3>${kpi.name}</h3>
        <p>${kpi.definition}</p>
        ${inputFields}
        <input type="hidden" id="kpiName_${index}" value="${kpi.name}">
        <input type="hidden" id="kpiType_${index}" value="${kpi.type}">
      `;
      kpiContainer.appendChild(kpiDiv);
    });
    
    submitBtn.style.display = 'block';
    statusDiv.textContent = '';
  }

  // دالة لحساب النتيجة
  window.calculateResult = function(index) {
    const type = document.getElementById(`kpiType_${index}`).value;
    let result = 0;
    if (type === 'معادلة') {
      const numerator = parseFloat(document.getElementById(`numerator_${index}`).value) || 0;
      const denominator = parseFloat(document.getElementById(`denominator_${index}`).value) || 0;
      result = denominator !== 0 ? (numerator / denominator) * 100 : 0;
    } else {
      result = parseFloat(document.getElementById(`value_${index}`).value) || 0;
    }
    document.getElementById(`result_${index}`).textContent = result.toFixed(2);
  };

  // عند الضغط على زر "حفظ وإرسال"
  submitBtn.addEventListener('click', () => {
    if (!confirm('هل أنت متأكد من حفظ البيانات وإرسالها؟')) {
      return;
    }

    const kpiData = [];
    const kpiItems = kpiContainer.querySelectorAll('.kpi-item');
    kpiItems.forEach((item, index) => {
      const type = document.getElementById(`kpiType_${index}`).value;
      const name = document.getElementById(`kpiName_${index}`).value;
      const result = parseFloat(document.getElementById(`result_${index}`).textContent) || 0;
      
      let dataEntry = { name, result };
      if (type === 'معادلة') {
        dataEntry.numerator = parseFloat(document.getElementById(`numerator_${index}`).value) || 0;
        dataEntry.denominator = parseFloat(document.getElementById(`denominator_${index}`).value) || 0;
      } else {
        dataEntry.value = parseFloat(document.getElementById(`value_${index}`).value) || 0;
      }
      kpiData.push(dataEntry);
    });

    const payload = {
      department: departmentSelect.value,
      year: yearSelect.value,
      month: monthSelect.value,
      kpiData: kpiData
    };

    statusDiv.textContent = 'جاري الحفظ...';
    submitBtn.disabled = true;

    google.script.run
      .withSuccessHandler(response => {
        statusDiv.textContent = response.message;
        statusDiv.style.color = response.success ? 'green' : 'red';
        submitBtn.disabled = false;
        if (response.success) {
          kpiContainer.innerHTML = '';
          submitBtn.style.display = 'none';
        }
      })
      .withFailureHandler(error => {
        statusDiv.textContent = 'حدث خطأ غير متوقع: ' + error.message;
        statusDiv.style.color = 'red';
        submitBtn.disabled = false;
      })
      .submitData(payload);
  });
});