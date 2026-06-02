// ============ إدارة اللغة والوضع الداكن ============
const translations = {
  ar: {
    title: 'محول PDF إلى Word',
    subtitle: 'تحويل متقدم مع التنسيقات، الصور، والجداول',
    dropText: 'اسحب وأفلت ملف PDF هنا',
    dropOr: 'أو',
    browse: 'اختر ملفاً',
    convertBtn: 'تحويل إلى Word',
    processing: 'جارٍ معالجة الصفحات...',
    generating: 'إنشاء مستند Word...',
    success: 'تم التحويل بنجاح!',
    error: 'حدث خطأ. تأكد من أن الملف PDF صالح.',
    download: 'تحميل ملف Word',
    noFile: 'الرجاء اختيار ملف PDF أولاً'
  },
  en: {
    title: 'PDF to Word Converter',
    subtitle: 'Advanced conversion with formatting, images & tables',
    dropText: 'Drag & drop your PDF here',
    dropOr: 'or',
    browse: 'Browse Files',
    convertBtn: 'Convert to Word',
    processing: 'Processing pages...',
    generating: 'Generating Word document...',
    success: 'Conversion successful!',
    error: 'An error occurred. Ensure the file is a valid PDF.',
    download: 'Download Word File',
    noFile: 'Please select a PDF file first'
  }
};

let currentLang = localStorage.getItem('lang') || 'ar';
let darkMode = localStorage.getItem('darkMode') === 'true';

const applyLanguage = () => {
  document.documentElement.lang = currentLang;
  document.documentElement.dir = currentLang === 'ar' ? 'rtl' : 'ltr';
  document.querySelectorAll('[data-lang]').forEach(el => {
    const key = el.getAttribute('data-lang');
    if (translations[currentLang][key]) {
      el.textContent = translations[currentLang][key];
    }
  });
  document.getElementById('langToggle').querySelector('.lang-text').textContent =
    currentLang === 'ar' ? 'EN' : 'عربي';
};

const applyDarkMode = () => {
  document.body.classList.toggle('dark', darkMode);
  localStorage.setItem('darkMode', darkMode);
};

document.getElementById('langToggle').addEventListener('click', () => {
  currentLang = currentLang === 'ar' ? 'en' : 'ar';
  localStorage.setItem('lang', currentLang);
  applyLanguage();
});

document.getElementById('darkModeToggle').addEventListener('click', () => {
  darkMode = !darkMode;
  applyDarkMode();
});

applyLanguage();
applyDarkMode();

// ============ إعداد PDF.js ============
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

// ============ عناصر DOM ============
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const fileNameDisplay = document.getElementById('fileName');
const convertBtn = document.getElementById('convertBtn');
const statusMessage = document.getElementById('statusMessage');
const downloadLink = document.getElementById('downloadLink');
const spinner = document.getElementById('spinner');
const progressContainer = document.getElementById('progressContainer');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');

let selectedFile = null;

// ============ رفع الملف ============
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
  dropZone.addEventListener(eventName, (e) => {
    e.preventDefault();
    e.stopPropagation();
  });
});

dropZone.addEventListener('dragover', () => dropZone.classList.add('drag-over'));
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));

dropZone.addEventListener('drop', (e) => {
  dropZone.classList.remove('drag-over');
  const files = e.dataTransfer.files;
  if (files.length > 0 && files[0].type === 'application/pdf') {
    handleFile(files[0]);
  }
});

fileInput.addEventListener('change', (e) => {
  if (fileInput.files.length > 0) {
    handleFile(fileInput.files[0]);
  }
});

function handleFile(file) {
  selectedFile = file;
  fileNameDisplay.textContent = file.name;
  convertBtn.disabled = false;
  resetUI();
}

function resetUI() {
  statusMessage.style.display = 'none';
  downloadLink.style.display = 'none';
  progressContainer.style.display = 'none';
}

// ============ استخراج الصور من صفحة PDF ============
async function extractImagesFromPage(page) {
  const operatorList = await page.getOperatorList();
  const images = [];
  const commonObjs = page.commonObjs;

  for (let i = 0; i < operatorList.fnArray.length; i++) {
    const fn = operatorList.fnArray[i];
    const args = operatorList.argsArray[i];
    if (fn === pdfjsLib.OPS.paintJpegXObject || fn === pdfjsLib.OPS.paintImageXObject) {
      const imgName = args[0];
      try {
        const imgData = await new Promise((resolve, reject) => {
          commonObjs.get(imgName, (data) => {
            if (data) resolve(data);
            else reject('Image not found');
          });
        });
        // تحويل بيانات الصورة إلى Blob URL
        let blob;
        if (imgData.data instanceof Uint8Array) {
          // JPEG or PNG
          const mime = fn === pdfjsLib.OPS.paintJpegXObject ? 'image/jpeg' : 'image/png';
          blob = new Blob([imgData.data], { type: mime });
        } else if (imgData instanceof ImageBitmap || imgData instanceof HTMLCanvasElement) {
          // fallback
          const canvas = document.createElement('canvas');
          canvas.width = imgData.width;
          canvas.height = imgData.height;
          canvas.getContext('2d').drawImage(imgData, 0, 0);
          blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
        } else {
          continue;
        }
        const url = URL.createObjectURL(blob);
        images.push({ url, width: imgData.width || 100, height: imgData.height || 100 });
      } catch (e) {
        console.warn('تعذر استخراج الصورة:', imgName, e);
      }
    }
  }
  return images;
}

// ============ استخراج النصوص مع التنسيق ============
async function extractStyledTextItems(page) {
  const textContent = await page.getTextContent();
  const viewport = page.getViewport({ scale: 1 });
  const items = [];

  for (const item of textContent.items) {
    if (!item.str.trim()) continue;
    const tx = item.transform;
    const fontSize = Math.sqrt(tx[0] * tx[0] + tx[1] * tx[1]); // تقريباً حجم الخط
    const style = item.fontName ? item.fontName.toLowerCase() : '';
    const isBold = style.includes('bold') || (item.fontWeight && item.fontWeight > 500);
    const isItalic = style.includes('italic') || style.includes('oblique');
    const color = item.color || [0,0,0]; // مصفوفة RGB

    // الموضع النسبي في الصفحة (يمكن استخدامه للكشف عن الجداول)
    const x = tx[4];
    const y = viewport.height - tx[5] - fontSize;

    items.push({
      text: item.str,
      x: x,
      y: y,
      fontSize: Math.round(fontSize * 10) / 10,
      bold: isBold,
      italic: isItalic,
      color: color,
      width: item.width || fontSize * item.str.length * 0.5
    });
  }

  return items;
}

// ============ خوارزمية تجميع الجداول البسيطة ============
function detectTables(textItems) {
  if (textItems.length === 0) return { tables: [], standalone: textItems };

  // تجميع العناصر حسب الصف (إحداثي Y متقارب)
  const rowTolerance = 8;
  const rows = [];
  let currentRow = [];
  let lastY = textItems[0].y;

  for (const item of textItems) {
    if (Math.abs(item.y - lastY) <= rowTolerance) {
      currentRow.push(item);
    } else {
      if (currentRow.length > 0) rows.push(currentRow);
      currentRow = [item];
      lastY = item.y;
    }
  }
  if (currentRow.length > 0) rows.push(currentRow);

  // تحديد الصفوف التي تحتوي على أكثر من عمود (x متباعد)
  const tables = [];
  const standalone = [];

  for (const row of rows) {
    // فرز حسب x
    row.sort((a, b) => a.x - b.x);
    // حساب الفجوات
    let columns = 1;
    for (let i = 1; i < row.length; i++) {
      if (row[i].x - (row[i-1].x + row[i-1].width) > 15) {
        columns++;
      }
    }
    if (columns > 1 && row.length >= 2) {
      tables.push({ type: 'row', items: row, cols: columns });
    } else {
      standalone.push(...row);
    }
  }

  return { tables, standalone };
}

// ============ بناء مستند Word ============
async function buildWordDocument(pdfData) {
  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun, AlignmentType } = docx;
  const children = [];

  for (const pageData of pdfData) {
    // الصور أولاً (إن وجدت)
    for (const img of pageData.images) {
      try {
        const response = await fetch(img.url);
        const blob = await response.blob();
        const arrayBuffer = await blob.arrayBuffer();
        children.push(new Paragraph({
          children: [new ImageRun({
            data: arrayBuffer,
            transformation: {
              width: Math.min(img.width, 500),
              height: Math.min(img.height, 500)
            }
          })],
          alignment: AlignmentType.CENTER
        }));
      } catch (e) {
        console.warn('تعذر تضمين الصورة:', e);
      }
    }

    // النصوص والجداول
    const { tables, standalone } = pageData.textStructure;

    // معالجة الجداول
    for (const table of tables) {
      const rows = [];
      // تجميع عناصر الصف الواحد إلى خلايا
      const rowCells = [];
      let currentCell = [];
      let lastX = table.items[0].x;
      for (const item of table.items) {
        if (item.x - lastX > 20 && currentCell.length > 0) {
          rowCells.push(currentCell);
          currentCell = [];
        }
        currentCell.push(item);
        lastX = item.x + item.width;
      }
      if (currentCell.length > 0) rowCells.push(currentCell);

      const tableRow = new TableRow({
        children: rowCells.map(cellItems => {
          const textRuns = cellItems.map(item => new TextRun({
            text: item.text + ' ',
            bold: item.bold,
            italics: item.italic,
            size: Math.round(item.fontSize * 1.5) * 2, // تحويل تقريبي
            color: item.color ? `#${((1 << 24) + (Math.round(item.color[0]*255) << 16) + (Math.round(item.color[1]*255) << 8) + Math.round(item.color[2]*255)).toString(16).slice(1)}` : undefined
          }));
          return new TableCell({
            children: [new Paragraph({ children: textRuns })]
          });
        })
      });
      rows.push(tableRow);
      children.push(new Table({ rows }));
    }

    // النصوص العادية
    for (const item of standalone) {
      children.push(new Paragraph({
        children: [new TextRun({
          text: item.text,
          bold: item.bold,
          italics: item.italic,
          size: Math.round(item.fontSize * 1.5) * 2,
          color: item.color ? `#${((1 << 24) + (Math.round(item.color[0]*255) << 16) + (Math.round(item.color[1]*255) << 8) + Math.round(item.color[2]*255)).toString(16).slice(1)}` : undefined
        })]
      }));
    }
  }

  const doc = new Document({
    sections: [{ children }]
  });

  return await Packer.toBlob(doc);
}

// ============ عملية التحويل الرئيسية ============
convertBtn.addEventListener('click', async () => {
  if (!selectedFile) {
    showStatus(translations[currentLang].noFile, true);
    return;
  }

  convertBtn.disabled = true;
  spinner.style.display = 'inline-block';
  progressContainer.style.display = 'block';
  progressFill.style.width = '0%';
  progressText.textContent = '0%';
  statusMessage.style.display = 'none';
  downloadLink.style.display = 'none';

  try {
    const arrayBuffer = await selectedFile.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    const totalPages = pdf.numPages;
    const pdfData = [];

    for (let i = 1; i <= totalPages; i++) {
      const percent = Math.round((i / totalPages) * 50); // 50% للاستخراج
      progressFill.style.width = `${percent}%`;
      progressText.textContent = `${percent}%`;
      showStatus(translations[currentLang].processing + ` (${i}/${totalPages})`, false);

      const page = await pdf.getPage(i);
      const images = await extractImagesFromPage(page);
      const textItems = await extractStyledTextItems(page);
      const textStructure = detectTables(textItems);

      pdfData.push({
        pageNumber: i,
        images,
        textStructure
      });
    }

    showStatus(translations[currentLang].generating, false);
    progressFill.style.width = '70%';
    progressText.textContent = '70%';

    const blob = await buildWordDocument(pdfData);
    progressFill.style.width = '100%';
    progressText.textContent = '100%';

    const url = URL.createObjectURL(blob);
    downloadLink.href = url;
    downloadLink.download = selectedFile.name.replace(/\.pdf$/i, '') + '.docx';
    downloadLink.style.display = 'block';
    downloadLink.textContent = translations[currentLang].download;
    showStatus(translations[currentLang].success, false);
  } catch (err) {
    console.error(err);
    showStatus(translations[currentLang].error, true);
  } finally {
    convertBtn.disabled = false;
    spinner.style.display = 'none';
    progressContainer.style.display = 'none';
  }
});

function showStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.style.color = isError ? '#e63946' : 'var(--text-secondary)';
  statusMessage.style.display = 'block';
}
