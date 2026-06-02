// ============ إدارة اللغة والوضع الداكن ============
const translations = {
  ar: {
    title: 'محول PDF إلى Word',
    subtitle: 'حوّل ملفات PDF إلى مستندات Word بسهولة مع دعم كامل للغة العربية',
    dropText: 'اسحب وأفلت ملف PDF هنا',
    dropOr: 'أو',
    browse: 'اختر ملفاً',
    convertBtn: 'تحويل إلى Word',
    processing: 'جارٍ استخراج النصوص...',
    generating: 'يتم إنشاء ملف Word...',
    success: 'تم التحويل بنجاح!',
    error: 'حدث خطأ أثناء التحويل. تأكد من أن الملف PDF صالح.',
    download: 'تحميل ملف Word',
    noFile: 'الرجاء اختيار ملف PDF أولاً'
  },
  en: {
    title: 'PDF to Word Converter',
    subtitle: 'Easily convert PDFs to Word documents with full Arabic support',
    dropText: 'Drag & drop your PDF here',
    dropOr: 'or',
    browse: 'Browse Files',
    convertBtn: 'Convert to Word',
    processing: 'Extracting text...',
    generating: 'Generating Word document...',
    success: 'Conversion successful!',
    error: 'An error occurred. Make sure the file is a valid PDF.',
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

// ============ منطق رفع الملف ============
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const fileNameDisplay = document.getElementById('fileName');
const convertBtn = document.getElementById('convertBtn');
const statusMessage = document.getElementById('statusMessage');
const downloadLink = document.getElementById('downloadLink');
const spinner = document.getElementById('spinner');

let selectedFile = null;

// منع السلوك الافتراضي للسحب
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
  } else {
    showStatus(translations[currentLang].error, true);
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
  statusMessage.style.display = 'none';
  downloadLink.style.display = 'none';
}

// ============ التحويل ============
convertBtn.addEventListener('click', async () => {
  if (!selectedFile) {
    showStatus(translations[currentLang].noFile, true);
    return;
  }

  convertBtn.disabled = true;
  spinner.style.display = 'inline-block';
  convertBtn.querySelector('.btn-text').textContent = '...';

  try {
    showStatus(translations[currentLang].processing, false);

    const arrayBuffer = await selectedFile.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    let fullText = '';

    // استخراج النص من كل صفحة
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      const pageText = content.items.map(item => item.str).join(' ');
      fullText += pageText + '\n\n';
    }

    showStatus(translations[currentLang].generating, false);

    // إنشاء مستند Word باستخدام docx
    const { Document, Packer, Paragraph, TextRun } = docx;
    const paragraphs = fullText.split('\n').filter(para => para.trim() !== '').map(text => {
      return new Paragraph({
        children: [new TextRun({ text: text.trim(), lang: 'ar-SA' })],
        bidirectional: true,
        spacing: { after: 200 }
      });
    });

    const doc = new Document({
      sections: [{
        properties: {},
        children: paragraphs
      }]
    });

    const blob = await Packer.toBlob(doc);
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
    convertBtn.querySelector('.btn-text').textContent = translations[currentLang].convertBtn;
  }
});

function showStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.style.color = isError ? '#e63946' : 'var(--text-secondary)';
  statusMessage.style.display = 'block';
                            }
