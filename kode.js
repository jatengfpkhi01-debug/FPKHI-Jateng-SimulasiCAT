const SHEET_USERS = 'Users';
const SHEET_QUESTIONS = 'Questions';
const SHEET_SETTINGS = 'Settings';
const SHEET_SCHEDULES = 'Schedules';

function doGet() {
  // Hanya untuk menampilkan halaman jika dibuka dari link bawaan Google
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Simulasi CAT Haji 2026')
    .setFaviconUrl('https://upload.wikimedia.org/wikipedia/commons/thumb/c/c2/Coat_of_arms_of_Indonesia.svg/500px-Coat_of_arms_of_Indonesia.svg.png')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1.0, user-scalable=no');
}

// FUNGSI UTAMA UNTUK VERCEL (REST API POST)
function doPost(e) {
  try {
    // Parsing request JSON dari Vercel / Frontend external
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    let result = {};

    // Routing aksi (API Endpoint Logic)
    if (action === 'loginUser') {
       if (payload.data) {
         result = loginUser(payload.data); // Mode Register
       } else {
         result = loginUser(payload.nr, payload.password); // Mode Login
       }
    } else if (action === 'getExamSetup') {
       result = getExamSetup();
    } else if (action === 'submitExam') {
       result = submitExam(payload.nr, payload.answers, payload.finishTimestamp, payload.isCheat);
    } else if (action === 'getAdminData') {
       result = getAdminData();
    } else if (action === 'importQuestions') {
       result = importQuestions(payload.questionsArray);
    } else if (action === 'saveScheduleData') {
       result = saveScheduleData(payload.payload);
    } else if (action === 'saveEmailSettings') {
       result = saveEmailSettings(payload.subject, payload.template);
    } else {
       result = { success: false, message: 'Action tidak ditemukan.' };
    }

    // Mengembalikan data JSON ke Vercel
    return ContentService.createTextOutput(JSON.stringify(result))
          .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.message }))
          .setMimeType(ContentService.MimeType.JSON);
  }
}

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!ss.getSheetByName(SHEET_USERS)) {
    const sheet = ss.insertSheet(SHEET_USERS);
    sheet.appendRow(['NR', 'Nama', 'NIK', 'Profesi', 'Peminatan', 'Email', 'Password', 'Score', 'Status', 'WaktuSelesai']);
    sheet.getRange('A1:J1').setFontWeight('bold');
  }
  
  if (!ss.getSheetByName(SHEET_QUESTIONS)) {
    const sheet = ss.insertSheet(SHEET_QUESTIONS);
    sheet.appendRow(['ID', 'Soal', 'OpsiA', 'OpsiB', 'OpsiC', 'OpsiD', 'OpsiE', 'JawabanBenar']);
    sheet.getRange('A1:H1').setFontWeight('bold');
    sheet.appendRow(['Q1', 'Apa rukun haji yang pertama?', 'Ihram', 'Wukuf', 'Tawaf', 'Sa\'i', 'Mabit', 'Ihram']);
  }

  if (!ss.getSheetByName(SHEET_SETTINGS)) {
    const sheet = ss.insertSheet(SHEET_SETTINGS);
    sheet.appendRow(['Key', 'Value']);
    sheet.getRange('A1:B1').setFontWeight('bold');
    sheet.appendRow(['ExamDuration', '60']); // dalam menit
    sheet.appendRow(['EmailSubject', 'Pendaftaran Simulasi CAT Haji 2026']);
    sheet.appendRow(['EmailTemplate', 'Halo {{Nama}},\n\nSelamat, pendaftaran Anda berhasil.\nNR: {{NR}}\nPassword: {{Password}}']);
  }

  if (!ss.getSheetByName(SHEET_SCHEDULES)) {
    const sheet = ss.insertSheet(SHEET_SCHEDULES);
    sheet.appendRow(['ID', 'NamaSimulasi', 'WaktuMulai', 'WaktuSelesai']);
    sheet.getRange('A1:D1').setFontWeight('bold');
  }
}

function getSheetDataAsObjects(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const rows = data.slice(1);
  
  return rows.map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      let cellValue = row[index];
      // Force konversi Date Object menjadi String agar aman saat dikirim via JSON
      if (cellValue instanceof Date) {
        cellValue = cellValue.toISOString();
      }
      obj[header] = cellValue;
    });
    return obj;
  });
}

function loginUser(arg1, arg2) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_USERS);
    
    // Jika argument ke-2 ada isinya, berarti ini proses LOGIN (NR, Password)
    if (arg2 !== undefined) {
      const nr = arg1;
      const password = arg2;
      
      if (nr === 'ADMIN' && password === 'admin123') {
        return { success: true, role: 'admin', user: { name: 'Administrator', nr: 'ADMIN' } };
      }
      
      const users = getSheetDataAsObjects(SHEET_USERS);
      const user = users.find(u => u.NR === nr && u.Password === password);
      
      if (user) {
        return { 
          success: true, 
          role: 'user', 
          user: { nr: user.NR, name: user.Nama, score: user.Score, status: user.Status, finishTime: user.WaktuSelesai } 
        };
      } else {
        return { success: false, message: 'Nomor Register atau Password salah.' };
      }
    } 
    // Jika argument ke-2 kosong, berarti arg1 adalah Object Data REGISTER
    else {
      const data = arg1;
      const users = getSheetDataAsObjects(SHEET_USERS);
      
      if (users.some(u => u.NIK === data.nik || u.Email === data.email)) {
        return { success: false, message: 'NIK atau Email sudah terdaftar!' };
      }
      
      const newNR = 'NR' + Math.floor(10000000000000 + Math.random() * 90000000000000);
      sheet.appendRow([newNR, data.nama, data.nik, data.profesi, data.peminatan, data.email, data.password, '', 'Pending', '']);
      
      // Kirim Email (Abaikan jika gagal agar pendaftaran tidak error)
      try {
        const settings = getSheetDataAsObjects(SHEET_SETTINGS);
        const subjObj = settings.find(s => s.Key === 'EmailSubject');
        const tplObj = settings.find(s => s.Key === 'EmailTemplate');
        
        let subject = subjObj ? subjObj.Value : "Pendaftaran CAT";
        let body = tplObj ? tplObj.Value : "NR: {{NR}}, Pass: {{Password}}";
        
        body = body.replace('{{Nama}}', data.nama).replace('{{NR}}', newNR).replace('{{Password}}', data.password);
        MailApp.sendEmail(data.email, subject, body);
      } catch (e) {}
      
      return { success: true, message: `Pendaftaran berhasil. Cek email Anda untuk Nomor Register (NR).\nNR Anda: ${newNR}` };
    }
  } catch (e) {
    return { success: false, message: 'Terjadi kesalahan sistem: ' + e.message };
  }
}

function getExamSetup() {
  const questions = getSheetDataAsObjects(SHEET_QUESTIONS);
  const settings = getSheetDataAsObjects(SHEET_SETTINGS);
  const durationRow = settings.find(s => s.Key === 'ExamDuration');
  const durationMinutes = durationRow ? parseInt(durationRow.Value) : 60;

  // Hapus JawabanBenar sebelum dikirim ke Frontend untuk mencegah kecurangan (Cheat)
  const sanitizedQuestions = questions.map(q => ({
    id: q.ID,
    text: q.Soal,
    options: { A: q.OpsiA, B: q.OpsiB, C: q.OpsiC, D: q.OpsiD, E: q.OpsiE }
  }));

  return { questions: sanitizedQuestions, durationMinutes: durationMinutes };
}

function submitExam(nr, answers, finishTimestamp, isCheat) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetUsers = ss.getSheetByName(SHEET_USERS);
    const questions = getSheetDataAsObjects(SHEET_QUESTIONS);
    
    let correct = 0;
    let finalScore = 0;

    // Koreksi jawaban dilakukan di backend (aman)
    if (!isCheat) {
      questions.forEach(q => {
        const userAnswer = answers[q.ID] ? answers[q.ID].toString().trim().toLowerCase() : "";
        const realAnswer = q.JawabanBenar ? q.JawabanBenar.toString().trim().toLowerCase() : "";
        if (userAnswer === realAnswer && userAnswer !== "") {
          correct++;
        }
      });
      finalScore = ((correct / questions.length) * 100).toFixed(2);
    }
    
    // Update Score & Waktu Selesai
    const data = sheetUsers.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === nr) {
        sheetUsers.getRange(i + 1, 8).setValue(finalScore); 
        sheetUsers.getRange(i + 1, 9).setValue('Done');     
        sheetUsers.getRange(i + 1, 10).setValue(finishTimestamp.toString()); 
        break;
      }
    }
    
    return { success: true, score: finalScore, isCheat: isCheat };
  } catch (e) {
    throw new Error('Gagal memproses jawaban');
  }
}

function getAdminData() {
  const users = getSheetDataAsObjects(SHEET_USERS);
  const schedules = getSheetDataAsObjects(SHEET_SCHEDULES);
  return { users: users, schedules: schedules };
}

function importQuestions(questionsArray) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_QUESTIONS);
    
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 8).clearContent();
    }
    
    if (questionsArray && questionsArray.length > 0) {
      sheet.getRange(2, 1, questionsArray.length, 8).setValues(questionsArray);
    }
    return { success: true, message: 'Bank soal berhasil diupdate!' };
  } catch (e) {
    return { success: false, message: 'Gagal import: ' + e.message };
  }
}

function saveScheduleData(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_SCHEDULES);
    
    if (payload.id) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === payload.id) {
          sheet.getRange(i + 1, 2).setValue(payload.nama);
          sheet.getRange(i + 1, 3).setValue(payload.mulai);
          sheet.getRange(i + 1, 4).setValue(payload.selesai);
          break;
        }
      }
    } else {
      const newId = 'SCH' + Date.now();
      sheet.appendRow([newId, payload.nama, payload.mulai, payload.selesai]);
    }
    return { success: true, message: 'Jadwal Simulasi berhasil disimpan.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function saveEmailSettings(subject, template) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_SETTINGS);
    const data = sheet.getDataRange().getValues();
    
    let subjFound = false; let tplFound = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'EmailSubject') { sheet.getRange(i + 1, 2).setValue(subject); subjFound = true; }
      if (data[i][0] === 'EmailTemplate') { sheet.getRange(i + 1, 2).setValue(template); tplFound = true; }
    }
    if(!subjFound) sheet.appendRow(['EmailSubject', subject]);
    if(!tplFound) sheet.appendRow(['EmailTemplate', template]);
    
    return { success: true, message: 'Pengaturan Email Tersimpan' };
  } catch (e) { return { success: false, message: e.message }; }
}
