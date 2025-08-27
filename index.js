#!/usr/bin/env node

// ===== AUTO INSTALL MODULE =====
const { execSync } = require('child_process');

function ensureModule(moduleName) {
  try {
    require.resolve(moduleName);
  } catch (e) {
    console.log(`Module "${moduleName}" not found. Installing...`);
    execSync(`npm install ${moduleName}`, { stdio: 'inherit' });
    console.log(`Module "${moduleName}" installed successfully.`);
  }
}

// List semua module eksternal yang dipakai
const modules = [
  'dotenv',
  'fs',
  'path',
  'axios',
  'qs',
  'nodemailer',
  'html-to-text',
  'https-proxy-agent',
  'socks-proxy-agent'
];

// Validasi & install semua module
modules.forEach(ensureModule);

// ===== LOAD MODULE =====
require('dotenv').config();
const fs = require('fs');
const path = require('path');
const axios = require('axios');
const qs = require('qs');
const nodemailer = require('nodemailer');
const { htmlToText } = require('html-to-text');

// ===== LOCAL MODULES =====
const { randomHeaders } = require('./assets/header.js');
const { randomString } = require('./assets/random.js');
const { formatDate } = require('./assets/date.js');
const { randomCity } = require('./assets/city.js');
const { askYesNo, askNumber, closeInput } = require('./assets/input.js');
const { generateRandomName, generateRandomUPN } = require('./assets/randomUser.js');
const { generateOrderId } = require('./assets/order.js');
const { loadProxies, getNextProxy, handleProxyError } = require('./assets/proxy.js');
// ===== ENV VALIDATION =====
const requiredEnv = [
  'TENANT_ID','CLIENT_ID','CLIENT_SECRET','DOMAIN','LICENSE_SKU_ID',
  'RECIPIENTS_DIR','SUBJECTS_DIR','LETTER_DIR','FROMNAME_DIR','LINKS_DIR','PROXY_DIR',
  'SMTP_HOST','SMTP_PORT','HOSTNAME','LOG_FILE',
  'MAIL_ENCODING','MAIL_CHARSET',
  'CUSTOM_REPLY_TO','REPLY_TO',
  'CUSTOM_RETURN_PATH','RETURN_PATH',
  'USE_PROXY'
];

// Cek semua env wajib
for (const key of requiredEnv) {
  if (!process.env[key] || process.env[key].trim() === '') {
    console.error(`❌ Missing required env var: ${key}`);
    process.exit(1);
  }
}

// ===== ENV VARS =====
const {
  TENANT_ID,
  CLIENT_ID,
  CLIENT_SECRET,
  DOMAIN,
  LICENSE_SKU_ID,
  RECIPIENTS_DIR,
  SUBJECTS_DIR,
  LETTER_DIR,
  FROMNAME_DIR,
  LINKS_DIR,
  PROXY_DIR,
  SMTP_HOST,
  SMTP_PORT,
  HOSTNAME,
  LOG_FILE,
  DELAY_MS,
  EMAIL_DELAY_MS,
  MAILBOX_RETRY,
  MAILBOX_TIMEOUT_MS,
  RANDOM_PARAMETER,
  EMAIL_PRIORITY,
  MAIL_CHARSET,
  MAIL_ENCODING,
  CLEAR_DUPLICATE,
  CUSTOM_REPLY_TO,
  REPLY_TO,
  CUSTOM_RETURN_PATH,
  RETURN_PATH,
  USE_PROXY
} = process.env;

// ===== CONSTANTS =====
const FAILED_FILE = './failed_emails.txt';
const LOG_PATH = path.resolve(LOG_FILE);
const EMAIL_DELAY = Number(EMAIL_DELAY_MS) || 1000;
const DELAY = Number(DELAY_MS) || 5000;
const MAILBOX_TIMEOUT = Number(MAILBOX_TIMEOUT_MS) || 120000;
const RETRY_SEND = 5;

// ===== COLORS =====
const COLORS = { RESET:"\x1b[0m", RED:"\x1b[31m", GREEN:"\x1b[32m", YELLOW:"\x1b[33m", CYAN:"\x1b[36m", WHITE:"\x1b[37m", MAGENTA:"\x1b[35m" };

// ===== LOGGING =====
function LOG(line, type='INFO') {
  const now = new Date().toISOString().replace('T',' ').split('.')[0];
  fs.appendFileSync(LOG_PATH, `${now} | ${type} | ${line}\n`);

  let color = COLORS.WHITE;
  if (type==='ERROR') color = COLORS.RED;
  else if (type==='OK') color = COLORS.GREEN;
  else if (type==='WARN') color = COLORS.YELLOW;
  else if (type==='INFO') color = COLORS.CYAN;
  else if (type==='PROXY') color = COLORS.MAGENTA; // Magenta

  console.log(color + `[${type}] ${line}` + COLORS.RESET);
}

// ===== HELPERS =====
function readTxt(dir) {
  if (!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir).flatMap(f =>
    fs.readFileSync(path.join(dir,f),'utf8')
      .split(/\r?\n/)
      .map(s => s.trim())
      .filter(Boolean)
  );
}

function readHtml(dir) {
  if (!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir)
    .filter(f => f.endsWith('.html') || f.endsWith('.htm'))
    .map(f => ({ name:f, content:fs.readFileSync(path.join(dir,f),'utf8') }));
}

function chunkArray(arr, size) {
  const out = [];
  for (let i=0; i<arr.length; i+=size) out.push(arr.slice(i,i+size));
  return out;
}

// ===== TOKEN MGMT =====
let cachedToken = null;
let tokenExpiry = 0;

async function getValidToken() {
  const now = Date.now();
  if (!cachedToken || now >= tokenExpiry) {
    LOG('Fetch new token...');
    const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
    const data = qs.stringify({
      client_id: CLIENT_ID,
      client_secret: CLIENT_SECRET,
      scope: 'https://graph.microsoft.com/.default',
      grant_type: 'client_credentials'
    });
    const res = await axios.post(url, data, { headers:{'Content-Type':'application/x-www-form-urlencoded'} });
    cachedToken = res.data.access_token;
    tokenExpiry = now + (res.data.expires_in - 60)*1000;
    LOG('Token OK','OK');
  }
  return cachedToken;
}

async function graphRequest(method, pathUrl, data=null) {
  const token = await getValidToken();
  const url = `https://graph.microsoft.com/v1.0${pathUrl}`;
  return (await axios({ method, url, headers:{ Authorization:`Bearer ${token}` }, data })).data;
}

// ===== USER LIFECYCLE =====
async function createTempUser() {
  const { firstName,lastName,displayName } = generateRandomName();
  const upn = generateRandomUPN(DOMAIN);
  const uniq = upn.split('@')[0];
  const payload = {
    accountEnabled: true,
    displayName, givenName:firstName, surname:lastName,
    mailNickname: uniq, userPrincipalName: upn,
    passwordProfile: { forceChangePasswordNextSignIn:false, password:'Fawwaz19@' }
  };
  const user = await graphRequest('post','/users',payload);
  return { id:user.id, upn:user.userPrincipalName, password:'Fawwaz19@' };
}

async function updateUsageLocation(userId) {
  await graphRequest('patch', `/users/${userId}`, { usageLocation:'US' });
}

async function assignLicense(userId, skuId) {
  if (!skuId) return;
  try {
    await graphRequest('post', `/users/${userId}/assignLicense`, { addLicenses:[{skuId,disabledPlans:[]}], removeLicenses:[] });
    LOG(`License assigned: ${skuId}`,'OK');
  } catch(e) {
    LOG(`Assign license failed: ${e.message}`,'WARN');
  }
}

async function waitForMailbox(userId) {
  let attempt=0;
  let delay=5000;
  const start = Date.now();
  while (Date.now()-start < MAILBOX_TIMEOUT && attempt < (MAILBOX_RETRY||10)) {
    try {
      await graphRequest('get', `/users/${userId}/mailFolders/inbox`);
      LOG('Mailbox ready','OK');
      return true;
    } catch(e) {
      attempt++;
      LOG(`Mailbox not ready, attempt ${attempt}`,'WARN');
      await new Promise(r=>setTimeout(r, delay));
      delay = Math.min(delay*2,30000); // exponential backoff
    }
  }
  LOG('Mailbox not ready after retries','WARN');
  return false;
}

async function removeLicense(userId, skuId, user) {
  try { 
    await graphRequest('post', `/users/${userId}/assignLicense`, { addLicenses:[], removeLicenses:[skuId] });
    LOG(`License removed: ${skuId} for ${user.upn}`, 'WARN'); 
  } catch(e){ 
    LOG(`Remove license failed: ${e.message}`, 'WARN'); 
  }
}

async function deleteUser(userId, user) {
  try { 
    await graphRequest('delete', `/users/${userId}`);
    LOG(`User deleted: ${user.upn}`, 'WARN'); 
  } catch(e){ 
    LOG(`Delete user failed: ${e.message}`, 'ERROR'); 
  }
}
LOG(`============================================`, 'INFO');
// ===== MIME ENCODING =====
function encodeMimeWord(str) {
  const base64 = Buffer.from(str, 'utf8').toString('base64');
  return `=?UTF-8?B?${base64}?=`;
}

// ===== SEND MAIL =====
async function sendMailSMTP(user, fromName, recipients, subject, htmlBody, link, groupIndex, totalGroups, opts) {
  const { useAttachments, useTo } = opts;
  let smtpProxyAgent=null;
  if (USE_PROXY==='true') {
  ({ smtpProxyAgent } = getNextProxy((msg)=>LOG(msg,'PROXY')));
}

  // Encode FromName & Subject
  const encodedFromName = encodeMimeWord(fromName);
  const encodedSubject = encodeMimeWord(subject);
  const fromEnc = `"${encodedFromName}" <${user.upn}>`;
  
  // ===== attachments hanya yang dipakai di htmlBody =====
  const attachments = [];
  if (useAttachments) {
    const logoDir = './assets/logo';
    const cidsInHtml = [...htmlBody.matchAll(/cid:([a-zA-Z0-9_\-]+)@cid/g)].map(m => m[1]);
    fs.readdirSync(logoDir)
      .filter(f => f.endsWith('.png'))
      .forEach(file => {
        const name = path.parse(file).name;
        if (cidsInHtml.includes(name)) {
          attachments.push({
            filename: file,
            path: path.join(logoDir, file),
            cid: `${name}@cid`
          });
        }
      });
  }

const transporterOptions = {
  host: SMTP_HOST,
  port: Number(SMTP_PORT),
  secure: false,
  requireTLS: true,
  auth: { user: user.upn, pass: user.password },
  name: `mail-pl1-f${Math.floor(Math.random()*900+100)}.${HOSTNAME}`,
  tls: { minVersion: 'TLSv1.2', rejectUnauthorized: false },
  debug: true,
  agent: smtpProxyAgent 
};

  if (smtpProxyAgent) transporterOptions.agent = smtpProxyAgent;

  const transporter = nodemailer.createTransport(transporterOptions);
  const textBody = htmlToText(htmlBody,{wordwrap:130});

  let toRecipient='', bccRecipients=recipients;
  if (useTo && recipients.length>1) {
    toRecipient = recipients[Math.floor(Math.random()*recipients.length)];
    bccRecipients = recipients.filter(r=>r!==toRecipient);
  }

// === BUAT RETURN PATH ===
const returnPath = RETURN_PATH;

// === BUAT RANDOM HEADER ===
const headers = randomHeaders(DOMAIN);

// === BUAT PRIORITY ===
function parseEmailPriority(value){
  switch(String(value)){
    case '1': return 'high';
    case '3': return 'normal';
    case '5': return 'low';
    case 'high':
    case 'normal':
    case 'low':
      return value;
    default: return 'normal';
  }
}
const priorityValue = parseEmailPriority(EMAIL_PRIORITY);

// Hanya buat alternatives jika textBody berbeda dari htmlBody
const alternatives = textBody !== htmlBody ? [
  { content: textBody, charset: MAIL_CHARSET || 'utf-8' }
] : undefined;

// === REPLY TO Settings ===
let replyToAddress = user.upn; // default pakai email pengirim sementara
if (CUSTOM_REPLY_TO === 'true') {
  replyToAddress = REPLY_TO; // pakai email dari env
}

  for (let attempt=1; attempt<=RETRY_SEND; attempt++) {
    try {
await transporter.sendMail({
  from: fromEnc,
  to: toRecipient || '',
  bcc: bccRecipients,
  replyTo:replyToAddress,
  subject: encodedSubject,
  text: textBody,
  html: htmlBody,
  attachments,
  headers,
  priority:priorityValue,
  encoding:MAIL_ENCODING,
  //alternatives,
  envelope: CUSTOM_RETURN_PATH==='true' 
    ? { from:returnPath, to:[...bccRecipients,toRecipient].filter(Boolean) }
    : undefined
});
// ===== UTILITY PAD =====
const pad = (str, len=25) => String(str).padEnd(len, ' ');

// ===== LOG EMAIL STATUS =====
LOG('==================================================','INFO');
console.log(`${COLORS.CYAN}${pad('STATUS EMAIL :')}${COLORS.GREEN}Terkirim${COLORS.RESET}`);
console.log(`${COLORS.CYAN}${pad('FROM :')}${COLORS.WHITE}${pad(fromName + ' <' + user.upn + '>', 40)}${COLORS.RESET}`);
if(toRecipient) console.log(`${COLORS.CYAN}${pad('TO :')}${COLORS.WHITE}${pad(toRecipient, 35)}${COLORS.RESET}`);
console.log(`${COLORS.CYAN}${pad('REPLY-TO :')}${COLORS.WHITE}${pad(replyToAddress, 35)}${COLORS.RESET}`);
console.log(`${COLORS.CYAN}${pad('SUBJECT :')}${COLORS.WHITE}${pad(subject, 40)}${COLORS.RESET}`);
console.log(`${COLORS.CYAN}${pad('LINK :')}${COLORS.WHITE}${pad(link, 50)}${COLORS.RESET}`);
console.log(`${COLORS.CYAN}${pad('BCC :')}${COLORS.WHITE}${pad(bccRecipients.length + ' Penerima', 20)}${COLORS.RESET}`);
console.log(`${COLORS.CYAN}${pad('TOTAL BCC :')}${COLORS.WHITE}${pad(groupIndex + ' / ' + totalGroups, 10)}${COLORS.RESET}`);
LOG('==================================================','INFO');
      return;
    } catch(e) {
      LOG(`Send attempt ${attempt} failed: ${e.message}`,'ERROR');
      if (USE_PROXY==='true') handleProxyError(e,(msg)=>LOG(msg,'PROXY'));
      if (attempt<RETRY_SEND) await new Promise(r=>setTimeout(r,DELAY));
    }
  }
  fs.appendFileSync(FAILED_FILE, recipients.join('\n')+'\n');
}

// ===== MAIN FLOW =====
(async()=>{
  let tempUsers=[];
  try {
    console.clear();
    LOG('=== START SMTP FLOW ===');
// kalau pakai proxy, load dulu & cek hasilnya
if (USE_PROXY === 'true') {
  await loadProxies(PROXY_DIR, true, (msg)=>LOG(msg,'PROXY'));
} else {
  LOG('Proxy OFF.','PROXY');
}


// ===== PILIH FILE RECIPIENTS =====
const recipientFiles = fs.readdirSync(RECIPIENTS_DIR) .filter(f => f.endsWith('.txt'));
console.log('\nSelect Email list:');
recipientFiles.forEach((f,i) => console.log(`${i+1}. ${f}`));
const fileIndex = await askNumber(`Pilih file (1-${recipientFiles.length}): `, 1);
const selectedFile = recipientFiles[fileIndex-1];
const recipients = [...new Set(
  fs.readFileSync(path.join(RECIPIENTS_DIR, selectedFile), 'utf8')
    .split(/\r?\n/)
    .map(s => s.trim())
    .filter(r => r && /\S+@\S+\.\S+/.test(r))
)];
LOG(`Selected recipients file: ${selectedFile} (total: ${recipients.length})`,'INFO');
    const MSG_SIZE_INPUT = await askNumber('Jumlah email per BCC group (default 30): ',30);
    const USE_TO = await askYesNo('Aktifkan TO (Y/N): ');
    const USE_ATTACHMENTS = await askYesNo('Aktifkan attachments (Y/N): ');
    const EMAIL_DELAY_INPUT = await askNumber(`Delay kirim email (default ${EMAIL_DELAY/1000}s): `, EMAIL_DELAY/1000);
    closeInput();

    const subjects = readTxt(SUBJECTS_DIR);
    const letters = readHtml(LETTER_DIR);
    const fromNames = readTxt(FROMNAME_DIR);
    const links = readTxt(LINKS_DIR);
	if (!links.length) throw new Error('No links found');
	// ===== LINK BERURUTAN =====
	let linkIndex = 0;

    if (!recipients.length) throw new Error('No recipients found');
    const chunks = chunkArray(recipients, MSG_SIZE_INPUT);

    for (let i=0; i<chunks.length; i++) {
      const chunk = chunks[i];
      const user = await createTempUser();
      tempUsers.push(user);
      LOG(`User created: ${user.upn}`,'INFO');
      await updateUsageLocation(user.id);
      if (LICENSE_SKU_ID) await assignLicense(user.id, LICENSE_SKU_ID);
      await waitForMailbox(user.id);

      await new Promise(r=>setTimeout(r,EMAIL_DELAY_INPUT*1000));

      let subject = subjects[Math.floor(Math.random()*subjects.length)];
      let fromName = fromNames[Math.floor(Math.random()*fromNames.length)];
 // Ambil link asli
  let link = links[linkIndex];
  linkIndex = (linkIndex + 1) % links.length; // maju ke link berikutnya secara berurutan
  if (RANDOM_PARAMETER==='true') {
    const sep = link.includes('?') ? '&':'?';
    link += `${sep}tid=${randomString(8)}`;
  }

      let letter = letters[Math.floor(Math.random()*letters.length)].content;
      const orderId = generateOrderId();
      letter = letter.replace(/##date##/g,formatDate())
                     .replace(/##link##/g,link)
                     .replace(/##city##/g,randomCity())
                     .replace(/##orderId##/g,orderId);
      subject = subject.replace(/##orderId##/g,orderId)
					   .replace(/##date##/g,formatDate());

      await sendMailSMTP(user, fromName, chunk, subject, letter, link, i+1, chunks.length, { useAttachments:USE_ATTACHMENTS, useTo:USE_TO });

      if (LICENSE_SKU_ID)
		await removeLicense(user.id, LICENSE_SKU_ID, user);
		await deleteUser(user.id, user);
		tempUsers = tempUsers.filter(u=>u.id!==user.id);

      await new Promise(r=>setTimeout(r,DELAY));
    }

    LOG('=== FLOW DONE ===');
  } catch(e) {
    LOG(`FATAL: ${e.message}`,'ERROR');
  } finally {
    // Cleanup leftover users if crash
	for (const u of tempUsers) {
	try { await deleteUser(u.id, u); LOG(`Cleanup deleted user ${u.upn}`,'WARN'); } catch {}
	}
  }
})();