/** Kolli Builders Support API + private staff dashboard.
 * Deploy as Web app: Execute as Me. Access: Anyone (public customer API).
 * Staff functions require the password configured using setAdminPassword().
 */
const SPREADSHEET_ID = "15ZxB5aUkpr1xZHaXgdJvvAfNjarTY466AIUE4xPb1wA";
const TICKETS_SHEET = "Tickets";
const MESSAGES_SHEET = "TicketMessages";
const HEADERS = ["id","name","contact","category","priority","subject","details","status","resolution","createdAt","updatedAt"];
const MESSAGE_HEADERS = ["ticketId","sender","message","createdAt"];

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.page === "admin") return HtmlService.createHtmlOutputFromFile("admin").setTitle("Kolli Builders | Staff Support");
  if (p.action === "lookup") {
    const result = lookupCustomerTicket_(String(p.ticketId || "").trim(), String(p.contact || "").trim());
    const callback = String(p.callback || "");
    if (callback && /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(callback)) {
      return ContentService.createTextOutput(callback + "(" + JSON.stringify(result) + ");").setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return json_(result);
  }
  if (p.action === "health") return json_({ok:true, service:"Kolli Builders Support API"});
  return json_({ok:true, message:"Kolli Builders Support API is running."});
}

function doPost(e) {
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || "{}");

    if (payload.action === "createTicket" && payload.ticket) {
      const t = payload.ticket;
      if (!t.name || !t.contact || !t.subject || !t.details) return json_({ok:false,error:"Required fields are missing"});
      const id = String(t.id || ("TC-" + new Date().getFullYear() + "-" + Utilities.getUuid().slice(0,6).toUpperCase())).trim();
      const now = new Date().toISOString();
      const ticket = {id:id,name:String(t.name).slice(0,150),contact:String(t.contact).trim().slice(0,200),category:String(t.category||"General question").slice(0,100),priority:String(t.priority||"Normal"),subject:String(t.subject).slice(0,250),details:String(t.details).slice(0,10000),status:"Open",resolution:"",createdAt:now,updatedAt:now};
      const sh = getSheet_(TICKETS_SHEET, HEADERS);
      if (findTicketRow_(sh, id)) return json_({ok:false,error:"Ticket ID already exists"});
      sh.appendRow(HEADERS.map(k => safeCell_(ticket[k] || "")));
      addMessage_(id,"Customer","Ticket created: " + ticket.details,now);
      return json_({ok:true,ticketId:id});
    }

    if (payload.action === "customerReply") {
      return customerReply_(payload.ticketId, payload.contact, payload.message);
    }

    return json_({ok:false,error:"Invalid request"});
  } catch (err) {
    return json_({ok:false,error:String(err)});
  }
}

function setupSupportSheets() {
  getSheet_(TICKETS_SHEET, HEADERS);
  getSheet_(MESSAGES_SHEET, MESSAGE_HEADERS);
}

function setAdminPassword() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Set staff dashboard password", "Enter a strong password (at least 12 characters).", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const password = response.getResponseText();
  if (!password || password.length < 12) { ui.alert("Use a password with at least 12 characters and run again."); return; }
  PropertiesService.getScriptProperties().setProperty("ADMIN_PASSWORD_HASH", hash_(password));
  ui.alert("Staff password saved. Open the web app URL with ?page=admin");
}

function adminLogin(password) {
  const stored = PropertiesService.getScriptProperties().getProperty("ADMIN_PASSWORD_HASH");
  if (!stored) throw new Error("Staff password is not configured. Run setAdminPassword() first.");
  if (hash_(String(password || "")) !== stored) { Utilities.sleep(700); throw new Error("Incorrect password."); }
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put("staff:" + token, "ok", 21600);
  return {ok:true, token:token, expiresInSeconds:21600};
}
function adminListTickets(token) { requireAdmin_(token); return {ok:true,tickets:readTickets_()}; }
function adminListMessages(token, ticketId) { requireAdmin_(token); return {ok:true,messages:readMessages_(String(ticketId||"").trim())}; }
function adminReplyTicket(token, ticketId, message, status, resolution) {
  requireAdmin_(token);
  ticketId = String(ticketId||"").trim(); message = String(message||"").trim();
  if (!ticketId || !message) throw new Error("Ticket ID and reply are required.");
  const sh = getSheet_(TICKETS_SHEET, HEADERS); const row = findTicketRow_(sh,ticketId);
  if (!row) throw new Error("Ticket not found.");
  if (String(sh.getRange(row,8).getDisplayValue()).toLowerCase() === "closed") throw new Error("This ticket is closed.");
  const now = new Date().toISOString();
  addMessage_(ticketId,"Staff",message,now);
  sh.getRange(row,8).setValue(String(status || "In Progress"));
  sh.getRange(row,9).setValue(String(resolution || ""));
  sh.getRange(row,11).setValue(now);
  return {ok:true};
}
function adminSetStatus(token,ticketId,status,resolution) {
  requireAdmin_(token);
  const allowed = ["Open","In Progress","Resolved","Closed"];
  if (!allowed.includes(status)) throw new Error("Invalid status.");
  const sh=getSheet_(TICKETS_SHEET,HEADERS); const row=findTicketRow_(sh,String(ticketId||"").trim());
  if(!row) throw new Error("Ticket not found.");
  sh.getRange(row,8).setValue(status); sh.getRange(row,9).setValue(String(resolution||""));
  const now = new Date().toISOString(); sh.getRange(row,11).setValue(now);
  addMessage_(ticketId,"System","Status changed to " + status + (resolution ? ". Resolution: " + resolution : ""),now);
  return {ok:true};
}
function adminLogout(token) { CacheService.getScriptCache().remove("staff:"+String(token||"")); return {ok:true}; }

// Customer can continue the same conversation until staff changes status to Closed.
function customerReply_(ticketId, contact, message) {
  ticketId = String(ticketId || "").trim();
  contact = String(contact || "").trim();
  message = String(message || "").trim();
  if (!ticketId || !contact || !message) return json_({ok:false,error:"Ticket ID, contact, and message are required."});
  if (message.length > 5000) return json_({ok:false,error:"Message must be 5,000 characters or fewer."});
  const sh = getSheet_(TICKETS_SHEET, HEADERS);
  const row = findTicketRow_(sh, ticketId);
  if (!row) return json_({ok:false,error:"Ticket not found."});
  const values = sh.getRange(row,1,1,HEADERS.length).getDisplayValues()[0];
  if (String(values[2] || "").trim().toLowerCase() !== contact.toLowerCase()) return json_({ok:false,error:"Contact detail does not match this ticket."});
  if (String(values[7] || "").trim().toLowerCase() === "closed") return json_({ok:false,error:"This ticket is closed. You cannot send more messages."});
  const now = new Date().toISOString();
  addMessage_(ticketId,"Customer",message,now);
  sh.getRange(row,11).setValue(now);
  return json_({ok:true});
}

function lookupCustomerTicket_(ticketId, contact) {
  ticketId = String(ticketId || "").trim(); contact = String(contact || "").trim();
  if (!ticketId || !contact) return {ok:false,error:"Enter ticket ID and the email/phone used when submitting it."};
  const sh=getSheet_(TICKETS_SHEET,HEADERS); const row=findTicketRow_(sh,ticketId);
  if(!row) return {ok:false,error:"Ticket not found. Check the ticket ID."};
  const vals=sh.getRange(row,1,1,HEADERS.length).getDisplayValues()[0];
  if(String(vals[2] || "").trim().toLowerCase() !== contact.toLowerCase()) return {ok:false,error:"The contact detail does not match this ticket."};
  const ticket=rowToTicket_(vals); ticket.messages=readMessages_(ticket.id);
  return {ok:true,ticket:ticket};
}
function readTickets_() { const sh=getSheet_(TICKETS_SHEET,HEADERS); const values=sh.getDataRange().getDisplayValues(); return values.slice(1).filter(r=>r[0]).map(rowToTicket_).reverse(); }
function rowToTicket_(r) { const t={}; HEADERS.forEach((k,i)=>t[k]=r[i]||""); return t; }
function readMessages_(ticketId) { const sh=getSheet_(MESSAGES_SHEET,MESSAGE_HEADERS); const values=sh.getDataRange().getDisplayValues(); return values.slice(1).filter(r=>String(r[0]||"").trim()===String(ticketId||"").trim()).map(r=>({ticketId:r[0],sender:r[1],message:r[2],createdAt:r[3]})); }
function addMessage_(id,sender,message,createdAt) { getSheet_(MESSAGES_SHEET,MESSAGE_HEADERS).appendRow([String(id||"").trim(),sender,safeCell_(String(message).slice(0,10000)),createdAt]); }
function findTicketRow_(sh,id) { const last=sh.getLastRow(); if(last<2)return 0; const ids=sh.getRange(2,1,last-1,1).getDisplayValues(); const i=ids.findIndex(r=>String(r[0]||"").trim()===String(id||"").trim()); return i<0?0:i+2; }
function getSheet_(name,headers) { const ss=SpreadsheetApp.openById(SPREADSHEET_ID); let sh=ss.getSheetByName(name); if(!sh)sh=ss.insertSheet(name); if(sh.getLastRow()===0)sh.appendRow(headers); return sh; }
function requireAdmin_(token) { if(!token || CacheService.getScriptCache().get("staff:"+String(token))!=="ok") throw new Error("Session expired. Please sign in again."); }
function hash_(value) { const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,value,Utilities.Charset.UTF_8); return bytes.map(b=>(b<0?b+256:b).toString(16).padStart(2,"0")).join(""); }
function safeCell_(v) { const s=String(v==null?"":v).slice(0,10000); return /^[=+@]/.test(s) ? "'"+s : s; }
function json_(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }
