/* ====== შეტყობინებები ახალ მოთხოვნაზე: მეილი + WhatsApp ======
   ეს ფაილი დაამატეთ იმავე Google Apps Script პროექტში, სადაც Code.gs არის.

   პარამეტრები (Project Settings → Script Properties), კოდში არაფერი ჩაწეროთ:
     NOTIFY_EMAIL       — მეილი, სადაც მოთხოვნები მოვა (მაგ. you@gmail.com)
     WHATSAPP_PHONE     — WhatsApp ნომერი საერთაშორისო ფორმატით (მაგ. 9955XXXXXXXX)
     CALLMEBOT_APIKEY   — CallMeBot-ის გასაღები (იხ. apps-script/README.md)

   რომელიმე თუ ცარიელია, ის არხი უბრალოდ გამოტოვდება.
   შეცდომა შეტყობინებაში მოთხოვნის შენახვას არ აფერხებს. */

function notifyNewRequest(kind, info) {
  const props = PropertiesService.getScriptProperties();
  const text = formatRequest_(kind, info || {});
  const subject = kind === "group" ? "ახალი ჯავშანი ჯგუფში" : "ახალი ინდივიდუალური მოთხოვნა";

  const email = props.getProperty("NOTIFY_EMAIL");
  if (email) {
    try { MailApp.sendEmail(email, subject, text); }
    catch (e) { console.error("email notify failed: " + e); }
  }

  const phone = props.getProperty("WHATSAPP_PHONE");
  const apikey = props.getProperty("CALLMEBOT_APIKEY");
  if (phone && apikey) {
    try {
      const url = "https://api.callmebot.com/whatsapp.php?phone=" + encodeURIComponent(phone) +
        "&text=" + encodeURIComponent(subject + "\n\n" + text) + "&apikey=" + encodeURIComponent(apikey);
      const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (r.getResponseCode() !== 200) console.error("whatsapp notify failed: " + r.getContentText());
    } catch (e) { console.error("whatsapp notify failed: " + e); }
  }
}

function formatRequest_(kind, d) {
  const lines = [];
  if (d.name) lines.push("სახელი: " + d.name);
  if (d.phone) lines.push("ტელეფონი: " + d.phone);
  if (kind === "group") {
    if (d.label || d.group) lines.push("ჯგუფი: " + (d.label || d.group));
  } else {
    if (d.date) lines.push("თარიღი: " + d.date + (d.time ? " " + d.time : ""));
    if (d.people) lines.push("ადამიანი: " + d.people + (d.price ? " · " + d.price + " ₾" : ""));
    if (d.experience) lines.push("გამოცდილება: " + d.experience);
  }
  return lines.join("\n");
}

/* გაუშვით რედაქტორიდან ერთხელ: ითხოვს ნებართვებს და აგზავნის სატესტო შეტყობინებას. */
function testNotify() {
  notifyNewRequest("individual", {
    name: "ტესტი", phone: "555123456", date: "2026-10-10", time: "19:00",
    people: 2, price: 180, experience: "0-1 წელი",
  });
}
