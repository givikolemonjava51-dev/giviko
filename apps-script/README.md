# სერვერი: Google Apps Script

`Code.gs` არის საიტის სერვერი. ის ინახავს რეგისტრაციებს, ჯავშნებსა და მოთხოვნებს Google Sheets
ცხრილში და ყოველ ახალ მოთხოვნაზე ან ჯგუფში ჯავშანზე გიგზავნით შეტყობინებას მეილზე და WhatsApp-ში.

სანამ `index.html`-ში `API_URL` ცარიელია, საიტი დემო რეჟიმშია და მოთხოვნები არსად ინახება.

## ნაბიჯები
1. **პროექტის შექმნა.** https://script.google.com → New project. დაარქვით „Giviko“.
2. **კოდის ჩასმა.** წაშალეთ რედაქტორში არსებული კოდი და ჩასვით `Code.gs`-ის შიგთავსი.
3. **WhatsApp გასაღები.** გახსენით https://www.callmebot.com/blog/free-api-whatsapp-messages/,
   შეინახეთ იქ მითითებული ბოტის ნომერი კონტაქტებში და WhatsApp-ით გაუგზავნეთ
   `I allow callmebot to send me messages`. პასუხად მოვა apikey.
4. **პარამეტრები.** `Code.gs`-ის თავში შეავსეთ `ADMIN_PASSWORD`, `NOTIFY_EMAIL`,
   `WHATSAPP_PHONE` (995-ით) და `CALLMEBOT_APIKEY`. შეავსეთ მხოლოდ Apps Script-ში, არა GitHub-ზე.
5. **შემოწმება.** ზემოთ ფუნქციების სიიდან აირჩიეთ `setup` → Run → დაეთანხმეთ ნებართვებს.
   შეიქმნება ცხრილი „Giviko — მონაცემები“ თქვენს Google Drive-ში და მოვა სატესტო შეტყობინება.
6. **გამოქვეყნება.** Deploy → New deployment → ტიპი Web app.
   Execute as: **Me**, Who has access: **Anyone** → Deploy. დააკოპირეთ ბმული (…/exec).
7. **საიტთან დაკავშირება.** `index.html`-ში ჩასვით ეს ბმული: `const API_URL = "…/exec";`

კოდის შეცვლის შემდეგ: Deploy → Manage deployments → Edit → Version: New version → Deploy.
ბმული იგივე რჩება.

## სხვა
- ცხრილში „Reserved“ ფურცელზე შეგიძლიათ მიუთითოთ, ჯგუფში რამდენი ადგილია უკვე დაკავებული საიტის გარეშე.
- PIN-კოდები ცხრილში დაშიფრული (hash) სახით ინახება.
