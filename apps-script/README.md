# შეტყობინებები მეილზე და WhatsApp-ზე

როცა მომხმარებელი აგზავნის ინდივიდუალურ მოთხოვნას ან ჯავშნის ადგილს ჯგუფში,
სერვერი (Google Apps Script) გიგზავნით შეტყობინებას მეილზე და WhatsApp-ში.

## 1. ფაილის დამატება
Apps Script რედაქტორში, Code.gs-ის გვერდით, დაამატეთ ახალი ფაილი `Notify.gs`
და ჩასვით ამ საქაღალდის `Notify.gs`-ის შიგთავსი.

## 2. გამოძახება Code.gs-ში
Code.gs-ში, იმ ადგილას, სადაც ახალი მოთხოვნა ცხრილში ინახება (`individual` და `group`
მოქმედებები), შენახვის შემდეგ დაამატეთ ერთი ხაზი:

```js
// action: "individual"
notifyNewRequest("individual", { name, phone, date, time, people, price, experience });

// action: "group"
notifyNewRequest("group", { name, phone, group, label });
```

ცვლადების სახელები თქვენს კოდს მოარგეთ.

## 3. WhatsApp გასაღები (CallMeBot, უფასო)
1. თქვენს ტელეფონში შეინახეთ კონტაქტი **+34 694 29 84 96**.
2. WhatsApp-ით გაუგზავნეთ: `I allow callmebot to send me messages`
3. პასუხად მოგივათ apikey.

## 4. პარამეტრები
Apps Script → Project Settings → Script Properties:

| Property | მნიშვნელობა |
|---|---|
| `NOTIFY_EMAIL` | თქვენი მეილი |
| `WHATSAPP_PHONE` | ნომერი `995` კოდით, მაგ. `9955XXXXXXXX` |
| `CALLMEBOT_APIKEY` | CallMeBot-ის გასაღები |

## 5. შემოწმება და გამოქვეყნება
1. რედაქტორში აირჩიეთ ფუნქცია `testNotify` და დააჭირეთ Run, დაეთანხმეთ ნებართვებს.
   სატესტო შეტყობინება უნდა მოვიდეს მეილზე და WhatsApp-ში.
2. Deploy → Manage deployments → Edit → Version: New version → Deploy.
   Web App ბმული იგივე რჩება, index.html-ში არაფრის შეცვლა არ არის საჭირო.
