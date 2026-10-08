/*
 * Typo Fixer - Outlook add-in
 * Fixes common typos in the message body (your text only, not quoted replies).
 * Runs when you click "Fix typos" and automatically when you press Send.
 * To add words: add a line to CORRECTIONS below ("typo": "correction").
 */

const CORRECTIONS = {
  "teh": "the",
  "adn": "and",
  "taht": "that",
  "waht": "what",
  "wiht": "with",
  "hte": "the",
  "thier": "their",
  "jsut": "just",
  "abotu": "about",
  "yuo": "you",
  "recieve": "receive",
  "recieved": "received",
  "definately": "definitely",
  "seperate": "separate",
  "occured": "occurred",
  "occurence": "occurrence",
  "untill": "until",
  "wich": "which",
  "becuase": "because",
  "beacuse": "because",
  "alot": "a lot",
  "accomodate": "accommodate",
  "acheive": "achieve",
  "adress": "address",
  "agressive": "aggressive",
  "apparant": "apparent",
  "arguement": "argument",
  "basicly": "basically",
  "beleive": "believe",
  "buisness": "business",
  "calender": "calendar",
  "carefull": "careful",
  "cieling": "ceiling",
  "collegue": "colleague",
  "comming": "coming",
  "commited": "committed",
  "completly": "completely",
  "concious": "conscious",
  "decison": "decision",
  "desparate": "desperate",
  "dissapoint": "disappoint",
  "embarass": "embarrass",
  "enviroment": "environment",
  "existance": "existence",
  "experiance": "experience",
  "familar": "familiar",
  "finaly": "finally",
  "foriegn": "foreign",
  "freind": "friend",
  "fullfill": "fulfill",
  "goverment": "government",
  "gaurd": "guard",
  "happend": "happened",
  "harrass": "harass",
  "heighth": "height",
  "immediatly": "immediately",
  "independant": "independent",
  "intresting": "interesting",
  "knowlege": "knowledge",
  "liason": "liaison",
  "libary": "library",
  "lenght": "length",
  "maintainance": "maintenance",
  "managment": "management",
  "mispell": "misspell",
  "neccessary": "necessary",
  "noticable": "noticeable",
  "occassion": "occasion",
  "oppurtunity": "opportunity",
  "paralell": "parallel",
  "persistant": "persistent",
  "posession": "possession",
  "prefered": "preferred",
  "probaly": "probably",
  "profesional": "professional",
  "publically": "publicly",
  "realy": "really",
  "reccomend": "recommend",
  "refered": "referred",
  "relevent": "relevant",
  "religous": "religious",
  "remeber": "remember",
  "resistence": "resistance",
  "responsability": "responsibility",
  "rythm": "rhythm",
  "scedule": "schedule",
  "successfull": "successful",
  "suprise": "surprise",
  "tommorow": "tomorrow",
  "truely": "truly",
  "unfortunatly": "unfortunately",
  "wierd": "weird",
  "writting": "writing",
  "thnaks": "thanks",
  "pelase": "please",
  "dont": "don't",
  "cant": "can't",
  "didnt": "didn't",
  "doesnt": "doesn't",
  "isnt": "isn't"
};

const KEYS = Object.keys(CORRECTIONS).sort((a, b) => b.length - a.length);
const PATTERN = new RegExp(
  "(^|[^\\w'\u2019])(" + KEYS.join("|") + ")(?![\\w'\u2019])",
  "gi"
);

function matchCase(original, replacement) {
  if (original.length > 1 && original === original.toUpperCase()) {
    return replacement.toUpperCase();
  }
  if (original[0] !== original[0].toLowerCase()) {
    return replacement[0].toUpperCase() + replacement.slice(1);
  }
  return replacement;
}

function fixText(text) {
  let count = 0;
  const out = text.replace(PATTERN, (m, pre, word) => {
    const fix = CORRECTIONS[word.toLowerCase()];
    if (!fix) return m;
    count++;
    return pre + matchCase(word, fix);
  });
  return { text: out, count };
}

function isQuoted(node) {
  for (let el = node.parentElement; el; el = el.parentElement) {
    const tag = el.tagName;
    if (tag === "BLOCKQUOTE" || tag === "PRE" || tag === "CODE" || tag === "A") return true;
    if (el.id === "divRplyFwdMsg" || el.id === "appendonsend" || el.id === "mail-editor-reference-message-container") return true;
  }
  return false;
}

function fixHtml(html) {
  const doc = new DOMParser().parseFromString(html, "text/html");
  const walker = doc.createTreeWalker(doc.body, NodeFilter.SHOW_TEXT);
  const nodes = [];
  while (walker.nextNode()) nodes.push(walker.currentNode);
  let total = 0;
  for (const node of nodes) {
    if (isQuoted(node)) continue;
    const r = fixText(node.nodeValue);
    if (r.count > 0) {
      node.nodeValue = r.text;
      total += r.count;
    }
  }
  return { html: "<!DOCTYPE html>" + doc.documentElement.outerHTML, count: total };
}

function getBodyHtml() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (r) =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error)
    );
  });
}

function setBodyHtml(html) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, (r) =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error)
    );
  });
}

async function applyFixes() {
  const html = await getBodyHtml();
  const result = fixHtml(html);
  if (result.count > 0) await setBodyHtml(result.html);
  return result.count;
}

function notify(message) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("typoFixer", {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: message,
    icon: "Icon.16x16",
    persistent: false,
  });
}

function fixTypos(event) {
  applyFixes()
    .then((n) => notify(n ? "Fixed " + n + " typo" + (n === 1 ? "" : "s") + "." : "No typos found."))
    .catch(() => notify("Could not fix typos."))
    .finally(() => event.completed());
}

function onMessageSendHandler(event) {
  applyFixes()
    .catch(() => {})
    .then(() => event.completed({ allowEvent: true }));
}

Office.onReady(() => {});
Office.actions.associate("fixTypos", fixTypos);
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
