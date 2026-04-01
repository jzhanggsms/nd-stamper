// ============================================================
// ND Doc Stamper — commands.js
//
// Two entry points:
//   onDocumentBeforeSave — fires automatically on every save
//   stampNow             — fires when user clicks ribbon button
// ============================================================

var ND_PROPS = {
  docNumber:  ["NDDocumentNumber", "ND_ID",     "NDDocNum",     "_ND_ID"],
  version:    ["NDVersion",        "ND_VER",    "NDDocVersion"],
  subVersion: ["NDSubVersion",     "ND_SUBVER", "NDSubVer"]
};

var STAMP_BOOKMARK = "NDDocStamp";
var NOTIFICATION_URL = "https://jzhanggsms.github.io/nd-stamper/notification.html";

// -----------------------------------------------------------
// Helpers
// -----------------------------------------------------------
function findProp(props, keys) {
  for (var i = 0; i < keys.length; i++) {
    var val = props[keys[i]];
    if (val !== undefined && val !== null && val !== "") return val;
  }
  return null;
}

function buildStamp(props) {
  var docNumber  = findProp(props, ND_PROPS.docNumber);
  if (!docNumber) return null;

  var version    = findProp(props, ND_PROPS.version);
  var subVersion = findProp(props, ND_PROPS.subVersion);

  var vStr = version ? (". v" + version + (subVersion ? ("." + subVersion) : "")) : "";
  return docNumber + vStr;
}

function showDialog(url) {
  Office.context.ui.displayDialogAsync(
    url,
    { height: 20, width: 30, displayInIframe: true },
    function(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        var dialog = result.value;
        // Auto-close after 3 seconds
        setTimeout(function() { dialog.close(); }, 3000);
      }
    }
  );
}

// -----------------------------------------------------------
// Shared stamp logic
// -----------------------------------------------------------
async function runStamp(event, isManual) {
  try {
    await Word.run(async function(ctx) {

      // 1. Load custom properties
      var customProps = ctx.document.properties.customProperties;
      customProps.load("items");
      await ctx.sync();

      var props = {};
      customProps.items.forEach(function(p) { props[p.key] = p.value; });

      // 2. Build stamp — bail if not an ND doc
      var stampText = buildStamp(props);
      if (!stampText) {
        if (isManual) {
          // Show "not a NetDocuments document" message to user
          showDialog(NOTIFICATION_URL + "?msg=not-nd");
        }
        event.completed();
        return;
      }

      // 3. Check for existing stamp bookmark
      var bookmarks = ctx.document.bookmarks;
      bookmarks.load("items");
      await ctx.sync();

      var existing = null;
      for (var i = 0; i < bookmarks.items.length; i++) {
        if (bookmarks.items[i].name === STAMP_BOOKMARK) {
          existing = bookmarks.items[i];
          break;
        }
      }

      if (existing) {
        // UPDATE — overwrite with fresh stamp text
        var bmRange = existing.getRange();
        bmRange.insertText(stampText, "Replace");
        await ctx.sync();

      } else {
        // INSERT — first stamp on this document
        var body = ctx.document.body;
        var endRange = body.getRange("End");

        var sep = endRange.insertParagraph("----------------------------------------", "Before");
        sep.font.size = 8;
        sep.font.color = "#AAAAAA";
        sep.spaceBefore = 12;
        sep.spaceAfter = 0;

        var p = endRange.insertParagraph(stampText, "Before");
        p.font.size = 8;
        p.font.color = "#555555";
        p.font.bold = true;
        p.spaceAfter = 0;
        p.spaceBefore = 0;

        await ctx.sync();

        var stampRange = p.getRange();
        stampRange.insertBookmark(STAMP_BOOKMARK);
        await ctx.sync();
      }

      // Show success notification if manually triggered
      if (isManual) {
        showDialog(NOTIFICATION_URL + "?msg=stamped&text=" + encodeURIComponent(stampText));
      }
    });

  } catch (err) {
    console.error("[ND Stamper]", err);
  } finally {
    event.completed();
  }
}

// -----------------------------------------------------------
// Entry point 1 — auto fires on save
// -----------------------------------------------------------
async function onDocumentBeforeSave(event) {
  await runStamp(event, false);
}

// -----------------------------------------------------------
// Entry point 2 — ribbon button "Stamp ND Info"
// -----------------------------------------------------------
async function stampNow(event) {
  await runStamp(event, true);
}

Office.actions.associate("onDocumentBeforeSave", onDocumentBeforeSave);
Office.actions.associate("stampNow", stampNow);