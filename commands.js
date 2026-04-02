// ============================================================
// ND Doc Stamper — commands.js
//
// Stamps doc number + version and today's date into the
// primary footer on every save.
//
// Footer layout:
//   Line 1 — left empty for Word's built-in page number
//   Line 2 — doc number + version  (bold, 8pt)
//   Line 3 — today's date          (italic, 8pt)
// ============================================================

var ND_PROPS = {
  docNumber: ["ndDocumentId", "NDDocID"],
  docIDFull: ["NDDocID"]
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
  var docNumber = findProp(props, ND_PROPS.docNumber);
  if (!docNumber) return null;

  var fullID   = findProp(props, ND_PROPS.docIDFull) || "";
  var verMatch = fullID.match(/v[\d.]+$/i);
  var vStr     = verMatch ? (". " + verMatch[0]) : "";

  return docNumber + vStr;
}

function getTodayStr() {
  return new Date().toLocaleDateString("en-US", {
    year: "numeric", month: "short", day: "numeric"
  });
}

function showDialog(url) {
  Office.context.ui.displayDialogAsync(
    url,
    { height: 20, width: 30, displayInIframe: true },
    function(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        var dialog = result.value;
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

      // 2. Build stamp text — bail if not an ND doc
      var stampText = buildStamp(props);
      if (!stampText) {
        if (isManual) {
          showDialog(NOTIFICATION_URL + "?msg=not-nd");
        }
        event.completed();
        return;
      }

      var dateText = getTodayStr();

      // 3. Get the primary footer of the first section
      var footer = ctx.document.sections.getFirst().getFooter("primary");
      var footerBody = footer.body;

      // 4. Check for existing stamp bookmark
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
        // UPDATE — replace bookmarked range with fresh stamp + date
        var bmRange = existing.getRange();
        bmRange.insertText(stampText + "\n" + dateText, "Replace");
        await ctx.sync();

      } else {
        // INSERT — build footer stamp for the first time
        // Footer already has one empty paragraph — leave it for page number.
        // Append stamp and date after it.

        // Line 2: doc number + version
        var p1 = footerBody.insertParagraph(stampText, "End");
        p1.font.size = 8;
        p1.font.bold = true;
        p1.font.color = "#555555";
        p1.alignment = Word.Alignment.left;
        p1.spaceAfter = 0;
        p1.spaceBefore = 0;

        // Line 3: today's date
        var p2 = footerBody.insertParagraph(dateText, "End");
        p2.font.size = 8;
        p2.font.italic = true;
        p2.font.color = "#555555";
        p2.alignment = Word.Alignment.left;
        p2.spaceAfter = 0;
        p2.spaceBefore = 0;

        await ctx.sync();

        // Bookmark spans both lines for future updates
        var stampRange = p1.getRange();
        stampRange.expandTo(p2.getRange());
        stampRange.insertBookmark(STAMP_BOOKMARK);
        await ctx.sync();
      }

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
// Entry point 2 — ribbon button
// -----------------------------------------------------------
async function stampNow(event) {
  await runStamp(event, true);
}

Office.actions.associate("onDocumentBeforeSave", onDocumentBeforeSave);
Office.actions.associate("stampNow", stampNow);
