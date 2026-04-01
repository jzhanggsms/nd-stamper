// ============================================================
// ND Doc Stamper — commands.js
//
// Two entry points:
//   onDocumentBeforeSave — fires automatically on every save
//   stampNow             — fires when user clicks ribbon button
//
// Both call the same shared runStamp() logic.
// ============================================================

var ND_PROPS = {
  docNumber:  ["NDDocumentNumber", "ND_ID",     "NDDocNum",     "_ND_ID"],
  version:    ["NDVersion",        "ND_VER",    "NDDocVersion"],
  subVersion: ["NDSubVersion",     "ND_SUBVER", "NDSubVer"]
};

var STAMP_BOOKMARK = "NDDocStamp";

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

// -----------------------------------------------------------
// Shared stamp logic — used by both entry points
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
          // Let user know if they clicked the button on a non-ND doc
          var body = ctx.document.body;
          body.load("text");
          await ctx.sync();
          Office.context.ui.displayDialogAsync(
            "https://jzhanggsms.github.io/nd-stamper/taskpane.html",
            { height: 30, width: 30 }
          );
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