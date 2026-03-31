// ============================================================
// ND Doc Stamper — commands.js
// Fires on OnDocumentBeforeSave.
// First save: inserts stamp at end of document.
// Subsequent saves: finds stamp via bookmark and updates it.
// Non-ND docs: does nothing.
//
// Stamp format:  ####-####-####. v#.#
// ============================================================

// -----------------------------------------------------------
// CONFIG — run the VBA macro in the README on a live ND doc
// to confirm exact property names for your ND setup.
// First match in each array wins.
// -----------------------------------------------------------
var ND_PROPS = {
  docNumber:  ["NDDocumentNumber", "ND_ID",     "NDDocNum",     "_ND_ID"],
  version:    ["NDVersion",        "ND_VER",    "NDDocVersion"],
  subVersion: ["NDSubVersion",     "ND_SUBVER", "NDSubVer"]
};

// Bookmark used to locate the stamp on subsequent saves.
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
// Entry point
// -----------------------------------------------------------
async function onDocumentBeforeSave(event) {
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
        event.completed();
        return;
      }

      // 3. Check for existing stamp via bookmark
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
        // UPDATE — overwrite bookmark range with fresh stamp text
        var bmRange = existing.getRange();
        bmRange.insertText(stampText, "Replace");
        await ctx.sync();

      } else {
        // INSERT — first time stamp on this document
        var body = ctx.document.body;
        var endRange = body.getRange("End");

        // Thin separator
        var sep = endRange.insertParagraph("----------------------------------------", "Before");
        sep.font.size = 8;
        sep.font.color = "#AAAAAA";
        sep.spaceBefore = 12;
        sep.spaceAfter = 0;

        // Stamp paragraph
        var p = endRange.insertParagraph(stampText, "Before");
        p.font.size = 8;
        p.font.color = "#555555";
        p.font.bold = true;
        p.spaceAfter = 0;
        p.spaceBefore = 0;

        await ctx.sync();

        // Bookmark the stamp so we can update it on future saves
        var stampRange = p.getRange();
        stampRange.insertBookmark(STAMP_BOOKMARK);
        await ctx.sync();
      }
    });

  } catch (err) {
    console.error("[ND Stamper]", err);
  } finally {
    // Always call completed — otherwise Word save will hang
    event.completed();
  }
}

Office.actions.associate("onDocumentBeforeSave", onDocumentBeforeSave);