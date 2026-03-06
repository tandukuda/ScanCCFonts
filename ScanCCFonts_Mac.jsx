// ============================================================
//  ScanCCFonts_Mac.jsx  —  Adobe Illustrator ExtendScript
//  ⚠️  macOS ONLY
//
//  Scans the Creative Cloud font folder + entitlements manifest,
//  then creates a new Illustrator document where each font is
//  applied to a text frame in its own typeface.
//
//  When you copy the saved .ai to another Mac and open it,
//  Creative Cloud will detect missing Adobe Fonts and sync them
//  automatically.
//
//  Font folder : ~/Library/Application Support/Adobe/CoreSync/plugins/livetype/.[e,r,t,u,w,x]
//  Manifest    : ~/Library/Application Support/Adobe/CoreSync/plugins/livetype/.c/entitlements.xml
//
//  Note: CC font subfolders on Mac are hidden (dot-prefixed).
//
//  How to run:
//    Illustrator menu → File > Scripts > Other Script…
//    then select this file.
// ============================================================

(function () {
  // ── Paths ────────────────────────────────────────────────
  // On macOS the livetype subfolders are hidden (dot-prefixed).
  var HOME = $.getenv("HOME");
  var LIVETYPE =
    HOME + "/Library/Application Support/Adobe/CoreSync/plugins/livetype/";
  var SUBFOLDERS = [".e", ".r", ".t", ".u", ".w", ".x"];
  var MANIFEST = LIVETYPE + ".c/entitlements.xml";

  // ── Helpers ──────────────────────────────────────────────
  function stripExt(filename) {
    return filename.replace(/\.[^.]+$/, "");
  }

  function isFontFile(filename) {
    return /\.(otf|ttf|woff|woff2|pfb|pfm|fon|bdf)$/i.test(filename);
  }

  function matchAll(str, pattern) {
    var results = [];
    var m;
    pattern.lastIndex = 0;
    while ((m = pattern.exec(str)) !== null) {
      if (m[1]) results.push(m[1]);
    }
    return results;
  }

  function trim(str) {
    return str.replace(/^\s+|\s+$/g, "");
  }

  // ── 1. Scan font subfolders ──────────────────────────────
  var folderFonts = {};

  for (var i = 0; i < SUBFOLDERS.length; i++) {
    var dir = new Folder(LIVETYPE + SUBFOLDERS[i]);
    if (!dir.exists) continue;
    var items = dir.getFiles();
    for (var j = 0; j < items.length; j++) {
      var item = items[j];
      if (!(item instanceof File)) continue;
      var fname = decodeURI(item.name);
      if (!isFontFile(fname)) continue;
      folderFonts[stripExt(fname)] = true;
    }
  }

  // ── 2. Parse entitlements.xml ────────────────────────────
  var manifestFonts = {};
  var manifestFile = new File(MANIFEST);

  if (manifestFile.exists) {
    manifestFile.encoding = "UTF-8";
    manifestFile.open("r");
    var xml = manifestFile.read();
    manifestFile.close();

    var patterns = [
      /postscriptName["\s]*[:=]["\s]*["']?([A-Za-z0-9_\-\. ]+)["']?/gi,
      /<postscriptName>([^<]+)<\/postscriptName>/gi,
      /fullName["\s]*[:=]["\s]*["']?([A-Za-z0-9_\-\. ]+)["']?/gi,
      /<fullName>([^<]+)<\/fullName>/gi,
      /familyName["\s]*[:=]["\s]*["']?([A-Za-z0-9_\-\. ]+)["']?/gi,
      /<familyName>([^<]+)<\/familyName>/gi,
      /name="([A-Za-z0-9_\-\. ]+)"/gi,
    ];

    for (var p = 0; p < patterns.length; p++) {
      var matches = matchAll(xml, patterns[p]);
      for (var k = 0; k < matches.length; k++) {
        var val = trim(matches[k]);
        if (val.length > 1) manifestFonts[val] = true;
      }
    }
  }

  // ── 3. Merge & sort ──────────────────────────────────────
  var allFonts = {};
  var key;
  for (key in folderFonts) allFonts[key] = true;
  for (key in manifestFonts) allFonts[key] = true;

  var sortedNames = [];
  for (key in allFonts) sortedNames.push(key);
  sortedNames.sort(function (a, b) {
    return a.toLowerCase() < b.toLowerCase() ? -1 : 1;
  });

  if (sortedNames.length === 0) {
    alert(
      "No fonts found. Check that the Creative Cloud livetype folder exists:\n" +
        LIVETYPE,
    );
    return;
  }

  // ── Utility: gray color (defined early, used in addFontRow) ─
  function makeGray(val) {
    var c = new RGBColor();
    c.red = val;
    c.green = val;
    c.blue = val;
    return c;
  }

  // ── 4. Document layout settings ─────────────────────────
  var COLS = 8; // text frame columns
  var ROW_H = 28; // height per row (pt)
  var COL_W = 300; // width per column (pt)
  var MARGIN = 40; // page margin (pt)
  var FONT_SIZE = 11; // preview font size (pt)
  var LABEL_SIZE = 8; // font name label size (pt)
  var ROWS_PER_PAGE = 36;

  var DOC_W = COLS * COL_W + MARGIN * 2;
  var DOC_H = ROWS_PER_PAGE * ROW_H + MARGIN * 2 + 60; // +60 for header

  // ── 5. Create document ───────────────────────────────────
  var docPreset = new DocumentPreset();
  docPreset.width = DOC_W;
  docPreset.height = DOC_H;
  docPreset.units = RulerUnits.Points;
  docPreset.colorMode = DocumentColorSpace.RGB;
  docPreset.title = "CC Font Collection";

  var doc = app.documents.addDocument("Print", docPreset);
  var layer = doc.layers[0];
  layer.name = "Fonts";

  // ── 6. Build a lookup map: display name → TextFont object ─
  // app.textFonts exposes every currently active font.
  // We index by (a) the font's own .name, and (b) a normalised
  // version so we can match "Acumin Pro Italic" → postscript
  // name "AcuminPro-Italic" reliably.
  var fontMap = {}; // normalised key → TextFont

  function normalise(str) {
    return str.replace(/[\s\-_]/g, "").toLowerCase();
  }

  for (var fi = 0; fi < app.textFonts.length; fi++) {
    var tf = app.textFonts[fi];
    // Index by PostScript name (exact)
    fontMap[tf.name] = tf;
    // Index by normalised PostScript name
    fontMap[normalise(tf.name)] = tf;
    // Index by normalised display name (family + style combined)
    try {
      var displayKey = normalise(tf.family + tf.style);
      fontMap[displayKey] = tf;
    } catch (e) {}
  }

  function findFont(displayName) {
    // 1. Direct PostScript match
    if (fontMap[displayName]) return fontMap[displayName];
    // 2. Normalised match (strips spaces, dashes, case)
    var nk = normalise(displayName);
    if (fontMap[nk]) return fontMap[nk];
    // 3. Fallback: getByName (throws if missing)
    try {
      return app.textFonts.getByName(displayName);
    } catch (e) {}
    return null;
  }

  // ── 7. Helper: add one text frame row ────────────────────
  function addFontRow(fontName, rowIndex) {
    var col = rowIndex % COLS;
    var row = Math.floor(rowIndex / COLS);
    var xBase = MARGIN + col * COL_W;
    var yBase = DOC_H - MARGIN - 60 - row * ROW_H;

    var previewFrame = layer.textFrames.add();
    previewFrame.contents = fontName;
    previewFrame.left = xBase;
    previewFrame.top = yBase;
    previewFrame.width = COL_W - 10;
    previewFrame.height = ROW_H * 0.65;

    var previewRange = previewFrame.textRange;
    previewRange.size = FONT_SIZE;

    var matched = findFont(fontName);
    if (matched) {
      previewRange.textFont = matched;
      return true;
    } else {
      // Font reference embedded but not active on this machine
      var gray = makeGray(160);
      previewRange.fillColor = gray;
      return false;
    }
  }

  // ── 7. Page header ───────────────────────────────────────
  var headerFrame = layer.textFrames.add();
  headerFrame.contents =
    "Adobe CC Font Collection  —  " + sortedNames.length + " fonts";
  headerFrame.left = MARGIN;
  headerFrame.top = DOC_H - MARGIN;
  headerFrame.width = DOC_W - MARGIN * 2;
  headerFrame.height = 50;

  var hr = headerFrame.textRange;
  hr.size = 18;
  try {
    hr.textFont = app.textFonts.getByName("ArialMT");
  } catch (e) {}

  // Separator line
  var line = layer.pathItems.add();
  line.setEntirePath([
    [MARGIN, DOC_H - MARGIN - 52],
    [DOC_W - MARGIN, DOC_H - MARGIN - 52],
  ]);
  line.stroked = true;
  line.filled = false;
  line.strokeWidth = 0.5;

  // ── 8. Add a text frame per font ─────────────────────────
  var applied = 0;
  var missing = 0;

  for (var n = 0; n < sortedNames.length; n++) {
    var name = sortedNames[n];
    var result = addFontRow(name, n);
    if (result) applied++;
    else missing++;
  }

  // ── 9. Footer note ───────────────────────────────────────
  var footerFrame = layer.textFrames.add();
  footerFrame.contents =
    "Fonts shown in gray (" +
    missing +
    ") are referenced but not currently active on this machine. " +
    "Active on this machine: " +
    applied +
    ". " +
    "Open this file on the target Mac — Creative Cloud will auto-sync Adobe Fonts.";
  footerFrame.left = MARGIN;
  footerFrame.top = MARGIN + 28;
  footerFrame.width = DOC_W - MARGIN * 2;
  footerFrame.height = 24;

  var fr = footerFrame.textRange;
  fr.size = 7;

  alert(
    "Done!\n\n" +
      "Total fonts referenced : " +
      sortedNames.length +
      "\n" +
      "Active on this machine : " +
      applied +
      "\n" +
      "Inactive / gray        : " +
      missing +
      "\n\n" +
      "Save this file as a regular .ai and copy it to your other Mac.\n" +
      "When opened, Creative Cloud will auto-sync all Adobe Fonts.",
  );
})();
