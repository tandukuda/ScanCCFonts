// ============================================================
//  ScanCCFonts.jsx  —  Adobe Illustrator ExtendScript
//  ⚠️  Cross-Platform (Windows & macOS)
//
//  Scans the Creative Cloud font folder + entitlements manifest,
//  then creates a new Illustrator document where each font is
//  applied to a text frame in its own typeface.
//
//  When you copy the saved .ai to another computer and open it,
//  Creative Cloud will detect missing Adobe Fonts and sync them
//  automatically.
//
//  How to run:
//    Illustrator menu → File > Scripts > Other Script…
//    then select this file.
// ============================================================

(function () {
  // ── 1. OS Detection & Paths ──────────────────────────────
  var isMac = $.os.match(/macintosh/i);

  var LIVETYPE = isMac
    ? $.getenv("HOME") +
      "/Library/Application Support/Adobe/CoreSync/plugins/livetype/"
    : $.getenv("APPDATA") + "/Adobe/CoreSync/plugins/livetype/";

  // Mac uses dot-prefixed hidden folders, Windows does not
  var SUBFOLDERS = isMac
    ? [".e", ".r", ".t", ".u", ".w", ".x"]
    : ["e", "r", "t", "u", "w", "x"];

  var MANIFEST = isMac
    ? LIVETYPE + ".c/entitlements.xml"
    : LIVETYPE + "c/entitlements.xml";

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

  function makeGray(val) {
    var c = new RGBColor();
    c.red = val;
    c.green = val;
    c.blue = val;
    return c;
  }

  function normalise(str) {
    return str.replace(/[\s\-_]/g, "").toLowerCase();
  }

  // ── 2. Scan font subfolders ──────────────────────────────
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

  // ── 3. Parse entitlements.xml ────────────────────────────
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

  // ── 4. Merge & sort ──────────────────────────────────────
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

  // ── 5. Document layout & dynamic height settings ─────────
  var COLS = 8; // text frame columns
  var ROW_H = 28; // height per row (pt)
  var COL_W = 300; // width per column (pt)
  var MARGIN = 40; // page margin (pt)
  var FONT_SIZE = 11; // preview font size (pt)

  var totalFonts = sortedNames.length;
  var totalRowsNeeded = Math.ceil(totalFonts / COLS);

  // Calculate width and dynamic height
  var DOC_W = COLS * COL_W + MARGIN * 2;
  var calculatedHeight = totalRowsNeeded * ROW_H + MARGIN * 2 + 100; // +100 for header/footer padding
  var DOC_H = Math.max(600, calculatedHeight); // Enforce a minimum height of 600pt

  // ── 6. Create document ───────────────────────────────────
  var docPreset = new DocumentPreset();
  docPreset.width = DOC_W;
  docPreset.height = DOC_H;
  docPreset.units = RulerUnits.Points;
  docPreset.colorMode = DocumentColorSpace.RGB;
  docPreset.title = "CC Font Collection";

  var doc = app.documents.addDocument("Print", docPreset);
  var layer = doc.layers[0];
  layer.name = "Fonts";

  // ── 7. Build lookup map: display name → TextFont object ──
  var fontMap = {};

  for (var fi = 0; fi < app.textFonts.length; fi++) {
    var tf = app.textFonts[fi];
    fontMap[tf.name] = tf;
    fontMap[normalise(tf.name)] = tf;
    try {
      var displayKey = normalise(tf.family + tf.style);
      fontMap[displayKey] = tf;
    } catch (e) {}
  }

  function findFont(displayName) {
    if (fontMap[displayName]) return fontMap[displayName];
    var nk = normalise(displayName);
    if (fontMap[nk]) return fontMap[nk];
    try {
      return app.textFonts.getByName(displayName);
    } catch (e) {}
    return null;
  }

  // ── 8. Helper: add one text frame row ────────────────────
  function addFontRow(fontName, rowIndex) {
    var col = rowIndex % COLS;
    var row = Math.floor(rowIndex / COLS);
    var xBase = MARGIN + col * COL_W;
    var yBase = DOC_H - MARGIN - 60 - row * ROW_H;

    var previewFrame = layer.textFrames.add();
    previewFrame.contents = fontName;
    previewFrame.left = xBase;
    previewFrame.top = yBase;

    // Width and height properties removed here to prevent point text stretching

    var previewRange = previewFrame.textRange;

    // Force 100% scaling to fix deformation
    previewRange.characterAttributes.size = FONT_SIZE;
    previewRange.characterAttributes.horizontalScale = 100;
    previewRange.characterAttributes.verticalScale = 100;

    var matched = findFont(fontName);
    if (matched) {
      previewRange.characterAttributes.textFont = matched;
      return true;
    } else {
      var gray = makeGray(160);
      previewRange.characterAttributes.fillColor = gray;
      return false;
    }
  }

  // ── 9. Page header ───────────────────────────────────────
  var headerFrame = layer.textFrames.add();
  headerFrame.contents =
    "Adobe CC Font Collection  —  " + sortedNames.length + " fonts";
  headerFrame.left = MARGIN;
  headerFrame.top = DOC_H - MARGIN;

  var hr = headerFrame.textRange;
  hr.characterAttributes.size = 18;
  hr.characterAttributes.horizontalScale = 100;
  hr.characterAttributes.verticalScale = 100;
  try {
    hr.characterAttributes.textFont = app.textFonts.getByName("ArialMT");
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

  // ── 10. Add a text frame per font ────────────────────────
  var applied = 0;
  var missing = 0;

  for (var n = 0; n < sortedNames.length; n++) {
    var name = sortedNames[n];
    var result = addFontRow(name, n);
    if (result) applied++;
    else missing++;
  }

  // ── 11. Footer note ──────────────────────────────────────
  var footerFrame = layer.textFrames.add();
  footerFrame.contents =
    "Fonts shown in gray (" +
    missing +
    ") are referenced but not currently active on this machine. " +
    "Active on this machine: " +
    applied +
    ". " +
    "Open this file on the target computer — Creative Cloud will auto-sync Adobe Fonts. by tandukuda.";
  footerFrame.left = MARGIN;
  footerFrame.top = MARGIN + 28;

  var fr = footerFrame.textRange;
  fr.characterAttributes.size = 7;
  fr.characterAttributes.horizontalScale = 100;
  fr.characterAttributes.verticalScale = 100;

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
      "Save this file as a regular .ai and copy it to your other computer.\n" +
      "When opened, Creative Cloud will auto-sync all Adobe Fonts.",
  );
})();
