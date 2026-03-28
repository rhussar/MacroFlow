/**
 * extract-imagemso.js
 *
 * Extracts ImageMSO icons from Microsoft Excel via COM automation.
 * Run once with Excel installed to generate icon PNGs for the sprite sheet.
 *
 * Usage:
 *   node scripts/extract-imagemso.js
 *
 * Prerequisites:
 *   - Microsoft Excel installed
 *   - npm install winax sharp (dev dependencies)
 *
 * Output:
 *   - scripts/imagemso-raw/  (individual BMP files)
 *   - assets/imagemso-sprite.png (composite sprite sheet)
 *   - src/features/icons/imagemso-manifest.json (name → position mapping)
 */

const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const RAW_DIR = path.join(__dirname, 'imagemso-raw');
const ASSETS_DIR = path.join(__dirname, '..', 'assets');
const MANIFEST_DIR = path.join(__dirname, '..', 'src', 'features', 'icons');
const ICON_SIZE = 32;
const SPRITE_COLS = 40;

// Comprehensive list of the most useful ImageMSO identifiers from Office 2016.
// This is a curated set of ~500 commonly used icons.
const IMAGEMSO_NAMES = [
  // File operations
  'FileSave', 'FileSaveAs', 'FileOpen', 'FileClose', 'FileNew', 'FilePrint',
  'FilePrintPreview', 'FilePrintQuick', 'FileProperties', 'FileSendAsAttachment',
  'FileBackup', 'FilePermissions', 'FileEncrypt', 'FileManageVersions',
  // Edit / Clipboard
  'Copy', 'Cut', 'Paste', 'PasteSpecial', 'PasteValues', 'PasteFormulas',
  'PasteFormatting', 'PasteLink', 'PasteAsPicture', 'PasteTranspose',
  'Undo', 'Redo', 'Clear', 'ClearAll', 'ClearFormats', 'ClearContents',
  'Delete', 'DeleteRows', 'DeleteColumns', 'DeleteCells', 'DeleteSheet',
  'SelectAll', 'Find', 'FindNext', 'Replace', 'GoTo',
  // Formatting
  'Bold', 'Italic', 'Underline', 'Strikethrough', 'Subscript', 'Superscript',
  'FontSize', 'FontColorPicker', 'TextHighlightColorPicker',
  'AlignLeft', 'AlignCenter', 'AlignRight', 'AlignJustify',
  'AlignTop', 'AlignMiddle', 'AlignBottom',
  'MergeCells', 'MergeCellsAcross', 'UnmergeCells',
  'WrapText', 'IndentIncrease', 'IndentDecrease',
  'BordersAll', 'BordersOutside', 'BordersInside', 'BordersNone',
  'BorderTop', 'BorderBottom', 'BorderLeft', 'BorderRight',
  'FillColor', 'FontColor', 'FormatPainter',
  // Numbers
  'NumberFormatGeneral', 'NumberFormatNumber', 'NumberFormatCurrency',
  'NumberFormatAccounting', 'NumberFormatDate', 'NumberFormatTime',
  'NumberFormatPercentage', 'NumberFormatFraction', 'NumberFormatScientific',
  'NumberFormatText', 'DecimalsIncrease', 'DecimalsDecrease',
  'ThousandsSeparator',
  // Insert
  'TableInsert', 'PivotTableInsert', 'ChartInsert', 'PictureInsert',
  'ShapesInsert', 'TextBoxInsert', 'HeaderFooterInsert', 'PageBreakInsert',
  'HyperlinkInsert', 'ObjectInsert', 'SymbolInsert', 'CommentInsert',
  'FunctionInsert', 'NameDefine', 'NameManager', 'NameCreate',
  'DiagramInsert', 'SmartArtInsert', 'IconInsert', 'ScreenClipping',
  // Data
  'SortAscending', 'SortDescending', 'SortDialog', 'SortClear',
  'FilterToggle', 'FilterClear', 'FilterAdvanced', 'FilterReapply',
  'GroupRows', 'UngroupRows', 'GroupColumns', 'UngroupColumns',
  'Subtotals', 'DataValidation', 'ConsolidateDialog',
  'TextToColumns', 'RemoveDuplicates', 'FlashFill',
  'ConnectionsDialog', 'RefreshAll', 'RefreshStatus',
  'FromWeb', 'FromText', 'FromAccess', 'FromOtherSources',
  // Formulas
  'AutoSum', 'AutoSumAverage', 'AutoSumCount', 'AutoSumMax', 'AutoSumMin',
  'FormulaInsertFunction', 'TracePrecedents', 'TraceDependents',
  'TraceRemoveAllArrows', 'ShowFormulas', 'ErrorChecking',
  'EvaluateFormula', 'CalculateNow', 'CalculateSheet',
  'WatchWindow', 'NamedCellReference',
  // Charts
  'ChartColumn', 'ChartBar', 'ChartLine', 'ChartPie', 'ChartArea',
  'ChartScatter', 'ChartBubble', 'ChartStock', 'ChartSurface',
  'ChartRadar', 'ChartDoughnut', 'ChartCombo', 'ChartFunnel',
  'ChartWaterfall', 'ChartSunburst', 'ChartTreemap', 'ChartHistogram',
  'ChartBoxWhisker', 'ChartMap',
  'ChartTypeChange', 'ChartMoveChart', 'ChartDataTable', 'ChartLegend',
  'ChartTitle', 'ChartAxisTitles', 'ChartGridlines',
  // Review
  'SpellingAndGrammar', 'Thesaurus', 'TranslateMenu',
  'CommentNew', 'CommentEdit', 'CommentDelete', 'CommentNext', 'CommentPrevious',
  'TrackChangesMenu', 'AcceptChange', 'RejectChange',
  'ProtectSheet', 'ProtectWorkbook', 'AllowEditRanges',
  // View
  'ViewNormalViewExcel', 'ViewPageLayoutView', 'ViewPageBreakPreview',
  'ViewCustomViews', 'ViewFullScreenView',
  'Zoom100', 'ZoomIn', 'ZoomOut', 'ZoomDialog',
  'FreezeTopRow', 'FreezeFirstColumn', 'FreezePanes', 'UnfreezePanes',
  'WindowNew', 'WindowArrange', 'WindowSplitToggle',
  'WindowSwitchWindows', 'WindowHide', 'WindowUnhide',
  'GridlinesToggle', 'HeadingsToggle', 'FormulaBarToggle',
  // Page Layout
  'PageMarginsNarrow', 'PageMarginsNormal', 'PageMarginsWide',
  'PageOrientationPortrait', 'PageOrientationLandscape',
  'PageSizeA4', 'PageSizeLetter', 'PrintAreaSet', 'PrintAreaClear',
  'PageBreakInsertExcel', 'PageBreakRemove', 'PageBreakReset',
  'PrintTitles', 'SheetBackground', 'ScaleToFitWidth', 'ScaleToFitHeight',
  // Shapes & Drawing
  'ShapeRectangle', 'ShapeOval', 'ShapeTriangle', 'ShapeArrowRight',
  'ShapeArrowLeft', 'ShapeArrowUp', 'ShapeArrowDown',
  'ShapeStar5Points', 'ShapeHeart', 'ShapeFlowchartProcess',
  'ShapeFlowchartDecision', 'ShapeFlowchartTerminator',
  'ShapeCallout', 'ShapeConnectorElbow', 'ShapeConnectorStraight',
  'ShapeFreeform', 'ShapeScribble',
  'ShapeFillColorPicker', 'ShapeOutlineColorPicker', 'ShapeEffectsMenu',
  'ObjectBringToFront', 'ObjectSendToBack', 'ObjectBringForward', 'ObjectSendBackward',
  'ObjectAlignLeft', 'ObjectAlignCenter', 'ObjectAlignRight',
  'ObjectAlignTop', 'ObjectAlignMiddle', 'ObjectAlignBottom',
  'ObjectRotateRight90', 'ObjectRotateLeft90', 'ObjectFlipHorizontal', 'ObjectFlipVertical',
  // Developer / Macros
  'MacroRecord', 'MacroPlay', 'MacroSecurity', 'MacrosDialog',
  'VisualBasic', 'ViewCode', 'DesignMode', 'ControlProperties',
  'ActiveXCheckBox', 'ActiveXTextBox', 'ActiveXCommandButton',
  'ActiveXListBox', 'ActiveXComboBox', 'ActiveXSpinButton',
  'ActiveXScrollBar', 'ActiveXOptionButton', 'ActiveXLabel',
  'ActiveXImage', 'ActiveXToggleButton',
  // Conditional Formatting
  'ConditionalFormattingMenu', 'ConditionalFormattingColorScalesGallery',
  'ConditionalFormattingDataBarsGallery', 'ConditionalFormattingIconSetsGallery',
  'ConditionalFormattingHighlightCellsRulesMenu',
  'ConditionalFormattingTopBottomRulesMenu',
  'ConditionalFormattingNewRule', 'ConditionalFormattingClearRules',
  'ConditionalFormattingManageRules',
  // Tables / Lists
  'TableStylesGallery', 'TableHeaderRow', 'TableTotalRow',
  'TableBandedRows', 'TableBandedColumns',
  'TableFirstColumn', 'TableLastColumn',
  'TableConvertToRange', 'TableSelect', 'TableResize',
  // Pivot Table
  'PivotTableActiveField', 'PivotTableFieldList', 'PivotTableFieldSettings',
  'PivotTableExpandField', 'PivotTableCollapseField',
  'PivotTableRefresh', 'PivotTableMoveUp', 'PivotTableMoveDown',
  'PivotTableMoveToBeginning', 'PivotTableMoveToEnd',
  'PivotTableCalculatedField', 'PivotTableCalculatedItem',
  'PivotTableClearAllFilters', 'PivotTableGroupSelection',
  'PivotTableUngroupSelection', 'PivotTableOptions',
  // General UI
  'Info', 'Help', 'About', 'Options', 'CheckMark', 'CancelRequest',
  'Star', 'FlagForFollowUp', 'CategorizeContact',
  'Happy', 'Sad', 'WarningNextItem', 'CriticalUpdate',
  'Lock', 'Unlock', 'Key', 'Shield',
  'Calendar', 'Clock', 'Alarm', 'Timer',
  'Globe', 'Home', 'Mail', 'Phone', 'Camera',
  'Lightbulb', 'Wrench', 'Gear', 'Hammer',
  'Magnifier', 'Binoculars', 'Compass',
  'Folder', 'FolderOpen', 'FolderCreate',
  'Tag', 'Bookmark', 'Pin', 'Pushpin',
  'User', 'Group', 'Meeting', 'PersonalFolder',
  'Dollar', 'Euro', 'Calculator',
  'Database', 'Server', 'Cloud',
  'ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight',
  'ArrowUpRight', 'ArrowDownRight',
  'PlayAudio', 'PauseAudio', 'StopAudio', 'RecordAudio',
  'VolumeUp', 'VolumeDown', 'VolumeMute',
  // Colors & Themes
  'ColorBlack', 'ColorBlue', 'ColorRed', 'ColorGreen',
  'ColorYellow', 'ColorOrange', 'ColorPurple', 'ColorWhite',
  'ThemeColors', 'PageColor',
  // Status indicators
  'TaskAccepted', 'TaskDeclined', 'TaskComplete',
  'TaskInProgress', 'TaskWaiting', 'TaskDeferred',
  'AppointmentColor0', 'AppointmentColor1', 'AppointmentColor2',
  'AppointmentColor3', 'AppointmentColor4', 'AppointmentColor5',
  'AppointmentColor6', 'AppointmentColor7', 'AppointmentColor8',
  'AppointmentColor9', 'AppointmentColor10',
  // Miscellaneous useful
  'PrintPreviewClose', 'PrintPreviewNextPage', 'PrintPreviewPreviousPage',
  'FileNewFromExisting', 'FileNewFromTemplate',
  'ReviewNewComment', 'ReviewDeleteComment', 'ReviewShowOrHideComment',
  'ReviewShowAllComments', 'InkToggle', 'InkEraser',
  'Accessibility', 'EquationInsert', 'SignatureLineInsert',
  'DropCapInsert', 'WatermarkInsert', 'PageBorderAndShading',
  'ColumnsDialog', 'LineNumbers', 'BreakPage',
  'EnvelopesAndLabels', 'MailMergeStartStep',
  'AddressBook', 'ContactsFolder',
  'JournalEntry', 'NoteInsert', 'TaskCreate',
  'ReminderDismiss', 'ReminderDismissAll',
];

async function main() {
  console.log(`\n=== ImageMSO Icon Extraction ===`);
  console.log(`  Icons to extract: ${IMAGEMSO_NAMES.length}`);
  console.log(`  Icon size: ${ICON_SIZE}x${ICON_SIZE}`);
  console.log(`  Output: ${RAW_DIR}\n`);

  // Ensure output directories exist
  if (!fs.existsSync(RAW_DIR)) fs.mkdirSync(RAW_DIR, { recursive: true });
  if (!fs.existsSync(ASSETS_DIR)) fs.mkdirSync(ASSETS_DIR, { recursive: true });
  if (!fs.existsSync(MANIFEST_DIR)) fs.mkdirSync(MANIFEST_DIR, { recursive: true });

  // Step 1: Write the icon name list for VBA to read
  const nameListPath = path.join(RAW_DIR, '_names.txt');
  fs.writeFileSync(nameListPath, IMAGEMSO_NAMES.join('\n'), 'utf8');

  // Step 2: Create a VBS script that opens Excel, injects a VBA macro, and runs it.
  // The VBA macro runs INSIDE Excel where stdole.SavePicture works correctly.
  const vbsPath = path.join(RAW_DIR, '_extract.vbs');
  const outputDirEscaped = RAW_DIR.replace(/\\/g, '\\\\');
  const nameListEscaped = nameListPath.replace(/\\/g, '\\\\');

  const vbsScript = `Option Explicit

Dim xl, wb, vbProj, vbMod

' Start Excel
Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False

' Need to enable programmatic access to VBA project
' This may require: Trust Center > Macro Settings > Trust access to VBA project object model
Set wb = xl.Workbooks.Add

' Inject VBA module
On Error Resume Next
Set vbProj = wb.VBProject
If Err.Number <> 0 Then
  WScript.Echo "ERROR: Cannot access VBA project. Please enable:"
  WScript.Echo "  Excel > File > Options > Trust Center > Trust Center Settings"
  WScript.Echo "  > Macro Settings > Trust access to the VBA project object model"
  wb.Close False
  xl.Quit
  WScript.Quit 1
End If
On Error GoTo 0

Set vbMod = vbProj.VBComponents.Add(1) ' 1 = vbext_ct_StdModule
vbMod.CodeModule.AddFromString _
  "Sub ExtractIcons()" & vbCrLf & _
  "  Dim fso As Object" & vbCrLf & _
  "  Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf & _
  "  Dim outputDir As String" & vbCrLf & _
  "  outputDir = ""${outputDirEscaped}""" & vbCrLf & _
  "  Dim nameFile As Object" & vbCrLf & _
  "  Set nameFile = fso.OpenTextFile(""${nameListEscaped}"", 1)" & vbCrLf & _
  "  Dim allText As String" & vbCrLf & _
  "  allText = nameFile.ReadAll" & vbCrLf & _
  "  nameFile.Close" & vbCrLf & _
  "  Dim names() As String" & vbCrLf & _
  "  names = Split(allText, vbLf)" & vbCrLf & _
  "  Dim i As Long, okCount As Long, errCount As Long" & vbCrLf & _
  "  okCount = 0: errCount = 0" & vbCrLf & _
  "  Dim iconName As String, savePath As String" & vbCrLf & _
  "  Dim pic As stdole.IPictureDisp" & vbCrLf & _
  "  For i = 0 To UBound(names)" & vbCrLf & _
  "    iconName = Trim$(names(i))" & vbCrLf & _
  "    If Len(iconName) > 0 Then" & vbCrLf & _
  "      savePath = outputDir & ""/"" & iconName & "".bmp""" & vbCrLf & _
  "      If Not fso.FileExists(savePath) Then" & vbCrLf & _
  "        On Error Resume Next" & vbCrLf & _
  "        Set pic = Application.CommandBars.GetImageMso(iconName, ${ICON_SIZE}, ${ICON_SIZE})" & vbCrLf & _
  "        If Err.Number = 0 And Not pic Is Nothing Then" & vbCrLf & _
  "          stdole.SavePicture pic, savePath" & vbCrLf & _
  "          If Err.Number = 0 Then" & vbCrLf & _
  "            okCount = okCount + 1" & vbCrLf & _
  "          Else" & vbCrLf & _
  "            errCount = errCount + 1" & vbCrLf & _
  "          End If" & vbCrLf & _
  "        Else" & vbCrLf & _
  "          errCount = errCount + 1" & vbCrLf & _
  "        End If" & vbCrLf & _
  "        On Error GoTo 0" & vbCrLf & _
  "        Set pic = Nothing" & vbCrLf & _
  "      Else" & vbCrLf & _
  "        okCount = okCount + 1" & vbCrLf & _
  "      End If" & vbCrLf & _
  "    End If" & vbCrLf & _
  "  Next i" & vbCrLf & _
  "  Dim resultFile As Object" & vbCrLf & _
  "  Set resultFile = fso.CreateTextFile(outputDir & ""/_result.txt"", True)" & vbCrLf & _
  "  resultFile.WriteLine ""ok="" & okCount" & vbCrLf & _
  "  resultFile.WriteLine ""failed="" & errCount" & vbCrLf & _
  "  resultFile.Close" & vbCrLf & _
  "End Sub"

' Run the macro
xl.Run wb.Name & "!ExtractIcons"

' Cleanup
wb.Close False
xl.Quit
Set xl = Nothing

' Read and display result
Dim resultFso, resultFile2, resultLine
Set resultFso = CreateObject("Scripting.FileSystemObject")
If resultFso.FileExists("${outputDirEscaped}/_result.txt") Then
  Set resultFile2 = resultFso.OpenTextFile("${outputDirEscaped}/_result.txt", 1)
  Do While Not resultFile2.AtEndOfStream
    WScript.Echo resultFile2.ReadLine
  Loop
  resultFile2.Close
Else
  WScript.Echo "No result file — macro may not have run."
End If
`;

  fs.writeFileSync(vbsPath, vbsScript, 'utf8');

  // Step 3: Run the VBS script
  console.log('Extracting icons from Excel via VBA macro...');
  console.log('(If this is your first time, you may need to enable VBA project access in Excel Trust Center)\n');
  try {
    execSync(`cscript //NoLogo "${vbsPath}"`, {
      stdio: 'inherit',
      timeout: 600000 // 10 minute timeout
    });
  } catch (err) {
    console.error('Extraction failed:', err.message);
    console.error('\nIf you see a VBA project access error, do this once:');
    console.error('  1. Open Excel');
    console.error('  2. File > Options > Trust Center > Trust Center Settings');
    console.error('  3. Macro Settings > check "Trust access to the VBA project object model"');
    console.error('  4. Click OK, close Excel, and run this script again.');
    process.exit(1);
  }

  // Read result
  const resultPath = path.join(RAW_DIR, '_result.txt');
  if (fs.existsSync(resultPath)) {
    const result = fs.readFileSync(resultPath, 'utf8');
    console.log('\n' + result);
    fs.unlinkSync(resultPath);
  }

  // Step 4: Build sprite sheet
  console.log('\nBuilding sprite sheet...');
  await buildSpriteSheet();

  // Cleanup
  try {
    fs.unlinkSync(vbsPath);
    fs.unlinkSync(nameListPath);
  } catch { /* ignore */ }

  console.log('\nDone!');
}

async function buildSpriteSheet() {
  let sharp;
  try {
    sharp = require('sharp');
  } catch {
    console.error('sharp not installed. Run: npm install --save-dev sharp');
    console.log('Skipping sprite sheet generation. Run build-icon-sprite.js after installing sharp.');
    return;
  }

  // Read all extracted BMP files
  const files = fs.readdirSync(RAW_DIR)
    .filter(f => f.endsWith('.bmp'))
    .sort();

  if (files.length === 0) {
    console.error('No BMP files found in', RAW_DIR);
    return;
  }

  console.log(`  Found ${files.length} icons`);

  const rows = Math.ceil(files.length / SPRITE_COLS);
  const width = SPRITE_COLS * ICON_SIZE;
  const height = rows * ICON_SIZE;

  // Build manifest and composite operations
  const manifest = {};
  const composites = [];

  let validIdx = 0;
  const skipped = [];

  // Helper: decode legacy BMP (24bpp, bottom-up) to raw RGBA buffer with transparency
  function decodeBmp(buf) {
    if (buf[0] !== 0x42 || buf[1] !== 0x4D) return null; // not a BMP
    const dataOffset = buf.readUInt32LE(10);
    const w = buf.readInt32LE(18);
    const h = Math.abs(buf.readInt32LE(22));
    const bpp = buf.readUInt16LE(28);
    if (bpp !== 24 && bpp !== 32) return null;
    const channels = bpp / 8;
    const rowSize = Math.ceil((w * channels) / 4) * 4;

    // First pass: detect the background color from the corners
    // Office icons use a consistent background color (usually white or magenta)
    function getPixel(px, py) {
      const srcRow = h - 1 - py;
      const off = dataOffset + srcRow * rowSize + px * channels;
      return [buf[off + 2], buf[off + 1], buf[off + 0]]; // RGB
    }
    const corners = [getPixel(0, 0), getPixel(w - 1, 0), getPixel(0, h - 1), getPixel(w - 1, h - 1)];
    // Use the most common corner color as background
    const cornerKey = (c) => `${c[0]},${c[1]},${c[2]}`;
    const counts = {};
    for (const c of corners) {
      const k = cornerKey(c);
      counts[k] = (counts[k] || 0) + 1;
    }
    let bgKey = null, bgCount = 0;
    for (const [k, count] of Object.entries(counts)) {
      if (count > bgCount) { bgKey = k; bgCount = count; }
    }
    const bg = bgKey.split(',').map(Number);
    const tolerance = 10; // allow slight color variation

    const rgba = Buffer.alloc(w * h * 4);
    for (let y = 0; y < h; y++) {
      const srcRow = h - 1 - y; // BMP is bottom-up
      for (let x = 0; x < w; x++) {
        const srcOff = dataOffset + srcRow * rowSize + x * channels;
        const dstOff = (y * w + x) * 4;
        const r = buf[srcOff + 2];
        const g = buf[srcOff + 1];
        const b = buf[srcOff + 0];
        rgba[dstOff + 0] = r;
        rgba[dstOff + 1] = g;
        rgba[dstOff + 2] = b;

        // Make background color transparent
        const dr = Math.abs(r - bg[0]);
        const dg = Math.abs(g - bg[1]);
        const db = Math.abs(b - bg[2]);
        if (dr <= tolerance && dg <= tolerance && db <= tolerance) {
          rgba[dstOff + 3] = 0; // transparent
        } else {
          rgba[dstOff + 3] = 255; // opaque
        }
      }
    }
    return { data: rgba, width: w, height: h, channels: 4 };
  }

  for (let idx = 0; idx < files.length; idx++) {
    const file = files[idx];
    const name = file.replace('.bmp', '');
    const filePath = path.join(RAW_DIR, file);
    const imgBuffer = fs.readFileSync(filePath);

    if (imgBuffer.length < 100) {
      skipped.push(name);
      continue;
    }

    let pngBuffer;
    try {
      const decoded = decodeBmp(imgBuffer);
      if (!decoded) {
        skipped.push(name);
        continue;
      }
      pngBuffer = await sharp(decoded.data, { raw: { width: decoded.width, height: decoded.height, channels: decoded.channels } })
        .resize(ICON_SIZE, ICON_SIZE, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } })
        .png()
        .toBuffer();
    } catch {
      skipped.push(name);
      continue;
    }

    const col = validIdx % SPRITE_COLS;
    const row = Math.floor(validIdx / SPRITE_COLS);
    manifest[name] = { col, row, idx: validIdx };

    composites.push({
      input: pngBuffer,
      left: col * ICON_SIZE,
      top: row * ICON_SIZE
    });
    validIdx++;
  }

  if (skipped.length > 0) {
    console.log(`  Skipped ${skipped.length} unreadable icons: ${skipped.slice(0, 5).join(', ')}${skipped.length > 5 ? '...' : ''}`);
  }

  if (composites.length === 0) {
    console.error('  No valid icons to build sprite sheet from.');
    return;
  }

  // Recalculate dimensions based on actual valid icons
  const actualRows = Math.ceil(composites.length / SPRITE_COLS);
  const actualWidth = SPRITE_COLS * ICON_SIZE;
  const actualHeight = actualRows * ICON_SIZE;

  console.log(`  Building sprite: ${composites.length} icons, ${actualWidth}x${actualHeight}`);

  // Create the sprite sheet
  const spriteSheet = sharp({
    create: {
      width: actualWidth,
      height: actualHeight,
      channels: 4,
      background: { r: 0, g: 0, b: 0, alpha: 0 }
    }
  }).composite(composites).png();

  const spritePath = path.join(ASSETS_DIR, 'imagemso-sprite.png');
  await spriteSheet.toFile(spritePath);
  console.log(`  Sprite sheet: ${spritePath} (${actualWidth}x${actualHeight})`);

  // Write manifest
  const manifestPath = path.join(MANIFEST_DIR, 'imagemso-manifest.json');
  fs.writeFileSync(manifestPath, JSON.stringify(manifest, null, 2), 'utf8');
  console.log(`  Manifest: ${manifestPath} (${Object.keys(manifest).length} entries)`);

  // Also write the sprite metadata for the CSS
  const metaPath = path.join(MANIFEST_DIR, 'imagemso-meta.json');
  fs.writeFileSync(metaPath, JSON.stringify({
    iconSize: ICON_SIZE,
    cols: SPRITE_COLS,
    rows: actualRows,
    width: actualWidth,
    height: actualHeight,
    count: composites.length
  }, null, 2), 'utf8');
  console.log(`  Metadata: ${metaPath}`);
}

main().catch(err => {
  console.error('Fatal error:', err);
  process.exit(1);
});
