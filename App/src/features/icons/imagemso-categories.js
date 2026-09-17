/**
 * imagemso-categories.js
 *
 * Curated categories for the icon picker's "Popular" tab.
 * Only includes icons that were successfully extracted.
 */

export const ICON_CATEGORIES = [
  {
    label: 'File',
    icons: [
      'FileSave', 'FileSaveAs', 'FileOpen', 'FileClose', 'FileNew',
      'FilePrint', 'FilePrintPreview', 'FilePrintQuick', 'FileProperties'
    ]
  },
  {
    label: 'Edit',
    icons: [
      'Copy', 'Cut', 'Paste', 'PasteValues', 'PasteFormulas',
      'PasteFormatting', 'PasteLink', 'PasteAsPicture', 'PasteTranspose',
      'Undo', 'Redo', 'Clear', 'ClearAll', 'ClearFormats',
      'Delete', 'DeleteRows', 'DeleteColumns', 'DeleteCells',
      'SelectAll', 'FindNext', 'GoTo', 'FormatPainter'
    ]
  },
  {
    label: 'Format',
    icons: [
      'Bold', 'Italic', 'Underline', 'Strikethrough', 'Subscript', 'Superscript',
      'FontColorPicker', 'TextHighlightColorPicker',
      'AlignLeft', 'AlignCenter', 'AlignRight', 'AlignJustify',
      'MergeCells', 'MergeCellsAcross', 'UnmergeCells',
      'WrapText', 'IndentIncrease', 'IndentDecrease',
      'BordersAll', 'BorderTop', 'BorderBottom', 'BorderLeft', 'BorderRight',
      'ShapeFillColorPicker', 'ShapeOutlineColorPicker'
    ]
  },
  {
    label: 'Data',
    icons: [
      'SortDialog', 'SortClear', 'FilterClear', 'FilterReapply',
      'DataValidation', 'RemoveDuplicates',
      'GroupColumns', 'RefreshAll', 'RefreshStatus'
    ]
  },
  {
    label: 'Charts & Tables',
    icons: [
      'ChartInsert', 'ChartTitle', 'ChartLegend', 'ChartGridlines',
      'ChartAxisTitles', 'ChartDataTable',
      'TableInsert', 'TableConvertToRange', 'TableSelect', 'TableResize',
      'TableStylesGallery', 'PivotTableInsert', 'PivotTableOptions',
      'PivotTableExpandField', 'PivotTableFieldSettings',
      'PivotTableGroupSelection'
    ]
  },
  {
    label: 'Formulas',
    icons: [
      'AutoSum',
      'TracePrecedents', 'TraceDependents', 'TraceRemoveAllArrows',
      'ShowFormulas', 'ErrorChecking', 'CalculateNow', 'CalculateSheet',
      'NameDefine', 'NameManager', 'WatchWindow'
    ]
  },
  {
    label: 'Review',
    icons: [
      'SpellingAndGrammar', 'Thesaurus', 'TranslateMenu',
      'ReviewNewComment',
      'ReviewShowOrHideComment', 'ReviewShowAllComments',
      'Accessibility'
    ]
  },
  {
    label: 'View',
    icons: [
      'ViewNormalViewExcel', 'ViewPageLayoutView', 'ViewCustomViews',
      'ViewFullScreenView', 'Zoom100', 'ZoomIn', 'ZoomOut', 'ZoomDialog',
      'FreezeTopRow', 'FreezeFirstColumn', 'FreezePanes',
      'WindowNew', 'WindowSplitToggle', 'WindowHide', 'WindowUnhide'
    ]
  },
  {
    label: 'Shapes & Objects',
    icons: [
      'ShapeRectangle', 'ShapeOval', 'ShapeHeart', 'ShapeFreeform', 'ShapeScribble',
      'ShapeEffectsMenu', 'SmartArtInsert', 'TextBoxInsert', 'ScreenClipping',
      'ObjectBringToFront', 'ObjectSendToBack', 'ObjectBringForward', 'ObjectSendBackward',
      'ObjectRotateRight90', 'ObjectRotateLeft90', 'ObjectFlipHorizontal', 'ObjectFlipVertical'
    ]
  },
  {
    label: 'Macros & Dev',
    icons: [
      'MacroRecord', 'MacroPlay', 'MacroSecurity',
      'VisualBasic', 'ViewCode', 'DesignMode', 'ControlProperties',
      'ActiveXCheckBox', 'ActiveXTextBox', 'ActiveXComboBox',
      'ActiveXListBox', 'ActiveXSpinButton', 'ActiveXScrollBar',
      'ActiveXLabel', 'ActiveXImage', 'ActiveXToggleButton'
    ]
  },
  {
    label: 'Conditional Formatting',
    icons: [
      'ConditionalFormattingMenu', 'ConditionalFormattingColorScalesGallery',
      'ConditionalFormattingDataBarsGallery', 'ConditionalFormattingIconSetsGallery',
      'ConditionalFormattingNewRule'
    ]
  },
  {
    label: 'Colors',
    icons: [
      'AppointmentColor0', 'AppointmentColor1', 'AppointmentColor2',
      'AppointmentColor3', 'AppointmentColor4',
      'AppointmentColor6', 'AppointmentColor7', 'AppointmentColor8',
      'AppointmentColor9', 'AppointmentColor10',
      'ColorBlack', 'ColorBlue', 'ColorRed', 'ColorGreen',
      'ColorYellow', 'ColorPurple', 'ColorWhite'
    ]
  },
  {
    label: 'Insert',
    icons: [
      'HyperlinkInsert', 'SymbolInsert', 'HeaderFooterInsert',
      'SignatureLineInsert', 'ColumnsDialog', 'EnvelopesAndLabels',
      'InkEraser'
    ]
  },
  {
    label: 'General',
    icons: [
      'Info', 'Help', 'About', 'CancelRequest',
      'Lock', 'Magnifier', 'Folder', 'Pushpin', 'Calculator',
      'Camera', 'AddressBook', 'TaskComplete'
    ]
  },
  {
    label: 'Numbers',
    icons: [
      'NumberFormatPercentage', 'NumberFormatFraction', 'NumberFormatText',
      'DecimalsIncrease', 'DecimalsDecrease'
    ]
  },
  {
    label: 'Page Layout',
    icons: [
      'PageOrientationPortrait', 'PageOrientationLandscape',
      'PageSizeA4', 'PageSizeLetter', 'PrintTitles', 'SheetBackground',
      'PrintPreviewClose'
    ]
  }
];

/** Flat list of all "popular" icon names for the default tab. */
export const POPULAR_ICONS = ICON_CATEGORIES.flatMap(cat => cat.icons);
