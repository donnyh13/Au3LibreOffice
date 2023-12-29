#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7
#include-once

; #INDEX# =======================================================================================================================
; Title .........: Libre Office Writer (LOWriter) Constants for the Libre Office Writer UDF.
; AutoIt Version : v3.3.16.1
; Description ...: COnstants for various functions in the Libre Office Writer UDF.
; Author(s) .....: donnyh13
; Dll ...........:
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; ===============================================================================================================================

;Cursor Data Related Constants
Global Const $LOW_CURDATA_BODY_TEXT = 1, _
		$LOW_CURDATA_FRAME = 2, _
		$LOW_CURDATA_CELL = 3, _
		$LOW_CURDATA_FOOTNOTE = 4, _
		$LOW_CURDATA_ENDNOTE = 5

;Cursor Type Related Constants
Global Const $LOW_CURTYPE_TEXT_CURSOR = 1, _
		$LOW_CURTYPE_TABLE_CURSOR = 2, _
		$LOW_CURTYPE_VIEW_CURSOR = 3

;Path Convert Constants.
Global Const $LOW_PATHCONV_AUTO_RETURN = 0, _
		$LOW_PATHCONV_OFFICE_RETURN = 1, _
		$LOW_PATHCONV_PCPATH_RETURN = 2

;Printer Duplex Constants.
Global Const $LOW_DUPLEX_UNKNOWN = 0, _
		$LOW_DUPLEX_OFF = 1, _
		$LOW_DUPLEX_LONG = 2, _
		$LOW_DUPLEX_SHORT = 3

;Printer Paper Orientation Constants.
Global Const $LOW_PAPER_PORTRAIT = 0, _
		$LOW_PAPER_LANDSCAPE = 1

;Paper Size Constants.
Global Const $LOW_PAPER_A3 = 0, _
		$LOW_PAPER_A4 = 1, _
		$LOW_PAPER_A5 = 2, _
		$LOW_PAPER_B4 = 3, _
		$LOW_PAPER_B5 = 4, _
		$LOW_PAPER_LETTER = 5, _
		$LOW_PAPER_LEGAL = 6, _
		$LOW_PAPER_TABLOID = 7, _
		$LOW_PAPER_USER_DEFINED = 8

;LO Print Comments Constants.
Global Const $LOW_PRINT_NOTES_NONE = 0, _
		$LOW_PRINT_NOTES_ONLY = 1, _
		$LOW_PRINT_NOTES_END = 2, _
		$LOW_PRINT_NOTES_NEXT_PAGE = 3

;LO Cursor Constants.
Global Enum $LOW_VIEWCUR_GO_DOWN, $LOW_VIEWCUR_GO_UP, $LOW_VIEWCUR_GO_LEFT, $LOW_VIEWCUR_GO_RIGHT, $LOW_VIEWCUR_GOTO_END_OF_LINE, _
		$LOW_VIEWCUR_GOTO_START_OF_LINE, $LOW_VIEWCUR_JUMP_TO_FIRST_PAGE, $LOW_VIEWCUR_JUMP_TO_LAST_PAGE, $LOW_VIEWCUR_JUMP_TO_PAGE, _
		$LOW_VIEWCUR_JUMP_TO_NEXT_PAGE, $LOW_VIEWCUR_JUMP_TO_PREV_PAGE, $LOW_VIEWCUR_JUMP_TO_END_OF_PAGE, $LOW_VIEWCUR_JUMP_TO_START_OF_PAGE, _
		$LOW_VIEWCUR_SCREEN_DOWN, $LOW_VIEWCUR_SCREEN_UP, $LOW_VIEWCUR_GOTO_START, $LOW_VIEWCUR_GOTO_END

Global Enum $LOW_TEXTCUR_COLLAPSE_TO_START, $LOW_TEXTCUR_COLLAPSE_TO_END, $LOW_TEXTCUR_GO_LEFT, $LOW_TEXTCUR_GO_RIGHT, _
		$LOW_TEXTCUR_GOTO_START, $LOW_TEXTCUR_GOTO_END, $LOW_TEXTCUR_GOTO_NEXT_WORD, $LOW_TEXTCUR_GOTO_PREV_WORD, $LOW_TEXTCUR_GOTO_END_OF_WORD, _
		$LOW_TEXTCUR_GOTO_START_OF_WORD, $LOW_TEXTCUR_GOTO_NEXT_SENTENCE, $LOW_TEXTCUR_GOTO_PREV_SENTENCE, $LOW_TEXTCUR_GOTO_END_OF_SENTENCE, _
		$LOW_TEXTCUR_GOTO_START_OF_SENTENCE, $LOW_TEXTCUR_GOTO_NEXT_PARAGRAPH, $LOW_TEXTCUR_GOTO_PREV_PARAGRAPH, _
		$LOW_TEXTCUR_GOTO_END_OF_PARAGRAPH, $LOW_TEXTCUR_GOTO_START_OF_PARAGRAPH

Global Enum $LOW_TABLECUR_GO_LEFT, $LOW_TABLECUR_GO_RIGHT, $LOW_TABLECUR_GO_UP, $LOW_TABLECUR_GO_DOWN, $LOW_TABLECUR_GOTO_START, _
		$LOW_TABLECUR_GOTO_END

;Break Type
Global Const $LOW_BREAK_NONE = 0, $LOW_BREAK_COLUMN_BEFORE = 1, $LOW_BREAK_COLUMN_AFTER = 2, $LOW_BREAK_COLUMN_BOTH = 3, _
		$LOW_BREAK_PAGE_BEFORE = 4, $LOW_BREAK_PAGE_AFTER = 5, $LOW_BREAK_PAGE_BOTH = 6

;Horizontal Orientation
Global Const $LOW_ORIENT_HORI_NONE = 0, $LOW_ORIENT_HORI_RIGHT = 1, $LOW_ORIENT_HORI_CENTER = 2, $LOW_ORIENT_HORI_LEFT = 3, _
		$LOW_ORIENT_HORI_FULL = 6, $LOW_ORIENT_HORI_LEFT_AND_WIDTH = 7

;Color in Long Color Format
Global Const $LOW_COLOR_OFF = -1, $LOW_COLOR_BLACK = 0, $LOW_COLOR_WHITE = 16777215, $LOW_COLOR_LGRAY = 11711154, $LOW_COLOR_GRAY = 8421504, _
		$LOW_COLOR_DKGRAY = 3355443, $LOW_COLOR_YELLOW = 16776960, $LOW_COLOR_GOLD = 16760576, $LOW_COLOR_ORANGE = 16744448, _
		$LOW_COLOR_BRICK = 16728064, $LOW_COLOR_RED = 16711680, $LOW_COLOR_MAGENTA = 12517441, $LOW_COLOR_PURPLE = 8388736, _
		$LOW_COLOR_INDIGO = 5582989, $LOW_COLOR_BLUE = 2777241, $LOW_COLOR_TEAL = 1410150, $LOW_COLOR_GREEN = 43315, $LOW_COLOR_LIME = 8508442, _
		$LOW_COLOR_BROWN = 9127187

;Border Style
Global Const $LOW_BORDERSTYLE_NONE = 0x7FFF, $LOW_BORDERSTYLE_SOLID = 0, $LOW_BORDERSTYLE_DOTTED = 1, $LOW_BORDERSTYLE_DASHED = 2, _
		$LOW_BORDERSTYLE_DOUBLE = 3, $LOW_BORDERSTYLE_THINTHICK_SMALLGAP = 4, $LOW_BORDERSTYLE_THINTHICK_MEDIUMGAP = 5, _
		$LOW_BORDERSTYLE_THINTHICK_LARGEGAP = 6, $LOW_BORDERSTYLE_THICKTHIN_SMALLGAP = 7, $LOW_BORDERSTYLE_THICKTHIN_MEDIUMGAP = 8, _
		$LOW_BORDERSTYLE_THICKTHIN_LARGEGAP = 9, $LOW_BORDERSTYLE_EMBOSSED = 10, $LOW_BORDERSTYLE_ENGRAVED = 11, $LOW_BORDERSTYLE_OUTSET = 12, _
		$LOW_BORDERSTYLE_INSET = 13, $LOW_BORDERSTYLE_FINE_DASHED = 14, $LOW_BORDERSTYLE_DOUBLE_THIN = 15, $LOW_BORDERSTYLE_DASH_DOT = 16, _
		$LOW_BORDERSTYLE_DASH_DOT_DOT = 17

;Border Width
Global Const $LOW_BORDERWIDTH_HAIRLINE = 2, $LOW_BORDERWIDTH_VERY_THIN = 18, $LOW_BORDERWIDTH_THIN = 26, $LOW_BORDERWIDTH_MEDIUM = 53, _
		$LOW_BORDERWIDTH_THICK = 79, $LOW_BORDERWIDTH_EXTRA_THICK = 159

;Vertical Orientation
Global Const $LOW_ORIENT_VERT_NONE = 0, $LOW_ORIENT_VERT_TOP = 1, $LOW_ORIENT_VERT_CENTER = 2, $LOW_ORIENT_VERT_BOTTOM = 3, _
		$LOW_ORIENT_VERT_CHAR_TOP = 4, $LOW_ORIENT_VERT_CHAR_CENTER = 5, $LOW_ORIENT_VERT_CHAR_BOTTOM = 6, $LOW_ORIENT_VERT_LINE_TOP = 7, _
		$LOW_ORIENT_VERT_LINE_CENTER = 8, $LOW_ORIENT_VERT_LINE_BOTTOM = 9

;Tab Alignment
Global Const $LOW_TAB_ALIGN_LEFT = 0, $LOW_TAB_ALIGN_CENTER = 1, $LOW_TAB_ALIGN_RIGHT = 2, $LOW_TAB_ALIGN_DECIMAL = 3, _
		$LOW_TAB_ALIGN_DEFAULT = 4

;Underline/Overline
Global Const $LOW_UNDERLINE_NONE = 0, $LOW_UNDERLINE_SINGLE = 1, $LOW_UNDERLINE_DOUBLE = 2, $LOW_UNDERLINE_DOTTED = 3, _
		$LOW_UNDERLINE_DONT_KNOW = 4, $LOW_UNDERLINE_DASH = 5, $LOW_UNDERLINE_LONG_DASH = 6, $LOW_UNDERLINE_DASH_DOT = 7, _
		$LOW_UNDERLINE_DASH_DOT_DOT = 8, $LOW_UNDERLINE_SML_WAVE = 9, $LOW_UNDERLINE_WAVE = 10, $LOW_UNDERLINE_DBL_WAVE = 11, $LOW_UNDERLINE_BOLD = 12, _
		$LOW_UNDERLINE_BOLD_DOTTED = 13, $LOW_UNDERLINE_BOLD_DASH = 14, $LOW_UNDERLINE_BOLD_LONG_DASH = 15, $LOW_UNDERLINE_BOLD_DASH_DOT = 16, _
		$LOW_UNDERLINE_BOLD_DASH_DOT_DOT = 17, $LOW_UNDERLINE_BOLD_WAVE = 18

;Strikeout
Global Const $LOW_STRIKEOUT_NONE = 0, $LOW_STRIKEOUT_SINGLE = 1, $LOW_STRIKEOUT_DOUBLE = 2, $LOW_STRIKEOUT_DONT_KNOW = 3, _
		$LOW_STRIKEOUT_BOLD = 4, $LOW_STRIKEOUT_SLASH = 5, $LOW_STRIKEOUT_X = 6

;Relief
Global Const $LOW_RELIEF_NONE = 0, $LOW_RELIEF_EMBOSSED = 1, $LOW_RELIEF_ENGRAVED = 2

;Case
Global Const $LOW_CASEMAP_NONE = 0, $LOW_CASEMAP_UPPER = 1, $LOW_CASEMAP_LOWER = 2, $LOW_CASEMAP_TITLE = 3, $LOW_CASEMAP_SM_CAPS = 4

;Shadow
Global Const $LOW_SHADOW_NONE = 0, $LOW_SHADOW_TOP_LEFT = 1, $LOW_SHADOW_TOP_RIGHT = 2, $LOW_SHADOW_BOTTOM_LEFT = 3, _
		$LOW_SHADOW_BOTTOM_RIGHT = 4

;Posture/Italic
Global Const $LOW_POSTURE_NONE = 0, $LOW_POSTURE_OBLIQUE = 1, $LOW_POSTURE_ITALIC = 2, $LOW_POSTURE_DontKnow = 3, _
		$LOW_POSTURE_REV_OBLIQUE = 4, $LOW_POSTURE_REV_ITALIC = 5

;Weight/Bold
Global Const $LOW_WEIGHT_DONT_KNOW = 0, $LOW_WEIGHT_THIN = 50, $LOW_WEIGHT_ULTRA_LIGHT = 60, $LOW_WEIGHT_LIGHT = 75, $LOW_WEIGHT_SEMI_LIGHT = 90, _
		$LOW_WEIGHT_NORMAL = 100, $LOW_WEIGHT_SEMI_BOLD = 110, $LOW_WEIGHT_BOLD = 150, $LOW_WEIGHT_ULTRA_BOLD = 175, $LOW_WEIGHT_BLACK = 200

;Outline
Global Const $LOW_OUTLINE_BODY = 0, $LOW_OUTLINE_LEVEL_1 = 1, $LOW_OUTLINE_LEVEL_2 = 2, $LOW_OUTLINE_LEVEL_3 = 3, $LOW_OUTLINE_LEVEL_4 = 4, _
		$LOW_OUTLINE_LEVEL_5 = 5, $LOW_OUTLINE_LEVEL_6 = 6, $LOW_OUTLINE_LEVEL_7 = 7, $LOW_OUTLINE_LEVEL_8 = 8, $LOW_OUTLINE_LEVEL_9 = 9, _
		$LOW_OUTLINE_LEVEL_10 = 10

;Line Spacing
Global Const $LOW_LINE_SPC_MODE_PROP = 0, $LOW_LINE_SPC_MODE_MIN = 1, $LOW_LINE_SPC_MODE_LEADING = 2, $LOW_LINE_SPC_MODE_FIX = 3

;Paragraph Horizontal Align
Global Const $LOW_PAR_ALIGN_HOR_LEFT = 0, $LOW_PAR_ALIGN_HOR_RIGHT = 1, $LOW_PAR_ALIGN_HOR_JUSTIFIED = 2, $LOW_PAR_ALIGN_HOR_CENTER = 3, _
		$LOW_PAR_ALIGN_HOR_STRETCH = 4 ;HoriAlign 4 does nothing??

;Paragraph Vertical Align
Global Const $LOW_PAR_ALIGN_VERT_AUTO = 0, $LOW_PAR_ALIGN_VERT_BASELINE = 1, $LOW_PAR_ALIGN_VERT_TOP = 2, $LOW_PAR_ALIGN_VERT_CENTER = 3, _
		$LOW_PAR_ALIGN_VERT_BOTTOM = 4

;Paragraph Last Line Alignment
Global Const $LOW_PAR_LAST_LINE_START = 0, $LOW_PAR_LAST_LINE_JUSTIFIED = 2, $LOW_PAR_LAST_LINE_CENTER = 3

;Text Direction
Global Const $LOW_TXT_DIR_LR_TB = 0, $LOW_TXT_DIR_RL_TB = 1, $LOW_TXT_DIR_TB_RL = 2, $LOW_TXT_DIR_TB_LR = 3, $LOW_TXT_DIR_CONTEXT = 4, _
		$LOW_TXT_DIR_BT_LR = 5

;Control Character
Global Const $LOW_CON_CHAR_PAR_BREAK = 0, $LOW_CON_CHAR_LINE_BREAK = 1, $LOW_CON_CHAR_HARD_HYPHEN = 2, $LOW_CON_CHAR_SOFT_HYPHEN = 3, _
		$LOW_CON_CHAR_HARD_SPACE = 4, $LOW_CON_CHAR_APPEND_PAR = 5

;Cell Type
Global Const $LOW_CELL_TYPE_EMPTY = 0, $LOW_CELL_TYPE_VALUE = 1, $LOW_CELL_TYPE_TEXT = 2, $LOW_CELL_TYPE_FORMULA = 3

;Paper Width in uM
Global Const $LOW_PAPER_WIDTH_A6 = 10490, $LOW_PAPER_WIDTH_A5 = 14808, $LOW_PAPER_WIDTH_A4 = 21006, $LOW_PAPER_WIDTH_A3 = 29693, _
		$LOW_PAPER_WIDTH_B6ISO = 12497, $LOW_PAPER_WIDTH_B5ISO = 17602, $LOW_PAPER_WIDTH_B4ISO = 24994, $LOW_PAPER_WIDTH_LETTER = 21590, _
		$LOW_PAPER_WIDTH_LEGAL = 21590, $LOW_PAPER_WIDTH_LONG_BOND = 21590, $LOW_PAPER_WIDTH_TABLOID = 27940, $LOW_PAPER_WIDTH_B6JIS = 12801, _
		$LOW_PAPER_WIDTH_B5JIS = 18212, $LOW_PAPER_WIDTH_B4JIS = 25705, $LOW_PAPER_WIDTH_16KAI = 18390, $LOW_PAPER_WIDTH_32KAI = 13005, _
		$LOW_PAPER_WIDTH_BIG_32KAI = 13995, $LOW_PAPER_WIDTH_DLENVELOPE = 10998, $LOW_PAPER_WIDTH_C6ENVELOPE = 11405, _
		$LOW_PAPER_WIDTH_C6_5_ENVELOPE = 11405, $LOW_PAPER_WIDTH_C5ENVELOPE = 16205, $LOW_PAPER_WIDTH_C4ENVELOPE = 22911, _
		$LOW_PAPER_WIDTH_6_3_4ENVELOPE = 9208, $LOW_PAPER_WIDTH_7_3_4ENVELOPE = 9855, $LOW_PAPER_WIDTH_9ENVELOPE = 9843, _
		$LOW_PAPER_WIDTH_10ENVELOPE = 10490, $LOW_PAPER_WIDTH_11ENVELOPE = 11430, $LOW_PAPER_WIDTH_12ENVELOPE = 12065, _
		$LOW_PAPER_WIDTH_JAP_POSTCARD = 10008

;Paper Height in uM
Global Const $LOW_PAPER_HEIGHT_A6 = 14808, $LOW_PAPER_HEIGHT_A5 = 21006, $LOW_PAPER_HEIGHT_A4 = 29693, $LOW_PAPER_HEIGHT_A3 = 42012, _
		$LOW_PAPER_HEIGHT_B6ISO = 17602, $LOW_PAPER_HEIGHT_B5ISO = 24994, $LOW_PAPER_HEIGHT_B4ISO = 35306, $LOW_PAPER_HEIGHT_LETTER = 27940, _
		$LOW_PAPER_HEIGHT_LEGAL = 35560, $LOW_PAPER_HEIGHT_LONG_BOND = 33020, $LOW_PAPER_HEIGHT_TABLOID = 43180, $LOW_PAPER_HEIGHT_B6JIS = 18200, _
		$LOW_PAPER_HEIGHT_B5JIS = 25705, $LOW_PAPER_HEIGHT_B4JIS = 36398, $LOW_PAPER_HEIGHT_16KAI = 26010, $LOW_PAPER_HEIGHT_32KAI = 18390, _
		$LOW_PAPER_HEIGHT_BIG_32KAI = 20295, $LOW_PAPER_HEIGHT_DLENVELOPE = 21996, $LOW_PAPER_HEIGHT_C6ENVELOPE = 16205, _
		$LOW_PAPER_HEIGHT_C6_5_ENVELOPE = 22911, $LOW_PAPER_HEIGHT_C5ENVELOPE = 22911, $LOW_PAPER_HEIGHT_C4ENVELOPE = 32410, _
		$LOW_PAPER_HEIGHT_6_3_4ENVELOPE = 16510, $LOW_PAPER_HEIGHT_7_3_4ENVELOPE = 19050, $LOW_PAPER_HEIGHT_9ENVELOPE = 22543, _
		$LOW_PAPER_HEIGHT_10ENVELOPE = 24130, $LOW_PAPER_HEIGHT_11ENVELOPE = 26365, $LOW_PAPER_HEIGHT_12ENVELOPE = 27940, _
		$LOW_PAPER_HEIGHT_JAP_POSTCARD = 14808

;Gradient Names
Global Const $LOW_GRAD_NAME_PASTEL_BOUQUET = "Pastel Bouquet", $LOW_GRAD_NAME_PASTEL_DREAM = "Pastel Dream", _
		$LOW_GRAD_NAME_BLUE_TOUCH = "Blue Touch", $LOW_GRAD_NAME_BLANK_W_GRAY = "Blank with Gray", $LOW_GRAD_NAME_SPOTTED_GRAY = "Spotted Gray", _
		$LOW_GRAD_NAME_LONDON_MIST = "London Mist", $LOW_GRAD_NAME_TEAL_TO_BLUE = "Teal to Blue", $LOW_GRAD_NAME_MIDNIGHT = "Midnight", _
		$LOW_GRAD_NAME_DEEP_OCEAN = "Deep Ocean", $LOW_GRAD_NAME_SUBMARINE = "Submarine", $LOW_GRAD_NAME_GREEN_GRASS = "Green Grass", _
		$LOW_GRAD_NAME_NEON_LIGHT = "Neon Light", $LOW_GRAD_NAME_SUNSHINE = "Sunshine", $LOW_GRAD_NAME_PRESENT = "Present", _
		$LOW_GRAD_NAME_MAHOGANY = "Mahogany"

;Page Layout
Global Const $LOW_PAGE_LAYOUT_ALL = 0, $LOW_PAGE_LAYOUT_LEFT = 1, $LOW_PAGE_LAYOUT_RIGHT = 2, $LOW_PAGE_LAYOUT_MIRRORED = 3

;Numbering Style Type
Global Const $LOW_NUM_STYLE_CHARS_UPPER_LETTER = 0, $LOW_NUM_STYLE_CHARS_LOWER_LETTER = 1, $LOW_NUM_STYLE_ROMAN_UPPER = 2, _
		$LOW_NUM_STYLE_ROMAN_LOWER = 3, $LOW_NUM_STYLE_ARABIC = 4, $LOW_NUM_STYLE_NUMBER_NONE = 5, $LOW_NUM_STYLE_CHAR_SPECIAL = 6, _
		$LOW_NUM_STYLE_PAGE_DESCRIPTOR = 7, $LOW_NUM_STYLE_BITMAP = 8, $LOW_NUM_STYLE_CHARS_UPPER_LETTER_N = 9, _
		$LOW_NUM_STYLE_CHARS_LOWER_LETTER_N = 10, $LOW_NUM_STYLE_TRANSLITERATION = 11, $LOW_NUM_STYLE_NATIVE_NUMBERING = 12, _
		$LOW_NUM_STYLE_FULLWIDTH_ARABIC = 13, $LOW_NUM_STYLE_CIRCLE_NUMBER = 14, $LOW_NUM_STYLE_NUMBER_LOWER_ZH = 15, _
		$LOW_NUM_STYLE_NUMBER_UPPER_ZH = 16, $LOW_NUM_STYLE_NUMBER_UPPER_ZH_TW = 17, $LOW_NUM_STYLE_TIAN_GAN_ZH = 18, _
		$LOW_NUM_STYLE_DI_ZI_ZH = 19, $LOW_NUM_STYLE_NUMBER_TRADITIONAL_JA = 20, $LOW_NUM_STYLE_AIU_FULLWIDTH_JA = 21, _
		$LOW_NUM_STYLE_AIU_HALFWIDTH_JA = 22, $LOW_NUM_STYLE_IROHA_FULLWIDTH_JA = 23, $LOW_NUM_STYLE_IROHA_HALFWIDTH_JA = 24, _
		$LOW_NUM_STYLE_NUMBER_UPPER_KO = 25, $LOW_NUM_STYLE_NUMBER_HANGUL_KO = 26, $LOW_NUM_STYLE_HANGUL_JAMO_KO = 27, _
		$LOW_NUM_STYLE_HANGUL_SYLLABLE_KO = 28, $LOW_NUM_STYLE_HANGUL_CIRCLED_JAMO_KO = 29, $LOW_NUM_STYLE_HANGUL_CIRCLED_SYLLABLE_KO = 30, _
		$LOW_NUM_STYLE_CHARS_ARABIC = 31, $LOW_NUM_STYLE_CHARS_THAI = 32, $LOW_NUM_STYLE_CHARS_HEBREW = 33, $LOW_NUM_STYLE_CHARS_NEPALI = 34, _
		$LOW_NUM_STYLE_CHARS_KHMER = 35, $LOW_NUM_STYLE_CHARS_LAO = 36, $LOW_NUM_STYLE_CHARS_TIBETAN = 37, _
		$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_BG = 38, $LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_BG = 39, _
		$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_N_BG = 40, $LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_N_BG = 41, _
		$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_RU = 42, $LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_RU = 43, _
		$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_N_RU = 44, $LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_N_RU = 45, _
		$LOW_NUM_STYLE_CHARS_PERSIAN = 46, $LOW_NUM_STYLE_CHARS_MYANMAR = 47, $LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_SR = 48, _
		$LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_SR = 49, $LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_N_SR = 50, _
		$LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_N_SR = 51, $LOW_NUM_STYLE_CHARS_GREEK_UPPER_LETTER = 52, _
		$LOW_NUM_STYLE_CHARS_GREEK_LOWER_LETTER = 53, $LOW_NUM_STYLE_CHARS_ARABIC_ABJAD = 54, $LOW_NUM_STYLE_CHARS_PERSIAN_WORD = 55, _
		$LOW_NUM_STYLE_NUMBER_HEBREW = 56, $LOW_NUM_STYLE_NUMBER_ARABIC_INDIC = 57, $LOW_NUM_STYLE_NUMBER_EAST_ARABIC_INDIC = 58, _
		$LOW_NUM_STYLE_NUMBER_INDIC_DEVANAGARI = 59, $LOW_NUM_STYLE_TEXT_NUMBER = 60, $LOW_NUM_STYLE_TEXT_CARDINAL = 61, _
		$LOW_NUM_STYLE_TEXT_ORDINAL = 62, $LOW_NUM_STYLE_SYMBOL_CHICAGO = 63, $LOW_NUM_STYLE_ARABIC_ZERO = 64, $LOW_NUM_STYLE_ARABIC_ZERO3 = 65, _
		$LOW_NUM_STYLE_ARABIC_ZERO4 = 66, $LOW_NUM_STYLE_ARABIC_ZERO5 = 67, $LOW_NUM_STYLE_SZEKELY_ROVAS = 68, _
		$LOW_NUM_STYLE_NUMBER_DIGITAL_KO = 69, $LOW_NUM_STYLE_NUMBER_DIGITAL2_KO = 70, $LOW_NUM_STYLE_NUMBER_LEGAL_KO = 71

;Line Style
Global Const $LOW_LINE_STYLE_NONE = 0, $LOW_LINE_STYLE_SOLID = 1, $LOW_LINE_STYLE_DOTTED = 2, $LOW_LINE_STYLE_DASHED = 3

;Vertical Alignment
Global Const $LOW_ALIGN_VERT_TOP = 0, $LOW_ALIGN_VERT_MIDDLE = 1, $LOW_ALIGN_VERT_BOTTOM = 2

;Horizontal Alignment
Global Const $LOW_ALIGN_HORI_LEFT = 0, $LOW_ALIGN_HORI_CENTER = 1, $LOW_ALIGN_HORI_RIGHT = 2

;Gradient Type
Global Const $LOW_GRAD_TYPE_OFF = -1, $LOW_GRAD_TYPE_LINEAR = 0, $LOW_GRAD_TYPE_AXIAL = 1, $LOW_GRAD_TYPE_RADIAL = 2, _
		$LOW_GRAD_TYPE_ELLIPTICAL = 3, $LOW_GRAD_TYPE_SQUARE = 4, $LOW_GRAD_TYPE_RECT = 5

;Follow By
Global Const $LOW_FOLLOW_BY_TABSTOP = 0, $LOW_FOLLOW_BY_SPACE = 1, $LOW_FOLLOW_BY_NOTHING = 2, $LOW_FOLLOW_BY_NEWLINE = 3

;Cursor Status
Global Enum $LOW_CURSOR_STAT_IS_COLLAPSED, $LOW_CURSOR_STAT_IS_START_OF_WORD, $LOW_CURSOR_STAT_IS_END_OF_WORD, _
		$LOW_CURSOR_STAT_IS_START_OF_SENTENCE, $LOW_CURSOR_STAT_IS_END_OF_SENTENCE, $LOW_CURSOR_STAT_IS_START_OF_PAR, _
		$LOW_CURSOR_STAT_IS_END_OF_PAR, $LOW_CURSOR_STAT_IS_START_OF_LINE, $LOW_CURSOR_STAT_IS_END_OF_LINE, $LOW_CURSOR_STAT_GET_PAGE, _
		$LOW_CURSOR_STAT_GET_RANGE_NAME

;Relative to
Global Const $LOW_RELATIVE_ROW = -1, $LOW_RELATIVE_PARAGRAPH = 0, $LOW_RELATIVE_PARAGRAPH_TEXT = 1, _
		$LOW_RELATIVE_CHARACTER = 2, $LOW_RELATIVE_PAGE_LEFT = 3, $LOW_RELATIVE_PAGE_RIGHT = 4, $LOW_RELATIVE_PARAGRAPH_LEFT = 5, _
		$LOW_RELATIVE_PARAGRAPH_RIGHT = 6, $LOW_RELATIVE_PAGE = 7, $LOW_RELATIVE_PAGE_PRINT = 8, $LOW_RELATIVE_TEXT_LINE = 9, _
		$LOW_RELATIVE_PAGE_PRINT_BOTTOM = 10, $LOW_RELATIVE_PAGE_PRINT_TOP = 11

;Anchor Type
Global Const $LOW_ANCHOR_AT_PARAGRAPH = 0, $LOW_ANCHOR_AS_CHARACTER = 1, $LOW_ANCHOR_AT_PAGE = 2, $LOW_ANCHOR_AT_FRAME = 3, _
		$LOW_ANCHOR_AT_CHARACTER = 4

;Wrap Type
Global Const $LOW_WRAP_MODE_NONE = 0, $LOW_WRAP_MODE_THROUGH = 1, $LOW_WRAP_MODE_PARALLEL = 2, $LOW_WRAP_MODE_DYNAMIC = 3, _
		$LOW_WRAP_MODE_LEFT = 4, $LOW_WRAP_MODE_RIGHT = 5

;Text Adjust
Global Const $LOW_TXT_ADJ_VERT_TOP = 0, $LOW_TXT_ADJ_VERT_CENTER = 1, $LOW_TXT_ADJ_VERT_BOTTOM = 2, $LOW_TXT_ADJ_VERT_BLOCK = 3

;Frame Target
Global Const $LOW_FRAME_TARGET_NONE = "", $LOW_FRAME_TARGET_TOP = "_top", $LOW_FRAME_TARGET_PARENT = "_parent", _
		$LOW_FRAME_TARGET_BLANK = "_blank", $LOW_FRAME_TARGET_SELF = "_self"

;Footnote Count type
Global Const $LOW_FOOTNOTE_COUNT_PER_PAGE = 0, $LOW_FOOTNOTE_COUNT_PER_CHAP = 1, $LOW_FOOTNOTE_COUNT_PER_DOC = 2

;Page Number Type
Global Const $LOW_PAGE_NUM_TYPE_PREV = 0, $LOW_PAGE_NUM_TYPE_CURRENT = 1, $LOW_PAGE_NUM_TYPE_NEXT = 2

;Field Chapter Display Type
Global Const $LOW_FIELD_CHAP_FRMT_NAME = 0, $LOW_FIELD_CHAP_FRMT_NUMBER = 1, $LOW_FIELD_CHAP_FRMT_NAME_NUMBER = 2, _
		$LOW_FIELD_CHAP_FRMT_NO_PREFIX_SUFFIX = 3, $LOW_FIELD_CHAP_FRMT_DIGIT = 4

;User Data Field Type
Global Const $LOW_FIELD_USER_DATA_COMPANY = 0, $LOW_FIELD_USER_DATA_FIRST_NAME = 1, $LOW_FIELD_USER_DATA_NAME = 2, _
		$LOW_FIELD_USER_DATA_SHORTCUT = 3, $LOW_FIELD_USER_DATA_STREET = 4, $LOW_FIELD_USER_DATA_COUNTRY = 5, $LOW_FIELD_USER_DATA_ZIP = 6, _
		$LOW_FIELD_USER_DATA_CITY = 7, $LOW_FIELD_USER_DATA_TITLE = 8, $LOW_FIELD_USER_DATA_POSITION = 9, $LOW_FIELD_USER_DATA_PHONE_PRIVATE = 10, _
		$LOW_FIELD_USER_DATA_PHONE_COMPANY = 11, $LOW_FIELD_USER_DATA_FAX = 12, $LOW_FIELD_USER_DATA_EMAIL = 13, $LOW_FIELD_USER_DATA_STATE = 14

;File Name Field Type
Global Const $LOW_FIELD_FILENAME_FULL_PATH = 0, $LOW_FIELD_FILENAME_PATH = 1, $LOW_FIELD_FILENAME_NAME = 2, $LOW_FIELD_FILENAME_NAME_AND_EXT = 3, _
		$LOW_FIELD_FILENAME_CATEGORY = 4, $LOW_FIELD_FILENAME_TEMPLATE_NAME = 5

;Format Key Type
Global Const $LOW_FORMAT_KEYS_ALL = 0, $LOW_FORMAT_KEYS_DEFINED = 1, $LOW_FORMAT_KEYS_DATE = 2, $LOW_FORMAT_KEYS_TIME = 4, _
		$LOW_FORMAT_KEYS_DATE_TIME = 6, $LOW_FORMAT_KEYS_CURRENCY = 8, $LOW_FORMAT_KEYS_NUMBER = 16, $LOW_FORMAT_KEYS_SCIENTIFIC = 32, _
		$LOW_FORMAT_KEYS_FRACTION = 64, $LOW_FORMAT_KEYS_PERCENT = 128, $LOW_FORMAT_KEYS_TEXT = 256, $LOW_FORMAT_KEYS_LOGICAL = 1024, _
		$LOW_FORMAT_KEYS_UNDEFINED = 2048, $LOW_FORMAT_KEYS_EMPTY = 4096, $LOW_FORMAT_KEYS_DURATION = 8196

;Reference Field Type
Global Const $LOW_FIELD_REF_TYPE_REF_MARK = 0, $LOW_FIELD_REF_TYPE_SEQ_FIELD = 1, $LOW_FIELD_REF_TYPE_BOOKMARK = 2, _
		$LOW_FIELD_REF_TYPE_FOOTNOTE = 3, $LOW_FIELD_REF_TYPE_ENDNOTE = 4

;Type of Reference
Global Const $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED = 0, $LOW_FIELD_REF_USING_CHAPTER = 1, $LOW_FIELD_REF_USING_REF_TEXT = 2, _
		$LOW_FIELD_REF_USING_ABOVE_BELOW = 3, $LOW_FIELD_REF_USING_PAGE_NUM_STYLED = 4, $LOW_FIELD_REF_USING_CAT_AND_NUM = 5, _
		$LOW_FIELD_REF_USING_CAPTION = 6, $LOW_FIELD_REF_USING_NUMBERING = 7, $LOW_FIELD_REF_USING_NUMBER = 8, _
		$LOW_FIELD_REF_USING_NUMBER_NO_CONT = 9, $LOW_FIELD_REF_USING_NUMBER_CONT = 10

;Count Field Type
Global Enum $LOW_FIELD_COUNT_TYPE_CHARACTERS = 0, $LOW_FIELD_COUNT_TYPE_IMAGES, $LOW_FIELD_COUNT_TYPE_OBJECTS, _
		$LOW_FIELD_COUNT_TYPE_PAGES, $LOW_FIELD_COUNT_TYPE_PARAGRAPHS, $LOW_FIELD_COUNT_TYPE_TABLES, $LOW_FIELD_COUNT_TYPE_WORDS

;Regular Field Types
Global Enum Step *2 $LOW_FIELD_TYPE_ALL = 1, $LOW_FIELD_TYPE_COMMENT, $LOW_FIELD_TYPE_AUTHOR, $LOW_FIELD_TYPE_CHAPTER, _
		$LOW_FIELD_TYPE_CHAR_COUNT, $LOW_FIELD_TYPE_COMBINED_CHAR, $LOW_FIELD_TYPE_COND_TEXT, $LOW_FIELD_TYPE_DATE_TIME, _
		$LOW_FIELD_TYPE_INPUT_LIST, $LOW_FIELD_TYPE_EMB_OBJ_COUNT, $LOW_FIELD_TYPE_SENDER, $LOW_FIELD_TYPE_FILENAME, $LOW_FIELD_TYPE_SHOW_VAR, _
		$LOW_FIELD_TYPE_INSERT_REF, $LOW_FIELD_TYPE_IMAGE_COUNT, $LOW_FIELD_TYPE_HIDDEN_PAR, $LOW_FIELD_TYPE_HIDDEN_TEXT, $LOW_FIELD_TYPE_INPUT, _
		$LOW_FIELD_TYPE_PLACEHOLDER, $LOW_FIELD_TYPE_MACRO, $LOW_FIELD_TYPE_PAGE_COUNT, $LOW_FIELD_TYPE_PAGE_NUM, $LOW_FIELD_TYPE_PAR_COUNT, _
		$LOW_FIELD_TYPE_SHOW_PAGE_VAR, $LOW_FIELD_TYPE_SET_PAGE_VAR, $LOW_FIELD_TYPE_SCRIPT, $LOW_FIELD_TYPE_SET_VAR, _
		$LOW_FIELD_TYPE_TABLE_COUNT, $LOW_FIELD_TYPE_TEMPLATE_NAME, $LOW_FIELD_TYPE_URL, $LOW_FIELD_TYPE_WORD_COUNT

;Advanced Field Types
Global Enum Step *2 $LOW_FIELDADV_TYPE_ALL = 1, $LOW_FIELDADV_TYPE_BIBLIOGRAPHY, $LOW_FIELDADV_TYPE_DATABASE, _
		$LOW_FIELDADV_TYPE_DATABASE_SET_NUM, $LOW_FIELDADV_TYPE_DATABASE_NAME, $LOW_FIELDADV_TYPE_DATABASE_NEXT_SET, _
		$LOW_FIELDADV_TYPE_DATABASE_NAME_OF_SET, $LOW_FIELDADV_TYPE_DDE, $LOW_FIELDADV_TYPE_INPUT_USER, $LOW_FIELDADV_TYPE_USER

;Document Information Field Types
Global Enum Step *2 $LOW_FIELD_DOCINFO_TYPE_ALL = 1, $LOW_FIELD_DOCINFO_TYPE_MOD_AUTH, $LOW_FIELD_DOCINFO_TYPE_MOD_DATE_TIME, _
		$LOW_FIELD_DOCINFO_TYPE_CREATE_AUTH, $LOW_FIELD_DOCINFO_TYPE_CREATE_DATE_TIME, $LOW_FIELD_DOCINFO_TYPE_CUSTOM, _
		$LOW_FIELD_DOCINFO_TYPE_COMMENTS, $LOW_FIELD_DOCINFO_TYPE_EDIT_TIME, $LOW_FIELD_DOCINFO_TYPE_KEYWORDS, _
		$LOW_FIELD_DOCINFO_TYPE_PRINT_AUTH, $LOW_FIELD_DOCINFO_TYPE_PRINT_DATE_TIME, $LOW_FIELD_DOCINFO_TYPE_REVISION, _
		$LOW_FIELD_DOCINFO_TYPE_SUBJECT, $LOW_FIELD_DOCINFO_TYPE_TITLE

;Placeholder Type
Global Const $LOW_FIELD_PLACEHOLD_TYPE_TEXT = 0, $LOW_FIELD_PLACEHOLD_TYPE_TABLE = 1, $LOW_FIELD_PLACEHOLD_TYPE_FRAME = 2, _
		$LOW_FIELD_PLACEHOLD_TYPE_GRAPHIC = 3, $LOW_FIELD_PLACEHOLD_TYPE_OBJECT = 4