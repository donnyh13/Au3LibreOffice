#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7
#include-once

; #INDEX# =======================================================================================================================
; Title .........: Libre Office Writer Constants for the Libre Office UDF.
; AutoIt Version : v3.3.16.1
; Description ...: Constants for various functions in the Libre Office UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; ===============================================================================================================================

; Sleep Divisor $__LOWCONST_SLEEP_DIV
; In applicable functions this is used for adjusting how frequent a sleep occurs in loops.
; For any number above 0 the number of times a loop has completed is divided by $__LOWCONST_SLEEP_DIV. If you find some functions cause momentary freeze ups, a recommended value is 15.
; Set to 0 for no pause in a loop.
Global Const $__LOWCONST_SLEEP_DIV = 0

#Tidy_ILC_Pos=65

; Fill Style Type Constants
Global Enum _
		$__LOWCONST_FILL_STYLE_OFF, _                           ; 0 Fillstyle is off.
		$__LOWCONST_FILL_STYLE_SOLID, _                         ; 1 Fillstyle is a solid color.
		$__LOWCONST_FILL_STYLE_GRADIENT, _                      ; 2 Fillstyle is a gradient color.
		$__LOWCONST_FILL_STYLE_HATCH, _                         ; 3 Fillstyle is a Hatch style color.
		$__LOWCONST_FILL_STYLE_BITMAP                           ; 4 Fillstyle is a Bitmap.

; Horizontal Alignment
Global Const _
		$LOW_ALIGN_HORI_LEFT = 0, _                             ; Horizontally align the object to the Left.
		$LOW_ALIGN_HORI_CENTER = 1, _                           ; Horizontally align the object to the Center.
		$LOW_ALIGN_HORI_RIGHT = 2                               ; Horizontally align the object to the Right.

; Vertical Alignment
Global Const _
		$LOW_ALIGN_VERT_TOP = 0, _                              ; Vertically Align the object to the Top.
		$LOW_ALIGN_VERT_MIDDLE = 1, _                           ; Vertically Align the object to the Middle.
		$LOW_ALIGN_VERT_BOTTOM = 2                              ; Vertically Align the object to the Bottom.

; Anchor Type
Global Const _
		$LOW_ANCHOR_AT_PARAGRAPH = 0, _                         ; Anchors the object to the current paragraph.
		$LOW_ANCHOR_AS_CHARACTER = 1, _                         ; Anchors the Object as character. The height of the current line is resized to match the height of the selection.
		$LOW_ANCHOR_AT_PAGE = 2, _                              ; Anchors the Object to the current page.
		$LOW_ANCHOR_AT_FRAME = 3, _                             ; Anchors the object to the surrounding frame.
		$LOW_ANCHOR_AT_CHARACTER = 4                            ; Anchors the Object to a character.

; Break Type
Global Const _
		$LOW_BREAK_NONE = 0, _                                  ; No column or page break is applied.
		$LOW_BREAK_COLUMN_BEFORE = 1, _                         ; A column break is applied before the current Paragraph. The current Paragraph, therefore, is the first in the column.
		$LOW_BREAK_COLUMN_AFTER = 2, _                          ; A column break is applied after the current Paragraph. The current Paragraph, therefore, is the last in the column.
		$LOW_BREAK_COLUMN_BOTH = 3, _                           ; A column break is applied before and after the current Paragraph. The current Paragraph, therefore, is the only Paragraph in the column.
		$LOW_BREAK_PAGE_BEFORE = 4, _                           ; A page break is applied before the current Paragraph. The current Paragraph, therefore, is the first on the page.
		$LOW_BREAK_PAGE_AFTER = 5, _                            ; A page break is applied after the current Paragraph. The current Paragraph, therefore, is the last on the page.
		$LOW_BREAK_PAGE_BOTH = 6                                ; A page break is applied before and after the current Paragraph. The current Paragraph, therefore, is the only paragraph on the page.

; Border Style
Global Const _
		$LOW_BORDERSTYLE_NONE = 0x7FFF, _                       ; No border line.
		$LOW_BORDERSTYLE_SOLID = 0, _                           ; Solid border line.
		$LOW_BORDERSTYLE_DOTTED = 1, _                          ; Dotted border line.
		$LOW_BORDERSTYLE_DASHED = 2, _                          ; Dashed border line.
		$LOW_BORDERSTYLE_DOUBLE = 3, _                          ; Double border line.
		$LOW_BORDERSTYLE_THINTHICK_SMALLGAP = 4, _              ; Double border line with a thin line outside and a thick line inside separated by a small gap.
		$LOW_BORDERSTYLE_THINTHICK_MEDIUMGAP = 5, _             ; Double border line with a thin line outside and a thick line inside separated by a medium gap.
		$LOW_BORDERSTYLE_THINTHICK_LARGEGAP = 6, _              ; Double border line with a thin line outside and a thick line inside separated by a large gap.
		$LOW_BORDERSTYLE_THICKTHIN_SMALLGAP = 7, _              ; Double border line with a thick line outside and a thin line inside separated by a small gap.
		$LOW_BORDERSTYLE_THICKTHIN_MEDIUMGAP = 8, _             ; Double border line with a thick line outside and a thin line inside separated by a medium gap.
		$LOW_BORDERSTYLE_THICKTHIN_LARGEGAP = 9, _              ; Double border line with a thick line outside and a thin line inside separated by a large gap.
		$LOW_BORDERSTYLE_EMBOSSED = 10, _                       ; 3D embossed border line.
		$LOW_BORDERSTYLE_ENGRAVED = 11, _                       ; 3D engraved border line.
		$LOW_BORDERSTYLE_OUTSET = 12, _                         ; Outset border line.
		$LOW_BORDERSTYLE_INSET = 13, _                          ; Inset border line.
		$LOW_BORDERSTYLE_FINE_DASHED = 14, _                    ; Finely dashed border line.
		$LOW_BORDERSTYLE_DOUBLE_THIN = 15, _                    ; Double border line consisting of two fixed thin lines separated by a variable gap.
		$LOW_BORDERSTYLE_DASH_DOT = 16, _                       ; Line consisting of a repetition of one dash and one dot.
		$LOW_BORDERSTYLE_DASH_DOT_DOT = 17                      ; Line consisting of a repetition of one dash and 2 dots.

; Border Width
Global Const _
		$LOW_BORDERWIDTH_HAIRLINE = 2, _                        ; Hairline Border line width.
		$LOW_BORDERWIDTH_VERY_THIN = 18, _                      ; Very Thin Border line width.
		$LOW_BORDERWIDTH_THIN = 26, _                           ; Thin Border line width.
		$LOW_BORDERWIDTH_MEDIUM = 53, _                         ; Medium Border line width.
		$LOW_BORDERWIDTH_THICK = 79, _                          ; Thick Border line width.
		$LOW_BORDERWIDTH_EXTRA_THICK = 159                      ; Extra Thick Border line width.

; Case Constants
Global Const _
		$LOW_CASEMAP_NONE = 0, _                                ; The case of the characters is unchanged.
		$LOW_CASEMAP_UPPER = 1, _                               ; All characters are put in upper case.
		$LOW_CASEMAP_LOWER = 2, _                               ; All characters are put in lower case.
		$LOW_CASEMAP_TITLE = 3, _                               ; The first character of each word is put in upper case.
		$LOW_CASEMAP_SM_CAPS = 4                                ; All characters are put in upper case, but with a smaller font height.

; Cell Type
Global Const _
		$LOW_CELL_TYPE_EMPTY = 0, _                             ; Cell is empty.
		$LOW_CELL_TYPE_VALUE = 1, _                             ; Cell contains a value.
		$LOW_CELL_TYPE_TEXT = 2, _                              ; Cell contains text.
		$LOW_CELL_TYPE_FORMULA = 3                              ; Cell contains a formula.

; Color in Long Color Format
Global Const _
		$LOW_COLOR_OFF = -1, _                                  ; Turn Color off, or to automatic mode.
		$LOW_COLOR_BLACK = 0, _                                 ; Black color.
		$LOW_COLOR_WHITE = 16777215, _                          ; White color.
		$LOW_COLOR_LGRAY = 11711154, _                          ; Light Gray color.
		$LOW_COLOR_GRAY = 8421504, _                            ; Gray color.
		$LOW_COLOR_DKGRAY = 3355443, _                          ; Dark Gray color.
		$LOW_COLOR_YELLOW = 16776960, _                         ; Yellow color.
		$LOW_COLOR_GOLD = 16760576, _                           ; Gold color.
		$LOW_COLOR_ORANGE = 16744448, _                         ; Orange color.
		$LOW_COLOR_BRICK = 16728064, _                          ; Brick color.
		$LOW_COLOR_RED = 16711680, _                            ; Red color.
		$LOW_COLOR_MAGENTA = 12517441, _                        ; Magenta color.
		$LOW_COLOR_PURPLE = 8388736, _                          ; Purple color.
		$LOW_COLOR_INDIGO = 5582989, _                          ; Indigo color.
		$LOW_COLOR_BLUE = 2777241, _                            ; Blue color.
		$LOW_COLOR_TEAL = 1410150, _                            ; Teal color.
		$LOW_COLOR_GREEN = 43315, _                             ; Green color.
		$LOW_COLOR_LIME = 8508442, _                            ; Lime color.
		$LOW_COLOR_BROWN = 9127187                              ; Brown color.

; Color Mode Constants.
Global Const _
		$LOW_COLORMODE_STANDARD = 0, _                          ; The graphic is rendered in the default color style of the output device.
		$LOW_COLORMODE_GRAYSCALE = 1, _                         ; The graphic is rendered in grayscale on the output device.
		$LOW_COLORMODE_BLACK_WHITE = 2, _                       ; The graphic is rendered in black and white only.
		$LOW_COLORMODE_WATERMARK = 3                            ; The graphic is rendered in a watermark like style.

; Control Character
Global Const _
		$LOW_CON_CHAR_PAR_BREAK = 0, _                          ; A new paragraph.
		$LOW_CON_CHAR_LINE_BREAK = 1, _                         ; A new line in a paragraph.
		$LOW_CON_CHAR_HARD_HYPHEN = 2, _                        ; A dash but prevents this position from being hyphenated.
		$LOW_CON_CHAR_SOFT_HYPHEN = 3, _                        ; Defines a preferred hyphenation point if the word must be split at the end of a line.
		$LOW_CON_CHAR_HARD_SPACE = 4, _                         ; Insert a space that prevents two words from splitting at a line break.
		$LOW_CON_CHAR_APPEND_PAR = 5                            ; Appends a new paragraph.

; Cursor Data Related Constants
Global Const _
		$LOW_CURDATA_BODY_TEXT = 1, _                           ; Cursor is currently in the Body Text.
		$LOW_CURDATA_FRAME = 2, _                               ; Cursor is currently in a Text Frame.
		$LOW_CURDATA_CELL = 3, _                                ; Cursor is currently in a Text Table Cell.
		$LOW_CURDATA_FOOTNOTE = 4, _                            ; Cursor is currently in a Footnote.
		$LOW_CURDATA_ENDNOTE = 5, _                             ; Cursor is currently in a Endnote.
		$LOW_CURDATA_HEADER_FOOTER = 6                          ; Cursor is currently in a Header or Footer.

; Cursor Status
Global Enum _
		$LOW_CURSOR_STAT_IS_COLLAPSED, _                        ; Test if the start and end positions are the same for a cursor selection, meaning the cursor has nothing selected..
		$LOW_CURSOR_STAT_IS_START_OF_WORD, _                    ; Test if a cursor is at the start of a word. Returns True if so.
		$LOW_CURSOR_STAT_IS_END_OF_WORD, _                      ; Test if a cursor is at the end of a word. Returns True if so.
		$LOW_CURSOR_STAT_IS_START_OF_SENTENCE, _                ; Test if a cursor is at the start of a sentence. Returns True if so.
		$LOW_CURSOR_STAT_IS_END_OF_SENTENCE, _                  ; Test if a cursor is at the end of a sentence. Returns True if so.
		$LOW_CURSOR_STAT_IS_START_OF_PAR, _                     ; Test if a cursor is at the start of a paragraph. Returns True if so.
		$LOW_CURSOR_STAT_IS_END_OF_PAR, _                       ; Test if a cursor is at the End of a paragraph. Returns True if so.
		$LOW_CURSOR_STAT_IS_START_OF_LINE, _                    ; Test if a cursor is at the start of the line. Returns True if so.
		$LOW_CURSOR_STAT_IS_END_OF_LINE, _                      ; Test if a cursor is at the end of the line. Returns True if so.
		$LOW_CURSOR_STAT_GET_PAGE, _                            ; Return the current page number the cursor is in as an integer.
		$LOW_CURSOR_STAT_GET_RANGE_NAME                         ; Return the cell range selected by a cursor as a string. For example, “B3:D5”.

; Cursor Type Related Constants
Global Const _
		$LOW_CURTYPE_TEXT_CURSOR = 1, _                         ; Cursor is a TextCursor type.
		$LOW_CURTYPE_TABLE_CURSOR = 2, _                        ; Cursor is a TableCursor type.
		$LOW_CURTYPE_VIEW_CURSOR = 3, _                         ; Cursor is a ViewCursor type.
		$LOW_CURTYPE_PARAGRAPH = 4, _                           ; Object is a Paragraph Object.
		$LOW_CURTYPE_TEXT_PORTION = 5                           ; Object is a Paragraph Text Portion Object.

; Printer Duplex Constants.
Global Const _
		$LOW_DUPLEX_UNKNOWN = 0, _                              ; Duplex mode setting is unknown.
		$LOW_DUPLEX_OFF = 1, _                                  ; Duplex mode is off.
		$LOW_DUPLEX_LONG = 2, _                                 ; Duplex mode is on, flip on Long edge.
		$LOW_DUPLEX_SHORT = 3                                   ; Duplex mode is on, flip on Short edge.

; Advanced Field Types
Global Enum Step *2 _
		$LOW_FIELD_ADV_TYPE_ALL = 1, _                          ; All of the below listed Fields will be returned.
		$LOW_FIELD_ADV_TYPE_BIBLIOGRAPHY, _                     ; A Bibliography Field, found in Fields dialog, Database tab.
		$LOW_FIELD_ADV_TYPE_DATABASE, _                         ; A Database Field, found in Fields dialog, Database tab.
		$LOW_FIELD_ADV_TYPE_DATABASE_SET_NUM, _                 ; A Database Field, found in Fields dialog, Database tab.
		$LOW_FIELD_ADV_TYPE_DATABASE_NAME, _                    ; A Database Field, found in Fields dialog, Database tab.
		$LOW_FIELD_ADV_TYPE_DATABASE_NEXT_SET, _                ; A Database Field, found in Fields dialog, Database tab.
		$LOW_FIELD_ADV_TYPE_DATABASE_NAME_OF_SET, _             ; A Database Field, found in Fields dialog, Database tab.
		$LOW_FIELD_ADV_TYPE_DDE, _                              ; A DDE Field, found in Fields dialog, Variables tab.
		$LOW_FIELD_ADV_TYPE_INPUT_USER, _                       ; ?
		$LOW_FIELD_ADV_TYPE_USER                                ; A User Field, found in Fields dialog, Variables tab.

; Field Chapter Display Type
Global Const _
		$LOW_FIELD_CHAP_FRMT_NAME = 0, _                        ; The title of the chapter is displayed.
		$LOW_FIELD_CHAP_FRMT_NUMBER = 1, _                      ; The number including prefix and suffix of the chapter is displayed.
		$LOW_FIELD_CHAP_FRMT_NAME_NUMBER = 2, _                 ; The title and number, with prefix and suffix of the chapter are displayed.
		$LOW_FIELD_CHAP_FRMT_NO_PREFIX_SUFFIX = 3, _            ; The name and number of the chapter are displayed.
		$LOW_FIELD_CHAP_FRMT_DIGIT = 4                          ; The number of the chapter is displayed.

; Count Field Type
Global Enum _
		$LOW_FIELD_COUNT_TYPE_CHARACTERS = 0, _                 ; Count field is a Character Count type field.
		$LOW_FIELD_COUNT_TYPE_IMAGES, _                         ; Count field is an Image Count type field.
		$LOW_FIELD_COUNT_TYPE_OBJECTS, _                        ; Count field is an Object Count type field.
		$LOW_FIELD_COUNT_TYPE_PAGES, _                          ; Count field is a Page Count type field.
		$LOW_FIELD_COUNT_TYPE_PARAGRAPHS, _                     ; Count field is a Paragraph Count type field.
		$LOW_FIELD_COUNT_TYPE_TABLES, _                         ; Count field is a Table Count type field.
		$LOW_FIELD_COUNT_TYPE_WORDS                             ; Count field is a Word Count type field.

; Document Information Field Types
Global Enum Step *2 _
		$LOW_FIELD_DOCINFO_TYPE_ALL = 1, _                      ; Returns a list of all field types listed below.
		$LOW_FIELD_DOCINFO_TYPE_MOD_AUTH, _                     ; A Modified By Author Field, found in Fields dialog, DocInformation Tab, Modified Type.
		$LOW_FIELD_DOCINFO_TYPE_MOD_DATE_TIME, _                ; A Modified Date/Time Field, found in Fields dialog, DocInformation Tab, Modified Type.
		$LOW_FIELD_DOCINFO_TYPE_CREATE_AUTH, _                  ; A Created By Author Field, found in Fields dialog, DocInformation Tab, Created Type.
		$LOW_FIELD_DOCINFO_TYPE_CREATE_DATE_TIME, _             ; A Created Date/Time Field, found in Fields dialog, DocInformation Tab, Created Type.
		$LOW_FIELD_DOCINFO_TYPE_CUSTOM, _                       ; A Custom Field, found in Fields dialog, DocInformation Tab.
		$LOW_FIELD_DOCINFO_TYPE_COMMENTS, _                     ; A Comments Field, found in Fields dialog, DocInformation Tab.
		$LOW_FIELD_DOCINFO_TYPE_EDIT_TIME, _                    ; A Total Editing Time Field, found in Fields dialog, DocInformation Tab.
		$LOW_FIELD_DOCINFO_TYPE_KEYWORDS, _                     ; A Keywords Field, found in Fields dialog, DocInformation Tab.
		$LOW_FIELD_DOCINFO_TYPE_PRINT_AUTH, _                   ; A Printed By Author Field, found in Fields dialog, DocInformation Tab, Last Printed Type.
		$LOW_FIELD_DOCINFO_TYPE_PRINT_DATE_TIME, _              ; A Printed Date/Time Field, found in Fields dialog, DocInformation Tab, Last Printed Type.
		$LOW_FIELD_DOCINFO_TYPE_REVISION, _                     ; A Revision Number Field, found in Fields dialog, DocInformation Tab.
		$LOW_FIELD_DOCINFO_TYPE_SUBJECT, _                      ; A Subject Field, found in Fields dialog, DocInformation Tab.
		$LOW_FIELD_DOCINFO_TYPE_TITLE                           ; A Title Field, found in Fields dialog, DocInformation Tab.

; File Name Field Type
Global Const _
		$LOW_FIELD_FILENAME_FULL_PATH = 0, _                    ; The content of the URL is completely displayed.
		$LOW_FIELD_FILENAME_PATH = 1, _                         ; Only the path of the file is displayed.
		$LOW_FIELD_FILENAME_NAME = 2, _                         ; Only the name of the file without the file extension is displayed.
		$LOW_FIELD_FILENAME_NAME_AND_EXT = 3, _                 ; The file name including the file extension is displayed.
		$LOW_FIELD_FILENAME_CATEGORY = 4, _                     ; The Category of the Template is displayed.
		$LOW_FIELD_FILENAME_TEMPLATE_NAME = 5                   ; The Template Name is displayed.

; Placeholder Type
Global Const _
		$LOW_FIELD_PLACEHOLD_TYPE_TEXT = 0, _                   ; The field represents a piece of text.
		$LOW_FIELD_PLACEHOLD_TYPE_TABLE = 1, _                  ; The field initiates the insertion of a text table.
		$LOW_FIELD_PLACEHOLD_TYPE_FRAME = 2, _                  ; The field initiates the insertion of a text frame.
		$LOW_FIELD_PLACEHOLD_TYPE_GRAPHIC = 3, _                ; The field initiates the insertion of a graphic object.
		$LOW_FIELD_PLACEHOLD_TYPE_OBJECT = 4                    ; The field initiates the insertion of an embedded object.

; Reference Field Type
Global Const _
		$LOW_FIELD_REF_TYPE_REF_MARK = 0, _                     ; The source is referencing a reference mark.
		$LOW_FIELD_REF_TYPE_SEQ_FIELD = 1, _                    ; The source is referencing a number sequence field.
		$LOW_FIELD_REF_TYPE_BOOKMARK = 2, _                     ; The source is referencing a bookmark.
		$LOW_FIELD_REF_TYPE_FOOTNOTE = 3, _                     ; The source is referencing a footnote.
		$LOW_FIELD_REF_TYPE_ENDNOTE = 4                         ; The source is referencing an endnote.

; Type of Reference
Global Const _
		$LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED = 0, _           ; The page number is displayed using Arabic numbers.
		$LOW_FIELD_REF_USING_CHAPTER = 1, _                     ; The number of the chapter is displayed.
		$LOW_FIELD_REF_USING_REF_TEXT = 2, _                    ; The reference text is displayed.
		$LOW_FIELD_REF_USING_ABOVE_BELOW = 3, _                 ; The reference is displayed as one of the words, "above" or "below".
		$LOW_FIELD_REF_USING_PAGE_NUM_STYLED = 4, _             ; The page number is displayed using the numbering type defined in the page style of the reference position.
		$LOW_FIELD_REF_USING_CAT_AND_NUM = 5, _                 ; Inserts the category (caption type) and the number of the reference target.
		$LOW_FIELD_REF_USING_CAPTION = 6, _                     ; Inserts the caption label of the reference target.
		$LOW_FIELD_REF_USING_NUMBERING = 7, _                   ; Inserts the caption number of the reference target.
		$LOW_FIELD_REF_USING_NUMBER = 8, _                      ;  Inserts the number of the heading or numbered paragraph, including superior levels.
		$LOW_FIELD_REF_USING_NUMBER_NO_CONT = 9, _              ; Inserts only the number of the heading or numbered paragraph.
		$LOW_FIELD_REF_USING_NUMBER_CONT = 10                   ; Inserts the number of the heading or numbered paragraph, including all superior levels.

; Regular Field Types
Global Enum Step *2 _
		$LOW_FIELD_TYPE_ALL = 1, _                              ; Returns a list of all field types listed below.
		$LOW_FIELD_TYPE_COMMENT, _                              ; A Comment Field. As Found at Insert > Comment
		$LOW_FIELD_TYPE_AUTHOR, _                               ; A Author field, found in the Fields Dialog, Document tab.
		$LOW_FIELD_TYPE_CHAPTER, _                              ; A Chapter field, found in the Fields Dialog, Document tab.
		$LOW_FIELD_TYPE_CHAR_COUNT, _                           ; A Character Count field, found in the Fields Dialog, Document tab, Statistics Type.
		$LOW_FIELD_TYPE_COMBINED_CHAR, _                        ; A Combined Character field, found in the Fields Dialog, Functions tab.
		$LOW_FIELD_TYPE_COND_TEXT, _                            ; A Conditional Text field, found in the Fields Dialog, Functions tab.
		$LOW_FIELD_TYPE_DATE_TIME, _                            ; A Date/Time field, found in the Fields Dialog, Document tab, Date Type and Time Type..
		$LOW_FIELD_TYPE_INPUT_LIST, _                           ; A Input List field, found in the Fields Dialog, Functions tab.
		$LOW_FIELD_TYPE_EMB_OBJ_COUNT, _                        ; A Object Count field, found in the Fields Dialog, Document tab, Statistics Type.
		$LOW_FIELD_TYPE_SENDER, _                               ; A Sender field, found in the Fields Dialog, Document tab.
		$LOW_FIELD_TYPE_FILENAME, _                             ; A File Name field, found in the Fields Dialog, Document tab.
		$LOW_FIELD_TYPE_SHOW_VAR, _                             ; A Show Variable field, found in the Fields Dialog, Variables tab.
		$LOW_FIELD_TYPE_INSERT_REF, _                           ; A Insert Reference field, found in the Fields Dialog, Cross-References tab. [Includes: "Insert Reference", "Headings", "Numbered Paragraphs", "Drawing", "Bookmarks", "Footnotes", "Endnotes", etc.]
		$LOW_FIELD_TYPE_IMAGE_COUNT, _                          ; A Image Count field, found in the Fields Dialog, Document tab, Statistics Type.
		$LOW_FIELD_TYPE_HIDDEN_PAR, _                           ; A Hidden Paragraph field, found in the Fields Dialog, Functions tab.
		$LOW_FIELD_TYPE_HIDDEN_TEXT, _                          ; A Hidden Text field, found in the Fields Dialog, Functions tab.
		$LOW_FIELD_TYPE_INPUT, _                                ; A Input field, found in the Fields Dialog, Functions tab.
		$LOW_FIELD_TYPE_PLACEHOLDER, _                          ; A Placeholder field, found in the Fields Dialog, Functions tab.
		$LOW_FIELD_TYPE_MACRO, _                                ; A Execute Macro field, found in the Fields Dialog, Functions tab.
		$LOW_FIELD_TYPE_PAGE_COUNT, _                           ; A Page Count field, found in the Fields Dialog, Document tab, Statistics Type.
		$LOW_FIELD_TYPE_PAGE_NUM, _                             ; A Page Number (Unstyled) field, found in the Fields Dialog, Document tab.
		$LOW_FIELD_TYPE_PAR_COUNT, _                            ; A Paragraph Count field, found in the Fields Dialog, Document tab, Statistics Type..
		$LOW_FIELD_TYPE_SHOW_PAGE_VAR, _                        ; A Show Page Variable field, found in the Fields Dialog, Variables tab.
		$LOW_FIELD_TYPE_SET_PAGE_VAR, _                         ; A Set Page Variable field, found in the Fields Dialog, Variables tab.
		$LOW_FIELD_TYPE_SCRIPT, _                               ;
		$LOW_FIELD_TYPE_SET_VAR, _                              ; A Set Variable field, found in the Fields Dialog, Variables tab..
		$LOW_FIELD_TYPE_TABLE_COUNT, _                          ; A Table Count field, found in the Fields Dialog, Document tab, Statistics Type.
		$LOW_FIELD_TYPE_TEMPLATE_NAME, _                        ; A Templates field, found in the Fields Dialog, Document tab.
		$LOW_FIELD_TYPE_URL, _                                  ;
		$LOW_FIELD_TYPE_WORD_COUNT                              ; A Word Count field, found in the Fields Dialog, Document tab, Statistics Type.

; User Data Field Type
Global Const _
		$LOW_FIELD_USER_DATA_COMPANY = 0, _                     ; The field shows the company name.
		$LOW_FIELD_USER_DATA_FIRST_NAME = 1, _                  ; The field shows the first name.
		$LOW_FIELD_USER_DATA_NAME = 2, _                        ; The field shows the name.
		$LOW_FIELD_USER_DATA_SHORTCUT = 3, _                    ; The field shows the initials.
		$LOW_FIELD_USER_DATA_STREET = 4, _                      ; The field shows the street.
		$LOW_FIELD_USER_DATA_COUNTRY = 5, _                     ; The field shows the country.
		$LOW_FIELD_USER_DATA_ZIP = 6, _                         ; The field shows the zip code.
		$LOW_FIELD_USER_DATA_CITY = 7, _                        ; The field shows the city.
		$LOW_FIELD_USER_DATA_TITLE = 8, _                       ; The field shows the title.
		$LOW_FIELD_USER_DATA_POSITION = 9, _                    ; The field shows the position.
		$LOW_FIELD_USER_DATA_PHONE_PRIVATE = 10, _              ; The field shows the number of the private phone.
		$LOW_FIELD_USER_DATA_PHONE_COMPANY = 11, _              ; The field shows the number of the business phone.
		$LOW_FIELD_USER_DATA_FAX = 12, _                        ; The field shows the fax number.
		$LOW_FIELD_USER_DATA_EMAIL = 13, _                      ; The field shows the e-Mail.
		$LOW_FIELD_USER_DATA_STATE = 14                         ; The field shows the state.

; Follow By
Global Const _
		$LOW_FOLLOW_BY_TABSTOP = 0, _                           ; A Tab will follow the Numbering Style Number.
		$LOW_FOLLOW_BY_SPACE = 1, _                             ; A Space will follow the Numbering Style Number.
		$LOW_FOLLOW_BY_NOTHING = 2, _                           ; Nothing will follow the Numbering Style Number.
		$LOW_FOLLOW_BY_NEWLINE = 3                              ; A Newline will follow the Numbering Style Number.

; Footnote Count type
Global Const _
		$LOW_FOOTNOTE_COUNT_PER_PAGE = 0, _                     ; Restarts the numbering of footnotes at the top of each page. This option is only available if End of Doc is set to False.
		$LOW_FOOTNOTE_COUNT_PER_CHAP = 1, _                     ; Restarts the numbering of footnotes at the beginning of each chapter.
		$LOW_FOOTNOTE_COUNT_PER_DOC = 2                         ; Numbers the footnotes in the document sequentially.

; Format Key Type
Global Const _
		$LOW_FORMAT_KEYS_ALL = 0, _                             ; Returns All number formats.
		$LOW_FORMAT_KEYS_DEFINED = 1, _                         ; Returns Only user-defined number formats.
		$LOW_FORMAT_KEYS_DATE = 2, _                            ; Returns Date formats.
		$LOW_FORMAT_KEYS_TIME = 4, _                            ; Returns Time formats.
		$LOW_FORMAT_KEYS_DATE_TIME = 6, _                       ; Returns Number formats which contain date and time.
		$LOW_FORMAT_KEYS_CURRENCY = 8, _                        ; Returns Currency formats.
		$LOW_FORMAT_KEYS_NUMBER = 16, _                         ; Returns Decimal number formats.
		$LOW_FORMAT_KEYS_SCIENTIFIC = 32, _                     ; Returns Scientific number formats.
		$LOW_FORMAT_KEYS_FRACTION = 64, _                       ; Returns Number formats for fractions.
		$LOW_FORMAT_KEYS_PERCENT = 128, _                       ; Returns Percentage number formats.
		$LOW_FORMAT_KEYS_TEXT = 256, _                          ; Returns Text number formats.
		$LOW_FORMAT_KEYS_LOGICAL = 1024, _                      ; Returns Boolean number formats.
		$LOW_FORMAT_KEYS_UNDEFINED = 2048, _                    ; Returns Is used as a return value if no format exists.
		$LOW_FORMAT_KEYS_EMPTY = 4096, _                        ; Returns Empty Number formats (?)
		$LOW_FORMAT_KEYS_DURATION = 8196                        ; Returns Duration number formats.

; Frame Target
Global Const _
		$LOW_FRAME_TARGET_NONE = "", _                          ;
		$LOW_FRAME_TARGET_TOP = "_top", _                       ; File opens in the topmost frame in the hierarchy.
		$LOW_FRAME_TARGET_PARENT = "_parent", _                 ; File opens in the parent frame of the current frame. If there is no parent frame, the current frame is used.
		$LOW_FRAME_TARGET_BLANK = "_blank", _                   ; File opens in a new page.
		$LOW_FRAME_TARGET_SELF = "_self"                        ; File opens in the current frame.

; Gradient Names
Global Const _
		$LOW_GRAD_NAME_PASTEL_BOUQUET = "Pastel Bouquet", _     ; The "Pastel Bouquet" Gradient Preset.
		$LOW_GRAD_NAME_PASTEL_DREAM = "Pastel Dream", _         ; The "Pastel Dream" Gradient Preset.
		$LOW_GRAD_NAME_BLUE_TOUCH = "Blue Touch", _             ; The "Blue Touch" Gradient Preset.
		$LOW_GRAD_NAME_BLANK_W_GRAY = "Blank with Gray", _      ; The "Blank with Gray" Gradient Preset.
		$LOW_GRAD_NAME_SPOTTED_GRAY = "Spotted Gray", _         ; The "Spotted Gray" Gradient Preset.
		$LOW_GRAD_NAME_LONDON_MIST = "London Mist", _           ; The "London Mist" Gradient Preset.
		$LOW_GRAD_NAME_TEAL_TO_BLUE = "Teal to Blue", _         ; The "Teal to Blue" Gradient Preset.
		$LOW_GRAD_NAME_MIDNIGHT = "Midnight", _                 ; The "Midnight" Gradient Preset.
		$LOW_GRAD_NAME_DEEP_OCEAN = "Deep Ocean", _             ; The "Deep Ocean" Gradient Preset.
		$LOW_GRAD_NAME_SUBMARINE = "Submarine", _               ; The "Submarine" Gradient Preset.
		$LOW_GRAD_NAME_GREEN_GRASS = "Green Grass", _           ; The "Green Grass" Gradient Preset.
		$LOW_GRAD_NAME_NEON_LIGHT = "Neon Light", _             ; The "Neon Light" Gradient Preset.
		$LOW_GRAD_NAME_SUNSHINE = "Sunshine", _                 ; The "Sunshine" Gradient Preset.
		$LOW_GRAD_NAME_PRESENT = "Present", _                   ; The "Present" Gradient Preset.
		$LOW_GRAD_NAME_MAHOGANY = "Mahogany"                    ; The "Mahogany" Gradient Preset.

; Gradient Type
Global Const _
		$LOW_GRAD_TYPE_OFF = -1, _                              ; Turn the Gradient off.
		$LOW_GRAD_TYPE_LINEAR = 0, _                            ; Linear type Gradient
		$LOW_GRAD_TYPE_AXIAL = 1, _                             ; Axial type Gradient
		$LOW_GRAD_TYPE_RADIAL = 2, _                            ; Radial type Gradient
		$LOW_GRAD_TYPE_ELLIPTICAL = 3, _                        ; Elliptical type Gradient
		$LOW_GRAD_TYPE_SQUARE = 4, _                            ; Square type Gradient
		$LOW_GRAD_TYPE_RECT = 5                                 ; Rectangle type Gradient

; Line Spacing
Global Const _
		$LOW_LINE_SPC_MODE_PROP = 0, _                          ; Specifies the height value as a proportional value. Min 6% Max 65,535%. (without percentage sign)
		$LOW_LINE_SPC_MODE_MIN = 1, _                           ; Specifies the height as the minimum line height. [Minimum/At least in L.O. U.I.]  Min 0, Max 10008 Micrometers (uM)
		$LOW_LINE_SPC_MODE_LEADING = 2, _                       ; Specifies the height value as the distance to the previous line. Min 0, Max 10008 Micrometers (uM)
		$LOW_LINE_SPC_MODE_FIX = 3                              ; Specifies the height value as a fixed line height. Min 51 Micrometers, Max 10008 Micrometers (uM)

; Line Style
Global Const _
		$LOW_LINE_STYLE_NONE = 0, _                             ; No line.
		$LOW_LINE_STYLE_SOLID = 1, _                            ; A solid line.
		$LOW_LINE_STYLE_DOTTED = 2, _                           ; A dotted line.
		$LOW_LINE_STYLE_DASHED = 3                              ; A Dashed line.

; Numbering Style Type
Global Const _
		$LOW_NUM_STYLE_CHARS_UPPER_LETTER = 0, _                ; Numbering is put in upper case letters. ("A, B, C, D)
		$LOW_NUM_STYLE_CHARS_LOWER_LETTER = 1, _                ; Numbering is in lower case letters. (a, b, c, d)
		$LOW_NUM_STYLE_ROMAN_UPPER = 2, _                       ; Numbering is in Roman numbers with upper case letters. (I, II, III)
		$LOW_NUM_STYLE_ROMAN_LOWER = 3, _                       ; Numbering is in Roman numbers with lower case letters. (i, ii, iii).
		$LOW_NUM_STYLE_ARABIC = 4, _                            ; Numbering is in Arabic numbers. (1, 2, 3, 4),
		$LOW_NUM_STYLE_NUMBER_NONE = 5, _                       ; Numbering is invisible.
		$LOW_NUM_STYLE_CHAR_SPECIAL = 6, _                      ; Use a character from a specified font.
		$LOW_NUM_STYLE_PAGE_DESCRIPTOR = 7, _                   ; Numbering is specified in the page style.
		$LOW_NUM_STYLE_BITMAP = 8, _                            ; Numbering is displayed as a bitmap graphic.
		$LOW_NUM_STYLE_CHARS_UPPER_LETTER_N = 9, _              ; Numbering is put in upper case letters. (A, B, Y, Z, AA, BB)
		$LOW_NUM_STYLE_CHARS_LOWER_LETTER_N = 10, _             ; Numbering is put in lower case letters. (a, b, y, z, aa, bb)
		$LOW_NUM_STYLE_TRANSLITERATION = 11, _                  ; A transliteration module will be used to produce numbers in Chinese, Japanese, etc.
		$LOW_NUM_STYLE_NATIVE_NUMBERING = 12, _                 ; The NativeNumberSupplier service will be called to produce numbers in native languages.
		$LOW_NUM_STYLE_FULLWIDTH_ARABIC = 13, _                 ; Numbering for full width Arabic number.
		$LOW_NUM_STYLE_CIRCLE_NUMBER = 14, _                    ; Bullet for Circle Number.
		$LOW_NUM_STYLE_NUMBER_LOWER_ZH = 15, _                  ; Numbering for Chinese lower case number.
		$LOW_NUM_STYLE_NUMBER_UPPER_ZH = 16, _                  ; Numbering for Chinese upper case number.
		$LOW_NUM_STYLE_NUMBER_UPPER_ZH_TW = 17, _               ; Numbering for Traditional Chinese upper case number.
		$LOW_NUM_STYLE_TIAN_GAN_ZH = 18, _                      ; Bullet for Chinese Tian Gan.
		$LOW_NUM_STYLE_DI_ZI_ZH = 19, _                         ; Bullet for Chinese Di Zi.
		$LOW_NUM_STYLE_NUMBER_TRADITIONAL_JA = 20, _            ; Numbering for Japanese traditional number.
		$LOW_NUM_STYLE_AIU_FULLWIDTH_JA = 21, _                 ; Bullet for Japanese AIU fullwidth.
		$LOW_NUM_STYLE_AIU_HALFWIDTH_JA = 22, _                 ; Bullet for Japanese AIU halfwidth.
		$LOW_NUM_STYLE_IROHA_FULLWIDTH_JA = 23, _               ; Bullet for Japanese IROHA fullwidth.
		$LOW_NUM_STYLE_IROHA_HALFWIDTH_JA = 24, _               ; Bullet for Japanese IROHA halfwidth.
		$LOW_NUM_STYLE_NUMBER_UPPER_KO = 25, _                  ; Numbering for Korean upper case number.
		$LOW_NUM_STYLE_NUMBER_HANGUL_KO = 26, _                 ; Numbering for Korean Hangul number.
		$LOW_NUM_STYLE_HANGUL_JAMO_KO = 27, _                   ; Bullet for Korean Hangul Jamo.
		$LOW_NUM_STYLE_HANGUL_SYLLABLE_KO = 28, _               ; Bullet for Korean Hangul Syllable.
		$LOW_NUM_STYLE_HANGUL_CIRCLED_JAMO_KO = 29, _           ; Bullet for Korean Hangul Circled Jamo.
		$LOW_NUM_STYLE_HANGUL_CIRCLED_SYLLABLE_KO = 30, _       ; Bullet for Korean Hangul Circled Syllable.
		$LOW_NUM_STYLE_CHARS_ARABIC = 31, _                     ; Numbering in Arabic alphabet letters.
		$LOW_NUM_STYLE_CHARS_THAI = 32, _                       ; Numbering in Thai alphabet letters.
		$LOW_NUM_STYLE_CHARS_HEBREW = 33, _                     ; Numbering in Hebrew alphabet letters.
		$LOW_NUM_STYLE_CHARS_NEPALI = 34, _                     ; Numbering in Nepali alphabet letters.
		$LOW_NUM_STYLE_CHARS_KHMER = 35, _                      ; Numbering in Khmer alphabet letters.
		$LOW_NUM_STYLE_CHARS_LAO = 36, _                        ; Numbering in Lao alphabet letters.
		$LOW_NUM_STYLE_CHARS_TIBETAN = 37, _                    ; Numbering in Tibetan/Dzongkha alphabet letters.
		$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_BG = 38, _   ; Numbering in Cyrillic alphabet upper case letters.
		$LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_BG = 39, _   ; Numbering in Cyrillic alphabet lower case letters.
		$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_N_BG = 40, _ ; Numbering in Cyrillic alphabet upper case letters.
		$LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_N_BG = 41, _ ; Numbering in Cyrillic alphabet upper case letters.
		$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_RU = 42, _   ; Numbering in Russian Cyrillic alphabet upper case letters.
		$LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_RU = 43, _   ; Numbering in Russian Cyrillic alphabet lower case letters.
		$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_N_RU = 44, _ ; Numbering in Russian Cyrillic alphabet upper case letters.
		$LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_N_RU = 45, _ ; Numbering in Russian Cyrillic alphabet upper case letters.
		$LOW_NUM_STYLE_CHARS_PERSIAN = 46, _                    ; Numbering in Persian alphabet letters.
		$LOW_NUM_STYLE_CHARS_MYANMAR = 47, _                    ; Numbering in Myanmar alphabet letters.
		$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_SR = 48, _   ; Numbering in Serbian Cyrillic alphabet upper case letters.
		$LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_SR = 49, _   ; Numbering in Russian Serbian alphabet lower case letters.
		$LOW_NUM_STYLE_CHARS_CYRILLIC_UPPER_LETTER_N_SR = 50, _ ; Numbering in Serbian Cyrillic alphabet upper case letters.
		$LOW_NUM_STYLE_CHARS_CYRILLIC_LOWER_LETTER_N_SR = 51, _ ; Numbering in Serbian Cyrillic alphabet upper case letters.
		$LOW_NUM_STYLE_CHARS_GREEK_UPPER_LETTER = 52, _         ; Numbering in Greek alphabet upper case letters.
		$LOW_NUM_STYLE_CHARS_GREEK_LOWER_LETTER = 53, _         ; Numbering in Greek alphabet lower case letters.
		$LOW_NUM_STYLE_CHARS_ARABIC_ABJAD = 54, _               ; Numbering in Arabic alphabet using abjad sequence.
		$LOW_NUM_STYLE_CHARS_PERSIAN_WORD = 55, _               ; Numbering in Persian words.
		$LOW_NUM_STYLE_NUMBER_HEBREW = 56, _                    ; Numbering in Hebrew numerals.
		$LOW_NUM_STYLE_NUMBER_ARABIC_INDIC = 57, _              ; Numbering in Arabic-Indic numerals.
		$LOW_NUM_STYLE_NUMBER_EAST_ARABIC_INDIC = 58, _         ; Numbering in East Arabic-Indic numerals.
		$LOW_NUM_STYLE_NUMBER_INDIC_DEVANAGARI = 59, _          ; Numbering in Indic Devanagari numerals.
		$LOW_NUM_STYLE_TEXT_NUMBER = 60, _                      ; Numbering in ordinal numbers of the language of the text node. (1st, 2nd, 3rd)
		$LOW_NUM_STYLE_TEXT_CARDINAL = 61, _                    ; Numbering in cardinal numbers of the language of the text node. (One, Two)
		$LOW_NUM_STYLE_TEXT_ORDINAL = 62, _                     ; Numbering in ordinal numbers of the language of the text node. (First, Second)
		$LOW_NUM_STYLE_SYMBOL_CHICAGO = 63, _                   ; Footnoting symbols according the University of Chicago style.
		$LOW_NUM_STYLE_ARABIC_ZERO = 64, _                      ; Numbering is in Arabic numbers, padded with zero to have a length of at least two. (01, 02)
		$LOW_NUM_STYLE_ARABIC_ZERO3 = 65, _                     ; Numbering is in Arabic numbers, padded with zero to have a length of at least three.
		$LOW_NUM_STYLE_ARABIC_ZERO4 = 66, _                     ; Numbering is in Arabic numbers, padded with zero to have a length of at least four.
		$LOW_NUM_STYLE_ARABIC_ZERO5 = 67, _                     ; Numbering is in Arabic numbers, padded with zero to have a length of at least five.
		$LOW_NUM_STYLE_SZEKELY_ROVAS = 68, _                    ; Numbering is in Szekely rovas (Old Hungarian) numerals.
		$LOW_NUM_STYLE_NUMBER_DIGITAL_KO = 69, _                ; Numbering is in Korean Digital number.
		$LOW_NUM_STYLE_NUMBER_DIGITAL2_KO = 70, _               ; Numbering is in Korean Digital Number, reserved "koreanDigital2".
		$LOW_NUM_STYLE_NUMBER_LEGAL_KO = 71                     ; Numbering is in Korean Legal Number, reserved "koreanLegal".

; Horizontal Orientation
Global Const _
		$LOW_ORIENT_HORI_NONE = 0, _                            ; No hard alignment is applied. Equal to "From Left" in L.O. U.I.
		$LOW_ORIENT_HORI_RIGHT = 1, _                           ; The object is aligned at the right side.
		$LOW_ORIENT_HORI_CENTER = 2, _                          ; The object is aligned at the middle.
		$LOW_ORIENT_HORI_LEFT = 3, _                            ; The object is aligned at the left side.
		$LOW_ORIENT_HORI_FULL = 6, _                            ; The table uses the full space (for text tables only).
		$LOW_ORIENT_HORI_LEFT_AND_WIDTH = 7                     ;  The left offset and the width of the table are defined.

; Vertical Orientation
Global Const _
		$LOW_ORIENT_VERT_NONE = 0, _                            ; No hard alignment. The same as "From Top"/From Bottom" in L.O. U.I., the only difference is the combination setting of Vertical Relation.
		$LOW_ORIENT_VERT_TOP = 1, _                             ; Aligned at the top.
		$LOW_ORIENT_VERT_CENTER = 2, _                          ; Aligned at the center.
		$LOW_ORIENT_VERT_BOTTOM = 3, _                          ; Aligned at the bottom.
		$LOW_ORIENT_VERT_CHAR_TOP = 4, _                        ; Aligned at the top of a character. Available only when anchor is set to "As character". Equal to L.O. UI setting of "Vertical" = Top, and "To" = Character.
		$LOW_ORIENT_VERT_CHAR_CENTER = 5, _                     ; Aligned at the center of a character. Available only when anchor is set to "As character". Equal to L.O. UI setting of "Vertical" = Center, and "To" = Character.
		$LOW_ORIENT_VERT_CHAR_BOTTOM = 6, _                     ; Aligned at the bottom of a character. Available only when anchor is set to "As character". Equal to L.O. UI setting of "Vertical" = Center, and "To" = Character.
		$LOW_ORIENT_VERT_LINE_TOP = 7, _                        ; Aligned at the top of the line. Available only when anchor is set to "As character". Equal to L.O. UI setting of "Vertical" = Top, and "To" = Row.
		$LOW_ORIENT_VERT_LINE_CENTER = 8, _                     ; Aligned at the center of the line. Available only when anchor is set to "As character". Equal to L.O. UI setting of "Vertical" = Center, and "To" = Row.
		$LOW_ORIENT_VERT_LINE_BOTTOM = 9                        ; Aligned at the bottom of the line. Available only when anchor is set to "As character". Equal to L.O. UI setting of "Vertical" = Center, and "To" = Row.

; Outline
Global Const _
		$LOW_OUTLINE_BODY = 0, _                                ; The paragraph belongs to the body text.
		$LOW_OUTLINE_LEVEL_1 = 1, _                             ; The paragraph belongs to the level 1 outline level.
		$LOW_OUTLINE_LEVEL_2 = 2, _                             ; The paragraph belongs to the level 2 outline level.
		$LOW_OUTLINE_LEVEL_3 = 3, _                             ; The paragraph belongs to the level 3 outline level.
		$LOW_OUTLINE_LEVEL_4 = 4, _                             ; The paragraph belongs to the level 4 outline level.
		$LOW_OUTLINE_LEVEL_5 = 5, _                             ; The paragraph belongs to the level 5 outline level.
		$LOW_OUTLINE_LEVEL_6 = 6, _                             ; The paragraph belongs to the level 6 outline level.
		$LOW_OUTLINE_LEVEL_7 = 7, _                             ; The paragraph belongs to the level 7 outline level.
		$LOW_OUTLINE_LEVEL_8 = 8, _                             ; The paragraph belongs to the level 8 outline level.
		$LOW_OUTLINE_LEVEL_9 = 9, _                             ; The paragraph belongs to the level 9 outline level.
		$LOW_OUTLINE_LEVEL_10 = 10                              ; The paragraph belongs to the level 10 outline level.

; Page Layout
Global Const _
		$LOW_PAGE_LAYOUT_ALL = 0, _                             ; Page style shows both odd(Right) and even(Left) pages. With left and right margins.
		$LOW_PAGE_LAYOUT_LEFT = 1, _                            ; Page style shows only even(Left) pages. Odd pages are shown as blank pages. With left and right margins.
		$LOW_PAGE_LAYOUT_RIGHT = 2, _                           ; Page style shows only odd(Right) pages. Even pages are shown as blank pages. With left and right margins.
		$LOW_PAGE_LAYOUT_MIRRORED = 3                           ; Page style shows both odd(Right) and even(Left) pages with inner and outer margins.

; Page Number Type
Global Const _
		$LOW_PAGE_NUM_TYPE_PREV = 0, _                          ; The Previous Page's page number.
		$LOW_PAGE_NUM_TYPE_CURRENT = 1, _                       ; The current page number.
		$LOW_PAGE_NUM_TYPE_NEXT = 2                             ; The Next Page's page number.

; Paper Size Constants.
Global Const _
		$LOW_PAPER_A3 = 0, _                                    ; A3 Paper size.
		$LOW_PAPER_A4 = 1, _                                    ; A4 Paper size.
		$LOW_PAPER_A5 = 2, _                                    ; A5 Paper size.
		$LOW_PAPER_B4 = 3, _                                    ; B4 Paper size.
		$LOW_PAPER_B5 = 4, _                                    ; B5 Paper size.
		$LOW_PAPER_LETTER = 5, _                                ; Letter Paper size.
		$LOW_PAPER_LEGAL = 6, _                                 ; Legal Paper size.
		$LOW_PAPER_TABLOID = 7, _                               ; Tabloid Paper size.
		$LOW_PAPER_USER_DEFINED = 8                             ; Paper size is User-Defined.

; Paper Height in uM
Global Const _
		$LOW_PAPER_HEIGHT_A6 = 14808, _                         ; A6 paper height in Micrometers.
		$LOW_PAPER_HEIGHT_A5 = 21006, _                         ; A5 paper height in Micrometers.
		$LOW_PAPER_HEIGHT_A4 = 29693, _                         ; A4 paper height in Micrometers.
		$LOW_PAPER_HEIGHT_A3 = 42012, _                         ; A3 paper height in Micrometers.
		$LOW_PAPER_HEIGHT_B6ISO = 17602, _                      ; B6ISO paper height in Micrometers.
		$LOW_PAPER_HEIGHT_B5ISO = 24994, _                      ; B5ISO paper height in Micrometers.
		$LOW_PAPER_HEIGHT_B4ISO = 35306, _                      ; B4ISO paper height in Micrometers.
		$LOW_PAPER_HEIGHT_LETTER = 27940, _                     ; Letter paper height in Micrometers.
		$LOW_PAPER_HEIGHT_LEGAL = 35560, _                      ; Legal paper height in Micrometers.
		$LOW_PAPER_HEIGHT_LONG_BOND = 33020, _                  ; Long Bond paper height in Micrometers.
		$LOW_PAPER_HEIGHT_TABLOID = 43180, _                    ; Tabloid paper height in Micrometers.
		$LOW_PAPER_HEIGHT_B6JIS = 18200, _                      ; B6JIS paper height in Micrometers.
		$LOW_PAPER_HEIGHT_B5JIS = 25705, _                      ; B5JIS paper height in Micrometers.
		$LOW_PAPER_HEIGHT_B4JIS = 36398, _                      ; B4JIS paper height in Micrometers.
		$LOW_PAPER_HEIGHT_16KAI = 26010, _                      ; 16KAI paper height in Micrometers.
		$LOW_PAPER_HEIGHT_32KAI = 18390, _                      ; 32KAI paper height in Micrometers.
		$LOW_PAPER_HEIGHT_BIG_32KAI = 20295, _                  ; Big 32KAI paper height in Micrometers.
		$LOW_PAPER_HEIGHT_DLENVELOPE = 21996, _                 ; DL Envelope paper height in Micrometers.
		$LOW_PAPER_HEIGHT_C6ENVELOPE = 16205, _                 ; C6 Envelope paper height in Micrometers.
		$LOW_PAPER_HEIGHT_C6_5_ENVELOPE = 22911, _              ; C6/5 Envelope paper height in Micrometers.
		$LOW_PAPER_HEIGHT_C5ENVELOPE = 22911, _                 ; C5 Envelope paper height in Micrometers.
		$LOW_PAPER_HEIGHT_C4ENVELOPE = 32410, _                 ; C4 Envelope paper height in Micrometers.
		$LOW_PAPER_HEIGHT_6_3_4ENVELOPE = 16510, _              ; 6 3/4 Pound Envelope paper height in Micrometers.
		$LOW_PAPER_HEIGHT_7_3_4ENVELOPE = 19050, _              ; 7 3/4 Pound Envelope paper height in Micrometers.
		$LOW_PAPER_HEIGHT_9ENVELOPE = 22543, _                  ; 9 Pound Envelope paper height in Micrometers.
		$LOW_PAPER_HEIGHT_10ENVELOPE = 24130, _                 ; 10 Pound Envelope paper height in Micrometers.
		$LOW_PAPER_HEIGHT_11ENVELOPE = 26365, _                 ; 11 Pound Envelope paper height in Micrometers.
		$LOW_PAPER_HEIGHT_12ENVELOPE = 27940, _                 ; 12 Pound Envelope paper height in Micrometers.
		$LOW_PAPER_HEIGHT_JAP_POSTCARD = 14808                  ; Japanese Postcard paper height in Micrometers.

; Printer Paper Orientation Constants.
Global Const _
		$LOW_PAPER_ORIENT_PORTRAIT = 0, _                       ; Portrait Paper Orientation.
		$LOW_PAPER_ORIENT_LANDSCAPE = 1                         ; Landscape Paper Orientation.

; Paper Width in uM
Global Const _
		$LOW_PAPER_WIDTH_A6 = 10490, _                          ; A6 paper width in Micrometers.
		$LOW_PAPER_WIDTH_A5 = 14808, _                          ; A5 paper width in Micrometers.
		$LOW_PAPER_WIDTH_A4 = 21006, _                          ; A4 paper width in Micrometers.
		$LOW_PAPER_WIDTH_A3 = 29693, _                          ; A3 paper width in Micrometers.
		$LOW_PAPER_WIDTH_B6ISO = 12497, _                       ; B6ISO paper width in Micrometers.
		$LOW_PAPER_WIDTH_B5ISO = 17602, _                       ; B5ISO paper width in Micrometers.
		$LOW_PAPER_WIDTH_B4ISO = 24994, _                       ; B4ISO paper width in Micrometers.
		$LOW_PAPER_WIDTH_LETTER = 21590, _                      ; Letter paper width in Micrometers.
		$LOW_PAPER_WIDTH_LEGAL = 21590, _                       ; Legal paper width in Micrometers.
		$LOW_PAPER_WIDTH_LONG_BOND = 21590, _                   ; Long Bond paper width in Micrometers.
		$LOW_PAPER_WIDTH_TABLOID = 27940, _                     ; Tabloid paper width in Micrometers.
		$LOW_PAPER_WIDTH_B6JIS = 12801, _                       ; B6JIS paper width in Micrometers.
		$LOW_PAPER_WIDTH_B5JIS = 18212, _                       ; B5JIS paper width in Micrometers.
		$LOW_PAPER_WIDTH_B4JIS = 25705, _                       ; B4JIS paper width in Micrometers.
		$LOW_PAPER_WIDTH_16KAI = 18390, _                       ; 16KAI paper width in Micrometers.
		$LOW_PAPER_WIDTH_32KAI = 13005, _                       ; 32KAI paper width in Micrometers.
		$LOW_PAPER_WIDTH_BIG_32KAI = 13995, _                   ; Big 32KAI paper width in Micrometers.
		$LOW_PAPER_WIDTH_DLENVELOPE = 10998, _                  ; DL Envelope paper width in Micrometers.
		$LOW_PAPER_WIDTH_C6ENVELOPE = 11405, _                  ; C6 Envelope paper width in Micrometers.
		$LOW_PAPER_WIDTH_C6_5_ENVELOPE = 11405, _               ; C6/5 Envelope paper width in Micrometers.
		$LOW_PAPER_WIDTH_C5ENVELOPE = 16205, _                  ; C5 Envelope paper width in Micrometers.
		$LOW_PAPER_WIDTH_C4ENVELOPE = 22911, _                  ; C4 Envelope paper width in Micrometers.
		$LOW_PAPER_WIDTH_6_3_4ENVELOPE = 9208, _                ; 6 3/4 Pound Envelope paper width in Micrometers.
		$LOW_PAPER_WIDTH_7_3_4ENVELOPE = 9855, _                ; 7 3/4 Pound Envelope paper width in Micrometers.
		$LOW_PAPER_WIDTH_9ENVELOPE = 9843, _                    ; 9 Pound Envelope paper width in Micrometers.
		$LOW_PAPER_WIDTH_10ENVELOPE = 10490, _                  ; 10 Pound Envelope paper width in Micrometers.
		$LOW_PAPER_WIDTH_11ENVELOPE = 11430, _                  ; 11 Pound Envelope paper width in Micrometers.
		$LOW_PAPER_WIDTH_12ENVELOPE = 12065, _                  ; 12 Pound Envelope paper width in Micrometers.
		$LOW_PAPER_WIDTH_JAP_POSTCARD = 10008                   ; Japanese Postcard paper width in Micrometers.

; Paragraph Horizontal Align
Global Const _
		$LOW_PAR_ALIGN_HOR_LEFT = 0, _                          ; The Paragraph is left-aligned between the borders.
		$LOW_PAR_ALIGN_HOR_RIGHT = 1, _                         ; The Paragraph is right-aligned between the borders.
		$LOW_PAR_ALIGN_HOR_JUSTIFIED = 2, _                     ; The Paragraph is adjusted / stretched to both borders.
		$LOW_PAR_ALIGN_HOR_CENTER = 3, _                        ; The Paragraph is centered between the left and right borders.
		$LOW_PAR_ALIGN_HOR_STRETCH = 4                          ; HoriAlign 4 does nothing??

; Paragraph Vertical Align
Global Const _
		$LOW_PAR_ALIGN_VERT_AUTO = 0, _                         ; Automatic vertical alignment mode. In automatic mode, horizontal text is aligned to the baseline. The same applies to text that is rotated 90°. Text that is rotated 270 ° is aligned to the center.
		$LOW_PAR_ALIGN_VERT_BASELINE = 1, _                     ; The text is aligned to the baseline.
		$LOW_PAR_ALIGN_VERT_TOP = 2, _                          ; The text is aligned to the top.
		$LOW_PAR_ALIGN_VERT_CENTER = 3, _                       ; The text is aligned to the center.
		$LOW_PAR_ALIGN_VERT_BOTTOM = 4                          ; The text is aligned to bottom.

; Paragraph Last Line Alignment
Global Const _
		$LOW_PAR_LAST_LINE_START = 0, _                         ; The Paragraph is aligned either to the Left border or the right, depending on the current text direction.
		$LOW_PAR_LAST_LINE_JUSTIFIED = 2, _                     ; The Paragraph is adjusted to both borders / stretched.
		$LOW_PAR_LAST_LINE_CENTER = 3                           ; The Paragraph is centered between the left and right borders.

; Path Convert Constants.
Global Const _
		$LOW_PATHCONV_AUTO_RETURN = 0, _                        ; Automatically returns the opposite of the input path, determined by StringInStr search for either "File:///"(L.O.Office URL) or "[A-Z]:\" (Windows File Path).
		$LOW_PATHCONV_OFFICE_RETURN = 1, _                      ; Returns L.O. Office URL, even if the input is already in that format.
		$LOW_PATHCONV_PCPATH_RETURN = 2                         ; Returns Windows File Path, even if the input is already in that format.

; Posture/Italic
Global Const _
		$LOW_POSTURE_NONE = 0, _                                ; Specifies a font without slant.
		$LOW_POSTURE_OBLIQUE = 1, _                             ; Specifies an oblique font (slant not designed into the font).
		$LOW_POSTURE_ITALIC = 2, _                              ; Specifies an italic font (slant designed into the font).
		$LOW_POSTURE_DontKnow = 3, _                            ; Specifies a font with an unknown slant. For Read Only.
		$LOW_POSTURE_REV_OBLIQUE = 4, _                         ; Specifies a reverse oblique font (slant not designed into the font).
		$LOW_POSTURE_REV_ITALIC = 5                             ; Specifies a reverse italic font (slant designed into the font).

; LO Print Comments Constants.
Global Const _
		$LOW_PRINT_NOTES_NONE = 0, _                            ; Document contents are printed, without printing any Comments.
		$LOW_PRINT_NOTES_ONLY = 1, _                            ; Only Comments are printed, and NONE of the Document content.
		$LOW_PRINT_NOTES_END = 2, _                             ; Document content is printed with comments appended to a blank page at the end of the document.
		$LOW_PRINT_NOTES_NEXT_PAGE = 3                          ; Document content is printed and comments are appended to a blank page after the commented page.

; Relative to
Global Const _
		$LOW_RELATIVE_ROW = -1, _                               ; Position an object considering the row height.
		$LOW_RELATIVE_PARAGRAPH = 0, _                          ; The Object is placed considering the available paragraph space, including indent spacing. [Also called "Margin" or "Baseline" in L.O. UI]
		$LOW_RELATIVE_PARAGRAPH_TEXT = 1, _                     ; The Object is placed considering the available paragraph space, excluding indent spacing.
		$LOW_RELATIVE_CHARACTER = 2, _                          ; The Object is placed considering the available character space.
		$LOW_RELATIVE_PAGE_LEFT = 3, _                          ; The Object is placed considering the available space between the left page border and the left Paragraph border. [Same as Left Page Border in L.O. UI]
		$LOW_RELATIVE_PAGE_RIGHT = 4, _                         ; The Object is placed considering the available space between the Right page border and the Right Paragraph border. [Same as Right Page Border in L.O. UI]
		$LOW_RELATIVE_PARAGRAPH_LEFT = 5, _                     ; The Object is placed considering the available indent space to the left of the paragraph.
		$LOW_RELATIVE_PARAGRAPH_RIGHT = 6, _                    ; The Object is placed considering the available indent space to the right of the paragraph.
		$LOW_RELATIVE_PAGE = 7, _                               ; The Object is placed considering the available space between the right and left, or top and bottom page borders.
		$LOW_RELATIVE_PAGE_PRINT = 8, _                         ; The Object is placed considering the available space between the right and left, or top and bottom page margins. [Same as Page Text Area in L.O. UI]
		$LOW_RELATIVE_TEXT_LINE = 9, _                          ; The Object is placed considering the height of the line.
		$LOW_RELATIVE_PAGE_PRINT_BOTTOM = 10, _                 ; The Object is placed considering the space available in the page footer(?)
		$LOW_RELATIVE_PAGE_PRINT_TOP = 11                       ; The Object is placed considering the space available in the page header(?)

; Relief
Global Const _
		$LOW_RELIEF_NONE = 0, _                                 ; No relief is applied.
		$LOW_RELIEF_EMBOSSED = 1, _                             ; The font relief is embossed.
		$LOW_RELIEF_ENGRAVED = 2                                ; The font relief is engraved.

; Shadow
Global Const _
		$LOW_SHADOW_NONE = 0, _                                 ; No shadow is applied.
		$LOW_SHADOW_TOP_LEFT = 1, _                             ; Shadow is located along the upper and left sides.
		$LOW_SHADOW_TOP_RIGHT = 2, _                            ; Shadow is located along the upper and right sides.
		$LOW_SHADOW_BOTTOM_LEFT = 3, _                          ; Shadow is located along the lower and left sides.
		$LOW_SHADOW_BOTTOM_RIGHT = 4                            ; Shadow is located along the lower and right sides.

; Arrowhead Type Constants
Global Enum _
		$LOW_SHAPE_LINE_ARROW_TYPE_NONE, _                      ; 0 -- No Arrow head.
		$LOW_SHAPE_LINE_ARROW_TYPE_ARROW_SHORT, _               ; 1 --Short Arrow head.
		$LOW_SHAPE_LINE_ARROW_TYPE_CONCAVE_SHORT, _             ; 2 -- Short Concave Arrow head.
		$LOW_SHAPE_LINE_ARROW_TYPE_ARROW, _                     ; 3 -- Arrow head.
		$LOW_SHAPE_LINE_ARROW_TYPE_TRIANGLE, _                  ; 4 -- Triangle Arrow head.
		$LOW_SHAPE_LINE_ARROW_TYPE_CONCAVE, _                   ; 5 -- Concave Arrow head.
		$LOW_SHAPE_LINE_ARROW_TYPE_ARROW_LARGE, _               ; 6 -- Large Arrow head.
		$LOW_SHAPE_LINE_ARROW_TYPE_CIRCLE, _                    ; 7 -- Circle Arrow head.
		$LOW_SHAPE_LINE_ARROW_TYPE_SQUARE, _                    ; 8 -- Square Arrow head.
		$LOW_SHAPE_LINE_ARROW_TYPE_SQUARE_45, _                 ; 9 -- Square Arrow head rotated 45 degrees.
		$LOW_SHAPE_LINE_ARROW_TYPE_DIAMOND, _                   ; 10 -- Diamond Arrow head.
		$LOW_SHAPE_LINE_ARROW_TYPE_HALF_CIRCLE, _               ; 11 -- Half Circle Arrow head.
		$LOW_SHAPE_LINE_ARROW_TYPE_DIMENSIONAL_LINES, _         ; 12 -- Dimension Lines head.
		$LOW_SHAPE_LINE_ARROW_TYPE_DIMENSIONAL_LINE_ARROW, _    ; 13 -- Dimension Line Arrow head.
		$LOW_SHAPE_LINE_ARROW_TYPE_DIMENSION_LINE, _            ; 14 -- Dimension Line head.
		$LOW_SHAPE_LINE_ARROW_TYPE_LINE_SHORT, _                ; 15 -- Short Line head.
		$LOW_SHAPE_LINE_ARROW_TYPE_LINE, _                      ; 16 -- Line head.
		$LOW_SHAPE_LINE_ARROW_TYPE_TRIANGLE_UNFILLED, _         ; 17 -- Unfilled Triangle Arrow head.
		$LOW_SHAPE_LINE_ARROW_TYPE_DIAMOND_UNFILLED, _          ; 18 -- Unfilled Diamond Arrow head.
		$LOW_SHAPE_LINE_ARROW_TYPE_CIRCLE_UNFILLED, _           ; 19 -- Unfilled Circle Arrow head.
		$LOW_SHAPE_LINE_ARROW_TYPE_SQUARE_45_UNFILLED, _        ; 20 -- Unfilled Square Arrow head, rotated 45 degrees.
		$LOW_SHAPE_LINE_ARROW_TYPE_SQUARE_UNFILLED, _           ; 21 -- Unfilled Square Arrow head.
		$LOW_SHAPE_LINE_ARROW_TYPE_HALF_CIRCLE_UNFILLED, _      ; 22 -- Unfilled Half Circle Arrow head.
		$LOW_SHAPE_LINE_ARROW_TYPE_HALF_ARROW_LEFT, _           ; 23 -- Half Arrow left Arrow head.
		$LOW_SHAPE_LINE_ARROW_TYPE_HALF_ARROW_RIGHT, _          ; 24 -- Half Arrow right Arrow head.
		$LOW_SHAPE_LINE_ARROW_TYPE_REVERSED_ARROW, _            ; 25 -- Reversed Arrow head.
		$LOW_SHAPE_LINE_ARROW_TYPE_DOUBLE_ARROW, _              ; 26 -- Double Arrow head.
		$LOW_SHAPE_LINE_ARROW_TYPE_CF_ONE, _                    ; 27 -- CF One Arrow head.
		$LOW_SHAPE_LINE_ARROW_TYPE_CF_ONLY_ONE, _               ; 28 -- CF Only One Arrow head.
		$LOW_SHAPE_LINE_ARROW_TYPE_CF_MANY, _                   ; 29 -- CF Many Arrow head.
		$LOW_SHAPE_LINE_ARROW_TYPE_CF_MANY_ONE, _               ; 30 -- CF Many One Arrow head.
		$LOW_SHAPE_LINE_ARROW_TYPE_CF_ZERO_ONE, _               ; 31 -- CF Zero One Arrow head.
		$LOW_SHAPE_LINE_ARROW_TYPE_CF_ZERO_MANY                 ; 32 -- CF Zero Many Arrow head.

; Shape Line End Cap Constants.
Global Const _
		$LOW_SHAPE_LINE_CAP_FLAT = 0, _                         ; Also called Butt, the line will end without any additional shape.
		$LOW_SHAPE_LINE_CAP_ROUND = 1, _                        ; The line will get a half circle as additional cap.
		$LOW_SHAPE_LINE_CAP_SQUARE = 2                          ; The line uses a square for the line end.

; Shape Line Joint Constants.
Global Const _
		$LOW_SHAPE_LINE_JOINT_NONE = 0, _                       ; The joint between lines will not be connected.
		$LOW_SHAPE_LINE_JOINT_MIDDLE = 1, _                     ; The middle value between the joints is used. ## Note used?
		$LOW_SHAPE_LINE_JOINT_BEVEL = 2, _                      ; The edges of the thick lines will be joined by lines.
		$LOW_SHAPE_LINE_JOINT_MITER = 3, _                      ; The lines join at intersections.
		$LOW_SHAPE_LINE_JOINT_ROUND = 4                         ; The lines join with an arc.

; Shape Line Style Constants.
Global Enum _
		$LOW_SHAPE_LINE_STYLE_NONE, _                           ; 0 -- No Line is applied.
		$LOW_SHAPE_LINE_STYLE_CONTINUOUS, _                     ; 1 -- A Solid Line.
		$LOW_SHAPE_LINE_STYLE_DOT, _                            ; 2 -- A Dotted Line.
		$LOW_SHAPE_LINE_STYLE_DOT_ROUNDED, _                    ; 3 -- A Rounded Dotted Line.
		$LOW_SHAPE_LINE_STYLE_LONG_DOT, _                       ; 4 -- A Long Dotted Line.
		$LOW_SHAPE_LINE_STYLE_LONG_DOT_ROUNDED, _               ; 5 -- A Rounded Long Dotted Line.
		$LOW_SHAPE_LINE_STYLE_DASH, _                           ; 6 -- A Dashed Line.
		$LOW_SHAPE_LINE_STYLE_DASH_ROUNDED, _                   ; 7 -- A Rounded Dashed Line.
		$LOW_SHAPE_LINE_STYLE_LONG_DASH, _                      ; 8 -- A Long Dashed Line.
		$LOW_SHAPE_LINE_STYLE_LONG_DASH_ROUNDED, _              ; 9 -- A Rounded Long Dashed Line.
		$LOW_SHAPE_LINE_STYLE_DOUBLE_DASH, _                    ; 10 -- A Double Dashed Line.
		$LOW_SHAPE_LINE_STYLE_DOUBLE_DASH_ROUNDED, _            ; 11 -- A Rounded Double Dash.
		$LOW_SHAPE_LINE_STYLE_DASH_DOT, _                       ; 12 -- A Dashed and Dotted Line.
		$LOW_SHAPE_LINE_STYLE_DASH_DOT_ROUNDED, _               ; 13 -- A Rounded Dashed and Dotted Line.
		$LOW_SHAPE_LINE_STYLE_LONG_DASH_DOT, _                  ; 14 -- A Long Dashed and Dotted Line.
		$LOW_SHAPE_LINE_STYLE_LONG_DASH_DOT_ROUNDED, _          ; 15 -- A Rounded Long Dashed and Dotted Line.
		$LOW_SHAPE_LINE_STYLE_DOUBLE_DASH_DOT, _                ; 16 -- A Double Dash Dot Line.
		$LOW_SHAPE_LINE_STYLE_DOUBLE_DASH_DOT_ROUNDED, _        ; 17 -- A Rounded Double Dash Dot Line
		$LOW_SHAPE_LINE_STYLE_DASH_DOT_DOT, _                   ; 18 -- A Dash Dot Dot Line.
		$LOW_SHAPE_LINE_STYLE_DASH_DOT_DOT_ROUNDED, _           ; 19 -- A Rounded Dash Dot Dot Line.
		$LOW_SHAPE_LINE_STYLE_DOUBLE_DASH_DOT_DOT, _            ; 20 -- A Double Dash Dot Dot Line.
		$LOW_SHAPE_LINE_STYLE_DOUBLE_DASH_DOT_DOT_ROUNDED, _    ; 21 -- A Rounded Double Dash Dot Dot Line.
		$LOW_SHAPE_LINE_STYLE_ULTRAFINE_DOTTED, _               ; 22 -- A Ultrafine Dotted Line.
		$LOW_SHAPE_LINE_STYLE_FINE_DOTTED, _                    ; 23 -- A Fine Dotted Line.
		$LOW_SHAPE_LINE_STYLE_ULTRAFINE_DASHED, _               ; 24 -- A Ultrafine Dashed Line.
		$LOW_SHAPE_LINE_STYLE_FINE_DASHED, _                    ; 25 -- A Fine Dashed Line.
		$LOW_SHAPE_LINE_STYLE_DASHED, _                         ; 26 -- A Dashed Line.
		$LOW_SHAPE_LINE_STYLE_LINE_STYLE_9, _                   ; 27 -- Line Style 9.
		$LOW_SHAPE_LINE_STYLE_3_DASHES_3_DOTS, _                ; 28 -- A Line consisting of 3 Dashes and 3 Dots.
		$LOW_SHAPE_LINE_STYLE_ULTRAFINE_2_DOTS_3_DASHES, _      ; 29 -- A Ultrafine Line consisting of 2 Dots and 3 Dashes.
		$LOW_SHAPE_LINE_STYLE_2_DOTS_1_DASH, _                  ; 30 -- A Line consisting of 2 Dots and 1 Dash.
		$LOW_SHAPE_LINE_STYLE_LINE_WITH_FINE_DOTS               ; 31 -- A Line with Fine Dots.

; Polygon Flags
Global Const _
		$LOW_SHAPE_POINT_TYPE_NORMAL = 0, _                     ; the point is normal, from the curve discussion view.
		$LOW_SHAPE_POINT_TYPE_SMOOTH = 1, _                     ; the point is smooth, the first derivation from the curve discussion view.
		$LOW_SHAPE_POINT_TYPE_CONTROL = 2, _                    ; the point is a control point, to control the curve from the user interface.
		$LOW_SHAPE_POINT_TYPE_SYMMETRIC = 3                     ; the point is symmetric, the second derivation from the curve discussion view.

; Shape Type Constants.
Global Enum _
		$LOW_SHAPE_TYPE_ARROWS_ARROW_4_WAY, _                   ; 0 -- A Four-way Arrow.
		$LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_4_WAY, _           ; 1 -- A Four-way Callout Arrow.
		$LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_DOWN, _            ; 2 -- A Downward Callout Arrow.
		$LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_LEFT, _            ; 3 -- A Left hand Callout Arrow.
		$LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_LEFT_RIGHT, _      ; 4 -- A Left and Right Callout Arrow.
		$LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_RIGHT, _           ; 5 -- A Right hand Callout Arrow.
		$LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP, _              ; 6 -- A Upward Callout Arrow.
		$LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_DOWN, _         ; 7 -- A Upward and Downward Callout Arrow.
		$LOW_SHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_RIGHT, _        ; 8 -- Upward and Right hand Callout Arrow. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOW_SHAPE_TYPE_ARROWS_ARROW_CIRCULAR, _                ; 9 -- A Circular Arrow.
		$LOW_SHAPE_TYPE_ARROWS_ARROW_CORNER_RIGHT, _            ; 10 -- A Right hand Corner Arrow. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOW_SHAPE_TYPE_ARROWS_ARROW_DOWN, _                    ; 11 -- A Downward Arrow.
		$LOW_SHAPE_TYPE_ARROWS_ARROW_LEFT, _                    ; 12 -- A Left hand Arrow.
		$LOW_SHAPE_TYPE_ARROWS_ARROW_LEFT_RIGHT, _              ; 13 -- A Left and Right Arrow.
		$LOW_SHAPE_TYPE_ARROWS_ARROW_NOTCHED_RIGHT, _           ; 14 -- A Notched Right Arrow.
		$LOW_SHAPE_TYPE_ARROWS_ARROW_RIGHT, _                   ; 15 -- A Right hand Arrow.
		$LOW_SHAPE_TYPE_ARROWS_ARROW_RIGHT_OR_LEFT, _           ; 16 -- A Right or Left Arrow. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOW_SHAPE_TYPE_ARROWS_ARROW_S_SHAPED, _                ; 17 -- A "S"-Shaped Arrow. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOW_SHAPE_TYPE_ARROWS_ARROW_SPLIT, _                   ; 18 -- A Split Arrow. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOW_SHAPE_TYPE_ARROWS_ARROW_STRIPED_RIGHT, _           ; 19 -- A Striped Arrow.
		$LOW_SHAPE_TYPE_ARROWS_ARROW_UP, _                      ; 20 -- A Upward Arrow.
		$LOW_SHAPE_TYPE_ARROWS_ARROW_UP_DOWN, _                 ; 21 -- A Up and Down Arrow.
		$LOW_SHAPE_TYPE_ARROWS_ARROW_UP_RIGHT, _                ; 22 -- A Upward and Right hand Arrow.
		$LOW_SHAPE_TYPE_ARROWS_ARROW_UP_RIGHT_DOWN, _           ; 23 -- A Upward, Right hand and Downward Arrow. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOW_SHAPE_TYPE_ARROWS_CHEVRON, _                       ; 24 -- A Chevron Shape Arrow.
		$LOW_SHAPE_TYPE_ARROWS_PENTAGON, _                      ; 25 -- A Pentagon Shape Arrow.
		$LOW_SHAPE_TYPE_BASIC_ARC, _                            ; 26 -- An Arc Shape.
		$LOW_SHAPE_TYPE_BASIC_ARC_BLOCK, _                      ; 27 -- A Block Arc Shape.
		$LOW_SHAPE_TYPE_BASIC_CIRCLE, _                         ; 28 -- A Circle.
		$LOW_SHAPE_TYPE_BASIC_CIRCLE_PIE, _                     ; 29 -- A Pie Circle. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOW_SHAPE_TYPE_BASIC_CIRCLE_SEGMENT, _                 ; 30 -- A Segment Circle.
		$LOW_SHAPE_TYPE_BASIC_CROSS, _                          ; 31 -- A Cross Shape.
		$LOW_SHAPE_TYPE_BASIC_CUBE, _                           ; 32 -- A Cube Shape.
		$LOW_SHAPE_TYPE_BASIC_CYLINDER, _                       ; 33 -- A Cylinder Shape.
		$LOW_SHAPE_TYPE_BASIC_DIAMOND, _                        ; 34 -- A Diamond Shape.
		$LOW_SHAPE_TYPE_BASIC_ELLIPSE, _                        ; 35 -- An Ellipse Shape.
		$LOW_SHAPE_TYPE_BASIC_FOLDED_CORNER, _                  ; 36 -- A Paper Shape with a Folded Corner.
		$LOW_SHAPE_TYPE_BASIC_FRAME, _                          ; 37 -- A Frame Shape. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOW_SHAPE_TYPE_BASIC_HEXAGON, _                        ; 38 -- A Hexagon Shape.
		$LOW_SHAPE_TYPE_BASIC_OCTAGON, _                        ; 39 -- A Octagon Shape.
		$LOW_SHAPE_TYPE_BASIC_PARALLELOGRAM, _                  ; 40 -- A Parallelogram Shape.
		$LOW_SHAPE_TYPE_BASIC_RECTANGLE, _                      ; 41 -- A Rectangle.
		$LOW_SHAPE_TYPE_BASIC_RECTANGLE_ROUNDED, _              ; 42 -- A Rectangle with rounded corners.
		$LOW_SHAPE_TYPE_BASIC_REGULAR_PENTAGON, _               ; 43 -- A regular Pentagon.
		$LOW_SHAPE_TYPE_BASIC_RING, _                           ; 44 -- A Ring Shape.
		$LOW_SHAPE_TYPE_BASIC_SQUARE, _                         ; 45 -- A Square.
		$LOW_SHAPE_TYPE_BASIC_SQUARE_ROUNDED, _                 ; 46 -- A Square with rounded corners.
		$LOW_SHAPE_TYPE_BASIC_TRAPEZOID, _                      ; 47 -- A Trapezoid Shape.
		$LOW_SHAPE_TYPE_BASIC_TRIANGLE_ISOSCELES, _             ; 48 -- An Isosceles Triangle.
		$LOW_SHAPE_TYPE_BASIC_TRIANGLE_RIGHT, _                 ; 49 -- A Right Angle Triangle.
		$LOW_SHAPE_TYPE_CALLOUT_CLOUD, _                        ; 50 -- A Cloud Shaped Callout.
		$LOW_SHAPE_TYPE_CALLOUT_LINE_1, _                       ; 51 -- A Callout with Line style #1.
		$LOW_SHAPE_TYPE_CALLOUT_LINE_2, _                       ; 52 -- A Callout with Line style #2.
		$LOW_SHAPE_TYPE_CALLOUT_LINE_3, _                       ; 53 -- A Callout with Line style #3.
		$LOW_SHAPE_TYPE_CALLOUT_RECTANGULAR, _                  ; 54 -- A Rectangular Callout.
		$LOW_SHAPE_TYPE_CALLOUT_RECTANGULAR_ROUNDED, _          ; 55 -- A Rectangular Callout with rounded corners.
		$LOW_SHAPE_TYPE_CALLOUT_ROUND, _                        ; 56 -- A Round Callout.
		$LOW_SHAPE_TYPE_FLOWCHART_CARD, _                       ; 57 -- A Card Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_COLLATE, _                    ; 58 -- A Collate Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_CONNECTOR, _                  ; 59 -- A Connector Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_CONNECTOR_OFF_PAGE, _         ; 60 -- A Off-Page Connector Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_DATA, _                       ; 61 -- A Data Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_DECISION, _                   ; 62 -- A Decision Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_DELAY, _                      ; 63 -- A Delay Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_DIRECT_ACCESS_STORAGE, _      ; 64 -- A Direct Access Storage Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_DISPLAY, _                    ; 65 -- A Display Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_DOCUMENT, _                   ; 66 -- A Document Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_EXTRACT, _                    ; 67 -- A Extract Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_INTERNAL_STORAGE, _           ; 68 -- A Internal Storage Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_MAGNETIC_DISC, _              ; 69 -- A Magnetic Disc Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_MANUAL_INPUT, _               ; 70 -- A Manual Input Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_MANUAL_OPERATION, _           ; 71 -- A Manual Operation Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_MERGE, _                      ; 72 -- A Merge Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_MULTIDOCUMENT, _              ; 73 -- A Multi-document Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_OR, _                         ; 74 -- A Or Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_PREPARATION, _                ; 75 -- A Preparation Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_PROCESS, _                    ; 76 -- A Process Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_PROCESS_ALTERNATE, _          ; 77 -- A Alternate Process Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_PROCESS_PREDEFINED, _         ; 78 -- A Predefined Process Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_PUNCHED_TAPE, _               ; 79 -- A Punched Tape Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_SEQUENTIAL_ACCESS, _          ; 80 -- A Sequential Access Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_SORT, _                       ; 81 -- A Sort Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_STORED_DATA, _                ; 82 -- A Stored Data Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_SUMMING_JUNCTION, _           ; 83 -- A Summing Junction Flowchart.
		$LOW_SHAPE_TYPE_FLOWCHART_TERMINATOR, _                 ; 84 -- A Terminator Flowchart.
		$LOW_SHAPE_TYPE_LINE_CURVE, _                           ; 85 -- A Curve.
		$LOW_SHAPE_TYPE_LINE_CURVE_FILLED, _                    ; 86 -- A Filled Curve.
		$LOW_SHAPE_TYPE_LINE_FREEFORM_LINE, _                   ; 87 -- A Freeform Line.
		$LOW_SHAPE_TYPE_LINE_FREEFORM_LINE_FILLED, _            ; 88 -- A Filled Freeform Line.
		$LOW_SHAPE_TYPE_LINE_LINE, _                            ; 89 -- A Line.
		$LOW_SHAPE_TYPE_LINE_POLYGON, _                         ; 90 -- A Polygon.
		$LOW_SHAPE_TYPE_LINE_POLYGON_45, _                      ; 91 -- A 45 degree Polygon.
		$LOW_SHAPE_TYPE_LINE_POLYGON_45_FILLED, _               ; 92 -- A Filled 45 degree Polygon.
		$LOW_SHAPE_TYPE_STARS_4_POINT, _                        ; 93 -- A 4 Pointed Star.
		$LOW_SHAPE_TYPE_STARS_5_POINT, _                        ; 94 -- A 5 Pointed Star.
		$LOW_SHAPE_TYPE_STARS_6_POINT, _                        ; 95 -- A 6 Pointed Star. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOW_SHAPE_TYPE_STARS_6_POINT_CONCAVE, _                ; 96 -- A Concave 6 Pointed Star. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOW_SHAPE_TYPE_STARS_8_POINT, _                        ; 97 -- A 8 Pointed Star.
		$LOW_SHAPE_TYPE_STARS_12_POINT, _                       ; 98 -- A 12 Pointed Star. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOW_SHAPE_TYPE_STARS_24_POINT, _                       ; 99 -- A 24 Pointed Star.
		$LOW_SHAPE_TYPE_STARS_DOORPLATE, _                      ; 100 -- A Doorplate Shape.
		$LOW_SHAPE_TYPE_STARS_EXPLOSION, _                      ; 101 -- A Explosion Shape.
		$LOW_SHAPE_TYPE_STARS_SCROLL_HORIZONTAL, _              ; 102 -- A Horizontal Scroll.
		$LOW_SHAPE_TYPE_STARS_SCROLL_VERTICAL, _                ; 103 -- A Vertical Scroll.
		$LOW_SHAPE_TYPE_STARS_SIGNET, _                         ; 104 -- A Signet Shape. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOW_SHAPE_TYPE_SYMBOL_BEVEL_DIAMOND, _                 ; 105 -- A Diamond Bevel. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOW_SHAPE_TYPE_SYMBOL_BEVEL_OCTAGON, _                 ; 106 -- A Octagon Bevel. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOW_SHAPE_TYPE_SYMBOL_BEVEL_SQUARE, _                  ; 107 -- A Square Bevel.
		$LOW_SHAPE_TYPE_SYMBOL_BRACE_DOUBLE, _                  ; 108 -- A Double Brace.
		$LOW_SHAPE_TYPE_SYMBOL_BRACE_LEFT, _                    ; 109 -- A Left hand Brace.
		$LOW_SHAPE_TYPE_SYMBOL_BRACE_RIGHT, _                   ; 110 -- A Right hand Brace.
		$LOW_SHAPE_TYPE_SYMBOL_BRACKET_DOUBLE, _                ; 111 -- A Double Bracket.
		$LOW_SHAPE_TYPE_SYMBOL_BRACKET_LEFT, _                  ; 112 -- A Left hand Bracket.
		$LOW_SHAPE_TYPE_SYMBOL_BRACKET_RIGHT, _                 ; 113 -- A Right hand Bracket.
		$LOW_SHAPE_TYPE_SYMBOL_CLOUD, _                         ; 114 -- A Cloud Shape. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOW_SHAPE_TYPE_SYMBOL_FLOWER, _                        ; 115 -- A Flower Shape. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOW_SHAPE_TYPE_SYMBOL_HEART, _                         ; 116 -- A Heart Shape.
		$LOW_SHAPE_TYPE_SYMBOL_LIGHTNING, _                     ; 117 -- A Lightning Shape. ## Note: Lightning is visually different than the one available in L.O. Shapes U.I.
		$LOW_SHAPE_TYPE_SYMBOL_MOON, _                          ; 118 -- A Moon Shape.
		$LOW_SHAPE_TYPE_SYMBOL_SMILEY, _                        ; 119 -- A Smiley Shape.
		$LOW_SHAPE_TYPE_SYMBOL_SUN, _                           ; 120 -- A Sun Shape.
		$LOW_SHAPE_TYPE_SYMBOL_PROHIBITED, _                    ; 121 -- A Prohibited Shape.
		$LOW_SHAPE_TYPE_SYMBOL_PUZZLE                           ; 122 -- A Puzzle Piece Shape. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.

; Strikeout
Global Const _
		$LOW_STRIKEOUT_NONE = 0, _                              ; No strike out.
		$LOW_STRIKEOUT_SINGLE = 1, _                            ; Strike out the characters with a single line.
		$LOW_STRIKEOUT_DOUBLE = 2, _                            ; Strike out the characters with a double line.
		$LOW_STRIKEOUT_DONT_KNOW = 3, _                         ; The strikeout mode is not specified. For Read Only.
		$LOW_STRIKEOUT_BOLD = 4, _                              ; Strike out the characters with a bold line.
		$LOW_STRIKEOUT_SLASH = 5, _                             ; Strike out the characters with slashes.
		$LOW_STRIKEOUT_X = 6                                    ; Strike out the characters with X's.

; Tab Alignment
Global Const _
		$LOW_TAB_ALIGN_LEFT = 0, _                              ; Aligns the left edge of the text to the tab stop and extends the text to the right.
		$LOW_TAB_ALIGN_CENTER = 1, _                            ; Aligns the center of the text to the tab stop.
		$LOW_TAB_ALIGN_RIGHT = 2, _                             ; Aligns the right edge of the text to the tab stop and extends the text to the left of the tab stop.
		$LOW_TAB_ALIGN_DECIMAL = 3, _                           ; Aligns the decimal separator of a number to the center of the tab stop and text to the left of the tab.
		$LOW_TAB_ALIGN_DEFAULT = 4                              ; This setting is the default setting when no TabStops are present. Setting any Tabstop to this constant will make it disappear from the TabStop list. It is therefore only listed here for property reading purposes.

; LO TableCursor Movement Constants.
Global Enum _
		$LOW_TABLECUR_GO_LEFT, _                                ; Move the cursor left n cells.
		$LOW_TABLECUR_GO_RIGHT, _                               ; Move the cursor right n cells.
		$LOW_TABLECUR_GO_UP, _                                  ; Move the cursor up n cells.
		$LOW_TABLECUR_GO_DOWN, _                                ; Move the cursor down n cells.
		$LOW_TABLECUR_GOTO_START, _                             ; Move the cursor to the first cell.
		$LOW_TABLECUR_GOTO_END                                  ; Move the cursor to the last cell.

; LO TextCursor Movement Constants.
Global Enum _
		$LOW_TEXTCUR_COLLAPSE_TO_START, _                       ; Collapses the current selection to the start of the selection.
		$LOW_TEXTCUR_COLLAPSE_TO_END, _                         ; Collapses the current selection the to end of the selection.
		$LOW_TEXTCUR_GO_LEFT, _                                 ; Move the cursor left by n characters.
		$LOW_TEXTCUR_GO_RIGHT, _                                ; Move the cursor right by n characters.
		$LOW_TEXTCUR_GOTO_START, _                              ; Move the cursor to the start of the text.
		$LOW_TEXTCUR_GOTO_END, _                                ; Move the cursor to the end of the text.
		$LOW_TEXTCUR_GOTO_NEXT_WORD, _                          ; Move to the start of the next word.
		$LOW_TEXTCUR_GOTO_PREV_WORD, _                          ; Move to the end of the previous word.
		$LOW_TEXTCUR_GOTO_END_OF_WORD, _                        ; Move to the end of the current word.
		$LOW_TEXTCUR_GOTO_START_OF_WORD, _                      ; Move to the start of the current word.
		$LOW_TEXTCUR_GOTO_NEXT_SENTENCE, _                      ; Move to the start of the next sentence.
		$LOW_TEXTCUR_GOTO_PREV_SENTENCE, _                      ; Move to the end of the previous sentence.
		$LOW_TEXTCUR_GOTO_END_OF_SENTENCE, _                    ; Move to the end of the current sentence.
		$LOW_TEXTCUR_GOTO_START_OF_SENTENCE, _                  ; Move to the start of the current sentence.
		$LOW_TEXTCUR_GOTO_NEXT_PARAGRAPH, _                     ; Move to the start of the next paragraph.
		$LOW_TEXTCUR_GOTO_PREV_PARAGRAPH, _                     ; Move to the End of the previous paragraph.
		$LOW_TEXTCUR_GOTO_END_OF_PARAGRAPH, _                   ; Move to the end of the current paragraph.
		$LOW_TEXTCUR_GOTO_START_OF_PARAGRAPH                    ; Move to the start of the current paragraph.

; Text Adjust
Global Const _
		$LOW_TXT_ADJ_VERT_TOP = 0, _                            ; The top edge of the text is adjusted to the top edge of the object.
		$LOW_TXT_ADJ_VERT_CENTER = 1, _                         ; The text is centered inside the object.
		$LOW_TXT_ADJ_VERT_BOTTOM = 2, _                         ; The bottom edge of the text is adjusted to the bottom edge of the object.
		$LOW_TXT_ADJ_VERT_BLOCK = 3                             ;

; Text Direction
Global Const _
		$LOW_TXT_DIR_LR_TB = 0, _                               ; Text within lines is written left-to-right. Lines and blocks are placed top-to-bottom. Typically, this is the writing mode for normal "alphabetic" text.
		$LOW_TXT_DIR_RL_TB = 1, _                               ; Text within a line are written right-to-left. Lines and blocks are placed top-to-bottom. Typically, this writing mode is used in Arabic and Hebrew text.
		$LOW_TXT_DIR_TB_RL = 2, _                               ; Text within a line is written top-to-bottom. Lines and blocks are placed right-to-left. Typically, this writing mode is used in Chinese and Japanese text.
		$LOW_TXT_DIR_TB_LR = 3, _                               ; Text within a line is written top-to-bottom. Lines and blocks are placed left-to-right. Typically, this writing mode is used in Mongolian text.
		$LOW_TXT_DIR_CONTEXT = 4, _                             ; Obtain actual writing mode from the context of the object.
		$LOW_TXT_DIR_BT_LR = 5                                  ; text within a line is written bottom-to-top. Lines and blocks are placed left-to-right. (LibreOffice 6.3).

; Underline/Overline
Global Const _
		$LOW_UNDERLINE_NONE = 0, _                              ; No Underline or Overline style.
		$LOW_UNDERLINE_SINGLE = 1, _                            ; Single line Underline/Overline style.
		$LOW_UNDERLINE_DOUBLE = 2, _                            ; Double line Underline/Overline style.
		$LOW_UNDERLINE_DOTTED = 3, _                            ; Dotted line Underline/Overline style.
		$LOW_UNDERLINE_DONT_KNOW = 4, _                         ; Unknown Underline/Overline style, for read only.
		$LOW_UNDERLINE_DASH = 5, _                              ; Dashed line Underline/Overline style.
		$LOW_UNDERLINE_LONG_DASH = 6, _                         ; Long Dashed line Underline/Overline style.
		$LOW_UNDERLINE_DASH_DOT = 7, _                          ; Dash Dot line Underline/Overline style.
		$LOW_UNDERLINE_DASH_DOT_DOT = 8, _                      ; Dash Dot Dot line Underline/Overline style.
		$LOW_UNDERLINE_SML_WAVE = 9, _                          ; Small Wave line Underline/Overline style.
		$LOW_UNDERLINE_WAVE = 10, _                             ; Wave line Underline/Overline style.
		$LOW_UNDERLINE_DBL_WAVE = 11, _                         ; Double Wave line Underline/Overline style.
		$LOW_UNDERLINE_BOLD = 12, _                             ; Bold line Underline/Overline style.
		$LOW_UNDERLINE_BOLD_DOTTED = 13, _                      ; Bold Dotted line Underline/Overline style.
		$LOW_UNDERLINE_BOLD_DASH = 14, _                        ; Bold Dashed line Underline/Overline style.
		$LOW_UNDERLINE_BOLD_LONG_DASH = 15, _                   ; Bold Long Dash line Underline/Overline style.
		$LOW_UNDERLINE_BOLD_DASH_DOT = 16, _                    ; Bold Dash Dot line Underline/Overline style.
		$LOW_UNDERLINE_BOLD_DASH_DOT_DOT = 17, _                ; Bold Dash Dot Dot line Underline/Overline style.
		$LOW_UNDERLINE_BOLD_WAVE = 18                           ; Bold Wave line Underline/Overline style.

; LO ViewCursor Movement Constants.
Global Enum _
		$LOW_VIEWCUR_GO_DOWN, _                                 ; Move the cursor Down by n lines.
		$LOW_VIEWCUR_GO_UP, _                                   ; Move the cursor Up by n lines.
		$LOW_VIEWCUR_GO_LEFT, _                                 ; Move the cursor left by n characters.
		$LOW_VIEWCUR_GO_RIGHT, _                                ; Move the cursor right by n characters.
		$LOW_VIEWCUR_GOTO_END_OF_LINE, _                        ; Move the cursor to the end of the current line.
		$LOW_VIEWCUR_GOTO_START_OF_LINE, _                      ; Move the cursor to the start of the current line.
		$LOW_VIEWCUR_JUMP_TO_FIRST_PAGE, _                      ; Move the cursor to the first page.
		$LOW_VIEWCUR_JUMP_TO_LAST_PAGE, _                       ; Move the cursor to the Last page.
		$LOW_VIEWCUR_JUMP_TO_PAGE, _                            ; Jump to a specified page.
		$LOW_VIEWCUR_JUMP_TO_NEXT_PAGE, _                       ; Move the cursor to the Next page.
		$LOW_VIEWCUR_JUMP_TO_PREV_PAGE, _                       ; Move the cursor to the previous page.
		$LOW_VIEWCUR_JUMP_TO_END_OF_PAGE, _                     ; Move the cursor to the end of the current page.
		$LOW_VIEWCUR_JUMP_TO_START_OF_PAGE, _                   ; Move the cursor to the start of the current page.
		$LOW_VIEWCUR_SCREEN_DOWN, _                             ; Scroll the view forward by one visible page.
		$LOW_VIEWCUR_SCREEN_UP, _                               ; Scroll the view back by one visible page.
		$LOW_VIEWCUR_GOTO_START, _                              ; Move the cursor to the start of the document or Table.
		$LOW_VIEWCUR_GOTO_END                                   ; Move the cursor to the end of the document or Table.

; Weight/Bold
Global Const _
		$LOW_WEIGHT_DONT_KNOW = 0, _                            ; The font weight is not specified/unknown. For Read Only.
		$LOW_WEIGHT_THIN = 50, _                                ; A 50% (Thin) font weight.
		$LOW_WEIGHT_ULTRA_LIGHT = 60, _                         ; A 60% (Ultra Light) font weight.
		$LOW_WEIGHT_LIGHT = 75, _                               ; A 75% (Light) font weight.
		$LOW_WEIGHT_SEMI_LIGHT = 90, _                          ; A 90% (Semi-Light) font weight.
		$LOW_WEIGHT_NORMAL = 100, _                             ; A 100% (Normal) font weight.
		$LOW_WEIGHT_SEMI_BOLD = 110, _                          ; A 110% (Semi-Bold) font weight.
		$LOW_WEIGHT_BOLD = 150, _                               ; A 150% (Bold) font weight.
		$LOW_WEIGHT_ULTRA_BOLD = 175, _                         ; A 175% (Ultra-Bold) font weight.
		$LOW_WEIGHT_BLACK = 200                                 ; A 200% (Black) font weight.

; Wrap Type
Global Const _
		$LOW_WRAP_MODE_NONE = 0, _                              ; Places the Object on a separate line in the document.
		$LOW_WRAP_MODE_THROUGH = 1, _                           ; Places the Object in front of the text.
		$LOW_WRAP_MODE_PARALLEL = 2, _                          ; Wraps text on all four sides of the border frame of the Object. [Same as "Optimal"]
		$LOW_WRAP_MODE_DYNAMIC = 3, _                           ; Automatically wraps text to the left, to the right, or on all four sides of the border of the object. [Same as "Before"]
		$LOW_WRAP_MODE_LEFT = 4, _                              ; Wraps text on the left side of the object. [Same as "After"]
		$LOW_WRAP_MODE_RIGHT = 5                                ; Wraps text on the right side of the object.
