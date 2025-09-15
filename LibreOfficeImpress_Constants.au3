#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel
#include-once

; #INDEX# =======================================================================================================================
; Title .........: Libre Office Impress Constants for the Libre Office UDF.
; AutoIt Version : v3.3.16.1
; Description ...: Constants for various functions in the Libre Office UDF.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
; Note ..........: Descriptions for some Constants are taken from the LibreOffice SDK API documentation.
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; ===============================================================================================================================

; Sleep Divisor $__LOICONST_SLEEP_DIV
; In applicable functions this is used for adjusting how frequent a sleep occurs in loops.
; For any number above 0 the number of times a loop has completed is divided by $__LOICONST_SLEEP_DIV. If you find some functions cause momentary freeze ups, a recommended value is 15.
; Set to 0 for no pause in a loop.
Global Const $__LOICONST_SLEEP_DIV = 0

#Tidy_ILC_Pos=65

; Fill Style Type Constants
Global Enum _                                                   ; com.sun.star.drawing.FillStyle
		$LOI_AREA_FILL_STYLE_OFF, _                             ; 0 Fill Style is off.
		$LOI_AREA_FILL_STYLE_SOLID, _                           ; 1 Fill Style is a solid color.
		$LOI_AREA_FILL_STYLE_GRADIENT, _                        ; 2 Fill Style is a gradient color.
		$LOI_AREA_FILL_STYLE_HATCH, _                           ; 3 Fill Style is a Hatch style color.
		$LOI_AREA_FILL_STYLE_BITMAP                             ; 4 Fill Style is a Bitmap.

; Drawing Shape Type Constants.
Global Enum _
		$LOI_DRAWSHAPE_TYPE_3D_CONE, _                          ; 0 -- A 3D Cone.
		$LOI_DRAWSHAPE_TYPE_3D_CUBE, _                          ; 1 -- A 3D Cube.
		$LOI_DRAWSHAPE_TYPE_3D_CYLINDER, _                      ; 2 -- A 3D Cylinder.
		$LOI_DRAWSHAPE_TYPE_3D_HALF_SPHERE, _                   ; 3 -- A 3D Half-Sphere.
		$LOI_DRAWSHAPE_TYPE_3D_PYRAMID, _                       ; 4 -- A 3D Pyramid.
		$LOI_DRAWSHAPE_TYPE_3D_SHELL, _                         ; 5 -- A 3D Shell.
		$LOI_DRAWSHAPE_TYPE_3D_SPHERE, _                        ; 6 -- A 3D Sphere.
		$LOI_DRAWSHAPE_TYPE_3D_TORUS, _                         ; 7 -- A 3D Torus.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_4_WAY, _               ; 8 -- A Four-way Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_4_WAY, _       ; 9 -- A Four-way Callout Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_DOWN, _        ; 10 -- A Downward Callout Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_LEFT, _        ; 11 -- A Left hand Callout Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_LEFT_RIGHT, _  ; 12 -- A Left and Right Callout Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_RIGHT, _       ; 13 -- A Right hand Callout Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP, _          ; 14 -- A Upward Callout Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_DOWN, _     ; 15 -- A Upward and Downward Callout Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_RIGHT, _    ; 16 -- Upward and Right hand Callout Arrow. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CIRCULAR, _            ; 17 -- A Circular Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CORNER_RIGHT, _        ; 18 -- A Right hand Corner Arrow. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_DOWN, _                ; 19 -- A Downward Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_LEFT, _                ; 20 -- A Left hand Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_LEFT_RIGHT, _          ; 21 -- A Left and Right Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_NOTCHED_RIGHT, _       ; 22 -- A Notched Right Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_RIGHT, _               ; 23 -- A Right hand Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_RIGHT_OR_LEFT, _       ; 24 -- A Right or Left Arrow. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_S_SHAPED, _            ; 25 -- A "S"-Shaped Arrow. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_SPLIT, _               ; 26 -- A Split Arrow. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_STRIPED_RIGHT, _       ; 27 -- A Striped Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP, _                  ; 28 -- A Upward Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_DOWN, _             ; 29 -- A Up and Down Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT, _            ; 30 -- A Upward and Right hand Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT_DOWN, _       ; 31 -- A Upward, Right hand and Downward Arrow. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_ARROWS_CHEVRON, _                   ; 32 -- A Chevron Shape Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_PENTAGON, _                  ; 33 -- A Pentagon Shape Arrow.
		$LOI_DRAWSHAPE_TYPE_BASIC_ARC, _                        ; 34 -- An Arc Shape.
		$LOI_DRAWSHAPE_TYPE_BASIC_ARC_BLOCK, _                  ; 35 -- A Block Arc Shape.
		$LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE, _                     ; 36 -- A Circle.
		$LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_PIE, _                 ; 37 -- A Pie Circle. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_SEGMENT, _             ; 38 -- A Segment Circle.
		$LOI_DRAWSHAPE_TYPE_BASIC_CROSS, _                      ; 39 -- A Cross Shape.
		$LOI_DRAWSHAPE_TYPE_BASIC_CUBE, _                       ; 40 -- A Cube Shape.
		$LOI_DRAWSHAPE_TYPE_BASIC_CYLINDER, _                   ; 41 -- A Cylinder Shape.
		$LOI_DRAWSHAPE_TYPE_BASIC_DIAMOND, _                    ; 42 -- A Diamond Shape.
		$LOI_DRAWSHAPE_TYPE_BASIC_ELLIPSE, _                    ; 43 -- An Ellipse Shape.
		$LOI_DRAWSHAPE_TYPE_BASIC_FOLDED_CORNER, _              ; 44 -- A Paper Shape with a Folded Corner.
		$LOI_DRAWSHAPE_TYPE_BASIC_FRAME, _                      ; 45 -- A Frame Shape. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_BASIC_HEXAGON, _                    ; 46 -- A Hexagon Shape.
		$LOI_DRAWSHAPE_TYPE_BASIC_OCTAGON, _                    ; 47 -- A Octagon Shape.
		$LOI_DRAWSHAPE_TYPE_BASIC_PARALLELOGRAM, _              ; 48 -- A Parallelogram Shape.
		$LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE, _                  ; 49 -- A Rectangle.
		$LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE_ROUNDED, _          ; 50 -- A Rectangle with rounded corners.
		$LOI_DRAWSHAPE_TYPE_BASIC_REGULAR_PENTAGON, _           ; 51 -- A regular Pentagon.
		$LOI_DRAWSHAPE_TYPE_BASIC_RING, _                       ; 52 -- A Ring Shape.
		$LOI_DRAWSHAPE_TYPE_BASIC_SQUARE, _                     ; 53 -- A Square.
		$LOI_DRAWSHAPE_TYPE_BASIC_SQUARE_ROUNDED, _             ; 54 -- A Square with rounded corners.
		$LOI_DRAWSHAPE_TYPE_BASIC_TRAPEZOID, _                  ; 55 -- A Trapezoid Shape.
		$LOI_DRAWSHAPE_TYPE_BASIC_TRIANGLE_ISOSCELES, _         ; 56 -- An Isosceles Triangle.
		$LOI_DRAWSHAPE_TYPE_BASIC_TRIANGLE_RIGHT, _             ; 57 -- A Right Angle Triangle.
		$LOI_DRAWSHAPE_TYPE_CALLOUT_CLOUD, _                    ; 58 -- A Cloud Shaped Callout.
		$LOI_DRAWSHAPE_TYPE_CALLOUT_LINE_1, _                   ; 59 -- A Callout with Line style #1.
		$LOI_DRAWSHAPE_TYPE_CALLOUT_LINE_2, _                   ; 60 -- A Callout with Line style #2.
		$LOI_DRAWSHAPE_TYPE_CALLOUT_LINE_3, _                   ; 61 -- A Callout with Line style #3.
		$LOI_DRAWSHAPE_TYPE_CALLOUT_RECTANGULAR, _              ; 62 -- A Rectangular Callout.
		$LOI_DRAWSHAPE_TYPE_CALLOUT_RECTANGULAR_ROUNDED, _      ; 63 -- A Rectangular Callout with rounded corners.
		$LOI_DRAWSHAPE_TYPE_CALLOUT_ROUND, _                    ; 64 -- A Round Callout.
		$LOI_DRAWSHAPE_TYPE_CONNECTOR, _                        ; 65 -- A plain connector.
		$LOI_DRAWSHAPE_TYPE_CONNECTOR_ARROWS, _                 ; 66 -- A Plain connector with Arrows on both ends.
		$LOI_DRAWSHAPE_TYPE_CONNECTOR_CURVED, _                 ; 67 -- A Curved connector.
		$LOI_DRAWSHAPE_TYPE_CONNECTOR_CURVED_ARROWS, _          ; 68 -- A Curved connector with Arrows on both ends.
		$LOI_DRAWSHAPE_TYPE_CONNECTOR_CURVED_ENDS_ARROW, _      ; 69 -- A Curved connector with an Arrow on one end.
		$LOI_DRAWSHAPE_TYPE_CONNECTOR_ENDS_ARROW, _             ; 70 -- A Plain connector with an Arrow on one end.
		$LOI_DRAWSHAPE_TYPE_CONNECTOR_LINE, _                   ; 71 -- A Line connector.
		$LOI_DRAWSHAPE_TYPE_CONNECTOR_LINE_ARROWS, _            ; 72 -- A Line connector with Arrows on both ends.
		$LOI_DRAWSHAPE_TYPE_CONNECTOR_LINE_ENDS_ARROW, _        ; 73 -- A Line connector with an Arrow on one end.
		$LOI_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT, _               ; 74 -- A Straight connector.
		$LOI_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT_ARROWS, _        ; 75 -- A Straight connector with Arrows on both ends.
		$LOI_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT_ENDS_ARROW, _    ; 76 -- A Straight connector with an Arrow on one end.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_CARD, _                   ; 77 -- A Card Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_COLLATE, _                ; 78 -- A Collate Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_CONNECTOR, _              ; 79 -- A Connector Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_CONNECTOR_OFF_PAGE, _     ; 80 -- A Off-Page Connector Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_DATA, _                   ; 81 -- A Data Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_DECISION, _               ; 82 -- A Decision Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_DELAY, _                  ; 83 -- A Delay Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_DIRECT_ACCESS_STORAGE, _  ; 84 -- A Direct Access Storage Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_DISPLAY, _                ; 85 -- A Display Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_DOCUMENT, _               ; 86 -- A Document Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_EXTRACT, _                ; 87 -- A Extract Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_INTERNAL_STORAGE, _       ; 88 -- A Internal Storage Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_MAGNETIC_DISC, _          ; 89 -- A Magnetic Disc Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_MANUAL_INPUT, _           ; 90 -- A Manual Input Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_MANUAL_OPERATION, _       ; 91 -- A Manual Operation Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_MERGE, _                  ; 92 -- A Merge Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_MULTIDOCUMENT, _          ; 93 -- A Multi-document Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_OR, _                     ; 94 -- A Or Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_PREPARATION, _            ; 95 -- A Preparation Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_PROCESS, _                ; 96 -- A Process Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_PROCESS_ALTERNATE, _      ; 97 -- A Alternate Process Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_PROCESS_PREDEFINED, _     ; 98 -- A Predefined Process Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_PUNCHED_TAPE, _           ; 99 -- A Punched Tape Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_SEQUENTIAL_ACCESS, _      ; 100 -- A Sequential Access Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_SORT, _                   ; 101 -- A Sort Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_STORED_DATA, _            ; 102 -- A Stored Data Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_SUMMING_JUNCTION, _       ; 103 -- A Summing Junction Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_TERMINATOR, _             ; 104 -- A Terminator Flowchart.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_AIR_MAIL, _                ; 105 -- "Air Mail" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_ASPHALT, _                 ; 106 -- "Asphalt" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_BLUE, _                    ; 107 -- "Blue" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_BLUE_MOON, _               ; 108 -- "Blue Moon" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_BURN, _                    ; 109 -- "Burn" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_BUSTER, _                  ; 110 -- "Buster" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_DONUTS_DONUTS, _           ; 111 -- "Donuts" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_FASTER, _                  ; 112 -- "Faster" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_FAVORITE_2, _              ; 113 -- "Favorite 2" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_FAVORITE_16, _             ; 114 -- "Favorite 16" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_GRAY, _                    ; 115 -- "Gray" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_GREEN_TO_BLUE, _           ; 116 -- "Green To Blue" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_GHOST, _                   ; 117 -- "Ghost" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_GOLD_WAVE, _               ; 118 -- "Gold Wave" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_HEAVY_METAL, _             ; 119 -- "Heavy Metal" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_HULK, _                    ; 120 -- "Hulk" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_NEWSPAPER, _               ; 121 -- "Newspaper" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_NO_WAY, _                  ; 122 -- "No Way" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_NOTE, _                    ; 123 -- "Note" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_OPEN, _                    ; 124 -- "Open" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_OUTLINE, _                 ; 125 -- "Outline" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_OUTLINE_BLUE, _            ; 126 -- "Outline Blue" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_PLANET, _                  ; 127 -- "Planet" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_POLICE_DRAMA, _            ; 128 -- "Police Drama" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_POW, _                     ; 129 -- "Pow!" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_PURPLE_SOLID, _            ; 130 -- "Purple Solid" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_RETRO, _                   ; 131 -- "Retro" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_SHADOW, _                  ; 132 -- "Shadow" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_SHADOW_BLUE, _             ; 133 -- "Shadow Blue" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_SIMPLE, _                  ; 134 -- "Simple" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_SNOW, _                    ; 135 -- "Snow" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_STAR_WARS, _               ; 136 -- "Star Wars" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_STONE, _                   ; 137 -- "Stone" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_STYLE, _                   ; 138 -- "Style" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_TRICOLORE, _               ; 139 -- "Tricolore" Fontwork.
		$LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ARROWS, _           ; 140 -- A line with an Arrow on either end.
		$LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ARROW_CIRCLE, _     ; 141 -- A line that starts with an Arrow, and ends with a Circle.
		$LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ARROW_SQUARE, _     ; 142 -- A line that starts with an Arrow, and ends with a Square.
		$LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_CIRCLE_ARROW, _     ; 143 -- A line that starts with a Circle, and ends with an Arrow.
		$LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ENDS_ARROW, _       ; 144 -- A line that ends with an Arrow.
		$LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_SQUARE_ARROW, _     ; 145 -- A line that starts with a Square, and ends with an Arrow.
		$LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_STARTS_ARROW, _     ; 146 -- A line that starts with an Arrow.
		$LOI_DRAWSHAPE_TYPE_LINE_CURVE, _                       ; 147 -- A Curve.
		$LOI_DRAWSHAPE_TYPE_LINE_CURVE_FILLED, _                ; 148 -- A Filled Curve.
		$LOI_DRAWSHAPE_TYPE_LINE_DIMENSION, _                   ; 149 -- A dimension line.
		$LOI_DRAWSHAPE_TYPE_LINE_FREEFORM_LINE, _               ; 150 -- A Freeform Line.
		$LOI_DRAWSHAPE_TYPE_LINE_FREEFORM_LINE_FILLED, _        ; 151 -- A Filled Freeform Line.
		$LOI_DRAWSHAPE_TYPE_LINE_LINE, _                        ; 152 -- A Line.
		$LOI_DRAWSHAPE_TYPE_LINE_LINE_45, _                     ; 153 -- A 45º line.
		$LOI_DRAWSHAPE_TYPE_LINE_POLYGON, _                     ; 154 -- A Polygon.
		$LOI_DRAWSHAPE_TYPE_LINE_POLYGON_45, _                  ; 155 -- A 45 degree Polygon.
		$LOI_DRAWSHAPE_TYPE_LINE_POLYGON_45_FILLED, _           ; 156 -- A Filled 45 degree Polygon.
		$LOI_DRAWSHAPE_TYPE_LINE_POLYGON_FILLED, _              ; 157 --  A Filled Polygon.
		$LOI_DRAWSHAPE_TYPE_STARS_4_POINT, _                    ; 158 -- A 4 Pointed Star.
		$LOI_DRAWSHAPE_TYPE_STARS_5_POINT, _                    ; 159 -- A 5 Pointed Star.
		$LOI_DRAWSHAPE_TYPE_STARS_6_POINT, _                    ; 160 -- A 6 Pointed Star. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_STARS_6_POINT_CONCAVE, _            ; 161 -- A Concave 6 Pointed Star. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_STARS_8_POINT, _                    ; 162 -- A 8 Pointed Star.
		$LOI_DRAWSHAPE_TYPE_STARS_12_POINT, _                   ; 163 -- A 12 Pointed Star. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_STARS_24_POINT, _                   ; 164 -- A 24 Pointed Star.
		$LOI_DRAWSHAPE_TYPE_STARS_DOORPLATE, _                  ; 165 -- A Doorplate Shape.
		$LOI_DRAWSHAPE_TYPE_STARS_EXPLOSION, _                  ; 166 -- A Explosion Shape.
		$LOI_DRAWSHAPE_TYPE_STARS_SCROLL_HORIZONTAL, _          ; 167 -- A Horizontal Scroll.
		$LOI_DRAWSHAPE_TYPE_STARS_SCROLL_VERTICAL, _            ; 168 -- A Vertical Scroll.
		$LOI_DRAWSHAPE_TYPE_STARS_SIGNET, _                     ; 169 -- A Signet Shape. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_DIAMOND, _             ; 170 -- A Diamond Bevel. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_OCTAGON, _             ; 171 -- A Octagon Bevel. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_SQUARE, _              ; 172 -- A Square Bevel.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_BRACE_DOUBLE, _              ; 173 -- A Double Brace.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_BRACE_LEFT, _                ; 174 -- A Left hand Brace.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_BRACE_RIGHT, _               ; 175 -- A Right hand Brace.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_BRACKET_DOUBLE, _            ; 176 -- A Double Bracket.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_BRACKET_LEFT, _              ; 177 -- A Left hand Bracket.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_BRACKET_RIGHT, _             ; 178 -- A Right hand Bracket.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_CLOUD, _                     ; 179 -- A Cloud Shape. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_FLOWER, _                    ; 180 -- A Flower Shape. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_HEART, _                     ; 181 -- A Heart Shape.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_LIGHTNING, _                 ; 182 -- A Lightning Shape. ## Note: Lightning is visually different than the one available in L.O. Shapes U.I.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_MOON, _                      ; 183 -- A Moon Shape.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_SMILEY, _                    ; 184 -- A Smiley Shape.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_SUN, _                       ; 185 -- A Sun Shape.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_PROHIBITED, _                ; 186 -- A Prohibited Shape.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_PUZZLE                       ; 187 -- A Puzzle Piece Shape. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.

; Gradient Names
Global Const _
		$LOI_GRAD_NAME_PASTEL_BOUQUET = "Pastel Bouquet", _     ; The "Pastel Bouquet" Gradient Preset.
		$LOI_GRAD_NAME_PASTEL_DREAM = "Pastel Dream", _         ; The "Pastel Dream" Gradient Preset.
		$LOI_GRAD_NAME_BLUE_TOUCH = "Blue Touch", _             ; The "Blue Touch" Gradient Preset.
		$LOI_GRAD_NAME_BLANK_W_GRAY = "Blank with Gray", _      ; The "Blank with Gray" Gradient Preset.
		$LOI_GRAD_NAME_LONDON_MIST = "London Mist", _           ; The "London Mist" Gradient Preset.
		$LOI_GRAD_NAME_SUBMARINE = "Submarine", _               ; The "Submarine" Gradient Preset.
		$LOI_GRAD_NAME_MIDNIGHT = "Midnight", _                 ; The "Midnight" Gradient Preset.
		$LOI_GRAD_NAME_DEEP_OCEAN = "Deep Ocean", _             ; The "Deep Ocean" Gradient Preset.
		$LOI_GRAD_NAME_MAHOGANY = "Mahogany", _                 ; The "Mahogany" Gradient Preset.
		$LOI_GRAD_NAME_GREEN_GRASS = "Green Grass", _           ; The "Green Grass" Gradient Preset.
		$LOI_GRAD_NAME_NEON_LIGHT = "Neon Light", _             ; The "Neon Light" Gradient Preset.
		$LOI_GRAD_NAME_SUNSHINE = "Sunshine", _                 ; The "Sunshine" Gradient Preset.
		$LOI_GRAD_NAME_RAINBOW = "Rainbow", _                   ; The "Rainbow" Gradient Preset. L.O. 7.6+
		$LOI_GRAD_NAME_SUNRISE = "Sunrise", _                   ; The "Sunrise" Gradient Preset. L.O. 7.6+
		$LOI_GRAD_NAME_SUNDOWN = "Sundown"                      ; The "Sundown" Gradient Preset. L.O. 7.6+

; Gradient Type
Global Const _                                                  ; com.sun.star.awt.GradientStyle
		$LOI_GRAD_TYPE_OFF = -1, _                              ; Turn the Gradient off.
		$LOI_GRAD_TYPE_LINEAR = 0, _                            ; Linear type Gradient
		$LOI_GRAD_TYPE_AXIAL = 1, _                             ; Axial type Gradient
		$LOI_GRAD_TYPE_RADIAL = 2, _                            ; Radial type Gradient
		$LOI_GRAD_TYPE_ELLIPTICAL = 3, _                        ; Elliptical type Gradient
		$LOI_GRAD_TYPE_SQUARE = 4, _                            ; Square type Gradient
		$LOI_GRAD_TYPE_RECT = 5                                 ; Rectangle type Gradient

; Shape Type Constants.
Global Enum Step * 2 _
		$LOI_SHAPE_TYPE_DRAWING_SHAPE = 1, _                    ; 1 - All shapes, 3D Shapes, Basic Shapes, Block Arrows, Flowcharts, Callouts, Lines, Connectors, Fontwork etc.
		$LOI_SHAPE_TYPE_FORM_CONTROL, _                         ; 2 - Form Controls.
		$LOI_SHAPE_TYPE_IMAGE, _                                ; 4 - An Image, Barcode or QR code.
		$LOI_SHAPE_TYPE_MEDIA, _                                ; 8 - A Video or Audio shape.
		$LOI_SHAPE_TYPE_OLE2, _                                 ; 16 - An OLE2 shape, such as a Chart, Formula, or Barcode etc.
		$LOI_SHAPE_TYPE_TABLE, _                                ; 32 - A Table.
		$LOI_SHAPE_TYPE_TEXTBOX, _                              ; 64 - A Text Box, including Hyperlinks, and most Fields.
		$LOI_SHAPE_TYPE_TEXTBOX_SUBTITLE, _                     ; 128 - A Slide Subtitle Box.
		$LOI_SHAPE_TYPE_TEXTBOX_TITLE, _                        ; 256 - A Slide Title Text Box.
		$LOI_SHAPE_TYPE_ALL = 511                               ; All types above.

; Zoom Type Constants
Global Const _                                                  ; com.sun.star.view.DocumentZoomType
		$LOI_ZOOMTYPE_OPTIMAL = 0, _                            ; The page content width (excluding margins) at the current selection is fit into the view.
		$LOI_ZOOMTYPE_PAGE_WIDTH = 1, _                         ; The page width at the current selection is fit into the view.
		$LOI_ZOOMTYPE_ENTIRE_PAGE = 2, _                        ; A complete page of the document is fit into the view.
		$LOI_ZOOMTYPE_BY_VALUE = 3, _                           ; The Zoom property is relative, and set using Zoom Value.
		$LOI_ZOOMTYPE_PAGE_WIDTH_EXACT = 4                      ; The Page width at the current selection is fit into the view with the view ends exactly at the end of the page.
