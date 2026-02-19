#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1
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

#Tidy_ILC_Pos=70

; Vertical Alignment
Global Const _                                                       ; com.sun.star.style.VerticalAlignment
		$LOI_ALIGN_VERT_TOP = 0, _                                   ; Vertically Align the object to the Top.
		$LOI_ALIGN_VERT_MIDDLE = 1, _                                ; Vertically Align the object to the Middle.
		$LOI_ALIGN_VERT_BOTTOM = 2                                   ; Vertically Align the object to the Bottom.

; Animation Direction
Global Const _                                                       ; com.sun.star.drawing.TextAnimationDirection
		$LOI_ANIMATION_DIR_LEFT = 0, _                               ; The Animation begins at the Right and goes to the Left.
		$LOI_ANIMATION_DIR_RIGHT = 1, _                              ; The Animation begins at the Left and goes to the Right.
		$LOI_ANIMATION_DIR_UP = 2, _                                 ; The Animation begins at the Bottom and goes to the Top.
		$LOI_ANIMATION_DIR_DOWN = 3                                  ; The Animation begins at the Top and goes to the Bottom.

; Animation Kind
Global Const _                                                       ; com.sun.star.drawing.TextAnimationKind
		$LOI_ANIMATION_TYPE_NONE = 0, _                              ; No Animation is applied.
		$LOI_ANIMATION_TYPE_BLINK = 1, _                             ; The text switches its state from visible to invisible continuously.
		$LOI_ANIMATION_TYPE_SCROLL_THROUGH = 2, _                    ; The text scrolls.
		$LOI_ANIMATION_TYPE_SCROLL_ALTERNATE = 3, _                  ; The text scrolls from one side to the other and back.
		$LOI_ANIMATION_TYPE_SCROLL_IN = 4                            ; The text Scrolls from one side to the final position and stops there.

; Anchor Type
Global Const _                                                       ; com.sun.star.text.TextContentAnchorType
		$LOI_ANCHOR_AT_PARAGRAPH = 0, _                              ; Anchors the object to the current paragraph.
		$LOI_ANCHOR_AS_CHARACTER = 1, _                              ; Anchors the Object as character. The height of the current line is resized to match the height of the selection.
		$LOI_ANCHOR_AT_PAGE = 2, _                                   ; Anchors the Object to the current page.
		$LOI_ANCHOR_AT_FRAME = 3, _                                  ; Anchors the object to the surrounding frame.
		$LOI_ANCHOR_AT_CHARACTER = 4                                 ; Anchors the Object to a character.

; Fill Style Type Constants
Global Enum _                                                        ; com.sun.star.drawing.FillStyle
		$LOI_AREA_FILL_STYLE_OFF, _                                  ; 0 Fill Style is off.
		$LOI_AREA_FILL_STYLE_SOLID, _                                ; 1 Fill Style is a solid color.
		$LOI_AREA_FILL_STYLE_GRADIENT, _                             ; 2 Fill Style is a gradient color.
		$LOI_AREA_FILL_STYLE_HATCH, _                                ; 3 Fill Style is a Hatch style color.
		$LOI_AREA_FILL_STYLE_BITMAP                                  ; 4 Fill Style is a Bitmap.

; Shape Connector type
Global Const _                                                       ; com.sun.star.drawing.ConnectorType
		$LOI_DRAWSHAPE_CONNECTOR_TYPE_STANDARD = 0, _                ; The connector is drawn with three lines, with the middle line perpendicular to the other two.
		$LOI_DRAWSHAPE_CONNECTOR_TYPE_CURVE = 1, _                   ; The connector is drawn as a curve.
		$LOI_DRAWSHAPE_CONNECTOR_TYPE_STRAIGHT = 2, _                ; The connector is drawn as a straight line.
		$LOI_DRAWSHAPE_CONNECTOR_TYPE_LINE = 3                       ; The connector is drawn with three lines.

; Dimension Line Text Horizontal Position.
Global Const _                                                       ; com.sun.star.drawing.MeasureTextHorzPos
		$LOI_DRAWSHAPE_DIMENSION_TEXT_HORI_POS_AUTO = 0, _           ; Select the best horizontal position for the text Automatically.
		$LOI_DRAWSHAPE_DIMENSION_TEXT_HORI_POS_LEFT = 1, _           ; The text is positioned at the Left.
		$LOI_DRAWSHAPE_DIMENSION_TEXT_HORI_POS_CENTER = 2, _         ; The text is positioned at the Horizontal Center.
		$LOI_DRAWSHAPE_DIMENSION_TEXT_HORI_POS_RIGHT = 3             ; The text is positioned at the Right.

; Dimension Line Text Vertical Position.
Global Const _                                                       ; com.sun.star.drawing.MeasureTextVertPos
		$LOI_DRAWSHAPE_DIMENSION_TEXT_VERT_POS_AUTO = 0, _           ; Select the best vertical position for the text automatically
		$LOI_DRAWSHAPE_DIMENSION_TEXT_VERT_POS_TOP = 1, _            ; The text is positioned at the Top.
		$LOI_DRAWSHAPE_DIMENSION_TEXT_VERT_POS_BOTTOM = 3, _         ; The text is positioned at the Bottom. (2 is not used?)
		$LOI_DRAWSHAPE_DIMENSION_TEXT_VERT_POS_MIDDLE = 4            ; The text is positioned in the Vertical Middle.

; Dimension Line Unit type.
Global Const _                                                       ; ?
		$LOI_DRAWSHAPE_DIMENSION_UNIT_TYPE_OFF = -1, _               ; No Measurement units are used.
		$LOI_DRAWSHAPE_DIMENSION_UNIT_TYPE_AUTO = 0, _               ; Measurement units are automatically determined.
		$LOI_DRAWSHAPE_DIMENSION_UNIT_TYPE_MM = 1, _                 ; Measurement units are in Millimeters.
		$LOI_DRAWSHAPE_DIMENSION_UNIT_TYPE_CM = 2, _                 ; Measurement units are in Centimeters.
		$LOI_DRAWSHAPE_DIMENSION_UNIT_TYPE_METER = 3, _              ; Measurement units are in Meters.
		$LOI_DRAWSHAPE_DIMENSION_UNIT_TYPE_KILOMETER = 4, _          ; Measurement units are in Kilometers.
		$LOI_DRAWSHAPE_DIMENSION_UNIT_TYPE_POINT = 6, _              ; Measurement units are in Points.
		$LOI_DRAWSHAPE_DIMENSION_UNIT_TYPE_PICA = 7, _               ; Measurement units are in Picas.
		$LOI_DRAWSHAPE_DIMENSION_UNIT_TYPE_INCH = 8, _               ; Measurement units are in Inches.
		$LOI_DRAWSHAPE_DIMENSION_UNIT_TYPE_FOOT = 9, _               ; Measurement units are in Feet.
		$LOI_DRAWSHAPE_DIMENSION_UNIT_TYPE_MILES = 10, _             ; Measurement units are in Miles.
		$LOI_DRAWSHAPE_DIMENSION_UNIT_TYPE_CHAR = 14, _              ; Measurement units are in Characters.
		$LOI_DRAWSHAPE_DIMENSION_UNIT_TYPE_LINE = 15                 ; Measurement units are in Lines.

; Arrowhead Type Constants
Global Enum _
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_NONE, _                       ; 0 -- No Arrow head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_ARROW_SHORT, _                ; 1 --Short Arrow head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_CONCAVE_SHORT, _              ; 2 -- Short Concave Arrow head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_ARROW, _                      ; 3 -- Arrow head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_TRIANGLE, _                   ; 4 -- Triangle Arrow head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_CONCAVE, _                    ; 5 -- Concave Arrow head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_ARROW_LARGE, _                ; 6 -- Large Arrow head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_CIRCLE, _                     ; 7 -- Circle Arrow head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_SQUARE, _                     ; 8 -- Square Arrow head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_SQUARE_45, _                  ; 9 -- Square Arrow head rotated 45 degrees.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_DIAMOND, _                    ; 10 -- Diamond Arrow head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_HALF_CIRCLE, _                ; 11 -- Half Circle Arrow head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_DIMENSIONAL_LINES, _          ; 12 -- Dimension Lines head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_DIMENSIONAL_LINE_ARROW, _     ; 13 -- Dimension Line Arrow head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_DIMENSION_LINE, _             ; 14 -- Dimension Line head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_LINE_SHORT, _                 ; 15 -- Short Line head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_LINE, _                       ; 16 -- Line head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_TRIANGLE_UNFILLED, _          ; 17 -- Unfilled Triangle Arrow head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_DIAMOND_UNFILLED, _           ; 18 -- Unfilled Diamond Arrow head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_CIRCLE_UNFILLED, _            ; 19 -- Unfilled Circle Arrow head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_SQUARE_45_UNFILLED, _         ; 20 -- Unfilled Square Arrow head, rotated 45 degrees.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_SQUARE_UNFILLED, _            ; 21 -- Unfilled Square Arrow head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_HALF_CIRCLE_UNFILLED, _       ; 22 -- Unfilled Half Circle Arrow head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_HALF_ARROW_LEFT, _            ; 23 -- Half Arrow left Arrow head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_HALF_ARROW_RIGHT, _           ; 24 -- Half Arrow right Arrow head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_REVERSED_ARROW, _             ; 25 -- Reversed Arrow head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_DOUBLE_ARROW, _               ; 26 -- Double Arrow head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_CF_ONE, _                     ; 27 -- CF One Arrow head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_CF_ONLY_ONE, _                ; 28 -- CF Only One Arrow head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_CF_MANY, _                    ; 29 -- CF Many Arrow head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_CF_MANY_ONE, _                ; 30 -- CF Many One Arrow head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_CF_ZERO_ONE, _                ; 31 -- CF Zero One Arrow head.
		$LOI_DRAWSHAPE_LINE_ARROW_TYPE_CF_ZERO_MANY                  ; 32 -- CF Zero Many Arrow head.

; Shape Line End Cap Constants.
Global Const _                                                       ; com.sun.star.drawing.LineCap
		$LOI_DRAWSHAPE_LINE_CAP_FLAT = 0, _                          ; Also called Butt, the line will end without any additional shape.
		$LOI_DRAWSHAPE_LINE_CAP_ROUND = 1, _                         ; The line will get a half circle as additional cap.
		$LOI_DRAWSHAPE_LINE_CAP_SQUARE = 2                           ; The line uses a square for the line end.

; Shape Line Joint Constants.
Global Const _                                                       ; com.sun.star.drawing.LineJoint
		$LOI_DRAWSHAPE_LINE_JOINT_NONE = 0, _                        ; The joint between lines will not be connected.
		$LOI_DRAWSHAPE_LINE_JOINT_MIDDLE = 1, _                      ; The middle value between the joints is used. ## Note used?
		$LOI_DRAWSHAPE_LINE_JOINT_BEVEL = 2, _                       ; The edges of the thick lines will be joined by lines.
		$LOI_DRAWSHAPE_LINE_JOINT_MITER = 3, _                       ; The lines join at intersections.
		$LOI_DRAWSHAPE_LINE_JOINT_ROUND = 4                          ; The lines join with an arc.

; Shape Line Style Constants.
Global Enum _
		$LOI_DRAWSHAPE_LINE_STYLE_NONE, _                            ; 0 -- No Line is applied.
		$LOI_DRAWSHAPE_LINE_STYLE_CONTINUOUS, _                      ; 1 -- A Solid Line.
		$LOI_DRAWSHAPE_LINE_STYLE_DOT, _                             ; 2 -- A Dotted Line.
		$LOI_DRAWSHAPE_LINE_STYLE_DOT_ROUNDED, _                     ; 3 -- A Rounded Dotted Line.
		$LOI_DRAWSHAPE_LINE_STYLE_LONG_DOT, _                        ; 4 -- A Long Dotted Line.
		$LOI_DRAWSHAPE_LINE_STYLE_LONG_DOT_ROUNDED, _                ; 5 -- A Rounded Long Dotted Line.
		$LOI_DRAWSHAPE_LINE_STYLE_DASH, _                            ; 6 -- A Dashed Line.
		$LOI_DRAWSHAPE_LINE_STYLE_DASH_ROUNDED, _                    ; 7 -- A Rounded Dashed Line.
		$LOI_DRAWSHAPE_LINE_STYLE_LONG_DASH, _                       ; 8 -- A Long Dashed Line.
		$LOI_DRAWSHAPE_LINE_STYLE_LONG_DASH_ROUNDED, _               ; 9 -- A Rounded Long Dashed Line.
		$LOI_DRAWSHAPE_LINE_STYLE_DOUBLE_DASH, _                     ; 10 -- A Double Dashed Line.
		$LOI_DRAWSHAPE_LINE_STYLE_DOUBLE_DASH_ROUNDED, _             ; 11 -- A Rounded Double Dash.
		$LOI_DRAWSHAPE_LINE_STYLE_DASH_DOT, _                        ; 12 -- A Dashed and Dotted Line.
		$LOI_DRAWSHAPE_LINE_STYLE_DASH_DOT_ROUNDED, _                ; 13 -- A Rounded Dashed and Dotted Line.
		$LOI_DRAWSHAPE_LINE_STYLE_LONG_DASH_DOT, _                   ; 14 -- A Long Dashed and Dotted Line.
		$LOI_DRAWSHAPE_LINE_STYLE_LONG_DASH_DOT_ROUNDED, _           ; 15 -- A Rounded Long Dashed and Dotted Line.
		$LOI_DRAWSHAPE_LINE_STYLE_DOUBLE_DASH_DOT, _                 ; 16 -- A Double Dash Dot Line.
		$LOI_DRAWSHAPE_LINE_STYLE_DOUBLE_DASH_DOT_ROUNDED, _         ; 17 -- A Rounded Double Dash Dot Line
		$LOI_DRAWSHAPE_LINE_STYLE_DASH_DOT_DOT, _                    ; 18 -- A Dash Dot Dot Line.
		$LOI_DRAWSHAPE_LINE_STYLE_DASH_DOT_DOT_ROUNDED, _            ; 19 -- A Rounded Dash Dot Dot Line.
		$LOI_DRAWSHAPE_LINE_STYLE_DOUBLE_DASH_DOT_DOT, _             ; 20 -- A Double Dash Dot Dot Line.
		$LOI_DRAWSHAPE_LINE_STYLE_DOUBLE_DASH_DOT_DOT_ROUNDED, _     ; 21 -- A Rounded Double Dash Dot Dot Line.
		$LOI_DRAWSHAPE_LINE_STYLE_ULTRAFINE_DOTTED, _                ; 22 -- A Ultrafine Dotted Line.
		$LOI_DRAWSHAPE_LINE_STYLE_FINE_DOTTED, _                     ; 23 -- A Fine Dotted Line.
		$LOI_DRAWSHAPE_LINE_STYLE_ULTRAFINE_DASHED, _                ; 24 -- A Ultrafine Dashed Line.
		$LOI_DRAWSHAPE_LINE_STYLE_FINE_DASHED, _                     ; 25 -- A Fine Dashed Line.
		$LOI_DRAWSHAPE_LINE_STYLE_DASHED, _                          ; 26 -- A Dashed Line.
		$LOI_DRAWSHAPE_LINE_STYLE_SPARSE_DASH, _                     ; 27 -- A Sparse Dash. Before version 24.2 this was called Line Style 9.
		$LOI_DRAWSHAPE_LINE_STYLE_3_DASHES_3_DOTS, _                 ; 28 -- A Line consisting of 3 Dashes and 3 Dots.
		$LOI_DRAWSHAPE_LINE_STYLE_ULTRAFINE_2_DOTS_3_DASHES, _       ; 29 -- A Ultrafine Line consisting of 2 Dots and 3 Dashes.
		$LOI_DRAWSHAPE_LINE_STYLE_2_DOTS_1_DASH, _                   ; 30 -- A Line consisting of 2 Dots and 1 Dash.
		$LOI_DRAWSHAPE_LINE_STYLE_LINE_WITH_FINE_DOTS                ; 31 -- A Line with Fine Dots.

; Polygon Flags
Global Const _                                                       ; com.sun.star.drawing.PolygonFlags
		$LOI_DRAWSHAPE_POINT_TYPE_NORMAL = 0, _                      ; the point is normal, from the curve discussion view.
		$LOI_DRAWSHAPE_POINT_TYPE_SMOOTH = 1, _                      ; the point is smooth, the first derivation from the curve discussion view.
		$LOI_DRAWSHAPE_POINT_TYPE_CONTROL = 2, _                     ; the point is a control point, to control the curve from the user interface.
		$LOI_DRAWSHAPE_POINT_TYPE_SYMMETRIC = 3                      ; the point is symmetric, the second derivation from the curve discussion view.

; Drawing Shape Type Constants.
Global Enum _
		$LOI_DRAWSHAPE_TYPE_3D_CONE, _                               ; 0 -- A 3D Cone.
		$LOI_DRAWSHAPE_TYPE_3D_CUBE, _                               ; 1 -- A 3D Cube.
		$LOI_DRAWSHAPE_TYPE_3D_CYLINDER, _                           ; 2 -- A 3D Cylinder.
		$LOI_DRAWSHAPE_TYPE_3D_HALF_SPHERE, _                        ; 3 -- A 3D Half-Sphere.
		$LOI_DRAWSHAPE_TYPE_3D_PYRAMID, _                            ; 4 -- A 3D Pyramid.
		$LOI_DRAWSHAPE_TYPE_3D_SHELL, _                              ; 5 -- A 3D Shell.
		$LOI_DRAWSHAPE_TYPE_3D_SPHERE, _                             ; 6 -- A 3D Sphere.
		$LOI_DRAWSHAPE_TYPE_3D_TORUS, _                              ; 7 -- A 3D Torus.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_4_WAY, _                    ; 8 -- A Four-way Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_4_WAY, _            ; 9 -- A Four-way Callout Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_DOWN, _             ; 10 -- A Downward Callout Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_LEFT, _             ; 11 -- A Left hand Callout Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_LEFT_RIGHT, _       ; 12 -- A Left and Right Callout Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_RIGHT, _            ; 13 -- A Right hand Callout Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP, _               ; 14 -- A Upward Callout Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_DOWN, _          ; 15 -- A Upward and Downward Callout Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CALLOUT_UP_RIGHT, _         ; 16 -- Upward and Right hand Callout Arrow. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CIRCULAR, _                 ; 17 -- A Circular Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_CORNER_RIGHT, _             ; 18 -- A Right hand Corner Arrow. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_DOWN, _                     ; 19 -- A Downward Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_LEFT, _                     ; 20 -- A Left hand Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_LEFT_RIGHT, _               ; 21 -- A Left and Right Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_NOTCHED_RIGHT, _            ; 22 -- A Notched Right Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_RIGHT, _                    ; 23 -- A Right hand Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_RIGHT_OR_LEFT, _            ; 24 -- A Right or Left Arrow. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_S_SHAPED, _                 ; 25 -- A "S"-Shaped Arrow. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_SPLIT, _                    ; 26 -- A Split Arrow. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_STRIPED_RIGHT, _            ; 27 -- A Striped Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP, _                       ; 28 -- A Upward Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_DOWN, _                  ; 29 -- A Up and Down Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT, _                 ; 30 -- A Upward and Right hand Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_ARROW_UP_RIGHT_DOWN, _            ; 31 -- A Upward, Right hand and Downward Arrow. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_ARROWS_CHEVRON, _                        ; 32 -- A Chevron Shape Arrow.
		$LOI_DRAWSHAPE_TYPE_ARROWS_PENTAGON, _                       ; 33 -- A Pentagon Shape Arrow.
		$LOI_DRAWSHAPE_TYPE_BASIC_ARC, _                             ; 34 -- An Arc Shape.
		$LOI_DRAWSHAPE_TYPE_BASIC_ARC_BLOCK, _                       ; 35 -- A Block Arc Shape.
		$LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE, _                          ; 36 -- A Circle.
		$LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_PIE, _                      ; 37 -- A Pie Circle. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_BASIC_CIRCLE_SEGMENT, _                  ; 38 -- A Segment Circle.
		$LOI_DRAWSHAPE_TYPE_BASIC_CROSS, _                           ; 39 -- A Cross Shape.
		$LOI_DRAWSHAPE_TYPE_BASIC_CUBE, _                            ; 40 -- A Cube Shape.
		$LOI_DRAWSHAPE_TYPE_BASIC_CYLINDER, _                        ; 41 -- A Cylinder Shape.
		$LOI_DRAWSHAPE_TYPE_BASIC_DIAMOND, _                         ; 42 -- A Diamond Shape.
		$LOI_DRAWSHAPE_TYPE_BASIC_ELLIPSE, _                         ; 43 -- An Ellipse Shape.
		$LOI_DRAWSHAPE_TYPE_BASIC_FOLDED_CORNER, _                   ; 44 -- A Paper Shape with a Folded Corner.
		$LOI_DRAWSHAPE_TYPE_BASIC_FRAME, _                           ; 45 -- A Frame Shape. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_BASIC_HEXAGON, _                         ; 46 -- A Hexagon Shape.
		$LOI_DRAWSHAPE_TYPE_BASIC_OCTAGON, _                         ; 47 -- A Octagon Shape.
		$LOI_DRAWSHAPE_TYPE_BASIC_PARALLELOGRAM, _                   ; 48 -- A Parallelogram Shape.
		$LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE, _                       ; 49 -- A Rectangle.
		$LOI_DRAWSHAPE_TYPE_BASIC_RECTANGLE_ROUNDED, _               ; 50 -- A Rectangle with rounded corners.
		$LOI_DRAWSHAPE_TYPE_BASIC_REGULAR_PENTAGON, _                ; 51 -- A regular Pentagon.
		$LOI_DRAWSHAPE_TYPE_BASIC_RING, _                            ; 52 -- A Ring Shape.
		$LOI_DRAWSHAPE_TYPE_BASIC_SQUARE, _                          ; 53 -- A Square.
		$LOI_DRAWSHAPE_TYPE_BASIC_SQUARE_ROUNDED, _                  ; 54 -- A Square with rounded corners.
		$LOI_DRAWSHAPE_TYPE_BASIC_TRAPEZOID, _                       ; 55 -- A Trapezoid Shape.
		$LOI_DRAWSHAPE_TYPE_BASIC_TRIANGLE_ISOSCELES, _              ; 56 -- An Isosceles Triangle.
		$LOI_DRAWSHAPE_TYPE_BASIC_TRIANGLE_RIGHT, _                  ; 57 -- A Right Angle Triangle.
		$LOI_DRAWSHAPE_TYPE_CALLOUT_CLOUD, _                         ; 58 -- A Cloud Shaped Callout.
		$LOI_DRAWSHAPE_TYPE_CALLOUT_LINE_1, _                        ; 59 -- A Callout with Line style #1.
		$LOI_DRAWSHAPE_TYPE_CALLOUT_LINE_2, _                        ; 60 -- A Callout with Line style #2.
		$LOI_DRAWSHAPE_TYPE_CALLOUT_LINE_3, _                        ; 61 -- A Callout with Line style #3.
		$LOI_DRAWSHAPE_TYPE_CALLOUT_RECTANGULAR, _                   ; 62 -- A Rectangular Callout.
		$LOI_DRAWSHAPE_TYPE_CALLOUT_RECTANGULAR_ROUNDED, _           ; 63 -- A Rectangular Callout with rounded corners.
		$LOI_DRAWSHAPE_TYPE_CALLOUT_ROUND, _                         ; 64 -- A Round Callout.
		$LOI_DRAWSHAPE_TYPE_CONNECTOR, _                             ; 65 -- A plain connector.
		$LOI_DRAWSHAPE_TYPE_CONNECTOR_ARROWS, _                      ; 66 -- A Plain connector with Arrows on both ends.
		$LOI_DRAWSHAPE_TYPE_CONNECTOR_CURVED, _                      ; 67 -- A Curved connector.
		$LOI_DRAWSHAPE_TYPE_CONNECTOR_CURVED_ARROWS, _               ; 68 -- A Curved connector with Arrows on both ends.
		$LOI_DRAWSHAPE_TYPE_CONNECTOR_CURVED_ENDS_ARROW, _           ; 69 -- A Curved connector with an Arrow on one end.
		$LOI_DRAWSHAPE_TYPE_CONNECTOR_ENDS_ARROW, _                  ; 70 -- A Plain connector with an Arrow on one end.
		$LOI_DRAWSHAPE_TYPE_CONNECTOR_LINE, _                        ; 71 -- A Line connector.
		$LOI_DRAWSHAPE_TYPE_CONNECTOR_LINE_ARROWS, _                 ; 72 -- A Line connector with Arrows on both ends.
		$LOI_DRAWSHAPE_TYPE_CONNECTOR_LINE_ENDS_ARROW, _             ; 73 -- A Line connector with an Arrow on one end.
		$LOI_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT, _                    ; 74 -- A Straight connector.
		$LOI_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT_ARROWS, _             ; 75 -- A Straight connector with Arrows on both ends.
		$LOI_DRAWSHAPE_TYPE_CONNECTOR_STRAIGHT_ENDS_ARROW, _         ; 76 -- A Straight connector with an Arrow on one end.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_CARD, _                        ; 77 -- A Card Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_COLLATE, _                     ; 78 -- A Collate Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_CONNECTOR, _                   ; 79 -- A Connector Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_CONNECTOR_OFF_PAGE, _          ; 80 -- A Off-Page Connector Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_DATA, _                        ; 81 -- A Data Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_DECISION, _                    ; 82 -- A Decision Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_DELAY, _                       ; 83 -- A Delay Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_DIRECT_ACCESS_STORAGE, _       ; 84 -- A Direct Access Storage Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_DISPLAY, _                     ; 85 -- A Display Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_DOCUMENT, _                    ; 86 -- A Document Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_EXTRACT, _                     ; 87 -- A Extract Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_INTERNAL_STORAGE, _            ; 88 -- A Internal Storage Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_MAGNETIC_DISC, _               ; 89 -- A Magnetic Disc Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_MANUAL_INPUT, _                ; 90 -- A Manual Input Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_MANUAL_OPERATION, _            ; 91 -- A Manual Operation Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_MERGE, _                       ; 92 -- A Merge Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_MULTIDOCUMENT, _               ; 93 -- A Multi-document Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_OR, _                          ; 94 -- A Or Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_PREPARATION, _                 ; 95 -- A Preparation Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_PROCESS, _                     ; 96 -- A Process Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_PROCESS_ALTERNATE, _           ; 97 -- A Alternate Process Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_PROCESS_PREDEFINED, _          ; 98 -- A Predefined Process Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_PUNCHED_TAPE, _                ; 99 -- A Punched Tape Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_SEQUENTIAL_ACCESS, _           ; 100 -- A Sequential Access Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_SORT, _                        ; 101 -- A Sort Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_STORED_DATA, _                 ; 102 -- A Stored Data Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_SUMMING_JUNCTION, _            ; 103 -- A Summing Junction Flowchart.
		$LOI_DRAWSHAPE_TYPE_FLOWCHART_TERMINATOR, _                  ; 104 -- A Terminator Flowchart.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_AIR_MAIL, _                     ; 105 -- "Air Mail" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_ASPHALT, _                      ; 106 -- "Asphalt" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_BLUE, _                         ; 107 -- "Blue" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_BLUE_MOON, _                    ; 108 -- "Blue Moon" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_BURN, _                         ; 109 -- "Burn" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_BUSTER, _                       ; 110 -- "Buster" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_DONUTS_DONUTS, _                ; 111 -- "Donuts" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_FASTER, _                       ; 112 -- "Faster" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_FAVORITE_2, _                   ; 113 -- "Favorite 2" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_FAVORITE_16, _                  ; 114 -- "Favorite 16" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_GRAY, _                         ; 115 -- "Gray" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_GREEN_TO_BLUE, _                ; 116 -- "Green To Blue" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_GHOST, _                        ; 117 -- "Ghost" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_GOLD_WAVE, _                    ; 118 -- "Gold Wave" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_HEAVY_METAL, _                  ; 119 -- "Heavy Metal" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_HULK, _                         ; 120 -- "Hulk" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_NEWSPAPER, _                    ; 121 -- "Newspaper" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_NO_WAY, _                       ; 122 -- "No Way" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_NOTE, _                         ; 123 -- "Note" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_OPEN, _                         ; 124 -- "Open" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_OUTLINE, _                      ; 125 -- "Outline" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_OUTLINE_BLUE, _                 ; 126 -- "Outline Blue" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_PLANET, _                       ; 127 -- "Planet" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_POLICE_DRAMA, _                 ; 128 -- "Police Drama" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_POW, _                          ; 129 -- "Pow!" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_PURPLE_SOLID, _                 ; 130 -- "Purple Solid" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_RETRO, _                        ; 131 -- "Retro" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_SHADOW, _                       ; 132 -- "Shadow" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_SHADOW_BLUE, _                  ; 133 -- "Shadow Blue" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_SIMPLE, _                       ; 134 -- "Simple" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_SNOW, _                         ; 135 -- "Snow" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_STAR_WARS, _                    ; 136 -- "Star Wars" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_STONE, _                        ; 137 -- "Stone" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_STYLE, _                        ; 138 -- "Style" Fontwork.
		$LOI_DRAWSHAPE_TYPE_FONTWORK_TRICOLORE, _                    ; 139 -- "Tricolore" Fontwork.
		$LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ARROWS, _                ; 140 -- A line with an Arrow on either end.
		$LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ARROW_CIRCLE, _          ; 141 -- A line that starts with an Arrow, and ends with a Circle.
		$LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ARROW_SQUARE, _          ; 142 -- A line that starts with an Arrow, and ends with a Square.
		$LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_CIRCLE_ARROW, _          ; 143 -- A line that starts with a Circle, and ends with an Arrow.
		$LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_ENDS_ARROW, _            ; 144 -- A line that ends with an Arrow.
		$LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_SQUARE_ARROW, _          ; 145 -- A line that starts with a Square, and ends with an Arrow.
		$LOI_DRAWSHAPE_TYPE_LINE_ARROW_LINE_STARTS_ARROW, _          ; 146 -- A line that starts with an Arrow.
		$LOI_DRAWSHAPE_TYPE_LINE_CURVE, _                            ; 147 -- A Curve.
		$LOI_DRAWSHAPE_TYPE_LINE_CURVE_FILLED, _                     ; 148 -- A Filled Curve.
		$LOI_DRAWSHAPE_TYPE_LINE_DIMENSION, _                        ; 149 -- A dimension line.
		$LOI_DRAWSHAPE_TYPE_LINE_FREEFORM_LINE, _                    ; 150 -- A Freeform Line.
		$LOI_DRAWSHAPE_TYPE_LINE_FREEFORM_LINE_FILLED, _             ; 151 -- A Filled Freeform Line.
		$LOI_DRAWSHAPE_TYPE_LINE_LINE, _                             ; 152 -- A Line.
		$LOI_DRAWSHAPE_TYPE_LINE_LINE_45, _                          ; 153 -- A 45º line.
		$LOI_DRAWSHAPE_TYPE_LINE_POLYGON, _                          ; 154 -- A Polygon.
		$LOI_DRAWSHAPE_TYPE_LINE_POLYGON_45, _                       ; 155 -- A 45 degree Polygon.
		$LOI_DRAWSHAPE_TYPE_LINE_POLYGON_45_FILLED, _                ; 156 -- A Filled 45 degree Polygon.
		$LOI_DRAWSHAPE_TYPE_LINE_POLYGON_FILLED, _                   ; 157 --  A Filled Polygon.
		$LOI_DRAWSHAPE_TYPE_STARS_4_POINT, _                         ; 158 -- A 4 Pointed Star.
		$LOI_DRAWSHAPE_TYPE_STARS_5_POINT, _                         ; 159 -- A 5 Pointed Star.
		$LOI_DRAWSHAPE_TYPE_STARS_6_POINT, _                         ; 160 -- A 6 Pointed Star. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_STARS_6_POINT_CONCAVE, _                 ; 161 -- A Concave 6 Pointed Star. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_STARS_8_POINT, _                         ; 162 -- A 8 Pointed Star.
		$LOI_DRAWSHAPE_TYPE_STARS_12_POINT, _                        ; 163 -- A 12 Pointed Star. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_STARS_24_POINT, _                        ; 164 -- A 24 Pointed Star.
		$LOI_DRAWSHAPE_TYPE_STARS_DOORPLATE, _                       ; 165 -- A Doorplate Shape.
		$LOI_DRAWSHAPE_TYPE_STARS_EXPLOSION, _                       ; 166 -- A Explosion Shape.
		$LOI_DRAWSHAPE_TYPE_STARS_SCROLL_HORIZONTAL, _               ; 167 -- A Horizontal Scroll.
		$LOI_DRAWSHAPE_TYPE_STARS_SCROLL_VERTICAL, _                 ; 168 -- A Vertical Scroll.
		$LOI_DRAWSHAPE_TYPE_STARS_SIGNET, _                          ; 169 -- A Signet Shape. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_DIAMOND, _                  ; 170 -- A Diamond Bevel. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_OCTAGON, _                  ; 171 -- A Octagon Bevel. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_BEVEL_SQUARE, _                   ; 172 -- A Square Bevel.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_BRACE_DOUBLE, _                   ; 173 -- A Double Brace.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_BRACE_LEFT, _                     ; 174 -- A Left hand Brace.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_BRACE_RIGHT, _                    ; 175 -- A Right hand Brace.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_BRACKET_DOUBLE, _                 ; 176 -- A Double Bracket.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_BRACKET_LEFT, _                   ; 177 -- A Left hand Bracket.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_BRACKET_RIGHT, _                  ; 178 -- A Right hand Bracket.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_CLOUD, _                          ; 179 -- A Cloud Shape. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_FLOWER, _                         ; 180 -- A Flower Shape. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_HEART, _                          ; 181 -- A Heart Shape.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_LIGHTNING, _                      ; 182 -- A Lightning Shape. ## Note: Lightning is visually different than the one available in L.O. Shapes U.I.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_MOON, _                           ; 183 -- A Moon Shape.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_SMILEY, _                         ; 184 -- A Smiley Shape.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_SUN, _                            ; 185 -- A Sun Shape.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_PROHIBITED, _                     ; 186 -- A Prohibited Shape.
		$LOI_DRAWSHAPE_TYPE_SYMBOL_PUZZLE                            ; 187 -- A Puzzle Piece Shape. ## Not implemented into Libre Office SDK as of 7.3.4.2 or higher.

; Gradient Names
Global Const _
		$LOI_GRAD_NAME_PASTEL_BOUQUET = "Pastel Bouquet", _          ; The "Pastel Bouquet" Gradient Preset.
		$LOI_GRAD_NAME_PASTEL_DREAM = "Pastel Dream", _              ; The "Pastel Dream" Gradient Preset.
		$LOI_GRAD_NAME_BLUE_TOUCH = "Blue Touch", _                  ; The "Blue Touch" Gradient Preset.
		$LOI_GRAD_NAME_BLANK_W_GRAY = "Blank with Gray", _           ; The "Blank with Gray" Gradient Preset.
		$LOI_GRAD_NAME_LONDON_MIST = "London Mist", _                ; The "London Mist" Gradient Preset.
		$LOI_GRAD_NAME_SUBMARINE = "Submarine", _                    ; The "Submarine" Gradient Preset.
		$LOI_GRAD_NAME_MIDNIGHT = "Midnight", _                      ; The "Midnight" Gradient Preset.
		$LOI_GRAD_NAME_DEEP_OCEAN = "Deep Ocean", _                  ; The "Deep Ocean" Gradient Preset.
		$LOI_GRAD_NAME_MAHOGANY = "Mahogany", _                      ; The "Mahogany" Gradient Preset.
		$LOI_GRAD_NAME_GREEN_GRASS = "Green Grass", _                ; The "Green Grass" Gradient Preset.
		$LOI_GRAD_NAME_NEON_LIGHT = "Neon Light", _                  ; The "Neon Light" Gradient Preset.
		$LOI_GRAD_NAME_SUNSHINE = "Sunshine", _                      ; The "Sunshine" Gradient Preset.
		$LOI_GRAD_NAME_RAINBOW = "Rainbow", _                        ; The "Rainbow" Gradient Preset. L.O. 7.6+
		$LOI_GRAD_NAME_SUNRISE = "Sunrise", _                        ; The "Sunrise" Gradient Preset. L.O. 7.6+
		$LOI_GRAD_NAME_SUNDOWN = "Sundown"                           ; The "Sundown" Gradient Preset. L.O. 7.6+

; Gradient Type
Global Const _                                                       ; com.sun.star.awt.GradientStyle
		$LOI_GRAD_TYPE_OFF = -1, _                                   ; Turn the Gradient off.
		$LOI_GRAD_TYPE_LINEAR = 0, _                                 ; Linear type Gradient
		$LOI_GRAD_TYPE_AXIAL = 1, _                                  ; Axial type Gradient
		$LOI_GRAD_TYPE_RADIAL = 2, _                                 ; Radial type Gradient
		$LOI_GRAD_TYPE_ELLIPTICAL = 3, _                             ; Elliptical type Gradient
		$LOI_GRAD_TYPE_SQUARE = 4, _                                 ; Square type Gradient
		$LOI_GRAD_TYPE_RECT = 5                                      ; Rectangle type Gradient

; Horizontal Orientation
Global Const _                                                       ; com.sun.star.text.HoriOrientation
		$LOI_ORIENT_HORI_NONE = 0, _                                 ; No hard alignment is applied. Equal to "From Left" in L.O. U.I.
		$LOI_ORIENT_HORI_RIGHT = 1, _                                ; The object is aligned at the right side.
		$LOI_ORIENT_HORI_CENTER = 2, _                               ; The object is aligned at the middle.
		$LOI_ORIENT_HORI_LEFT = 3, _                                 ; The object is aligned at the left side.
		$LOI_ORIENT_HORI_FULL = 6, _                                 ; The table uses the full space (for text tables only).
		$LOI_ORIENT_HORI_LEFT_AND_WIDTH = 7                          ; The left offset and the width of the table are defined.

; Vertical Orientation
Global Const _                                                       ; com.sun.star.text.VertOrientation
		$LOI_ORIENT_VERT_NONE = 0, _                                 ; No hard alignment. The same as "From Top"/From Bottom" in L.O. U.I., the only difference is the combination setting of Vertical Relation.
		$LOI_ORIENT_VERT_TOP = 1, _                                  ; Aligned at the top.
		$LOI_ORIENT_VERT_CENTER = 2, _                               ; Aligned at the center.
		$LOI_ORIENT_VERT_BOTTOM = 3, _                               ; Aligned at the bottom.
		$LOI_ORIENT_VERT_CHAR_TOP = 4, _                             ; Aligned at the top of a character. Available only when anchor is set to "As character". Equal to L.O. UI setting of "Vertical" = Top, and "To" = Character.
		$LOI_ORIENT_VERT_CHAR_CENTER = 5, _                          ; Aligned at the center of a character. Available only when anchor is set to "As character". Equal to L.O. UI setting of "Vertical" = Center, and "To" = Character.
		$LOI_ORIENT_VERT_CHAR_BOTTOM = 6, _                          ; Aligned at the bottom of a character. Available only when anchor is set to "As character". Equal to L.O. UI setting of "Vertical" = Center, and "To" = Character.
		$LOI_ORIENT_VERT_LINE_TOP = 7, _                             ; Aligned at the top of the line. Available only when anchor is set to "As character". Equal to L.O. UI setting of "Vertical" = Top, and "To" = Row.
		$LOI_ORIENT_VERT_LINE_CENTER = 8, _                          ; Aligned at the center of the line. Available only when anchor is set to "As character". Equal to L.O. UI setting of "Vertical" = Center, and "To" = Row.
		$LOI_ORIENT_VERT_LINE_BOTTOM = 9                             ; Aligned at the bottom of the line. Available only when anchor is set to "As character". Equal to L.O. UI setting of "Vertical" = Center, and "To" = Row.

; Relative to
Global Const _                                                       ; com.sun.star.text.RelOrientation
		$LOI_RELATIVE_ROW = -1, _                                    ; Position an object considering the row height.
		$LOI_RELATIVE_PARAGRAPH = 0, _                               ; The Object is placed considering the available paragraph space, including indent spacing. [Also called "Margin" or "Baseline" in L.O. UI]
		$LOI_RELATIVE_PARAGRAPH_TEXT = 1, _                          ; The Object is placed considering the available paragraph space, excluding indent spacing.
		$LOI_RELATIVE_CHARACTER = 2, _                               ; The Object is placed considering the available character space.
		$LOI_RELATIVE_PAGE_LEFT = 3, _                               ; The Object is placed considering the available space between the left page border and the left Paragraph border. [Same as Left Page Border in L.O. UI]
		$LOI_RELATIVE_PAGE_RIGHT = 4, _                              ; The Object is placed considering the available space between the Right page border and the Right Paragraph border. [Same as Right Page Border in L.O. UI]
		$LOI_RELATIVE_PARAGRAPH_LEFT = 5, _                          ; The Object is placed considering the available indent space to the left of the paragraph.
		$LOI_RELATIVE_PARAGRAPH_RIGHT = 6, _                         ; The Object is placed considering the available indent space to the right of the paragraph.
		$LOI_RELATIVE_PAGE = 7, _                                    ; The Object is placed considering the available space between the right and left, or top and bottom page borders.
		$LOI_RELATIVE_PAGE_PRINT = 8, _                              ; The Object is placed considering the available space between the right and left, or top and bottom page margins. [Same as Page Text Area in L.O. UI]
		$LOI_RELATIVE_TEXT_LINE = 9, _                               ; The Object is placed considering the height of the line.
		$LOI_RELATIVE_PAGE_PRINT_BOTTOM = 10, _                      ; The Object is placed considering the space available in the page footer(?)
		$LOI_RELATIVE_PAGE_PRINT_TOP = 11                            ; The Object is placed considering the space available in the page header(?)

; Shape Type Constants.
Global Enum Step * 2 _
		$LOI_SHAPE_TYPE_DRAWING_SHAPE = 1, _                         ; 1 - All shapes, 3D Shapes, Basic Shapes, Block Arrows, Flowcharts, Callouts, Lines, Connectors, Fontwork etc.
		$LOI_SHAPE_TYPE_FORM_CONTROL, _                              ; 2 - Form Controls.
		$LOI_SHAPE_TYPE_IMAGE, _                                     ; 4 - An Image, Barcode or QR code.
		$LOI_SHAPE_TYPE_MEDIA, _                                     ; 8 - A Video or Audio shape.
		$LOI_SHAPE_TYPE_OLE2, _                                      ; 16 - An OLE2 shape, such as a Chart, Formula etc.
		$LOI_SHAPE_TYPE_TABLE, _                                     ; 32 - A Table.
		$LOI_SHAPE_TYPE_TEXTBOX, _                                   ; 64 - A Text Box, including Hyperlinks, and most Fields.
		$LOI_SHAPE_TYPE_TEXTBOX_SUBTITLE, _                          ; 128 - A Slide Subtitle Box.
		$LOI_SHAPE_TYPE_TEXTBOX_TITLE, _                             ; 256 - A Slide Title Text Box.
		$LOI_SHAPE_TYPE_ALL = 511                                    ; All types above.

; Slide layout arrangements.
Global Const _
		$LOI_SLIDE_LAYOUT_TITLE = 0, _                               ; The Slide will contain a Title textbox and a Subtitle textbox.
		$LOI_SLIDE_LAYOUT_TITLE_CONTENT = 1, _                       ; The Slide will contain a Title textbox and a content textbox.
		$LOI_SLIDE_LAYOUT_TITLE_2_CONTENT = 3, _                     ; The Slide will contain a Title textbox and two content textboxes.
		$LOI_SLIDE_LAYOUT_TITLE_CONTENT_AND_2_CONTENT = 12, _        ; The Slide will contain a Title textbox and a content textbox beside the two smaller content boxes.
		$LOI_SLIDE_LAYOUT_TITLE_CONTENT_OVER_CONTENT = 14, _         ; The Slide will contain a Title textbox two content textboxes one positioned over top the other.
		$LOI_SLIDE_LAYOUT_TITLE_2_CONTENT_AND_CONTENT = 15, _        ; The Slide will contain a Title textbox with two smaller content textboxes beside a third content text box.
		$LOI_SLIDE_LAYOUT_TITLE_2_CONTENT_OVER_CONTENT = 16, _       ; The Slide will contain a Title textbox with two smaller content textboxes over top of a third content text box.
		$LOI_SLIDE_LAYOUT_TITLE_4_CONTENT = 18, _                    ; The Slide will contain a Title textbox with four smaller content textboxes.
		$LOI_SLIDE_LAYOUT_TITLE_ONLY = 19, _                         ; The Slide will contain a Title textbox only.
		$LOI_SLIDE_LAYOUT_BLANK = 20, _                              ; The Slide will contain no textbox.
		$LOI_SLIDE_LAYOUT_CENTERED_TEXT = 32, _                      ; The Slide will contain a content textbox with centered text.
		$LOI_SLIDE_LAYOUT_TITLE_6_CONTENT = 34                       ; The Slide will contain a Title textbox with six smaller content textboxes.

; SlideShow Presentation Mode.
Global Enum _
		$LOI_SLIDESHOW_VIEW_MODE_FULL_SCREEN, _                      ; The SlideShow is Full Screen.
		$LOI_SLIDESHOW_VIEW_MODE_IN_WINDOW, _                        ; The SlideShow is displayed in the LibreOffice program window.
		$LOI_SLIDESHOW_VIEW_MODE_LOOP                                ; The SlideShow is looped after a set pause.

; SlideShow Pen Width
Global Const _
		$LOI_SLIDESHOW_PEN_WIDTH_VERY_THIN = 4, _                    ; A very thin width pen line for drawing with.
		$LOI_SLIDESHOW_PEN_WIDTH_THIN = 100, _                       ; A thin width pen line for drawing with.
		$LOI_SLIDESHOW_PEN_WIDTH_NORMAL = 150, _                     ; A normal width pen line for drawing with.
		$LOI_SLIDESHOW_PEN_WIDTH_THICK = 200, _                      ; A thick width pen line for drawing with.
		$LOI_SLIDESHOW_PEN_WIDTH_VERY_THICK = 400                    ; A very thick width pen line for drawing with.

; SlideShow active Presentation commands and queries.
Global Enum _
		$LOI_SLIDESHOW_PRES_QUERY_GET_CURRENT_SLIDE, _               ; Returns the Object for the slide that is currently displayed.
		$LOI_SLIDESHOW_PRES_QUERY_GET_CURRENT_SLIDE_INDEX, _         ; Returns the index of the current slide. Index is 0 based.
		$LOI_SLIDESHOW_PRES_QUERY_GET_NEXT_SLIDE_INDEX, _            ; Returns the index for the slide that is displayed next. Index is 0 based.
		$LOI_SLIDESHOW_PRES_QUERY_GET_SLIDE_BY_INDEX, _              ; Returns the Object for the slide at the index. Index is 0 based. Slides are in the order they will be displayed in the presentation which can be different than the orders of slides in the document. Not all slides must be present and each slide can be used more than once.
		$LOI_SLIDESHOW_PRES_QUERY_GET_SLIDE_COUNT, _                 ; Returns the number of slides in this slide show.
		$LOI_SLIDESHOW_PRES_QUERY_IS_ACTIVE, _                       ; Determines if the slide show is active. Returns TRUE for UI active slide show, FALSE otherwise.
		$LOI_SLIDESHOW_PRES_QUERY_IS_ENDLESS, _                      ; Returns TRUE if the slide show was started to run endlessly.
		$LOI_SLIDESHOW_PRES_QUERY_IS_FULLSCREEN, _                   ; Returns TRUE if the slide show was started in full-screen mode.
		$LOI_SLIDESHOW_PRES_QUERY_IS_PAUSED, _                       ; Returns TRUE if the slide show is currently paused.
		$LOI_SLIDESHOW_PRES_COMMAND_ACTIVATE, _                      ; Activates the user interface of this slide show.
		$LOI_SLIDESHOW_PRES_COMMAND_ACTIVATE_BLANK_SCREEN, _         ; >Expects Parameter: Pause Screen Color as a RGB Color Integer.< Pauses the slide show and blanks the screen in the given color. Call Resume to unpause the slide show.
		$LOI_SLIDESHOW_PRES_COMMAND_DEACTIVATE, _                    ; Can be called to deactivate the user interface of this slide show. (Doesn't seem to set IsActive to False!)
		$LOI_SLIDESHOW_PRES_COMMAND_ERASE_ALL_INK, _                 ; Clears ink drawing from the slideshow being played. L.O. 7.2+
		$LOI_SLIDESHOW_PRES_COMMAND_GOTO_FIRST_SLIDE, _              ; Goto and display the first slide.
		$LOI_SLIDESHOW_PRES_COMMAND_GOTO_LAST_SLIDE, _               ; Goto and display last slide. Remaining effects on the current slide will be skipped.
		$LOI_SLIDESHOW_PRES_COMMAND_GOTO_NEXT_EFFECT, _              ; Start next effects that wait on a generic trigger. If no generic triggers are waiting the next slide will be displayed.
		$LOI_SLIDESHOW_PRES_COMMAND_GOTO_NEXT_SLIDE, _               ; Goto and display next slide. Remaining effects on the current slide will be skipped.
		$LOI_SLIDESHOW_PRES_COMMAND_GOTO_PREV_EFFECT, _              ; Undo the last effects that were triggered by a generic trigger. If there is no previous effect that can be undone then the previous slide will be displayed.
		$LOI_SLIDESHOW_PRES_COMMAND_GOTO_PREV_SLIDE, _               ; Goto and display previous slide. Remaining effects on the current slide will be skipped.
		$LOI_SLIDESHOW_PRES_COMMAND_GOTO_SLIDE, _                    ; >Expects Parameter: Slide Object to jump to.< Jumps to the given slide. The slide can also be a slide that would normally not be shown during the current slide show.
		$LOI_SLIDESHOW_PRES_COMMAND_GOTO_SLIDE_BY_INDEX, _           ; >Expects Parameter: Slide's index to jump to.< Jumps to the slide at the given index. 0 based.
		$LOI_SLIDESHOW_PRES_COMMAND_GOTO_SLIDE_BY_NAME, _            ; >Expects Parameter: Slide's name to jump to.< Jumps to the slide with the given name.
		$LOI_SLIDESHOW_PRES_COMMAND_PAUSE, _                         ; Pauses the slide show. All effects are paused. The slide show continues on next user input or if resume is called.
		$LOI_SLIDESHOW_PRES_COMMAND_RESUME, _                        ; Resumes a paused slide show.
		$LOI_SLIDESHOW_PRES_COMMAND_STOP_SOUND                       ; Stop all currently played sounds

; SlideShow Presentation Range
Global Enum _
		$LOI_SLIDESHOW_RANGE_ALL, _                                  ; All the slides in the presentation are included in the SlideShow.
		$LOI_SLIDESHOW_RANGE_FROM, _                                 ; The SlideShow begins at the defined slide.
		$LOI_SLIDESHOW_RANGE_CUSTOM                                  ; A custom SlideShow order is followed.

; Horizontal Text Alignment
Global Const _                                                       ; com.sun.star.drawing.TextHorizontalAdjust
		$LOI_TEXT_ALIGN_HORI_LEFT = 0, _                             ; The left edge of the text is adjusted to the left edge of the shape.
		$LOI_TEXT_ALIGN_HORI_CENTER = 1, _                           ; The text is centered horizontally inside the shape.
		$LOI_TEXT_ALIGN_HORI_RIGHT = 2, _                            ; The right edge of the text is adjusted to the right edge of the shape.
		$LOI_TEXT_ALIGN_HORI_BLOCK = 3                               ; The text extends from the left to the right edge of the shape.

; Vertical Text Alignment
Global Const _                                                       ; com.sun.star.drawing.TextVerticalAdjust
		$LOI_TEXT_ALIGN_VERT_TOP = 0, _                              ; The top edge of the text is adjusted to the top edge of the shape.
		$LOI_TEXT_ALIGN_VERT_CENTER = 1, _                           ; The text is centered vertically inside the shape.
		$LOI_TEXT_ALIGN_VERT_BOTTOM = 2, _                           ; The bottom edge of the text is adjusted to the bottom edge of the shape.
		$LOI_TEXT_ALIGN_VERT_BLOCK = 3                               ; The text extends from the top to the bottom edge of the shape.

; Text Anchor Position
Global Enum _
		$LOI_TEXT_ANCHOR_TOP_LEFT, _                                 ; The text is positioned in the Upper-Left corner of the Shape.
		$LOI_TEXT_ANCHOR_TOP_CENTER, _                               ; The text is positioned in the Upper-Center of the Shape.
		$LOI_TEXT_ANCHOR_TOP_RIGHT, _                                ; The text is positioned in the Upper-Right of the Shape.
		$LOI_TEXT_ANCHOR_MIDDLE_LEFT, _                              ; The text is positioned in the Middle-Left corner of the Shape.
		$LOI_TEXT_ANCHOR_MIDDLE_CENTER, _                            ; The text is positioned in the Middle-Center of the Shape.
		$LOI_TEXT_ANCHOR_MIDDLE_RIGHT, _                             ; The text is positioned in the Middle-Right of the Shape.
		$LOI_TEXT_ANCHOR_BOTTOM_LEFT, _                              ; The text is positioned in the Lower-Left corner of the Shape.
		$LOI_TEXT_ANCHOR_BOTTOM_CENTER, _                            ; The text is positioned in the Lower-Center of the Shape.
		$LOI_TEXT_ANCHOR_BOTTOM_RIGHT                                ; The text is positioned in the Lower-Right of the Shape.

; Zoom Type Constants
Global Const _                                                       ; com.sun.star.view.DocumentZoomType
		$LOI_ZOOMTYPE_OPTIMAL = 0, _                                 ; The page content width (excluding margins) at the current selection is fit into the view.
		$LOI_ZOOMTYPE_PAGE_WIDTH = 1, _                              ; The page width at the current selection is fit into the view.
		$LOI_ZOOMTYPE_ENTIRE_PAGE = 2, _                             ; A complete page of the document is fit into the view.
		$LOI_ZOOMTYPE_BY_VALUE = 3, _                                ; The Zoom property is relative, and set using Zoom Value.
		$LOI_ZOOMTYPE_PAGE_WIDTH_EXACT = 4                           ; The Page width at the current selection is fit into the view with the view ends exactly at the end of the page.
