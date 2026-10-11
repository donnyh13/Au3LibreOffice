# UDF To Do

The following is a basic list of things needing done, added to, or looked at, in this UDF.

## UDF

Things pertaining to the **entire UDF**, or **equally to all sub-Components**.

- Add Hatch Background option
  - Will probably have to insert them in HatchTable, like I do for Gradients.
- Add Macros functions? Plus Integrate with all Functions that can activate a Macro?
- Make a global Open/Close func?
- Is it important to use "Get" / "set"? e.g. getText, etc? Instead of Text()?
  - <https://ask.libreoffice.org/t/calc-named-range-err-508/> See comment by JohnSUN
  - OOME 4.1 pg 306
- Add note that shape insertion/manipulation is sub-par etc.
- Add FileDelete to _Error func in examples to ones that make a file.
- Should transparency funcs be listed under other than "Area"?
- Look at "OpenNewView" when opening documents? file:///C:/Program%20Files/LibreOffice/sdk/docs/idl/ref/a17820.html
- ParaTopMargin etc seems to have max of 100,000 now?
- Look into loadStylesFromURL for importing styles (Writer/Calc). OO Dev Pg 877.
- Rename TransparencyGradient to Trans(p)? Gradient?
- Is it possible to add a global variable that tracks the current func name for use in debugging in COM Errors?
  - Would have to remove crumbs on exiting function?
- For Border functions, make a Constant value rather than 3 Booleans?
  - Better way to make _LOWriter_DocHeaderGetTextCursor and _LOWriter_DocFooterGetTextCursor decide where to make the cursor?
- Look into Writer Web?
  = "private:factory/swriter/web"
- Add Themes? (Format>Theme)
- For DocCLose, remove the auto-save option?

## Base

Things pertaining to **Base**.

- Can't find a clean way to see if ReportDoc was created Hidden in _LOBase_ReportDocVisible
  - Currently doesn't cause a COM Error though.

## Calc

Things pertaining to **Calc**.

- Find Optimal Width/Height "Add" property? Format>Rows/Columns>OptimalWidth/Height
- Add Shapes?
- Make easier way to delete header fields after reinserting header -- or way to re-identify it?
- Can't I eliminate the need for Field Maps by using the main Field Object and enumerating the text to find the field? Or why did I not do this?
  - I think I Can't because the Field's anchor is always covering the entire text area, thus how would I identify it?
  - Maybe could enumerate, and get TextField Object for text portion? Then compare Objects?

## Draw

Things pertaining to **Draw**.

- Way to make comments? Insert>Comment (DrawPage> CreateandInsertAnnotation
- For future reference:  The PresentationDocument service implements the DrawingDocument service. This means that every presentation document looks like a drawing document. To distinguish between the two document types, you must first check for a presentation (Impress) document and then check for a drawing document. OOME 4.1. Pg 562

## Impress

Things pertaining to **Impress**.

- Can't set animation event duration and delay, see StackOverflow "LibreOffice Impress macro to read a slide's animation event duration and delay times"
  - Maybe can, but it's tricky. Several layers deep in Slide's animation node.
- Affine Matrix transformation DOES NOT seem to work using Transformation. Setting it to known values retrieved from LO doesn't return the shape to correct positioning.
- This note is in ConnectorModify: Currently, it seems to be not possible to disconnect a shape from the Start or End programatically.
- This note is in _LOImpress_FieldCurrentDisplayGet
  = Both Slide Title and Slide Number fields may return "<slide-name?" or "<number>" respectively instead of their current display value. I don't know why.
- A5, A4, B6ISO, B5ISO, B4ISO, B6JIS, B4JIS, Big 32 Kai, DL Envelope, C6 Env, Jap Post card, page format height/width is different from Calc/Writer. Just for information. Maybe doesn't matter
- This Note is in Master Slide delete: Trying to delete a Master Slide that is used by a slide will result in a processing error. I currently have no way of checking if a master slide is free to be deleted.
  = I could make a manual check, cycling through all slides and comparing Master Slide Objects, to see if it is used?
- This note is in Master Slide Transparency Gradient func:
  = While these properties can be set successfully, LibreOffice doesn't seem to apply it to the master slide, even when done using the UI.
- This note: I am unable to find the properties to set for "FitObject to Paper Format", "Background covers margins", "Slide numbers", and "Paper tray".
- Notes: I have not found a way to import Master slide from the LibreOffice templates yet.
- Add DrawShape glue point modify etc
- Need way to insert Form Controls/Fields etc?
- Need Effects/Transitions for Slides and Shapes
  = Add note that currently cant set or retrieve Shape Animation settings.
- Impress examples use "Slide 1" etc names, in other languages, this may not work, unless I set the name specifically?
  = The name is needed as the slide name is used for slideshow functions.
  = Retrieve array of names?
- Need Clear Dir Formatting func/ability.
  - See if I can refine/fix this.? Otherwise skip it.
- Add copy shape etc? Also copy textContent? Look at methods of shapes
- Numbering Styles is missing Graphics option support, if it can be added/not too complex??
- Since I made ShapeExists search all slides to match LO, should I add ShapeGetObjByName function, and maybe ShapesGetNamed? What about ShapeGetParent (Slide).
- Maybe I can add an "Apply to all" setting to applicable funcs, by modifying the master page?

- For future reference:  The PresentationDocument service implements the DrawingDocument service. This means that every presentation document looks like a drawing document. To distinguish between the two document types, you must first check for a presentation (Impress) document and then check for a drawing document. OOME 4.1. Pg 562

## Writer

Things pertaining to **Writer**.

- This note is added to Form Controls : Setting $iBorder to $LOW_FORM_CON_BORDER_WITHOUT, will not trigger an error, but does not currently work. This is a known bug, <https://bugs.documentfoundation.org/show_bug.cgi?id=131196>
- It's not possible to set DirFrmt Transparency Gradient (can't set TransparenceName)
- Make _LOWriter_DirFrmtGetCurStyles also Set curr style? Or Delete it?
- Add predefined List Style settings (Numbering Styles).
- Numbering Styles is missing Graphics option support, if it can be added/not too complex??
- Method available for cursor.Text?? convertToTable -- can replace dispatch? -- See Apache OpenOffice Community Forum - [Solved] Writer convertToTable
