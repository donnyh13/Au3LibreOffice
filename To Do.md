# UDF To Do

The following is a basic list of things needing done, added to, or looked at, in this UDF.

## UDF

Things pertaining to the **entire UDF**, or **equally to all sub-Components**.

- Add Hatch Background option
- Add Macros functions? Plus Integrate with all Functions that can activate a Macro?
- Null Variables called in Close and Delete functions.
- Make a global Open/connectAll/connectCurrentCreate func
- Is it important to use "Get" / "set"? e.g. getText, etc? Instead of Text()?
	- <https://ask.libreoffice.org/t/calc-named-range-err-508/> See comment by JohnSUN
	- OOME 4.1 pg 306

## Writer

Things pertaining to **Writer**.

- This note is added to Form Controls: Setting $iBorder to $LOW_FORM_CON_BORDER_WITHOUT, will not trigger an error, but does not currently work. This is a known bug, <https://bugs.documentfoundation.org/show_bug.cgi?id=131196>

## Calc

Things pertaining to **Calc**.

- Find Optimal Width/Height "Add" property? Format>Rows/Columns>OptimalWidth/Height
- Add Shapes?
- Make easier way to delete header fields after reinserting header -- or way to re-identify it?

## Base

Things pertaining to **Base**.

## Impress

Things pertaining to **Impress**.

- Can't set animation event duration and delay, see StackOverflow "LibreOffice Impress macro to read a slide's animation event duration and delay times"
- Move _LOImpress_CursorInsertString to somewhere other than Helper.
- Can I help this? -- Warning! For some reason this function doesn't seem to set the modified status to True. Changes could be inadvertently lost due to this, if the user closes without saving.
- Affine Matrix transformation DOES NOT seem to work using Transformation. Setting it to known values retrieved from LO doesn't return the shape to correct positioning.
- This note is in ConnectorModify: Currently, it seems to be not possible to disconnect a shape from the Start or End programatically.
- Add DrawShape glue point modify etc

- For future reference:  The PresentationDocument service implements the DrawingDocument service. This means that every presentation document looks like a drawing document. To distinguish between the two document types, you must first check for a presentation (Impress) document and then check for a drawing document. OOME 4.1. Pg 562

## Draw

Things pertaining to **Draw**.

- Implement Draw.
- For future reference:  The PresentationDocument service implements the DrawingDocument service. This means that every presentation document looks like a drawing document. To distinguish between the two document types, you must first check for a presentation (Impress) document and then check for a drawing document. OOME 4.1. Pg 562
