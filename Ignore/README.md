![image](https://github.com/mlipok/Au3LibreOffice/assets/11089482/e7e0966b-2f25-41d9-b927-cb5661dc6c6b)

# Au3LibreOffice

## Description

This [AutoIt](https://www.autoitscript.com/) UDF for [LibreOffice API/SDK](https://api.libreoffice.org/) provides tools for the automation of tasks in [LibreOffice](https://www.libreoffice.org/) using the [AutoIt Scripting language](https://www.autoitscript.com/).

## Currently Supported

Au3LibreOffice UDF currently provides support for the following LibreOffice components:<br>
- **Writer**<br> 
- **Calc** ***In Development***
- ~~Draw~~ ***Pending Development*** 
<br>
Support for other components will be provided as time permits.<br>

## Release

https://github.com/mlipok/Au3LibreOffice/releases/latest

## Changes

Please see the [Changelog](CHANGELOG.md)

## License

Distributed under the MIT License. See the [LICENSE](LICENSE) for more information.

## Notes

- This UDF currently works **only** with the **INSTALLED** version of LibreOffice. The **Portable** version will **Not work**.
- For those using AutoIt versions **older** than **_3.3.16.1,_** one internal function used for “Saving as” and “Exporting” documents uses Maps, which will **Not** be recognized as proper syntax in AutoIt. 
- LibreOffice uses Micrometers for sizing internally, all functions in this UDF, unless otherwise stated, use Micrometers. A converter has been created for converting to/from Inches, Centimeters, Printer’s Points, and Millimeters to/from Micrometers, for all sizing needs. _ConvertFromMicrometer, and _ConvertToMicrometer. Either for Writer (LOWriter) or Calc (LOCalc).
- LibreOffice uses the Long color format for all color settings, A converter has also been created for converting from/to Hex; (R)ed, (G)reen, (Blue); (H)ue, (S)aturation, and (B)rightness; and (C)yan, (M)agenta, (Y)ellow, Blac(K); to/from long color format. _ConvertColorFromLong, and _ConvertColorToLong. Either for Writer (LOWriter) or Calc (LOCalc).
- This UDF was first made public here: https://www.autoitscript.com/forum/index.php?showtopic=210514

## Acknowledgements

- Opportunity by [GitHub](https://github.com)
- Scripting ability by [AutoIt](https://www.autoitscript.com/site/autoit/)
- Thanks to the authors of the Third-Party UDFs:
  - *OOo/LibOCalcUDF* by @GMK, @Leagnus, @Andy G, @mLipok. [OOo/LibOCalcUDF](https://www.autoitscript.com/forum/topic/151530-ooolibo-calc-udf/)
  - *“WriterDemo.au3”* by @mLipok. [WriterDemo](https://www.autoitscript.com/forum/topic/204665-libreopenoffice-writer/?do=findComment&comment=1471711)
  - *Printers Management UDF* by @jguinch. [Printers Management UDF](https://www.autoitscript.com/forum/topic/155485-printers-management-udf/)
  - *Word* UDF supplied with AutoIt by @water.
- Thanks to Andrew Pitonyak for his invaluable book on writing Open Office/ Libre Office Macros, and his Macro collection document.
  - [OpenOffice.org Macros Explained — OOME Third Edition](https://www.pitonyak.org/OOME_3_0.pdf)
  - [OpenOffice.org Macros Explained — OOME Fourth Edition](https://www.pitonyak.org/OOME_4_1.odt)
  - [Useful Macro Information For OpenOffice.org](https://www.pitonyak.org/AndrewMacro.pdf)
  - Andrew Pitonyak's website: https://www.pitonyak.org/oo.php
- Thanks to the maintainers
  - Thanks to [@mLipok](https://github.com/mLipok) for hosting this project on his GitHub. Not to mention his tireless energy and long hours of development and code review and clean-up.
  - Thanks to [@donnyh13](https://github.com/donnyh13) for the initial project creation and development.
  - **Big thanks** to all the hard-working contributors.

## Links 

[License]https://github.com/mlipok/Au3LibreOffice/tree/main/LICENSE <br>
[AutoIt](https://www.autoitscript.com/site/autoit/) <br>
[AutoIt Forum Post](https://www.autoitscript.com/forum/index.php?showtopic=210514) <br>


