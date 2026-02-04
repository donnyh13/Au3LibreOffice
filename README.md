![image](https://github.com/mlipok/Au3LibreOffice/assets/11089482/e7e0966b-2f25-41d9-b927-cb5661dc6c6b)

# Au3LibreOffice

## Description

This [AutoIt](https://www.autoitscript.com/) UDF for [LibreOffice API/SDK](https://api.libreoffice.org/) provides tools for the automation of tasks in [LibreOffice](https://www.libreoffice.org/) using the [AutoIt Scripting language](https://www.autoitscript.com/).

## Currently Supported

Au3LibreOffice UDF currently provides support for the following LibreOffice components:

- **Writer**
- **Calc**
- **Base**
- ~~Impress~~ ***Development Pending***

Support for other components will be provided as time permits.

## Release

<https://github.com/mlipok/Au3LibreOffice/releases/latest>

## Changes

Please see the [Changelog](CHANGELOG.md)

## License

Distributed under the MIT License. See the [LICENSE](LICENSE) for more information.

## Notes

- ~~This UDF currently works **only** with the **INSTALLED** version of LibreOffice. The **Portable** version will **Not work**.~~
- The ability to Automate LibreOffice Portable (and OpenOffice) has been added, and *Should* work correctly. It currently adds some temporary Registry entries however. See `_LO_InitializePortable`.
- For those using AutoIt versions **older** than ***3.3.16.1,*** some functions use Maps, which will **Not** be recognized as proper syntax in the older AutoIt versions.
- LibreOffice uses Hundredths of a Millimeter (100th MM) for sizing internally, all functions in this UDF, unless otherwise stated, use these units. A converter has been created for converting Inches, Centimeters, Printer’s Points, and Millimeters to/from 100th MMs, for all sizing needs. `_LO_UnitConvert`.
- LibreOffice uses RGB color value as a long integer for all color settings, A converter has also been created for converting from/to Hex; (R)ed, (G)reen, (Blue); (H)ue, (S)aturation, and (B)rightness; and (C)yan, (M)agenta, (Y)ellow, Blac(K); to/from long integer RGB color value. `_LO_ConvertColorFromLong`, and `_LO_ConvertColorToLong`.
- This UDF was originally posted here: <https://www.autoitscript.com/forum/index.php?showtopic=210514>
- Active development is currently taking place here: <https://github.com/mlipok/Au3LibreOffice>
- Support can be found at <https://github.com/mlipok/Au3LibreOffice/issues> Or <https://www.autoitscript.com/forum/index.php?showtopic=210514>.

## Acknowledgements

- Collaboration opportunity by [GitHub](https://github.com)
- Scripting ability by [AutoIt](https://www.autoitscript.com/site/autoit/)
- Thanks to the authors of these Third-Party UDFs:
  - *OOo/LibOCalcUDF* by @GMK, @Leagnus, @Andy G, @mLipok. [OOo/LibOCalcUDF](https://www.autoitscript.com/forum/topic/151530-ooolibo-calc-udf/)
  - *“WriterDemo.au3”* by @mLipok. [WriterDemo](https://www.autoitscript.com/forum/topic/204665-libreopenoffice-writer/?do=findComment&comment=1471711)
  - *Printers Management UDF* by @jguinch. [Printers Management UDF](https://www.autoitscript.com/forum/topic/155485-printers-management-udf/)
  - *Word* UDF supplied with AutoIt by @water.
- Thanks to Andrew Pitonyak for his valuable books on writing Open Office/Libre Office Macros, and his Macro collection document.
  - [OpenOffice.org Macros Explained — OOME Third Edition](https://www.pitonyak.org/OOME_3_0.pdf)
  - [OpenOffice.org Macros Explained — OOME Fourth Edition](https://www.pitonyak.org/OOME_4_1.odt)
  - [Useful Macro Information For OpenOffice.org](https://www.pitonyak.org/AndrewMacro.pdf)
  - [OpenOffice.org Base Macro Programming](https://www.pitonyak.org/database/AndrewBase.odt)
  - Andrew Pitonyak's website: <https://www.pitonyak.org/oo.php>
- Thanks to the following maintainers and contributors:
  - [@mLipok](https://github.com/mLipok) for hosting this project on his GitHub. As well as his tireless energy during the long hours of development, code review and clean-up.
  - [@donnyh13](https://github.com/donnyh13) for the initial project creation and further development.

## Links

[License](https://github.com/mlipok/Au3LibreOffice/tree/main/LICENSE)  
[AutoIt](https://www.autoitscript.com/site/autoit/)  
[AutoIt Forum Post](https://www.autoitscript.com/forum/index.php?showtopic=210514)  
