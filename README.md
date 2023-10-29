# Au3LibreOffice

## Description

This AutoIt UDF for LibreOffice API/SDK provides tools for the automation of tasks in LibreOffice using the AutoIt Scripting Program.

## Currently Supported

Au3LibreOffice UDF currently provides support for the following LibreOffice components:<br>
- **Writer**<br> 
- ~~Calc~~ ***In Development***
- ~~Draw~~ ***In Development*** 
<br>
Support for other components will be provided as time permits.<br>

## Release

https://github.com/mlipok/Au3LibreOffice/releases/latest

## Changes

Please see the [Changelog](CHANGELOG.md)

## License

Distributed under the MIT License. See [LICENSE](LICENSE.md) for more information.

## Notes

- This UDF currently works **only** with the **INSTALLED** version of LibreOffice. The **Portable** version will not work.
- For those using AutoIt versions **older** than **_3.3.16.1,_** one internal function used for “Saving as” and “Exporting” documents uses Maps, which will **Not** be recognized as proper syntax in AutoIt. 
- LibreOffice uses Micrometers for sizing internally, all functions, unless otherwise stated, use Micrometers. A converter has been created for converting to/from Inches, Centimeters, Printer’s Points, and Millimeters to/from Micrometers, for all sizing needs. _LOWriter_ConvertFromMicrometer, and  _LOWriter_ConvertToMicrometer.
- LibreOffice uses the Long color format for all color settings, A converter has also been created for converting from/to Hex; (R)ed, (G)reen, (Blue); (H)ue, (S)aturation, and (B)rightness; and (C)yan, (M)agenta, (Y)ellow, Blac(K); to/from long color format. _LOWriter_ConvertColorFromLong, and _LOWriter_ConvertColorToLong.
<br>
- This UDF was first made public here:<br>
https://www.autoitscript.com/forum/index.php?showtopic=210514

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
  - [Useful Macro Information For OpenOffice.org]()
- Thanks to the maintainers
  - Thanks to [@mLipok](https://github.com/mLipok) for hosting this project on his GitHub. Not to mention his tireless energy and long hours of development and code review and clean-up.
  - Thanks to [@donnyh13](https://github.com/donnyh13) for the initial project creation and development.
  - **Big thanks** to all the hard-working [contributors]

## Links 

[License]: https://github.com/mlipok/Au3LibreOffice/tree/main/LICENSE
[AutoIt](https://www.autoitscript.com/site/autoit/)
[AutoIt Forum Post](https://www.autoitscript.com/forum/index.php?showtopic=210514)


