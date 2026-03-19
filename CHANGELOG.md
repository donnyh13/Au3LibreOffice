# Changelog

All notable changes to ["Au3LibreOffice"](https://github.com/mlipok/Au3LibreOffice/tree/main) SDK/API will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
This project also adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

Go to [legend](#legend---types-of-changes) for further information about the types of changes.

## Releases

|    Version       |    Changes                         |    Download                 |     Released   |    Compare on GitHub       |
|:-----------------|:----------------------------------:|:---------------------------:|:--------------:|:---------------------------|
|    **v0.0.0.2**  | [Change Log](#0002---2023-07-16)   |                             | 2023-07-16     |                            |
|    **v0.0.0.1**  | [Change Log](#0001---2023-07-02)   | [v0.0.0.1][v0.0.0.1]        | 2023-07-02     |                            |

[To Top](#releases)

## [0.0.0.2] - 2023-07-16

### LibreOfficeWriter

#### Changed

- `_DocReplaceAllInRange` to have two methods of performing a Regular Expression find and replace.
- Method for skipping $atFindFormat and $atReplaceFormat, now uses an empty array called in each parameter to skip.
  - _DocReplaceAll,
  - _DocReplaceAllInRange.

#### Documented

- UDF version number in the UDF Header.
- Updated function documentation to reflect the changes.

#### Refactored

- Removed the if/else block in $atFindFormat parameter checking.
  - _DocReplaceAll,
  - _DocReplaceAllInRange,
  - _DocFindNext,
  - _DocFindAll,
  - _DocFindAllInRange.

[To Top](#releases)

## [0.0.0.1] - 2023-07-02

### LibreOfficeWriter

#### Added

- Initial UDF Release.

[To Top](#releases)

---

#### Legend - Types of changes

- `Added` for new features.
- `Changed` for changes in existing functionality.
- `Deprecated` for soon-to-be removed features.
- `Documented` for documentation only changes.
- `Fixed` for any bug fixes.
- `Refactored` for changes that neither fixes a bug nor adds a feature.
- `Removed` for now removed features.
- `Security` in case of vulnerabilities.
- `Styled` for changes like whitespaces, formatting, missing semicolons etc.

[To the top](#changelog)

---

[v0.0.0.1]: https://github.com/donnyh13/Au3LibreOffice/releases/tag/v0.0.0.1
