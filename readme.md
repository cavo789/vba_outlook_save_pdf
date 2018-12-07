![Banner](images/banner.jpg)

# Outlook - Save selected email as PDF

> Outlook macro to save a selected item in the pdf-format

## Description

Select an email from withing your Outlook client, click on a custom button of your ribbon and save the email in a specific folder of your hard disk.

## Table of Contents

- [Install](#install)
- [Usage](#usage)
- [License](#license)

## Install

Get a copy of the `module.bas` VBA code and copy it into your Outlook client.

- Press `ALT-F11` in Outlook to open the `Visual Basic Editor` (aka VBE) window.
- Create a new module and copy/paste the content of the `module.bas` file that you can found in this repository
- Close the VBE
- Right-click on your Outlook ribbon to customize it so you can add a new button. Assign the `SaveAsPDFfile` subroutine to that button.

Note :

- Requires Word 2007 SP2 or Word 2010
- Requires a reference to "Microsoft Word <version> Object Library" (version is 12.0 or 14.0)

To add them, in the VBE window, click on the `Tools` then `References`

## Authors

Original author : [Robert Sparnaaij](http://www.howto-outlook.com/howto/saveaspdf.htm)
Modified by : Christophe Avonture

## License

[MIT](LICENSE)
