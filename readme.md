# Outlook - Save selected emails as PDF

![Banner](./banner.svg)

> Outlook macro to save emails in as pdf files on your disk

## Description

Select one or more emails from within your Outlook client, click on a custom button of your ribbon and save them in a specific folder of your hard disk.

You can f.i. select 250 emails and in just a few clicks you can save them as pdf.

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

Note : you need to have Winword installed on your computer.

## Usage

1. Select one or more emails
2. Click on your `SaveAsPDFfile` button
3. A few popups will be displayed asking you for instance where to store the emails (as pdf files) and if you want to delete emails once saved as pdf or not.
4. That's it, wait a few and you'll get your mails saved on your disk.

![](images/demo.gif)

## License

[MIT](LICENSE)
