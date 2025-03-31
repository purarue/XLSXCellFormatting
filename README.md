# XLSXCellFormatting

A .NET CLI tool that segments rich text from an Excel file and saves it to JSON.

The [code here](./Program.cs) is licensed under the ['Unlicense License'](./LICENSE), feel free to use it as you wish/as an example for how to extract rich text, but `EPPlus` (the spreadsheet library this uses) itself is licensed under [Polyform Noncommercial license](https://www.epplussoftware.com/en/LicenseOverview/)

## Installation

- Install the `dotnet` SDK

1. Clone/Download dependencies/build:

```sh
git clone https://github.com/purarue/XLSXCellFormatting
cd ./XLSXCellFormatting
dotnet build
```

## Usage

```sh
dotnet run <input.xlsx> <output.json>
```

This generates a mapping like `{"<worksheet name>": [cell1, cell2, ...]}`, where each cell looks like this:

```json
{
  "Row": 26,
  "Column": 4,
  "FullText": "Here is some text\nThis is bold\nThis is italic\nThis is underlined.",
  "RichTextChildren": [
    {
      "Text": "Here is some text\n",
      "Bold": false,
      "Italic": false,
      "Underline": false
    },
    {
      "Text": "This is bold",
      "Bold": true,
      "Italic": false,
      "Underline": false
    },
    {
      "Text": "\n",
      "Bold": false,
      "Italic": false,
      "Underline": false
    },
    {
      "Text": "This is italic",
      "Bold": false,
      "Italic": true,
      "Underline": false
    },
    {
      "Text": "\n",
      "Bold": false,
      "Italic": false,
      "Underline": false
    },
    {
      "Text": "This is underlined.",
      "Bold": false,
      "Italic": false,
      "Underline": true
    }
  ]
}
```
