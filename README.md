# XLSXCellFormatting

A .NET CLI tool that segments rich text from an Excel file and saves it to JSON.

## Installation

- Install the `dotnet` SDK
- A compatible Excel file (e.g., `.xlsx`)

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

This generates a mapping like `{"<worksheet name>": [cell1, cell2]}`, where each cell looks like this:

```
{
  "Row": 26,
  "Column": 4,
  "FullText": "Here is some text\nThis is bold\nThis is italic\nThis is underlined.",
  "RichTextChildren": [
    {
      "Text": "Here is some text",
      "IsBold": false,
      "IsItalic": false,
      "IsUnderline": false
    },
    {
      "Text": "\n",
      "IsBold": false,
      "IsItalic": false,
      "IsUnderline": false
    },
    {
      "Text": "This is bold",
      "IsBold": true,
      "IsItalic": false,
      "IsUnderline": false
    },
    {
      "Text": "\n",
      "IsBold": false,
      "IsItalic": false,
      "IsUnderline": false
    },
    {
      "Text": "This is italic",
      "IsBold": false,
      "IsItalic": true,
      "IsUnderline": false
    },
    {
      "Text": "\n",
      "IsBold": false,
      "IsItalic": false,
      "IsUnderline": false
    },
    {
      "Text": "This is underlined.",
      "IsBold": false,
      "IsItalic": false,
      "IsUnderline": true
    }
  ]
}
```
