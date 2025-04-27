# BugarTest

BugarTest is a simple, interactive and dynamic physical fitness test app created with [Excel VBA App](https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/getting-started-with-vba-in-office) and [Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel).

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes.

### Prerequisites

The things you need before installing the software.

* [Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel) installed (any version that support VBA, like Excel 2007 or newer)
* [Macro](https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/getting-started-with-vba-in-office#macros-and-the-visual-basic-editor) enabled when opening the project
* [VBA](https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/getting-started-with-vba-in-office) enabled (normally built into Excel)
* Placing the file in a [Trusted Location](https://support.microsoft.com/en-us/office/change-macro-security-settings-in-excel-a97c09d2-c082-46b8-b19f-e8621e8fe373) to avoid security warnings
* [Git](https://git-scm.com/downloads) installed (optional but recommended)

### Installation

A step by step guide that will tell you how to get the development environment up and running.

Using Git

```
$ git clone https://github.com/raulmaulidhino-dev/bugar-test-excel-vba.git
$ cd bugar-test-excel-vba/src
```

Or you can also install the zip file and unzip it after installation.


## Usage


To use the app, you can open the Excel with Macro file named `BugarTest.xlsm` inside `src` folder in your Microsoft Excel and start experimenting!


### Exporting Components

To export all VBA components in a folder, please follow these steps:
1. Open the **Excel Workbook** containing the BugarTest VBA project you want to export
2. Press <code>Alt</code> + <code>F11</code> to open the **VBA Editor**
3. Look for a module file named `ComponentExporter`
4. Click it to see the module file
5. Press <code>F5</code> to run the exporting process
6. After the process finished, you will get a window telling you the folder path containing the exported components


### Importing Components

To import all VBA components in a folder, please follow these steps:
1. Open the **Excel Workbook** containing the BugarTest VBA project you want to import
2. Press <code>Alt</code> + <code>F11</code> to open the **VBA Editor**
3. Look for a module file named `ComponentImporter`
4. Click it to see the module file
5. Before running the file, make sure there is a folder named `components` (or you can change it in the file as you like) containing the VBA components (in .bas or .frm [with .frx] format)
6. Press <code>F5</code> to run the importing process
7. After the process finished, you will get a window telling you that the files were imported successfully!


Made with ❤️ in Indonesia 🇮🇩


MIT (Modified) © [Raul Maulidhino](https://rauldev.my.id)
