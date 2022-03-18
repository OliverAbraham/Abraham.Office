# Abraham.Office

This library is a Nuget package to manipulate docx files (Microsoft Office file format).
It works as a generator for docx files, by reading a docx template and replace tokens in the file.
The basic idea was to generate and email invoices.
To use Microsoft Word as a layout editor, I edit the template in word and then
use my nuget package to replace tokens by the actual values.
Then, use a docx to PDF converter (I'm using doxillion document converter).
Then I open up a new email window in Microsoft Outlook.
This gives me a very easy to use solution to generate and email invoices.


## Source code

The source code is hosted at https://github.com/OliverAbraham/Abraham.Office


## License

Licensed under GPL v3 license.
https://www.gnu.org/licenses/gpl-3.0.en.html


## Compatibility

The nuget package was build with DotNET 6.

## Example

For an example refer to project "Abraham.Office.Demo". It demonstrates:
- how to generate a docx file from a docx template
- how to convert it to a PDF. (you must have doxillion document converter installed)
- and open up the outlook "new email" window


## Getting started

Follow the demo project to see the implementation.
Inspect the Template.docx in the repository.
That's all!
