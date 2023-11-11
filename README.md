# Abraham.Office

![](https://img.shields.io/github/downloads/oliverabraham/Abraham.Office/total) ![](https://img.shields.io/github/license/oliverabraham/Abraham.Office) ![](https://img.shields.io/github/languages/count/oliverabraham/Abraham.Office) ![GitHub Repo stars](https://img.shields.io/github/stars/oliverabraham/Abraham.Office?label=repo%20stars) ![GitHub Repo stars](https://img.shields.io/github/stars/oliverabraham?label=user%20stars)


## OVERVIEW

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

Licensed under Apache licence.
https://www.apache.org/licenses/LICENSE-2.0


## Compatibility

The nuget package was build with DotNET 6.



## INSTALLATION

Install the Nuget package "Abraham.Office" into your application (from https://www.nuget.org).
Please refer to my demo application in this repository for more details.


## HOW TO INSTALL A NUGET PACKAGE
This is very simple:
- Start Visual Studio (with NuGet installed) 
- Right-click on your project's References and choose "Manage NuGet Packages..."
- Choose Online category from the left
- Enter the name of the nuget package to the top right search and hit enter
- Choose your package from search results and hit install
- Done!


or from NuGet Command-Line:

    Install-Package Abraham.Office





## AUTHOR

Oliver Abraham, mail@oliver-abraham.de, https://www.oliver-abraham.de

Please feel free to comment and suggest improvements!



## SOURCE CODE

The source code is hosted at:

https://github.com/OliverAbraham/Abraham.Office

The Nuget Package is hosted at: 

https://www.nuget.org/packages/Abraham.Office


## Getting started

For an example refer to project "Abraham.Office.Demo". It demonstrates:
- how to generate a docx file from a docx template
- how to convert it to a PDF. (you must have doxillion document converter installed)
- and open up the outlook "new email" window



## SCREENSHOTS


# MAKE A DONATION !
If you find this application useful, buy me a coffee!
I would appreciate a small donation on https://www.buymeacoffee.com/oliverabraham

<a href="https://www.buymeacoffee.com/app/oliverabraham" target="_blank"><img src="https://cdn.buymeacoffee.com/buttons/v2/default-yellow.png" alt="Buy Me A Coffee" style="height: 60px !important;width: 217px !important;" ></a>
