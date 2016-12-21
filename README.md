# EPPlus.TableAsEnumerable
Generic extension method enabling retrival ExcelTable rows as an enumeration of typed objects.

## Synopsis

This project adds an extension method to EPPlus ExcelTable objects that enables generic retrival of the data within. 

At the top of the file there should be a short introduction and/ or overview that explains **what** the project is. This description should match descriptions added for package managers (Gemspec, package.json, etc.)

## Code Example

Show what the library does as concisely as possible, developers should be able to figure out **how** your project solves their problem by looking at the code example. Make sure the API you are showing off is obvious, and that your code is short and concise.

## Motivation

[EPPLus](http://epplus.codeplex.com/) is a great tool allowing manipulation of Excel xlsx format files without Excel installed and interop, by directly manipulating the OpenXML format.

EPPLus lacks of some features, for example strongly typed access of data. This project aims to fill this gap partially by providing a way to read table objects stored in the worksheets as an enumerable of typed objects.

These types can be defined by the user, as common public classes containing publicly settable properties. Thes properties can be than mapped to columns in the table. It is important to notice that only classes with parameterless constructor can be used for this purpose.

It is not necessary to map all class properties, and it is not necessary either to map properties to all columns. Mapping can be done by either the name or the index of the column. Automatic mapping to the property name is also possible. 

## Installation

Provide code examples and explanations of how to get the project.

## API Reference

Depending on the size of the project, if it is small and simple enough the reference docs can be added to the README. For medium size to larger projects it is important to at least provide a link to where the API reference docs live.

## Tests

Describe and show how to run the tests with code examples.

## Contributors

Let people know how they can dive into the project, include important links to things like issue trackers, irc, twitter accounts if applicable.

## License

A short snippet describing the license (MIT, Apache, etc.)