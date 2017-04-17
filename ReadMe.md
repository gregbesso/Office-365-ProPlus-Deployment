## Synopsis

This project is my attempt to reign in the tasks of deploying Office 365 ProPlus components to a variety of different types of user workstations.

## Code Example

You'll have to check out my blog post, it's the only time I used this and it's the only documentation or example I have of this type of
setup. That is up at: https://gregbesso.wordpress.com/office-365/deployment/

## Motivation

I wanted to make a silent install package that could be used for anyone's computer regardless of what version of Office they had, whether they had 32-bit or 64-bit, or 
if they had other products such as Project/Visio also installed.


## Installation

This project includes a handful of files that are placed on a UNC network share, and go along with the Microsoft-provided Office Deployment Tool. You should first download 
and install that tool, and then copy/paste these files into the folders that the tool creates. Then put all of them together somewhere on your network that is accessible by 
the workstations.

## API Reference

No API here. <sounds of crickets>

## Tests

No testing info here. <sounds of crickets>

## Contributors

Just a solo script project by moi, Greg Besso. Hi there :-)

## License

Copyright (c) 2017 Greg Besso

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.