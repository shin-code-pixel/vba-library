# VBA Utility Library

A personal collection of VBA class modules, standard modules, and UserForms  
for Excel / Outlook / Access / PowerPoint.

This repository contains reusable VBA components that I use as a base  
for building tools and automation scripts.

---

## Contents

- `.xlsm` workbooks (sample / runnable tools)
- `.bas` standard modules
- `.cls` class modules
- `.frm` / `.frx` UserForms

Each folder represents an independent tool or utility set.

---

## How to use

1. Open the `.xlsm` workbook in the `workbook` folder  
2. Import required modules from the `src` folder:
   - `.bas` → Standard Module
   - `.cls` → Class Module
   - `.frm` + `.frx` → UserForm
3. Run the entry point macro (usually in `Mdl_Main`)

---

## Design policy

- `src/` folder is the **canonical source**
- `.xlsm` files are provided as runnable examples
- All modules are designed to be:
  - reference-free (no external libraries)
  - portable between projects
  - safe for copy & reuse

---

## Environment

- Microsoft Excel VBA (Office 365 / 2021)
- 64bit Windows
- No external dependencies

---

## License

Copyright (c) 2026 shin-code-pixel

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
