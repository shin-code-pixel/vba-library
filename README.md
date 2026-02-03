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

This repository is for personal and educational use.  
Feel free to reference and modify the code for your own projects.
