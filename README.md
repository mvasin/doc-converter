# PDF/DOC(X) Converter

`@mvasin/doc-converter` converts `.doc`, `.docx`, and `.pdf` resumes into Markdown files, optionally clearing prior output before writing.

## Usage

1.  Run without installation:

    ```bash
    npx @mvasin/doc-converter --input <file|directory> --output <file|directory> [--clear-output]
    ```

    - `--input` **(required)**: path to a single `.doc`, `.docx`, or `.pdf`, or a directory containing those files.
    - `--output` **(required)**: a file path when the input is a single file, or a directory path when the input is a folder.
    - `--clear-output` _(optional)_: when provided, deletes existing output files/folders before conversion; otherwise, existing `.md` files are preserved and conversion skips them.

2.  When testing locally without publishing, link the package and execute the bin:

    ```bash
    npm link
    doc-converter --input ./input/resume.docx --output ./resume.md
    ```

## Requirements

- Node.js (see `package.json` for engine/dep versions).
- LibreOffice (for headless `.doc` → `.docx` conversion).

## What it does under the hood

1. Cleans the `output/` directory so each run starts fresh.
2. Accepts `.doc`, `.docx`, and `.pdf` resumes from `output/`.
3. Converts `.doc` files to `.docx` via `soffice --headless --convert-to docx` (LibreOffice must be installed on the host).
4. Pipes every `.docx` through `mammoth` and `turndown` to emit Markdown, stripping inline images.
5. Extracts text from PDFs using `pdf-parse` and formats it with paragraph spacing for readability.
6. Logs success/failure for each file and leaves `.md` files in `output/`.

## NPM dependencies

- `mammoth`: reads `.docx` and produces HTML.
- `turndown`: converts Mammoth’s HTML into Markdown while allowing custom rules like image stripping.
- `pdf-parse`: pulls raw text from PDFs for Markdown output.
- Node.js built-ins: `fs`, `path`, `os`, and `child_process` for file handling and LibreOffice calls.
