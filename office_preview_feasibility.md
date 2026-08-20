# Feasibility Analysis: Using Open XML SDK for Office File Previews in PowerToys Peek

## 1. Feasibility Verdict: Partial / Highly Complex for Visual Previews
While the **Open XML SDK** is the official Microsoft .NET library for parsing and manipulating Office documents (`.docx`, `.pptx`, `.xlsx`), using it to generate **rich visual previews** (like rendering the document as it appears in Office) is **not directly feasible**.

### Why?
- **No Rendering Engine:** The Open XML SDK is designed for reading, writing, and manipulating the underlying XML structure of Office documents. It does not include a layout or rendering engine.
- **Complexity of Office Formats:** To accurately preview a document, you must calculate pagination, font rendering, text wrapping, and complex layouts (especially for `.pptx` and `.xlsx`). Recreating a rendering engine from scratch using the raw XML is an immensely complex task.

### What it *can* be used for:
- **Text Extraction:** Quickly extracting raw text from a `.docx` or `.xlsx` file to show a plain-text preview.
- **Metadata & Structure:** Reading document properties (author, title) or structure (number of slides, sheet names).
- **Embedded Images/Thumbnails:** Extracting embedded media or pre-generated thumbnails (if the document was saved with a thumbnail `docProps/thumbnail.jpeg`).

## 2. Alternative Approaches & Implementations
To actually display previews in Peek, you would need to convert the documents into a renderable format (like HTML, PDF, or an Image).

### Approach A: Use Open XML SDK + Open-Xml-PowerTools (For Text/HTML)
You can use the Open XML SDK alongside a library like `Open-Xml-PowerTools` (or community forks like `OpenXmlPowerTools`) to convert Word documents (`.docx`) to HTML, which Peek's `WebBrowserPreviewer` (WebView2) can render.
- **Pros:** Does not require Office to be installed.
- **Cons:** HTML conversion is often imperfect. Complex layouts, tables, and SmartArt may break or not render correctly. `.xlsx` and `.pptx` conversion is much harder or unsupported.

### Approach B: Rely on Windows Shell Preview Handlers (Existing capability)
Windows already provides `IPreviewHandler` interfaces. When Microsoft Office (or a standalone viewer) is installed, Windows automatically registers preview handlers for `.docx`, `.xlsx`, and `.pptx`. PowerToys Peek already uses `ShellPreviewHandlerPreviewer` which leverages these handlers.
- **Pros:** Perfect rendering fidelity (it uses Office's native renderers).
- **Cons:** Fails if the user does not have Office or a compatible preview handler installed.

### Approach C: Extract Embedded Thumbnails
Many Office documents are saved with an embedded thumbnail (if the "Save Thumbnail" option is checked in Office). The Open XML SDK can be used to extract this `thumbnail.jpeg` directly from the `.zip` / `.docx` package and display it.
- **Pros:** Extremely fast, uses native Peek `ImagePreviewer`.
- **Cons:** Not all documents have embedded thumbnails. It's only a single page/slide view.

## 3. Implementation Plan (If Proceeding with Open XML SDK)
If the goal is to provide a fallback text/basic HTML preview when no Shell Preview Handler is available, here is the plan:

### Step 1: Add Dependencies
- Add the `DocumentFormat.OpenXml` NuGet package to the `Peek.FilePreviewer` project.

### Step 2: Create a DocumentPreviewer Model
- Create an `OfficePreviewer` class that implements `IPreviewer`.
- Register `.docx`, `.xlsx`, `.pptx` in the supported file types.
- In `PreviewerFactory.cs`, add logic to route Office extensions to `OfficePreviewer`. This should act as a fallback if `ShellPreviewHandlerPreviewer.IsItemSupported()` fails or is not available.

### Step 3: Implement Extraction Logic
- **For `.docx`:** Use `WordprocessingDocument.Open()` to extract paragraphs and text streams. Build a basic HTML string or Markdown representation.
- **For `.xlsx`:** Use `SpreadsheetDocument.Open()` to extract cell values from the first few sheets and build an HTML `<table>`.
- **For `.pptx`:** Use `PresentationDocument.Open()` to extract slide text or embedded images/thumbnails.

### Step 4: Display the Preview
- Pass the generated HTML string or extracted text/image to the existing `WebBrowserPreviewer` (via a temporary HTML file).
- Ensure temporary files are cleaned up using `FilePreviewCommon.Helper.CleanupTempDirAsync`.

### Step 5: Handle Asynchronous Loading and Exceptions
- Office documents can be large. Wrap the Open XML SDK parsing in `Task.Run` to prevent blocking the UI thread.
- Handle `OpenXmlPackageException` (e.g., corrupted files, password-protected documents) gracefully by showing an `UnsupportedFilePreviewer` or an error state.
