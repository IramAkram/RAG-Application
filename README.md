# Doc Studio

A Streamlit document-Q&A starter with animated dark UI, session-local FAISS indexes,
multi-file uploads, streamed answers, persistent source excerpts and chat export.

## Run
Use Python 3.11 or 3.12 in a new virtual environment.

    python -m venv .venv
    # macOS/Linux: source .venv/bin/activate
    # Windows PowerShell: .venv\Scripts\Activate.ps1
    python -m pip install --upgrade pip
    pip install -r requirements.txt
    streamlit run app.py

Keep .streamlit/config.toml inside the project folder and run from that folder.
Enter an OpenAI API key in the sidebar. Upload files, then click Build knowledge base.
Dependencies are ranges, not a tested lockfile. Resolve them in a clean environment;
freeze the working versions for deployment. This project was syntax-checked, not
end-to-end tested against the OpenAI API or a browser.

## Formats
- PDF: text extraction with page numbers; no OCR.
- DOCX: paragraphs and tables; no reliable page numbering.
- DOC: optional server-installed LibreOffice/soffice conversion. On Debian/Ubuntu,
  install LibreOffice Writer separately (sudo apt-get install libreoffice-writer).
  If unavailable, convert the file to DOCX before uploading.
- CSV: inferred delimiter; header row assumed.
- XLSX/XLS: all sheets; header row assumed; row labels start at 2.
- PPTX: slide text and tables; no OCR or speaker notes.
- TXT/MD/JSON/HTML: extracted or normalized text. HTML is not rendered or fetched.

## Limits and important caveats
10 files, 20 MB/file, 2 million extracted characters/file, 5 million/workspace.
Character limits are checked after parsing and do not prevent archive bombs or parser
resource exhaustion. Public deployments need authentication, isolated parsing processes,
CPU/memory/time limits, file content validation, rate limits and abuse protection.
Treat all uploaded documents as untrusted. LibreOffice conversion runs as a local
process and should be isolated on an untrusted/public deployment.

Documents go to OpenAI as embedding inputs; retrieved excerpts also go to the chat
model. Follow your organization's data policy. Keys are passed explicitly to clients,
not written to process environment. Session memory is not durable or encrypted storage;
the host may have a default environment key configured. Do not share such a deployment
without authentication and usage controls.

The prompt requests grounded answers and source IDs, but cannot guarantee grounding,
correct citations, completeness, or prompt-injection resistance. Review the evidence.
Follow-ups use a separate query-rewriting API call. Every build sends embeddings and
can incur charges. Indexes are not shared across user sessions or persisted to disk.

RAG over spreadsheet rows is not a deterministic analysis engine. For exact totals,
filters, rankings and whole-sheet comparisons, add a dedicated validated pandas/SQL
analysis flow. Large tables can produce many small chunks and significant indexing cost.
CSS uses Streamlit internal test IDs and may need adjusting after upgrades. Reduced-motion
preferences disable animation. Legacy DOC conversion requires a server-side LibreOffice
binary; pip requirements do not install it.

## Reference documentation
https://docs.streamlit.io/develop/api-reference/chat/st.chat_input
https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html
