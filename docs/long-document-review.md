# Long-document review

The Issues tab reviews the complete document or exact selection in consecutive batches. The previous limits of 80 paragraphs / 12,000 characters for the body and 6,000 characters for a selection are now per-request defaults, rather than a cutoff for the entire operation. A paragraph larger than one request is split while retaining its original paragraph number. The Review tab also splits oversized items and rejoins all successful fragments before publishing a suggestion.

Use **Review document** on the Issues tab to collect findings. For tracked edits, use **Load document**, **Generate all**, then review the suggestions and accept the required changes. Existing character-level tracked writeback remains in use. A failed or cancelled fragment does not expose a partial paragraph as an applicable suggestion.

The UI reports reading and generation progress, completed/failed/unprocessed batches or items, and character coverage for issue review. Empty responses, invalid JSON, invented locations, duplicate suggestion IDs, and explicit output-length truncation are not counted as successful processing. A valid empty issue array is a completed batch with no findings. Cancellation reaches the model request. Document changes during issue review stop the queue, and results retain their document identity and original paragraph hash for later positioning.

## Copies, saved results and retry

Both review tabs provide **Review a copy**, **History / restore**, **Continue / retry failed**, and **Export review report**.

**Review a copy** asks for a new file name, saves the current in-memory document using Word's native `SaveAs2`, and makes the copy the active document. The original file is not overwritten. `.doc`, `.docx`, `.docm`, Word templates and RTF retain their native save format; an unsaved docx document containing a VBA project uses `.docm`. Existing formatting, document parts, comments and revisions remain under Word's native save handling. In the Review tab all supported document text areas are then loaded, ready for **Generate all**; in Issues, start **Review document** on the copy. Applying suggestions or adding review comments saves the working copy. Bulk apply saves once after restoring Word's previous tracking state.

The review plan, successful results, failed attempts and user actions are automatically saved on this computer. Review-card edits are also saved after a short typing delay, while the original model output remains available in the report. Each new full generation keeps a separate session with its model, processing dates, instructions and settings. A same-session continuation preserves those settings and processes only unfinished/failed items. Oversized review cards keep successful fragments, retry only the remaining fragments, and publish an applicable suggestion only after the complete card succeeds.

To continue after restarting Word, open the same saved document or review copy, choose **History / restore**, select the session, then choose **Continue / retry failed**. Recovery checks file identity, text in every supported area, paragraph structure and model/instruction/budget settings. Native review edits update the expected source stamp and record offset changes for remaining cards. Unrelated text changes require loading a new scope. Formatting-only changes are not part of this text stamp. Newly created unsaved documents support in-memory retry, but must be saved and reloaded before a session can be restored after closing Word. An original document with unsaved edits must also be saved before closing; creating a review copy handles that save as part of the workflow.

An older session can be exported directly from history even when the current body no longer matches. Markdown and JSON reports include the full processing range, source text, suggestions/findings, dates, model, safe settings, every work item, earlier failure attempts, stop/error history and user actions. The report does not apply edits automatically. It contains document text because it is a complete local review record; the user chooses where to save it. Provider credentials and endpoints are excluded from the configuration snapshot.

The cache is an encrypted per-item journal under `%LOCALAPPDATA%\WordOllama\ReviewSessions`, protected by Windows DPAPI for the current user. Completing a batch rewrites only that batch and the summary. Atomic replacements keep encrypted backups; an interrupted request becomes pending, and a missing payload is reported as an error. The saved cache is not portable to a different Windows account. These features use local files and do not add an online service; model requests still follow the selected provider. For offline review, select Ollama with a local endpoint.

## Built-in Agent (V2.5)

Tell the built-in Agent, for example: "Proofread the whole document, create a review copy, apply precise tracked corrections, and save the document and reports." The shared review queue handles context limits and progress; the Agent snapshot is not used to infer full coverage.

| Tool | Behavior |
| --- | --- |
| `review_document` | `kind` is `suggestions` (default) or `issues`; accepts `instruction`, `create_copy` (default true) and an optional new absolute `copy_path` |
| `list_review_sessions` | Lists current-document and related review-copy sessions with IDs, progress and paths |
| `resume_review` | Takes `session_id`; checks the opened document/copy and source stamp, retaining completed results and retrying unfinished work |
| `apply_review_revisions` | Applies completed suggestions as character-level tracked edits; requires Agent Tracked Changes mode |
| `save_review_document` | Saves the current review document and session; exports full Markdown/JSON reports by default (`export_report=false` skips export) |

Generation does not apply changes automatically. View Only and Propose Changes modes can generate suggestions; applying revisions requires Tracked Changes mode. Tool results report failed/pending batches and unapplied suggestions. If a saved review copy is not open, resume returns its path for opening rather than binding the session to a different file. The default Agent task allowance still applies to the Agent loop; configure a longer allowance before a long unattended run. A single review request follows the selected provider timeout.

## Revision authors

`ReviewReviewerName` defaults to **AI审批**. `ReviewReviewerIdentityMode` defaults to `Date`, yielding names such as `AI审批-2026-10-12`. `Fixed` uses only the name; `Session` adds the session date and an eight-character run ID. Settings provides these choices. AI writeback temporarily sets Word's user name/initials and local-user option, then restores them immediately, including on errors. A native review copy retains revision-author information on save. The original source file and existing revision authors are preserved.

## Settings

The **Long-document review and processing budgets** panel and `appconfig.json` support these settings. Existing configuration files receive the defaults below automatically.

| Setting | Default | Meaning |
| --- | ---: | --- |
| `ReviewReviewerName` | AI审批 | Base name for AI tracked edits and comments |
| `ReviewReviewerIdentityMode` | Date | Fixed, Date or Session |
| `ReviewIssueBatchMaxParagraphs` | 80 | Maximum paragraph fragments in an issue-review request |
| `ReviewIssueBatchMaxCharacters` | 12000 | Maximum body issue-review context characters, including paragraph labels |
| `ReviewSelectionBatchMaxCharacters` | 6000 | Maximum selection issue-review context characters |
| `ReviewSuggestionBatchMaxItems` | 8 | Maximum review cards in one suggestion request |
| `ReviewSuggestionBatchMaxCharacters` | 6000 | Maximum source characters per suggestion request/fragment |
| `ReviewMaxConcurrentRequests` | 3 | Concurrent suggestion requests; issue batches run sequentially |
| `ReviewOllamaOutputTokens` | 2048 | Requested local review output allowance, capped at half `OllamaMaxToken` |
| `OllamaTimeoutSeconds` | 600 | Regular local generate/stream request timeout |
| `LinterContextTokens` | 4096 | Silent Linter local context |
| `LinterTimeoutSeconds` | 300 | Silent Linter local request timeout |
| `AgentTaskTimeoutMinutes` | 30 | Agent task time allowance |

Local review uses `OllamaMaxToken` as `num_ctx` and sends a distinct `num_predict` output limit. Online review uses the configured `OnlineContextWindow` and `ChatGPTMaxToken`; its timeout remains `ProviderTimeoutSeconds`. Batches are reduced further to reserve room for the complete prompt, output and a safety margin. Suggestions also reserve output space for the revised source text. These are conservative character-based token estimates, not measurements from a model tokenizer. Increasing a setting does not expand a model's supported context or guarantee proofreading quality; reasoning tokens can also consume the model's output allowance.

## Current scope and verification

Both review flows cover the main body and tables, primary/first/even headers and footers in every section, footnotes/endnotes, linked text boxes, grouped shapes and drawing-canvas text. Shared headers and linked text-box stories are processed once. Addresses and offset changes are qualified by document area, so one story cannot accidentally edit another. Word fields, note references and structural markers are protected during precise writeback. Existing comments and revision history are retained, while image/chart text is outside Word text review. The Silent Linter remains a nearby-edit check. Macro entry points and folder-wide batch processing are intentionally excluded.

`tools/run-core-smoke-test.ps1` covers 12,000 original paragraph indexes, a selection over 100,000 characters, oversized single paragraphs, Unicode/CRLF boundaries, whitespace-preserving reassembly, prompt/output budgeting, and invalid/truncated model responses. It also checks encrypted session round trips, interruption and fragment recovery, retry coverage, offset-journal replay, archive separation, complete report export, and rejection of source/existing copy paths. These synthetic checks establish batching, persistence and failure handling; a 1,000-page document with a real Word host and local model still requires an end-to-end acceptance run.

On a machine with Word installed, run `powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File tools/run-review-word-smoke-test.ps1` after building Release. It uses a separate temporary Word instance to verify exact paragraph-start indexes past paragraph 80 and selection clipping inside a large paragraph, then closes only the test document and test instance.

The automated checks also cover encrypted story recovery, legacy V2.4 sessions, cross-story offset isolation, Agent workflow routing and field protection. Real Word/Ollama testing remains necessary to assess host behavior, performance and proofreading quality on very long documents.
