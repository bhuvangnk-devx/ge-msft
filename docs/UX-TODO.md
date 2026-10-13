# CN-GPT UX to-do

Agreed on 13 Oct 2026. Order is the build order.

## Now

1. **Result card after a change.** Say what changed in plain words ("Added totals in F2:F11,
   10 cells"), with **Show me** (jump to it in the document) and **Undo**.
2. **One approval for simple changes.** A clear, single change (one range, one format, one slide)
   goes straight to one approval card. The plan step stays for unclear or multi-step requests.
   Needs a security review before it ships.
3. **Plain language, in the user's language.**
   - Approval cards, steps and status describe the change ("Add 3 slides after slide 2"), not the
     command (`insert-slide`, `write-cells`, shape ids).
   - **Language follows the conversation.** If the user started in English, planning steps,
     approvals and messages are in English. If they started in Bahasa Indonesia, they are in Bahasa.
     There is no language setting.
4. **No raw errors in the chat.** A failed request shows a short message and a **Try again** button
   ("That didn't work. Try again." / "Maaf, belum berhasil. Coba lagi."). The error code and message
   go only to **Copy diagnostics** (already captured there, redacted).
5. **Resume after the pane closes or reloads mid-task.** On reopen: "Your last request (add 3 slides)
   was interrupted: 2 of 3 done. Continue?" Continue runs only what is left; nothing is redone.

## Next

6. **Auto-check uncertain writes.** When the host does not confirm a write (PowerPoint
   GeneralException), re-read the document and report "added" or "not added" instead of "outcome
   unknown".
7. **Silent retry** of one failed network call (dropped connection, expired token) before showing
   Try again.
8. **Check host support on open** and hide actions this Office version cannot do.
9. **Clear context chip:** "Using your selection: Sheet1!A1:E11" or "Nothing selected: using the whole
   sheet".

## Later (end of the list)

10. Helpful / not helpful on each answer.
11. Report a problem from the pane (question + diagnostics + optional note).
