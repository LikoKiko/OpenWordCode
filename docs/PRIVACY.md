# Privacy

Before sending an agent request, the task pane shows the selected provider's privacy note. The Core sends the provider only the instruction plus bounded context needed for that request. Cloud providers receive document context according to their service policies; local providers are marked local but depend on the configured endpoint.

The Core does not persist conversation history, document text, or uploaded attachment contents in 0.1.0. Attachments are held in memory for the active request and sent to the selected provider when the provider supports multimodal input. The in-memory change store is capped at 100 entries and is cleared when the Core restarts. Diagnostics contain provider ids, model ids, and connection state, not document contents, attachment contents, or secret values.
