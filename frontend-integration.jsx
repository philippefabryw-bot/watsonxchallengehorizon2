// ============================================================================
// FRONTEND CHANGES for App.jsx
// ----------------------------------------------------------------------------
// Replace the browser-side watsonx call with a call to your own /api backend.
// This removes: the WX_ENDPOINT constant, the iamToken state + token input UI,
// and the entire SSE-parsing block. No token in the browser, no CORS.
// ============================================================================


// 1. DELETE these — no longer needed:
//    - const WX_ENDPOINT = '...';
//    - const [iamToken, setIamToken] = useState('');
//    - the <input type="password" ... llm-token-input /> field in the JSX


// 2. REPLACE fetchLlmWorkbackPlan with this NON-STREAMING version
//    (simplest — matches /api/workback):

const fetchLlmWorkbackPlan = useCallback(async () => {
  if (!selectedUseCase) return;

  if (abortRef.current) abortRef.current.abort();
  const controller = new AbortController();
  abortRef.current = controller;

  setLlmPlan('');
  setLlmError(null);
  setLlmLoading(true);

  const usecaseVar = `${selectedUseCase.product} — ${selectedUseCase.title}`;

  try {
    const response = await fetch('/api/workback', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ usecase: usecaseVar }),
      signal: controller.signal,
    });

    if (!response.ok) {
      const text = await response.text().catch(() => response.statusText);
      throw new Error(`API error ${response.status}: ${text}`);
    }

    const data = await response.json();
    // watsonx generation (non-stream) shape: { results: [{ generated_text }] }
    const text =
      data?.results?.[0]?.generated_text ??
      data?.generated_text ??
      JSON.stringify(data, null, 2);
    setLlmPlan(text);
  } catch (err) {
    if (err.name !== 'AbortError') setLlmError(err.message);
  } finally {
    setLlmLoading(false);
  }
}, [selectedUseCase]);


// ============================================================================
// OPTIONAL: streaming version (token-by-token UX) — calls /api/workback/stream
// Use this INSTEAD of the above if you want the typewriter effect back.
// ============================================================================

const fetchLlmWorkbackPlanStreaming = useCallback(async () => {
  if (!selectedUseCase) return;

  if (abortRef.current) abortRef.current.abort();
  const controller = new AbortController();
  abortRef.current = controller;

  setLlmPlan('');
  setLlmError(null);
  setLlmLoading(true);

  const usecaseVar = `${selectedUseCase.product} — ${selectedUseCase.title}`;

  try {
    const response = await fetch('/api/workback/stream', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ usecase: usecaseVar }),
      signal: controller.signal,
    });

    if (!response.ok) {
      const text = await response.text().catch(() => response.statusText);
      throw new Error(`API error ${response.status}: ${text}`);
    }

    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';

    while (true) {
      const { value, done } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split('\n');
      buffer = lines.pop();
      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed || trimmed === 'data: [DONE]') continue;
        const jsonStr = trimmed.startsWith('data:') ? trimmed.slice(5).trim() : trimmed;
        try {
          const chunk = JSON.parse(jsonStr);
          const token =
            chunk?.results?.[0]?.generated_text ??
            chunk?.generated_text ??
            '';
          if (token) setLlmPlan((prev) => prev + token);
        } catch {
          /* ignore keep-alives / non-JSON lines */
        }
      }
    }
  } catch (err) {
    if (err.name !== 'AbortError') setLlmError(err.message);
  } finally {
    setLlmLoading(false);
  }
}, [selectedUseCase]);
