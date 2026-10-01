import { useId, useState } from 'react';

/** A Vertex AI model id as Gemini Enterprise takes it in `generationSpec.modelId`. */
const MODEL_ID = /^[a-z0-9][a-z0-9._-]{1,63}$/i;

export interface ModelSettingProps {
  /** The model in use now (undefined = the engine's default model). */
  current: string | undefined;
  /** The deployment's default, restored by "Use default". */
  defaultModelId: string | undefined;
  disabled?: boolean;
  /** Apply a model for later turns; undefined = back to the deployment default. */
  onChange: (modelId: string | undefined) => void;
}

/**
 * Settings field for the Gemini model each turn asks for. The engine checks the id: an unknown one
 * fails the next turn with "model id: … is invalid", so the field only checks its shape.
 */
export function ModelSetting({
  current,
  defaultModelId,
  disabled,
  onChange,
}: ModelSettingProps): JSX.Element {
  const inputId = useId();
  const [draft, setDraft] = useState(current ?? '');
  const [status, setStatus] = useState('');
  const value = draft.trim();
  const valid = MODEL_ID.test(value);

  const apply = (): void => {
    if (!valid) return;
    onChange(value);
    setStatus(`Using ${value} from the next message.`);
  };
  const reset = (): void => {
    onChange(undefined);
    setDraft(defaultModelId ?? '');
    setStatus(
      defaultModelId ? `Back to the default, ${defaultModelId}.` : 'Back to the engine default.',
    );
  };

  return (
    <section className="model-setting" aria-label="Model">
      <label htmlFor={inputId} className="eyebrow">
        Model
      </label>
      <div className="model-setting-row">
        <input
          id={inputId}
          type="text"
          spellCheck={false}
          autoComplete="off"
          value={draft}
          placeholder={defaultModelId ?? 'engine default'}
          disabled={disabled}
          aria-invalid={value !== '' && !valid}
          onChange={(e) => {
            setDraft(e.target.value);
            setStatus('');
          }}
          onKeyDown={(e) => {
            if (e.key === 'Enter') apply();
          }}
        />
        <button type="button" disabled={disabled || !valid || value === current} onClick={apply}>
          Use
        </button>
        <button type="button" disabled={disabled || current === defaultModelId} onClick={reset}>
          Use default
        </button>
      </div>
      <p className="muted small">
        {status ||
          (value !== '' && !valid
            ? 'A model id looks like gemini-3.8-flash.'
            : `Gemini model for each reply. Default: ${defaultModelId ?? 'engine default'}.`)}
      </p>
    </section>
  );
}
