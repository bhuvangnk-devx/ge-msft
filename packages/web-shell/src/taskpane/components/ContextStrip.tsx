import { useState } from 'react';
import type { ContextChip } from '../../controller.js';

interface ContextStripProps {
  chips: ContextChip[];
  disabled: boolean;
  onToggle: (id: string, attached: boolean) => void;
  onReveal: (id: string) => void;
  onRefresh: () => void;
}

/**
 * Source controls as a dropdown: one summary line, closed by default, that opens to show the attached
 * chips (and, from "+ N available" / "Manage", every attachable source). The controller remains the
 * sole attachment authority.
 */
export function ContextStrip({
  chips,
  disabled,
  onToggle,
  onReveal,
  onRefresh,
}: ContextStripProps): JSX.Element {
  const [open, setOpen] = useState(false);
  const [expanded, setExpanded] = useState(false);
  const attached = chips.filter((chip) => chip.attached);
  const nearby = chips.filter((chip) => !chip.attached);
  const visible = expanded ? chips : attached.slice(0, 3);
  return (
    <section className="context-strip" aria-label="Active context">
      <div className="context-strip-heading">
        <button
          type="button"
          className="context-strip-toggle"
          aria-expanded={open}
          aria-controls="active-context-chips"
          onClick={() => {
            setOpen(!open);
            if (open) setExpanded(false);
          }}
        >
          Attached context ·{' '}
          <strong>
            {attached.length} {attached.length === 1 ? 'source' : 'sources'}
          </strong>
          <span className="context-strip-caret" aria-hidden="true">
            {open ? '▴' : '▾'}
          </span>
        </button>
        <button
          type="button"
          className="text-control"
          aria-expanded={expanded}
          aria-controls="active-context-chips"
          onClick={() => {
            setExpanded(!expanded);
            setOpen(true);
          }}
        >
          {expanded ? 'Done' : nearby.length ? `+ ${nearby.length} available` : 'Manage'}
        </button>
      </div>
      {chips
        .filter((chip) => chip.notice)
        .map((chip) => (
          <p key={`notice-${chip.id}`} className="context-notice" role="note">
            {chip.notice}
          </p>
        ))}
      {open && (
        <div id="active-context-chips" className="smart-chips">
          {visible.map((chip) => (
            <span className="smart-chip" data-attached={chip.attached} key={chip.id}>
              <button
                type="button"
                className="smart-chip-title"
                disabled={disabled || (chip.attached && !chip.revealable)}
                title={chip.preview ?? chip.title}
                aria-label={chip.attached ? `Locate ${chip.title}` : `Use ${chip.title}`}
                onClick={() => (chip.attached ? onReveal(chip.id) : onToggle(chip.id, true))}
              >
                <span className="smart-chip-symbol" aria-hidden="true">
                  {chip.attached ? '◈' : '+'}
                </span>
                <span>{chip.title}</span>
              </button>
              {chip.attached && (
                <button
                  type="button"
                  className="smart-chip-remove"
                  disabled={disabled}
                  aria-label={`Remove ${chip.title} from context`}
                  onClick={() => onToggle(chip.id, false)}
                >
                  ×
                </button>
              )}
            </span>
          ))}
          {!expanded && attached.length > 3 && (
            <button type="button" className="text-control" onClick={() => setExpanded(true)}>
              +{attached.length - 3} more
            </button>
          )}
          {expanded && (
            <button type="button" className="text-control" disabled={disabled} onClick={onRefresh}>
              Refresh sources
            </button>
          )}
        </div>
      )}
    </section>
  );
}
