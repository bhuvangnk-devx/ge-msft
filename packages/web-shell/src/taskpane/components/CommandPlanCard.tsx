import type { CommandScope } from '@ge/contracts';
import type { PendingCommandPlan } from '../../controller.js';
import { uiStrings, type UiLanguage, type UiStrings } from '../../ui-language.js';

export interface CommandPlanCardProps {
  /** The conversation's language, for the card's own text. */
  language?: UiLanguage;
  /** The planner's high-level plan awaiting confirm; absent → nothing rendered. */
  pending?: PendingCommandPlan;
  onConfirm: () => void;
  onCancel: () => void;
}

/**
 * Where the plan works: "In <ref>", "Using your selection", or nothing. The ref is model-written, so
 * it is set apart from the card's own words (see {@link LabelValue}).
 */
function scopeLabel(scope: CommandScope | undefined, t: UiStrings): JSX.Element | undefined {
  if (scope?.ref) return <LabelValue label={t.planIn} value={scope.ref} />;
  return scope?.kind === 'selection' ? <>{t.planSelection}</> : undefined;
}

/**
 * A card label with a model-written value. The value sits in its own element with a distinct look,
 * so planner text ("everything (read-only)") cannot pass for the card's own wording.
 */
function LabelValue({ label, value }: { label: string; value: string }): JSX.Element {
  return (
    <>
      {label} <span className="model-value">{value}</span>
    </>
  );
}

/**
 * The planner-confirm card (EXPERIENCE.md §F). For a complex free-text actuating request, the planner
 * proposes a high-level {@link PendingCommandPlan} — the verb, the scope, the ordered steps, and any
 * exclusions/grounding — which the user confirms (or cancels) BEFORE the executor runs. This is the
 * legible "see the plan before it acts" stage; on confirm the executor then stages its own
 * effect-level gate (`PlanApprovalCard`). Nothing actuates from this card directly.
 */
export function CommandPlanCard({
  pending,
  onConfirm,
  onCancel,
  language = 'en',
}: CommandPlanCardProps): JSX.Element | null {
  if (!pending) return null;
  const { plan } = pending;
  const t = uiStrings(language);
  const where = scopeLabel(plan.scope, t);
  // Named sources only: bare keywords ("this", "unit") are planner vocabulary, not words for users.
  const sources = plan.ground.flatMap((g) => (g.ref ? [g.ref] : []));
  return (
    <section className="command-plan" role="group" aria-label="Proposed plan">
      <div className="command-plan-head">
        <span className="command-plan-title">{t.planTitle}</span>
        {where && <span className="command-plan-scope">{where}</span>}
      </div>
      <ol className="command-plan-steps">
        {plan.steps.map((step, i) => (
          <li key={i}>{step}</li>
        ))}
      </ol>
      {plan.excludes.length > 0 && (
        <ul className="command-plan-excludes" aria-label="Exclusions">
          {plan.excludes.map((ex, i) => (
            <li key={i}>
              <LabelValue label={t.planKeep} value={ex} />
            </li>
          ))}
        </ul>
      )}
      {sources.length > 0 && (
        <div className="command-plan-ground">
          <LabelValue label={t.planUsing} value={sources.join(', ')} />
        </div>
      )}
      <div className="command-plan-actions">
        <button
          type="button"
          className="command-plan-confirm"
          data-testid="command-plan-confirm"
          onClick={onConfirm}
        >
          {t.planGo}
        </button>
        <button
          type="button"
          className="command-plan-cancel"
          data-testid="command-plan-cancel"
          onClick={onCancel}
        >
          {t.cancel}
        </button>
      </div>
    </section>
  );
}
