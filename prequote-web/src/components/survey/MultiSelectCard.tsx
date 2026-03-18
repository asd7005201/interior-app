import { useState } from 'react';
import { Check } from 'lucide-react';
import type { SurveyOption } from '@/survey/types';

interface MultiSelectCardProps {
  option: SurveyOption;
  selected: boolean;
  onToggle: (optionCode: string) => void;
}

export default function MultiSelectCard({
  option,
  selected,
  onToggle,
}: MultiSelectCardProps) {
  const [justToggled, setJustToggled] = useState(false);
  const [imgError, setImgError] = useState(false);
  const hasImage = !!option.option_image_url && !imgError;

  const handleClick = () => {
    setJustToggled(true);
    onToggle(option.option_code);
    setTimeout(() => setJustToggled(false), 300);
  };

  /* ── Visual card layout (with image) ─────────────────────────── */
  if (hasImage) {
    return (
      <button
        type="button"
        onClick={handleClick}
        className={`
          group relative w-full text-left rounded-2xl overflow-hidden
          transition-all duration-200 ease-out
          focus:outline-none focus:ring-2 focus:ring-[#8b6d4b]/20
          ${
            selected
              ? 'ring-2 ring-[#8b6d4b] shadow-lg'
              : 'border border-[var(--survey-card-border)] hover:shadow-lg hover:-translate-y-0.5'
          }
          ${justToggled ? 'scale-[1.02]' : 'scale-100'}
        `}
        style={{
          transition: justToggled
            ? 'transform 0.2s cubic-bezier(0.34, 1.56, 0.64, 1), box-shadow 0.2s ease, border-color 0.2s ease'
            : 'transform 0.2s ease, box-shadow 0.2s ease, border-color 0.2s ease',
        }}
      >
        {/* Image area */}
        <div className="relative h-36 w-full overflow-hidden bg-[#f5f3f0]">
          <img
            src={option.option_image_url}
            alt={option.option_label}
            onError={() => setImgError(true)}
            className="h-full w-full object-cover transition-transform duration-500 group-hover:scale-105"
            loading="lazy"
          />
          <div className="absolute inset-0 bg-gradient-to-t from-black/40 via-transparent to-transparent" />

          {/* Badge */}
          {option.badge_text && (
            <span
              className="absolute top-2.5 right-2.5 inline-flex items-center rounded-full px-2.5 py-1 text-[10px] font-semibold text-white tracking-wide backdrop-blur-sm"
              style={{
                background: 'linear-gradient(135deg, rgba(139,109,75,0.85), rgba(185,167,158,0.85))',
              }}
            >
              {option.badge_text}
            </span>
          )}

          {/* Checkbox circle (top-left) */}
          <div
            className={`
              absolute top-2.5 left-2.5 flex h-6 w-6 items-center justify-center rounded-md border-[1.5px] transition-all duration-200
              ${selected ? 'border-[#8b6d4b] bg-[#8b6d4b] shadow-md' : 'border-white/70 bg-white/30 backdrop-blur-sm'}
            `}
          >
            {selected && <Check className="h-3.5 w-3.5 text-white" strokeWidth={2.5} />}
          </div>

          {/* Label over image bottom */}
          <div className="absolute bottom-0 left-0 right-0 px-4 pb-3 pt-6">
            <span className="text-sm font-sans font-semibold text-white drop-shadow-md leading-snug">
              {option.option_label}
            </span>
          </div>
        </div>

        {/* Description below image */}
        {option.option_description && (
          <div className="px-4 py-2.5 bg-white">
            <p className="text-[11px] font-sans font-light leading-relaxed text-[#696969]">
              {option.option_description}
            </p>
          </div>
        )}
      </button>
    );
  }

  /* ── Text-only card layout (no image) ────────────────────────── */
  return (
    <button
      type="button"
      onClick={handleClick}
      className={`
        relative w-full text-left rounded-xl overflow-hidden
        transition-all duration-200 ease-out
        focus:outline-none focus:ring-2 focus:ring-[#8b6d4b]/20
        ${
          selected
            ? 'border-l-[3px] border-l-[#8b6d4b] border-y border-r border-y-[var(--survey-card-border)] border-r-[var(--survey-card-border)] bg-white/80 backdrop-blur-sm shadow-sm'
            : 'border border-[var(--survey-card-border)] bg-white hover:border-[#b9a79e] hover:-translate-y-px hover:shadow-md'
        }
        ${justToggled ? 'scale-[1.02]' : 'scale-100'}
      `}
      style={{
        transition: justToggled
          ? 'transform 0.2s cubic-bezier(0.34, 1.56, 0.64, 1), box-shadow 0.2s ease, border-color 0.2s ease, background-color 0.2s ease'
          : 'transform 0.2s ease, box-shadow 0.2s ease, border-color 0.2s ease, background-color 0.2s ease',
      }}
    >
      {/* Badge in top-right */}
      {option.badge_text && (
        <span
          className="absolute top-2 right-2 inline-flex items-center rounded-full px-2 py-0.5 text-[10px] font-semibold text-white tracking-wide"
          style={{
            background: 'linear-gradient(135deg, #8b6d4b, #b9a79e)',
          }}
        >
          {option.badge_text}
        </span>
      )}

      <div className="flex items-center gap-3.5 px-4 py-3.5">
        {/* Checkbox indicator */}
        <div
          className={`
            flex h-4 w-4 shrink-0 items-center justify-center rounded-md border-[1.5px]
            transition-all duration-200
            ${selected ? 'border-[#8b6d4b] bg-[#8b6d4b]' : 'border-[#d0cbc6] bg-white'}
          `}
        >
          {selected && <Check className="h-2.5 w-2.5 text-white" strokeWidth={3} />}
        </div>

        <div className="flex-1 min-w-0">
          <span className="block text-sm font-sans font-medium leading-snug text-[#333]">
            {option.option_label}
          </span>
          {option.option_description && (
            <p className="mt-0.5 text-xs font-sans font-light leading-relaxed text-[#696969]">
              {option.option_description}
            </p>
          )}
        </div>
      </div>
    </button>
  );
}
