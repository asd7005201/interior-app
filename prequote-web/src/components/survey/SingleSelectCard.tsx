import { useState } from 'react';
import type { SurveyOption } from '@/survey/types';

interface SingleSelectCardProps {
  option: SurveyOption;
  selected: boolean;
  onSelect: (optionCode: string) => void;
}

export default function SingleSelectCard({
  option,
  selected,
  onSelect,
}: SingleSelectCardProps) {
  const [justSelected, setJustSelected] = useState(false);
  const [imgError, setImgError] = useState(false);
  const hasImage = !!option.option_image_url && !imgError;

  const handleClick = () => {
    setJustSelected(true);
    onSelect(option.option_code);
    setTimeout(() => setJustSelected(false), 300);
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
          ${justSelected ? 'scale-[1.02]' : 'scale-100'}
        `}
        style={{
          transition: justSelected
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
          {/* Dark overlay gradient for readability */}
          <div className="absolute inset-0 bg-gradient-to-t from-black/40 via-transparent to-transparent" />

          {/* Badge (over image) */}
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

          {/* Selected check circle */}
          {selected && (
            <div className="absolute top-2.5 left-2.5 flex h-6 w-6 items-center justify-center rounded-full bg-[#8b6d4b] shadow-md">
              <svg className="h-3.5 w-3.5 text-white" fill="none" stroke="currentColor" strokeWidth={2.5} viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7" />
              </svg>
            </div>
          )}

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
        ${justSelected ? 'scale-[1.02]' : 'scale-100'}
      `}
      style={{
        transition: justSelected
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
