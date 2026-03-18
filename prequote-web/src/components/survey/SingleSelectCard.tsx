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

  const handleClick = () => {
    setJustSelected(true);
    onSelect(option.option_code);
    // Reset the pop animation
    setTimeout(() => setJustSelected(false), 300);
  };

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
        {/* Thumbnail image */}
        {option.option_image_url && (
          <img
            src={option.option_image_url}
            alt=""
            className="h-14 w-14 shrink-0 rounded-xl object-cover"
            loading="lazy"
          />
        )}

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
