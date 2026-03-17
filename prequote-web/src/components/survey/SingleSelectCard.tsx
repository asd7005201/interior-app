import { Check } from 'lucide-react';
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
  return (
    <button
      type="button"
      onClick={() => onSelect(option.option_code)}
      className={`
        relative w-full text-left rounded-lg border px-4 py-3.5
        transition-all duration-200 ease-out
        hover:shadow-md hover:-translate-y-0.5
        focus:outline-none focus:ring-2 focus:ring-[#8b6d4b]/30
        ${
          selected
            ? 'border-[#8b6d4b] bg-[#8b6d4b]/[0.04] shadow-sm'
            : 'border-[#e8e4e0] bg-white hover:border-[#b9a79e]'
        }
      `}
    >
      <div className="flex items-start gap-3">
        {/* Radio indicator */}
        <div
          className={`
            mt-0.5 flex h-5 w-5 shrink-0 items-center justify-center rounded-full border-2
            transition-colors duration-200
            ${
              selected
                ? 'border-[#8b6d4b] bg-[#8b6d4b]'
                : 'border-[#d0cbc6] bg-white'
            }
          `}
        >
          {selected && <Check className="h-3 w-3 text-white" strokeWidth={3} />}
        </div>

        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-2">
            <span
              className={`
                text-sm font-sans font-medium leading-snug
                ${selected ? 'text-[#333]' : 'text-[#333]'}
              `}
            >
              {option.option_label}
            </span>
            {option.badge_text && (
              <span className="inline-flex items-center rounded-full bg-[#8b6d4b]/10 px-2 py-0.5 text-[10px] font-semibold text-[#8b6d4b] uppercase tracking-wide">
                {option.badge_text}
              </span>
            )}
          </div>
          {option.option_description && (
            <p className="mt-1 text-xs font-sans font-light leading-relaxed text-[#696969]">
              {option.option_description}
            </p>
          )}
        </div>
      </div>
    </button>
  );
}
